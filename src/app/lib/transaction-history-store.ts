import type { SupabaseClient } from '@supabase/supabase-js';
import { normalizePeriod, type TransactionRow } from './utils/transaction-history';

export interface TransactionVersion {
  id: string;
  period: string;
  created_at: string;
  rows: TransactionRow[];
}

export async function requireHistoryMember(client: SupabaseClient): Promise<string> {
  const { data: { user }, error } = await client.auth.getUser();
  if (error || !user) throw new Error('Vui lòng đăng nhập kho Transaction.');
  const membership = await client.from('transaction_history_members').select('user_id').eq('user_id', user.id).maybeSingle();
  if (membership.error) throw new Error(membership.error.message);
  if (!membership.data) throw new Error('Tài khoản chưa được cấp quyền truy cập kho payroll.');
  return user.id;
}

export async function loadLatestVersion(client: SupabaseClient, period: string): Promise<TransactionVersion | null> {
  await requireHistoryMember(client);
  const { data, error } = await client.from('transaction_monthly_versions')
    .select('id, period, created_at, rows').eq('period', `${normalizePeriod(period)}-01`)
    .order('id', { ascending: false }).limit(1).maybeSingle();
  if (error) throw new Error(error.message);
  return data ? {...data, id: String(data.id)} as TransactionVersion : null;
}

export class HistorySaveConflictError extends Error {}

export async function replaceVersionIfCurrent(
  client: SupabaseClient,
  expected: TransactionVersion,
  rows: TransactionRow[],
  requestId: string,
): Promise<string> {
  const period = expected.period.slice(0, 7);
  const latest = await loadLatestVersion(client, period);
  if (!latest || latest.id !== expected.id) {
    throw new HistorySaveConflictError(
      `Dữ liệu tháng ${period} đã thay đổi. Hãy Check STK & ID lại trước khi đồng bộ.`,
    );
  }
  return saveVersion(client, period, rows, requestId);
}

async function latestVersionIds(client: SupabaseClient, before?: string): Promise<Map<string, string>> {
  const latest = new Map<string, string>();
  let cursor: string | undefined;
  while (true) {
    let query = client.from('transaction_monthly_versions').select('id, period')
      .order('id', { ascending: false }).limit(200);
    if (before) query = query.lt('period', before);
    if (cursor) query = query.lt('id', cursor);
    const { data, error } = await query;
    if (error) throw new Error(error.message);
    if (!data?.length) break;
    for (const item of data) if (!latest.has(item.period)) latest.set(item.period, String(item.id));
    const nextCursor = String(data[data.length - 1].id);
    if (nextCursor === cursor) throw new Error('Không thể tải đầy đủ lịch sử Transaction. Hãy thử lại.');
    cursor = nextCursor;
  }
  return latest;
}

async function loadPinnedVersions(client: SupabaseClient, latest: Map<string, string>): Promise<TransactionVersion[]> {
  const versions: TransactionVersion[] = [];
  const ids = [...latest.values()];
  for (let offset = 0; offset < ids.length; offset += 4) {
    const batch = await Promise.all(ids.slice(offset, offset + 4).map(async id => {
      const { data, error } = await client.from('transaction_monthly_versions')
        .select('id, period, created_at, rows').eq('id', id).single();
      if (error) throw new Error(error.message);
      if (!data) throw new Error(`Không đọc được phiên bản nguồn #${id}.`);
      return { ...data, id: String(data.id) } as TransactionVersion;
    }));
    versions.push(...batch);
  }
  return versions.sort((a, b) => a.period.localeCompare(b.period));
}

/** Pin the latest saved snapshot for every earlier saved month. */
export async function loadAllPriorVersions(client: SupabaseClient, period: string): Promise<TransactionVersion[]> {
  await requireHistoryMember(client);
  return loadPinnedVersions(client, await latestVersionIds(client, `${normalizePeriod(period)}-01`));
}

/** Restore all saved months, reusing only pinned versions with an unchanged ID. */
export async function loadLatestTransactionMonths(
  client: SupabaseClient,
  known: Record<string, TransactionVersion> = {},
): Promise<TransactionVersion[]> {
  await requireHistoryMember(client);
  const latest = await latestVersionIds(client);
  const missing = new Map([...latest].filter(([period, id]) => known[period.slice(0, 7)]?.id !== id));
  const fetched = await loadPinnedVersions(client, missing);
  const byPeriod = new Map(fetched.map(version => [version.period, version]));
  const versions = [...latest].map(([period, id]) => {
    const version = byPeriod.get(period) || known[period.slice(0, 7)];
    if (!version || version.id !== id || version.period !== period) throw new Error('Không đọc được đầy đủ Batch Payment theo tháng.');
    return version;
  });
  if (!sameVersionIds(latest, await latestVersionIds(client))) {
    throw new HistorySaveConflictError('Batch Payment vừa được cập nhật. Bấm Đồng bộ ngay để tải lại đầy đủ các tháng.');
  }
  return versions.sort((a, b) => a.period.localeCompare(b.period));
}

function afterPeriod(period: string): string {
  const [year, month] = normalizePeriod(period).split('-').map(Number);
  return new Date(Date.UTC(year, month, 1)).toISOString().slice(0, 10);
}

function sameVersionIds(left: Map<string, string>, right: Map<string, string>): boolean {
  return left.size === right.size && [...left].every(([period, id]) => right.get(period) === id);
}

export interface TransactionCheckSource {
  currentVersion: TransactionVersion;
  versions: TransactionVersion[];
}

/** Lock/recheck donors and targets, then commit all selected months together. */
export async function replaceTransactionVersionsAtomically(
  client: SupabaseClient,
  source: TransactionCheckSource,
  changes: {version: TransactionVersion; rows: TransactionRow[]}[],
): Promise<{period: string; id: string}[]> {
  await requireHistoryMember(client);
  if (!changes.length) throw new Error('Chưa chọn tháng cập nhật.');
  const {data, error} = await client.rpc('replace_transaction_versions', {
    p_expected: [source.currentVersion, ...source.versions].map(version => ({period: version.period, id: version.id})),
    p_changes: changes.map(change => ({
      period: change.version.period, rows: change.rows, request_id: crypto.randomUUID(),
    })),
  });
  if (error?.code === 'P0002') throw new HistorySaveConflictError(error.message);
  if (error) throw new Error(error.message);
  return data as {period: string; id: string}[];
}

/** Every check reads current and prior months from Supabase afresh, including a consistency check. */
export async function loadTransactionCheckSource(client: SupabaseClient, period: string): Promise<TransactionCheckSource> {
  await requireHistoryMember(client);
  const currentPeriod = `${normalizePeriod(period)}-01`;
  const latest = await latestVersionIds(client, afterPeriod(period));
  if (!latest.has(currentPeriod)) throw new Error(`Tháng ${normalizePeriod(period)} chưa được lưu trên Supabase. Lưu sửa chỉ lưu dữ liệu trên máy; bấm Lưu tháng để cập nhật Supabase trước khi Check STK & ID.`);
  const snapshots = await loadPinnedVersions(client, latest);
  if (!sameVersionIds(latest, await latestVersionIds(client, afterPeriod(period)))) {
    throw new HistorySaveConflictError('Dữ liệu Supabase vừa được cập nhật trong lúc kiểm tra. Bấm Check STK & ID để lấy phiên bản mới nhất.');
  }
  const currentVersion = snapshots.find(snapshot => snapshot.period === currentPeriod);
  if (!currentVersion) throw new Error('Không đọc được dữ liệu tháng đang chọn từ Supabase.');
  return {currentVersion, versions: snapshots.filter(snapshot => snapshot.period < currentPeriod)};
}

/** Recheck donors as well as targets before copying any ID/account from a displayed report. */
export async function assertTransactionCheckCurrent(client: SupabaseClient, source: TransactionCheckSource): Promise<void> {
  await requireHistoryMember(client);
  const expected = new Map([source.currentVersion, ...source.versions].map(version => [version.period, version.id]));
  if (!sameVersionIds(expected, await latestVersionIds(client, afterPeriod(source.currentVersion.period.slice(0, 7))))) {
    throw new HistorySaveConflictError('Dữ liệu Supabase đã thay đổi sau lần kiểm tra. Bấm Check STK & ID lại trước khi đồng bộ.');
  }
}

export async function saveVersion(client: SupabaseClient, period: string, rows: TransactionRow[], requestId: string): Promise<string> {
  await requireHistoryMember(client);
  const { data, error } = await client.rpc('append_transaction_version', {
    p_period: `${normalizePeriod(period)}-01`, p_rows: rows, p_request_id: requestId,
  });
  if (error?.code === 'P0002') throw new HistorySaveConflictError(error.message);
  if (error) throw new Error(error.message);
  return String(data);
}
