import type { AppData } from '../../types';
import { normalizePeriod, selectPeriodRows, type TransactionRow } from './transaction-history';
import { withTransactionMonthOriginal } from './transaction-originals';

export interface TransactionSnapshotLike {
  period: string;
  rows: TransactionRow[];
}

function normalizeSnapshotPeriod(value: string): string {
  const raw = String(value || '').trim();
  const dated = raw.match(/^(\d{4}-\d{2})-\d{2}$/);
  return normalizePeriod(dated?.[1] ?? raw);
}

function snapshotPeriodRows(rows: TransactionRow[], month: string): TransactionRow[] {
  if (!rows.length) return [];
  return selectPeriodRows(rows, month);
}

function fingerprint(rows: TransactionRow[]): string {
  // JSONB does not retain object key order. Array/row order and every stored value matter.
  return JSON.stringify(rows, (_key, value) => value && typeof value === 'object' && !Array.isArray(value)
    ? Object.fromEntries(Object.entries(value).sort(([a], [b]) => a.localeCompare(b)))
    : value);
}

export function sameTransactionSnapshot(localRows: TransactionRow[], savedRows: TransactionRow[], month: string): boolean {
  try {
    return fingerprint(snapshotPeriodRows(localRows, month)) === fingerprint(snapshotPeriodRows(savedRows, month));
  } catch { return false; }
}

export function getSavedLocalTransactionSnapshots(
  appData: Pick<AppData, 'TransactionMonthCache'>,
  beforeMonth: string,
): TransactionSnapshotLike[] {
  let before: string;
  try { before = normalizePeriod(beforeMonth); } catch { return []; }

  return Object.entries(appData.TransactionMonthCache?.months || {})
    .flatMap(([period, snapshot]) => {
      try {
        const normalized = normalizeSnapshotPeriod(period);
        return normalized < before && snapshot.activity?.lastAction === 'saved'
          ? [{period: normalized, rows: snapshot.table.data as TransactionRow[]}]
          : [];
      } catch { return []; }
    })
    .sort((left, right) => left.period.localeCompare(right.period));
}

export function findMismatchedTransactionHistoryPeriods(
  localSnapshots: TransactionSnapshotLike[],
  cloudSnapshots: TransactionSnapshotLike[],
  beforeMonth: string,
): string[] {
  let before: string;
  try { before = normalizePeriod(beforeMonth); } catch { return []; }

  const cloudByPeriod = new Map<string, TransactionRow[]>();
  for (const snapshot of cloudSnapshots) {
    try { cloudByPeriod.set(normalizeSnapshotPeriod(snapshot.period), snapshot.rows); } catch { /* ignore invalid remote period */ }
  }

  const mismatched = new Set<string>();
  for (const snapshot of localSnapshots) {
    try {
      const period = normalizeSnapshotPeriod(snapshot.period);
      if (period >= before) continue;
      const cloudRows = cloudByPeriod.get(period);
      if (!cloudRows || !sameTransactionSnapshot(snapshot.rows, cloudRows, period)) mismatched.add(period);
    } catch { /* ignore invalid local cache keys */ }
  }
  return [...mismatched].sort();
}

/** Map snapshot-relative row indexes back to the complete local table without touching other months. */
export function replaceTransactionPeriod(rows: TransactionRow[], month: string, replacement: TransactionRow[]): TransactionRow[] {
  const period = normalizePeriod(month);
  const selected = snapshotPeriodRows(replacement, period);
  if (selected.length !== replacement.length) throw new Error('Dữ liệu thay thế phải thuộc đúng tháng đang chọn.');
  const output: TransactionRow[] = [];
  let inserted = false;
  for (const row of rows) {
    const raw = row['Tháng báo cáo'] || row._fileMonth;
    if (raw && normalizePeriod(raw) === period) {
      if (!inserted) output.push(...selected);
      inserted = true;
    } else output.push(row);
  }
  if (!inserted) output.push(...selected);
  return output;
}

function syncedTransactionActivity(
  activity: AppData['TransactionActivity'],
  savedAt: string,
): NonNullable<AppData['TransactionActivity']> {
  return {
    generatedAt: activity?.generatedAt || savedAt,
    lastSavedAt: savedAt,
    editCount: activity?.editCount || 0,
    saveVersion: (activity?.saveVersion || 0) + 1,
    lastAction: 'saved',
  };
}

/**
 * Apply verified Supabase snapshots to the matching local month cache.
 * Only the active reporting month is also written to BankExport.
 */
export function replaceTransactionSnapshotsInAppData(
  appData: AppData,
  snapshots: TransactionSnapshotLike[],
  savedAt = new Date().toISOString(),
): AppData {
  if (!snapshots.length) return appData;

  let currentPeriod = '';
  try { currentPeriod = normalizePeriod(appData.globalMonth || ''); }
  catch { currentPeriod = appData.TransactionMonthCache?.activePeriod || ''; }

  const normalized = new Map<string, TransactionRow[]>();
  for (const snapshot of snapshots) {
    const period = normalizeSnapshotPeriod(snapshot.period);
    const rows = snapshotPeriodRows(snapshot.rows, period);
    if (rows.length !== snapshot.rows.length) throw new Error('Dữ liệu đồng bộ phải thuộc đúng tháng nguồn.');
    normalized.set(period, rows);
  }

  const months = {...(appData.TransactionMonthCache?.months || {})};
  let bankExport = appData.BankExport;
  let transactionActivity = appData.TransactionActivity;
  let originals = appData;

  for (const [period, rows] of normalized) {
    const existing = months[period];
    const activity = syncedTransactionActivity(
      existing?.activity ?? (period === currentPeriod ? appData.TransactionActivity : undefined),
      savedAt,
    );
    const table = {...(existing?.table ?? appData.BankExport), data: rows};
    months[period] = {table, activity};
    originals = withTransactionMonthOriginal(originals, period, table);
    if (period === currentPeriod) {
      bankExport = table;
      transactionActivity = activity;
    }
  }

  return {
    ...appData,
    TableOriginals: originals.TableOriginals,
    BankExport: bankExport,
    TransactionActivity: transactionActivity,
    TransactionMonthCache: {
      activePeriod: currentPeriod || appData.TransactionMonthCache?.activePeriod || normalized.keys().next().value || '',
      months,
    },
  };
}
