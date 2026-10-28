import type { AppData } from '../../types';
import type { TransactionCheckSource, TransactionVersion } from '../transaction-history-store';
import { markTransactionSaved } from './transaction-activity';
import { normalizePeriod, selectPeriodRows, type TransactionRow } from './transaction-history';
import { replaceTransactionPeriod, sameTransactionSnapshot } from './transaction-snapshot';
import { withTransactionMonthOriginal } from './transaction-originals';

export type LocalTransactionMonths = NonNullable<AppData['TransactionMonthCache']>['months'];

/** A cloud-only check must not silently contradict a retained, saved Batch Payment. */
export function conflictingLocalTransactionPeriods(
  source: TransactionCheckSource,
  rows: TransactionRow[],
  month: string,
  localMonths: LocalTransactionMonths = {},
): string[] {
  const period = normalizePeriod(month);
  const conflicts = new Set<string>();
  if (!sameTransactionSnapshot(rows, source.currentVersion.rows, period)) conflicts.add(period);
  const cloudMonths = new Map(source.versions.map(version => [version.period.slice(0, 7), version]));
  for (const [localPeriod, snapshot] of Object.entries(localMonths)) {
    // Generated previews are not confirmed history. The active rows above take
    // precedence over their cache entry, which can lag until month navigation.
    if (localPeriod >= period || snapshot.activity?.lastAction === 'generated') continue;
    const cloud = cloudMonths.get(localPeriod);
    if (!cloud || !sameTransactionSnapshot(snapshot.table.data, cloud.rows, localPeriod)) {
      conflicts.add(localPeriod);
    }
  }
  return [...conflicts].sort();
}

/** Apply only explicitly loaded/resolved cloud versions, including historical months. */
export function applyLocalTransactionVersions(
  appData: AppData,
  versions: TransactionVersion[],
  savedAt = new Date().toISOString(),
): AppData {
  if (!versions.length) return appData;
  const activePeriod = normalizePeriod(appData.globalMonth);
  const months = {...appData.TransactionMonthCache?.months};
  let table = appData.BankExport;
  let activity = appData.TransactionActivity;
  let originals = appData;
  for (const version of versions) {
    const period = normalizePeriod(version.period.slice(0, 7));
    const selected = selectPeriodRows(version.rows, period);
    if (selected.length !== version.rows.length) throw new Error('Dữ liệu thay thế phải thuộc đúng tháng đã lưu.');
    const existing = months[period];
    const nextActivity = markTransactionSaved({
      ...appData,
      TransactionActivity: period === activePeriod ? activity : existing?.activity,
    }, savedAt);
    months[period] = {
      table: {...(existing?.table || appData.BankExport), data: selected},
      activity: nextActivity,
    };
    originals = withTransactionMonthOriginal(originals, period, months[period].table);
    if (period === activePeriod) {
      table = {...table, data: replaceTransactionPeriod(table.data, period, selected)};
      activity = nextActivity;
    }
  }
  return {
    ...appData,
    TableOriginals: originals.TableOriginals,
    BankExport: table,
    TransactionActivity: activity,
    TransactionMonthCache: {activePeriod, months},
  };
}
