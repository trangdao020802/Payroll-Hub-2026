import type { AppData } from '../../types';
import type { TransactionVersion } from '../transaction-history-store';
import { normalizePeriod, selectPeriodRows } from './transaction-history';
import { sameTransactionSnapshot } from './transaction-snapshot';
import { withTransactionMonthOriginal } from './transaction-originals';

/** Fill missing working months from verified history, preserving ongoing working edits. */
export function restoreSavedTransactionMonths(data: AppData, versions: TransactionVersion[]): AppData {
  if (data.TransactionHistoryRestoreDisabled) return data;
  const period = normalizePeriod(data.globalMonth || '03.2026');
  const previous = data.TransactionMonthCache;
  const months = {...previous?.months};
  const cloudMonths = {...previous?.cloudMonths};
  let next = data;
  let changed = false;
  for (const version of versions) {
    const key = normalizePeriod(version.period.slice(0, 7));
    if (cloudMonths[key]?.id !== version.id) {
      const rows = version.rows.length ? selectPeriodRows(version.rows, key) : [];
      if (rows.length !== version.rows.length) throw new Error('Batch Payment đã lưu có dòng không đúng tháng nguồn.');
      cloudMonths[key] = version;
      next = withTransactionMonthOriginal(next, key, {...data.BankExport, data: version.rows});
      changed = true;
    }
    const working = months[key];
    const oldCloud = previous?.cloudMonths?.[key];
    const clean = working && oldCloud && working.cloudVersionId === oldCloud.id && sameTransactionSnapshot(working.table.data, oldCloud.rows, key);
    const missing = !working || (!working.table.data.length && working.activity?.lastAction !== 'saved' && working.activity?.lastAction !== 'edited');
    if (missing || (clean && working.cloudVersionId !== version.id)) {
      const table = {...(working?.table || data.BankExport), data: version.rows};
      const activity: AppData['TransactionActivity'] = {
        generatedAt: version.created_at, lastSavedAt: version.created_at,
        editCount: 0, saveVersion: working?.activity?.saveVersion || 0, lastAction: 'saved',
      };
      months[key] = {table, activity, cloudVersionId: version.id};
      next = withTransactionMonthOriginal(next, key, table);
      changed = true;
    } else if (!working.cloudVersionId && sameTransactionSnapshot(working.table.data, version.rows, key)) {
      months[key] = {...working, cloudVersionId: version.id};
      changed = true;
    }
  }
  const selected = months[period];
  const currentRows = data.BankExport.data;
  const currentOwnsPeriod = currentRows.length > 0 && currentRows.every(row => {
    try { return normalizePeriod(row['Tháng báo cáo'] || row._fileMonth) === period; } catch { return false; }
  });
  const oldSelected = previous?.months[period];
  // A missing table plus an old global "saved" marker is not evidence of an
  // intentional empty month. An explicit retained empty month or a live edit is.
  const deliberateEmpty = !currentRows.length && previous?.activePeriod === period &&
    (data.TransactionActivity?.lastAction === 'edited' || (oldSelected && !oldSelected.table.data.length &&
      oldSelected.activity?.lastAction === 'saved'));
  const currentChanged = currentOwnsPeriod && !sameTransactionSnapshot(currentRows, oldSelected?.table.data || selected?.table.data || [], period);
  if (selected && !deliberateEmpty && !currentChanged &&
    (data.BankExport !== selected.table || data.TransactionActivity !== selected.activity || previous?.activePeriod !== period)) {
    next = {...next, BankExport: selected.table, TransactionActivity: selected.activity};
    changed = true;
  }
  if (!changed) return data;
  return {...next, TransactionMonthCache: {...previous, activePeriod: period, months, cloudMonths}};
}
