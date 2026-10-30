import type { AppData, TableData } from '../../types';
import { parsePayrollMonth, reconcileHoldTransactionRows, mergeDuplicateHoldRows } from './hold-carryover';
import { createDeductionsRowProjector, createDeductionsDisplayProjector, getDeductionsHeaders } from './deductions-view';
import { createDeductionsPrioritizer } from './deductions-display';
import { sortMissingDeductionsSourceNotesFirst } from './deductions-sheet-source';
import { calculateReconciliationTotals } from './reconciliation-sync';

type Locks = AppData['HoldCarrySnapshots'];
export interface DeductionsWriteOptions { unlockDeductionsMonth?: string }

export function deductionsRowMonth(row: TableData['data'][number], fallback?: string): string | undefined {
  return parsePayrollMonth(row?.['Tháng báo cáo'] || row?._fileMonth || row?.Tháng || fallback)?.dot;
}

export function isDeductionsMonthLocked(locks: Locks, month: unknown): boolean {
  const period = parsePayrollMonth(month)?.dot;
  return Boolean(period && Object.keys(locks || {}).some(key => parsePayrollMonth(key)?.dot === period));
}

/** Upgrade metadata-only saved months before any migration or write can change them. */
export function captureSavedDeductions(data: AppData): AppData {
  let snapshots = data.HoldCarrySnapshots;
  for (const [key, snapshot] of Object.entries(snapshots || {})) {
    if (snapshot.rows && snapshot.displayRows && snapshot.reconciliationTotals) continue;
    const month = parsePayrollMonth(key)?.dot;
    if (!month) continue;
    const rows = snapshot.rows || structuredClone((data.Hold_AE?.data || []).filter(row => deductionsRowMonth(row, data.globalMonth) === month));
    const sourceRows = [...(data.Hold_AE?.data || []).filter(row => deductionsRowMonth(row, data.globalMonth) !== month), ...rows];
    const projected = createDeductionsRowProjector(month)(sourceRows);
    const displayRows = reconcileHoldTransactionRows(projected).filter(row => {
      const operation = String(row['Nghiệp vụ'] || '').toUpperCase();
      return deductionsRowMonth(row) === month && !/BONUS|^B$|⏩|⏯/.test(operation);
    }).map(createDeductionsDisplayProjector());
    snapshots = {...snapshots, [key]: {...snapshot, rows,
      displayRows: snapshot.displayRows || structuredClone(createDeductionsPrioritizer()(sortMissingDeductionsSourceNotesFirst(mergeDuplicateHoldRows(displayRows)))),
      displayHeaders: snapshot.displayHeaders || getDeductionsHeaders(data.Hold_AE?.headers || []),
      reconciliationTotals: snapshot.reconciliationTotals || calculateReconciliationTotals({...data, Hold_AE: {...data.Hold_AE, data: projected}}, month),
    }};
  }
  return snapshots === data.HoldCarrySnapshots ? data : {...data, HoldCarrySnapshots: snapshots};
}

/** Every local write, including Undo, imports and automation, preserves saved months. */
export function protectSavedDeductions(previous: AppData, candidate: AppData, options: DeductionsWriteOptions = {}): AppData {
  const baseline = captureSavedDeductions(previous);
  const unlocked = parsePayrollMonth(options.unlockDeductionsMonth)?.dot;
  const locked = new Map(Object.entries(baseline.HoldCarrySnapshots || {}).flatMap(([key, snapshot]) => {
    const month = parsePayrollMonth(key)?.dot;
    return month && month !== unlocked ? [[month, {key, snapshot}] as const] : [];
  }));
  if (!locked.size) return candidate;
  const remaining = new Set(locked.keys());
  // Keep snapshot objects separate from writable rows, including legacy writers
  // that still mutate a row in place before returning their proposed table.
  const retained = new Map([...locked].map(([month, {snapshot}]) => {
    const previousRows = (previous.Hold_AE?.data || []).filter(row => deductionsRowMonth(row, previous.globalMonth) === month);
    return [month, JSON.stringify(previousRows) === JSON.stringify(snapshot.rows)
      ? previousRows : structuredClone(snapshot.rows!)];
  }));
  const rows: TableData['data'] = [];
  for (const row of candidate.Hold_AE?.data || []) {
    const month = deductionsRowMonth(row, candidate.globalMonth);
    if (!month || !locked.has(month)) { rows.push(row); continue; }
    if (remaining.delete(month)) rows.push(...retained.get(month)!);
  }
  for (const month of remaining) rows.push(...retained.get(month)!);
  const oldRows = previous.Hold_AE?.data || [];
  const proposedRows = candidate.Hold_AE?.data || [];
  const sameRows = (other: TableData['data']) => rows.length === other.length && rows.every((row, index) => row === other[index]);
  const table = sameRows(oldRows) ? previous.Hold_AE : sameRows(proposedRows) ? candidate.Hold_AE : {...candidate.Hold_AE, data: rows};
  let snapshots = candidate.HoldCarrySnapshots;
  for (const {key, snapshot} of locked.values()) {
    if (snapshots?.[key] !== snapshot) snapshots = {...snapshots, [key]: snapshot};
  }
  const result = table === candidate.Hold_AE && snapshots === candidate.HoldCarrySnapshots
    ? candidate : {...candidate, Hold_AE: table, HoldCarrySnapshots: snapshots};
  const keys = Object.keys(result) as (keyof AppData)[];
  return keys.length === Object.keys(previous).length && keys.every(key => result[key] === previous[key]) ? previous : result;
}

/** Cloud restoration follows the incoming lock/unlock state, never an old device's lock. */
export function restoreSavedDeductions(data: AppData): AppData {
  const captured = captureSavedDeductions(data);
  return protectSavedDeductions(captured, captured);
}
