import type { AppData, TableData } from '../../types';
import { deductionsRowMonth, isDeductionsMonthLocked } from './deductions-lock';
import { getHoldScopedIdentity, mergeDuplicateHoldRows, parsePayrollMonth, reconcileHoldTransactionRows, removeHoldCarryoverFromReport } from './hold-carryover';
import { buildBankExportRowsForMonth, calculateReconciliationTotals } from './reconciliation-sync';
import { markTransactionGenerated } from './transaction-activity';
import { withTransactionMonthOriginal } from './transaction-originals';

export interface MasterSourceScope {
  inputId: string;
  name: string;
  previousName?: string;
  month: string;
  replacementRequested?: boolean;
}

type Row = Record<string, unknown>;
type MonthCache = NonNullable<AppData['TransactionMonthCache']>;
type PendingUpdate = NonNullable<MonthCache['pendingMaster']>[string];

export function matchesMasterSource(row: Row, source: MasterSourceScope): boolean {
  if (parsePayrollMonth(row._fileMonth || row['Tháng báo cáo'] || row.month)?.dot !== source.month) return false;
  if (row._masterInputId) return row._masterInputId === source.inputId;
  const fileName = row['TÊN FILE'] || row._sourceFile;
  return Boolean(fileName && (fileName === source.name || fileName === source.previousName));
}

function periodOf(month: string | undefined): string {
  const parsed = parsePayrollMonth(month);
  return parsed ? `${parsed.year}-${String(parsed.month).padStart(2, '0')}` : '';
}

const importMetadata = new Set(['id', 'No', 'No.', '_masterInputId', '_fileMonth', '_recordId', '_uploadTimestamp', 'TÊN FILE', 'Tháng báo cáo']);

function sourceFingerprint(rows: Row[], source: MasterSourceScope): string {
  return JSON.stringify(rows.filter(row => matchesMasterSource(row, source)).map(row =>
    JSON.stringify(Object.entries(row).filter(([key]) => !importMetadata.has(key)).sort(([a], [b]) => a.localeCompare(b))),
  ).sort());
}

function hasUnappliedSavedSource(data: AppData, source: MasterSourceScope): boolean {
  const snapshot = Object.entries(data.HoldCarrySnapshots || {}).find(([month]) => parsePayrollMonth(month)?.dot === source.month)?.[1];
  if (!snapshot || !data.Hold_AE_Source) return false;
  const savedAt = Date.parse(snapshot.savedAt);
  const savedKeys = new Set((snapshot.rows || data.Hold_AE.data).map(row => getHoldScopedIdentity(row, source.month)));
  return data.Hold_AE_Source.data.some(row => {
    if (!matchesMasterSource(row, source) || !(Date.parse(String(row._uploadTimestamp || '')) > savedAt)) return false;
    const key = getHoldScopedIdentity(row, source.month);
    return Boolean(key && !savedKeys.has(key));
  });
}

function applyDeductionSources(data: AppData, month: string, sources: MasterSourceScope[]): AppData {
  if (!sources.length) return data;
  const owned = (row: Row) => deductionsRowMonth(row) === month && sources.some(source => matchesMasterSource(row, source));
  const incoming = new Map<string, Row>();
  const unkeyed: Row[] = [];
  for (const sourceRow of (data.Hold_AE_Source?.data || []).filter(owned)) {
    const row = structuredClone(sourceRow);
    const key = getHoldScopedIdentity(row, month);
    if (key) {
      row._recordId ||= key;
      const existing = incoming.get(key);
      const merged = existing && mergeDuplicateHoldRows([existing, row], {scopeByReportMonth: true});
      incoming.set(key, merged && merged.length === 1 ? merged[0] : row);
    } else unkeyed.push(row);
  }
  const previous = data.Hold_AE.data;
  const retained = previous.filter(row => {
    if (!owned(row) || row._holdCarryFromReportMonth) return true;
    const key = getHoldScopedIdentity(row, month);
    const replacement = incoming.get(key);
    if (replacement) {
      const merged = mergeDuplicateHoldRows([row, replacement], {scopeByReportMonth: true});
      if (merged.length === 1) incoming.set(key, merged[0]);
    }
    return false;
  });
  const resolved = reconcileHoldTransactionRows([...retained, ...incoming.values(), ...unkeyed]);
  return {...data, Hold_AE: {...data.Hold_AE, data: [
    ...previous.filter(row => deductionsRowMonth(row, data.globalMonth) !== month),
    ...resolved.filter(row => deductionsRowMonth(row, data.globalMonth) === month),
  ]}};
}

function refreshMasterBatch(data: AppData, month: string, allowEmpty: boolean): AppData {
  const period = periodOf(month);
  const current = periodOf(data.globalMonth);
  const bankRows = data.Bank_North_AE.data.filter(row => parsePayrollMonth(row._fileMonth || row['Tháng báo cáo'])?.dot === month);
  if (!bankRows.length && !allowEmpty) return data;
  const existing = data.TransactionMonthCache?.months[period];
  const table: TableData = {...(existing?.table || data.BankExport), data: buildBankExportRowsForMonth({
    ...data,
    Bank_North_AE: {...data.Bank_North_AE, data: bankRows},
    Sheet1_AE: {...data.Sheet1_AE, data: data.Sheet1_AE.data.filter(row => parsePayrollMonth(row._fileMonth || row['Tháng báo cáo'])?.dot === month)},
  }, month)};
  const activity = markTransactionGenerated({...data, TransactionActivity: period === current ? data.TransactionActivity : existing?.activity});
  const next: AppData = {
    ...data,
    ...(period === current ? {BankExport: table, TransactionActivity: activity} : {}),
    TransactionHistoryRestoreDisabled: false,
    TransactionMonthCache: {
      ...data.TransactionMonthCache,
      activePeriod: period === current ? period : data.TransactionMonthCache?.activePeriod || current,
      months: {...data.TransactionMonthCache?.months, [period]: {table, activity, masterDraft: true}},
    },
  };
  return withTransactionMonthOriginal(next, period, table);
}

function applyMonthUpdate(data: AppData, month: string, update: PendingUpdate): AppData {
  const period = periodOf(month);
  let next = applyDeductionSources(data, month, update.sources);
  next = refreshMasterBatch(next, month, update.refreshBatch);
  const pendingMaster = {...next.TransactionMonthCache?.pendingMaster};
  delete pendingMaster[period];
  if (next.TransactionMonthCache) {
    next = {...next, TransactionMonthCache: {...next.TransactionMonthCache, pendingMaster}};
  }
  return {...next, ReconciliationByMonth: {
    ...next.ReconciliationByMonth,
    [month]: {...calculateReconciliationTotals(next, month), generatedAt: new Date().toISOString()},
  }};
}

/** An actual source change starts a new working month. Reprocessing identical files preserves saved edits. */
export function applyMasterSourceRefresh(previous: AppData, imported: AppData, holdSources: MasterSourceScope[], bankSources: MasterSourceScope[]): AppData {
  const updates = new Map<string, PendingUpdate>();
  function updateFor(month: string): PendingUpdate {
    const existing = updates.get(month);
    if (existing) return existing;
    const pending = previous.TransactionMonthCache?.pendingMaster?.[periodOf(month)];
    const update = {sources: [...(pending?.sources || [])], refreshBatch: pending?.refreshBatch || false};
    updates.set(month, update);
    return update;
  }
  for (const source of holdSources) {
    const previousSource = previous.Hold_AE_Source?.data || previous.Hold_AE.data;
    if (!source.replacementRequested && !hasUnappliedSavedSource(previous, source) &&
      sourceFingerprint(previousSource, source) === sourceFingerprint(imported.Hold_AE_Source?.data || [], source)) continue;
    const update = updateFor(source.month);
    const existing = update.sources.findIndex(scope => scope.inputId === source.inputId);
    if (existing < 0) update.sources.push(source);
    else update.sources[existing] = source;
  }
  for (const source of bankSources) {
    if (source.replacementRequested || updates.has(source.month) ||
      sourceFingerprint(previous.Bank_North_AE.data, source) !== sourceFingerprint(imported.Bank_North_AE.data, source)) {
      updateFor(source.month).refreshBatch = true;
    }
  }

  let next = imported;
  for (const [month, update] of updates) {
    if (!isDeductionsMonthLocked(previous.HoldCarrySnapshots, month)) {
      next = applyMonthUpdate(next, month, update);
      continue;
    }
    const cache = next.TransactionMonthCache;
    next = {...next, TransactionMonthCache: {
      activePeriod: cache?.activePeriod || periodOf(next.globalMonth),
      months: cache?.months || {},
      ...cache,
      pendingMaster: {...cache?.pendingMaster, [periodOf(month)]: update},
    }};
  }
  return next;
}

export function pendingMasterUpdate(data: Pick<AppData, 'TransactionMonthCache'>, month: string | undefined): PendingUpdate | undefined {
  return data.TransactionMonthCache?.pendingMaster?.[periodOf(month)];
}

/** The existing Delete saved Hold action is the explicit unlock boundary. */
export function reopenMasterDeductionsMonth(data: AppData, month: string): AppData {
  const snapshots = Object.fromEntries(Object.entries(data.HoldCarrySnapshots || {}).filter(([key]) => periodOf(key) !== periodOf(month)));
  const next = {...data, HoldCarrySnapshots: snapshots, Hold_AE: {
    ...data.Hold_AE,
    data: removeHoldCarryoverFromReport(data.Hold_AE.data, month),
  }};
  const pending = pendingMasterUpdate(data, month);
  return pending ? applyMonthUpdate(next, month, pending) : next;
}
