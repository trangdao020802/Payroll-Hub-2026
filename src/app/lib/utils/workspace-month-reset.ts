import type { AppData, TableData } from '../../types';
import { normalizeDateFilterValue } from './data-utils';
import { isDeductionsMonthLocked } from './deductions-lock';
import { restoreGrossPayMonth } from './gross-pay-snapshot';
import { parsePayrollMonth } from './hold-carryover';
import { payrollRowMonth, replacePayrollMonthRows } from './payroll-month-rows';
import { normalizePivotMonth } from './pivot-utils';
import { calculateReconciliationTotals } from './reconciliation-sync';
import { restoreTableOriginals } from './table-originals';
import { normalizePeriod } from './transaction-history';

function rosterMonth(row: Record<string, unknown>): string | undefined {
  if (!row) return undefined;
  const reportMonth = payrollRowMonth(row);
  if (reportMonth) return reportMonth;
  const date = normalizeDateFilterValue(row.ngay || row.date || row.Date || row['Ngày'] || row.TK_Date || row['Session Date']);
  return date.match(/^\d{2}\/(\d{2})\/(\d{4})$/)?.slice(1).join('.');
}

function restoreMonthTable(table: TableData, original: TableData | undefined, month: string): TableData {
  const originalRows = original?.data.filter(row => payrollRowMonth(row) === month);
  if (!original || !originalRows?.length) return table;
  return {
    headers: [...new Set([...table.headers, ...original.headers])],
    data: replacePayrollMonthRows(table.data, month, structuredClone(originalRows)),
  };
}

export function restoreWorkspaceMonth(data: AppData, reportMonth: string, resetAt = new Date().toISOString()) {
  const month = parsePayrollMonth(reportMonth)?.dot;
  if (!month || month !== parsePayrollMonth(data.globalMonth || '03.2026')?.dot) {
    return {data, batchRestored: false};
  }

  const period = normalizePeriod(month);
  const originals = data.TableOriginals;
  const masterChanged = Boolean(data.TransactionMonthCache?.months[period]?.masterDraft || data.TransactionMonthCache?.pendingMaster?.[period]);
  const nextGross = restoreGrossPayMonth(data, month, masterChanged);
  const next: AppData = {
    ...nextGross,
    Bank_North_AE: restoreMonthTable(data.Bank_North_AE, originals?.Bank_North_AE, month),
    Hold_AE: isDeductionsMonthLocked(data.HoldCarrySnapshots, month) ? data.Hold_AE : restoreMonthTable(data.Hold_AE, originals?.Hold_AE, month),
    PivotMasterClearedMonths: data.PivotMasterClearedMonths?.filter(value => normalizePivotMonth(value) !== month),
    AuditClearedTables: {...data.AuditClearedTables, main: false, detail: false},
  };
  for (const field of ['Timesheet_Roster', 'Master_Roster', 'Q_Cache'] as const) {
    const originalRows = originals?.[field]?.filter(row => rosterMonth(row) === month);
    if (originalRows?.length) next[field] = replacePayrollMonthRows(data[field], month, structuredClone(originalRows), rosterMonth);
  }

  let restored = next;
  let batchRestored = false;
  try {
    restored = restoreTableOriginals(next, ['BankExport']);
    batchRestored = true;
  } catch {
    // No verified baseline for this month. Keep the working Batch Payment.
  }
  return {
    batchRestored,
    data: {
      ...restored,
      updatedAt: resetAt,
      ReconciliationByMonth: {
        ...restored.ReconciliationByMonth,
        [month]: {...calculateReconciliationTotals(restored, month), generatedAt: resetAt},
      },
    },
  };
}
