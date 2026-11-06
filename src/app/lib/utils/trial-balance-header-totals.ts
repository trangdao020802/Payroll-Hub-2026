import { parseMoneyToNumber } from "./data-utils";
import { trialBalancePeriod } from "./trial-balance-presentation";

export interface TrialBalanceHeaderTotals {
  payrollCost: number;
  hold: number;
  add: number;
  cancel: number;
}

type TrialBalanceTotalsRow = {
  id?: string;
  month?: string;
  reportMonth?: string;
  displayMonth?: string;
  customMonthDisplay?: string;
  thu?: unknown;
  chi?: unknown;
  rawHold?: unknown;
  rawAdd?: unknown;
  rawCancel?: unknown;
  hold?: unknown;
  add?: unknown;
  cancel?: unknown;
  _excludeFromTotals?: boolean;
  _isOpeningHold?: boolean;
};

function isBasePayroll(row: TrialBalanceTotalsRow): boolean {
  return !row.customMonthDisplay && !/_(adjustment|cancel|add|hold)/.test(row.id || "");
}

/** Sum earned payroll for a single BU. Hold and release rows settle amounts already earned. */
export function calculateTrialBalanceInternCost(
  rows: readonly TrialBalanceTotalsRow[],
  reportMonth: string,
): number {
  const period = trialBalancePeriod(reportMonth);
  if (!period) return 0;

  return rows.reduce((total, row) => {
    if (!row || row._excludeFromTotals || row._isOpeningHold) return total;
    const report = trialBalancePeriod(row.reportMonth || row.month || "");
    const origin = trialBalancePeriod(row.displayMonth || row.month || "");
    const isLegacyBonus = /_bonus\b|_bonus_/.test(row.id || "");
    if (report !== period || origin !== period || (!isBasePayroll(row) && !isLegacyBonus)) return total;
    return total + parseMoneyToNumber(row.thu);
  }, 0);
}

/**
 * Use raw movements after approval clears the provisional columns. HOLD and
 * CANCEL belong to the report's occurrence month; ADD releases older months.
 * Keep the unscoped form for callers without monthly row metadata.
 */
export function calculateTrialBalanceHeaderTotals(
  rows: readonly TrialBalanceTotalsRow[],
  reportMonth?: string,
): TrialBalanceHeaderTotals {
  return rows.reduce<TrialBalanceHeaderTotals>(
    (totals, row) => {
      if (!row || row._excludeFromTotals) return totals;
      const report = trialBalancePeriod(reportMonth || "");
      if (reportMonth && trialBalancePeriod(row.reportMonth || row.month || "") !== report) return totals;

      totals.payrollCost += parseMoneyToNumber(row.thu ?? 0);
      if (!reportMonth) {
        totals.hold += Math.abs(parseMoneyToNumber(row.hold ?? 0));
        totals.add += Math.abs(parseMoneyToNumber(row.add ?? 0));
        totals.cancel += Math.abs(parseMoneyToNumber(row.cancel ?? 0));
      } else {
        const origin = trialBalancePeriod(row.displayMonth || row.month || "");
        if (origin === report) {
          const isExplicitHold = /\bhold\b/i.test(row.customMonthDisplay || "") || /_hold/.test(row.id || "");
          if (isBasePayroll(row) || isExplicitHold) {
            totals.hold += Math.abs(parseMoneyToNumber(row.rawHold || (!row.customMonthDisplay ? row.chi : (row.hold || row.chi)) || 0));
          }
          const cancelVal = Math.abs(parseMoneyToNumber(row.rawCancel ?? row.cancel ?? 0));
          if (cancelVal > 0) {
            totals.cancel += cancelVal;
          }
        }
        if (origin > 0 && origin < report) totals.add += Math.abs(parseMoneyToNumber(row.rawAdd ?? row.add ?? 0));
      }
      return totals;
    },
    { payrollCost: 0, hold: 0, add: 0, cancel: 0 },
  );
}
