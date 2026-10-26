import { parseMoneyToNumber } from "./data-utils";
import { trialBalancePeriod } from "./trial-balance-presentation";

export interface TrialBalanceHeaderTotals {
  payrollCost: number;
  hold: number;
  add: number;
  cancel: number;
}

/**
 * Use raw movements after approval clears the provisional columns. HOLD and
 * CANCEL belong to the report's occurrence month; ADD releases older months.
 * Keep the unscoped form for callers without monthly row metadata.
 */
export function calculateTrialBalanceHeaderTotals(
  rows: Array<{
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
  }>,
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
          const isBasePayroll = !row.customMonthDisplay && !/_(adjustment|cancel|add|hold)/.test(row.id || "");
          const isExplicitHold = /\bhold\b/i.test(row.customMonthDisplay || "") || /_hold/.test(row.id || "");
          if (isBasePayroll || isExplicitHold) {
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
