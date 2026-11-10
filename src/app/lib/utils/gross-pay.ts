import { parseMoneyToNumber } from "./data-utils";
import { parsePayrollMonth } from "./hold-carryover";
import { payrollRowMonth } from "./payroll-month-rows";

export function normalizeGrossPayBusiness(value: unknown): string {
  const business = String(value || "").trim().toUpperCase();
  if (business === "AHN_HP") return "AHP";
  return business === "UNKNOWN" || !business ? "OTHER" : business;
}

/** Report totals come from employee rows in Gross Pay, before deductions. */
export function grossPayTotalsByMonth(rows: readonly Record<string, unknown>[], reportMonth?: string) {
  const selected = reportMonth === undefined ? undefined : parsePayrollMonth(reportMonth)?.dot;
  const totals = new Map<string, Map<string, number>>();
  if (reportMonth !== undefined && !selected) return totals;
  for (const row of rows) {
    if (!row || row._isSubtotal || row._isTotalRow ||
        !String(row["ID Number"] || row.id_number || row["Mã AE"] || row["Mã ae"] || "").trim()) continue;
    const month = payrollRowMonth(row);
    if (!month || (selected && month !== selected)) continue;
    const business = normalizeGrossPayBusiness(row.Business || row.BU);
    const rawAmount = [row["TOTAL PAYMENT"], row["Payment Amount"], row["Grand Total"], row["GRAND TOTAL"], row["Total Payment"]]
      .find(value => value != null && String(value).trim() !== "");
    const amount = parseMoneyToNumber(rawAmount);
    const monthTotals = totals.get(month) || new Map<string, number>();
    monthTotals.set(business, (monthTotals.get(business) || 0) + amount);
    totals.set(month, monthTotals);
  }
  return totals;
}

export const GROSS_PAY_CHARGE_COLUMNS = [
  "CHARGE TO LXO",
  "CHARGE TO EC",
  "CHARGE TO PT-DEMO",
  "Charge MKT Local",
  "CHARGE TO OTHER",
  "Charge Renewal Projects",
  "Charge Discovery Camp",
  "Charge Summer Outing",
  "Charge Summer Instructors",
  "Extra Summer Instructors",
] as const;

/**
 * Keep the Gross Pay total supplied by Sheet 1 as the source of truth. Some
 * workbooks contain additional charge columns that are not part of the
 * standard display schema, so rebuilding the total from a fixed list can
 * silently discard valid payroll amounts.
 */
export function resolveGrossPayTotal(row: Record<string, unknown>): number {
  const sourceTotal = parseMoneyToNumber(row["TOTAL PAYMENT"]);
  if (sourceTotal !== 0) return sourceTotal;

  return GROSS_PAY_CHARGE_COLUMNS.reduce(
    (sum, column) => sum + parseMoneyToNumber(row[column]),
    0,
  );
}

/**
 * Replace only the Extra Summer Instructors portion of a Gross Pay row.
 * Reapplying the same bonus is idempotent and never rebuilds TOTAL PAYMENT
 * from a partial list of charge columns.
 */
export function applyExtraSummerInstructorBonus(
  row: Record<string, unknown>,
  bonusAmount: unknown,
): Record<string, unknown> {
  const nextBonus = Math.abs(parseMoneyToNumber(bonusAmount));
  const previousBonus =
    Math.abs(parseMoneyToNumber(row["Extra Summer Instructors"])) ||
    Math.abs(
      parseMoneyToNumber(row["CHARGE TO EXTRA SUMMER INSTRUCTORS"]),
    );
  const currentTotal = resolveGrossPayTotal(row);

  return {
    ...row,
    "Extra Summer Instructors": nextBonus,
    "CHARGE TO EXTRA SUMMER INSTRUCTORS": nextBonus,
    "TOTAL PAYMENT": currentTotal - previousBonus + nextBonus,
  };
}
