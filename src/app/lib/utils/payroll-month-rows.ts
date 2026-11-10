import { parsePayrollMonth } from './hold-carryover';

type Row = Record<string, unknown>;

export function payrollRowMonth(row: Row): string | undefined {
  if (!row) return undefined;
  for (const value of [row['Tháng báo cáo'], row._fileMonth, row['Tháng'], row.Month, row.month, row['Report Month']]) {
    const month = parsePayrollMonth(value)?.dot;
    if (month) return month;
  }
}

/** Replace the selected month's rows without changing the order of the other months. */
export function replacePayrollMonthRows(rows: Row[], month: string, replacement: Row[], rowMonth = payrollRowMonth): Row[] {
  const result: Row[] = [];
  let inserted = false;
  for (const row of rows) {
    if (rowMonth(row) !== month) {
      result.push(row);
    } else if (!inserted) {
      for (const restored of replacement) result.push(restored);
      inserted = true;
    }
  }
  return inserted ? result : result.concat(replacement);
}
