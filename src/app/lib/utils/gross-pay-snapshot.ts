import type { AppData, TableData } from '../../types';
import { parsePayrollMonth } from './hold-carryover';
import { payrollRowMonth, replacePayrollMonthRows } from './payroll-month-rows';

type GrossPaySource = Pick<AppData, 'Sheet1_AE' | 'SavedGrossPayByMonth'>;

export function grossPayTableForMonth(table: TableData | undefined, reportMonth: string): TableData | undefined {
  const month = parsePayrollMonth(reportMonth)?.dot;
  if (!table || !month) return undefined;
  const data = table.data.filter(row => row && !row._isSubtotal && !row._isTotalRow && payrollRowMonth(row) === month);
  return data.length ? {headers: table.headers, data} : undefined;
}

export function grossPayReportTable(source: GrossPaySource, reportMonth: string): TableData | undefined {
  const month = parsePayrollMonth(reportMonth)?.dot;
  if (!month) return undefined;
  return grossPayTableForMonth(source.Sheet1_AE, month) ?? grossPayTableForMonth(source.SavedGrossPayByMonth?.[month], month);
}

export function saveGrossPayMonth(data: AppData, month: string, source: TableData | undefined): AppData {
  const period = parsePayrollMonth(month)?.dot;
  if (!period) return data;
  const table = grossPayTableForMonth(source, period);
  if (!table) return data;
  return {...data, SavedGrossPayByMonth: {...data.SavedGrossPayByMonth, [period]: structuredClone(table)}};
}

export function restoreGrossPayMonth(data: AppData, month: string, useMasterSource = false): AppData {
  const period = parsePayrollMonth(month)?.dot;
  if (!period) return data;
  const saved = useMasterSource ? undefined : grossPayTableForMonth(data.SavedGrossPayByMonth?.[period], period);
  const table = saved ?? grossPayTableForMonth(data.TableOriginals?.Sheet1_AE, period);
  if (!table) return data;
  return {
    ...data,
    Sheet1_AE: {
      headers: [...new Set([...data.Sheet1_AE.headers, ...table.headers])],
      data: replacePayrollMonthRows(data.Sheet1_AE.data, period, structuredClone(table.data)),
    },
  };
}
