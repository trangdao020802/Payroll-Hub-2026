import type { AppData, TableData } from '../../types';
import { normalizePeriod } from './transaction-history';

function periodOf(value: unknown): string | undefined {
  try { return normalizePeriod(String(value || '')); } catch { return undefined; }
}

/** A legacy table proves ownership only through explicit row months. */
function tableForPeriod(table: TableData | undefined, period: string): TableData | undefined {
  if (!table?.data.length) return undefined;
  const rows = table.data.filter(row => periodOf(row['Tháng báo cáo'] || row._fileMonth) === period);
  return rows.length ? {...table, data: rows} : undefined;
}

export function withTransactionMonthOriginal(data: AppData, period: string, table: TableData): AppData {
  const normalized = normalizePeriod(period);
  if (table.data.some(row => periodOf(row['Tháng báo cáo'] || row._fileMonth) !== normalized)) {
    throw new Error('Bản gốc Batch Payment phải thuộc đúng tháng đã chọn.');
  }
  return {
    ...data,
    TableOriginals: {
      ...data.TableOriginals,
      BankExportByMonth: {
        ...data.TableOriginals?.BankExportByMonth,
        [normalized]: structuredClone(table),
      },
    },
  };
}

/** Preserve the first pre-edit value separately for each month, never across navigation. */
export function trackTransactionMonthOriginal(previous: AppData, next: AppData, isEdit: boolean, isSource: boolean): AppData {
  if (previous.BankExport === next.BankExport) return next;
  const period = periodOf(next.globalMonth);
  if (!period) return next;
  if (isSource) {
    const table = tableForPeriod(next.BankExport, period);
    return table ? withTransactionMonthOriginal(next, period, table) : next;
  }
  if (!isEdit || periodOf(previous.globalMonth) !== period || next.TableOriginals?.BankExportByMonth?.[period]) return next;
  const table = tableForPeriod(previous.TableOriginals?.BankExport, period) || tableForPeriod(previous.BankExport, period);
  if (table) return withTransactionMonthOriginal(next, period, table);
  // An empty active month is also a valid baseline, but another month's empty
  // preview must not become this month's original during automatic navigation.
  return !previous.BankExport.data.length && previous.TransactionMonthCache?.activePeriod === period
    ? withTransactionMonthOriginal(next, period, previous.BankExport)
    : next;
}

export function getTransactionMonthOriginal(data: AppData): TableData {
  const period = normalizePeriod(data.globalMonth);
  // A verified saved month supersedes the generated pre-edit baseline. Reset
  // must not undo saved amounts or identity corrections after a cloud restore.
  const saved = data.TransactionMonthCache?.cloudMonths?.[period];
  if (saved) {
    if (normalizePeriod(saved.period.slice(0, 7)) !== period ||
        saved.rows.some(row => periodOf(row['Tháng báo cáo'] || row._fileMonth) !== period)) {
      throw new Error(`Batch Payment đã lưu không đúng tháng ${data.globalMonth}. Dữ liệu được giữ nguyên.`);
    }
    return structuredClone({...data.BankExport, data: saved.rows});
  }
  const original = data.TableOriginals?.BankExportByMonth?.[period];
  if (original) {
    if (original.data.some(row => periodOf(row['Tháng báo cáo'] || row._fileMonth) !== period)) {
      throw new Error(`Bản gốc Batch Payment không đúng tháng ${data.globalMonth}. Dữ liệu được giữ nguyên.`);
    }
    return structuredClone(original);
  }
  const legacy = tableForPeriod(data.TableOriginals?.BankExport, period);
  if (legacy) return structuredClone(legacy);
  throw new Error(`Chưa có bản gốc Batch Payment đúng tháng ${data.globalMonth}. Dùng Bỏ lọc để hiển thị lại các dòng; dữ liệu được giữ nguyên.`);
}
