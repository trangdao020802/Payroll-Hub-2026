import type { AppData, TableData } from '../../types';
import { getTransactionMonthOriginal, trackTransactionMonthOriginal } from './transaction-originals';
import { normalizePeriod } from './transaction-history';
import { replaceTransactionPeriod } from './transaction-snapshot';
import { markTransactionSaved } from './transaction-activity';
import { calculateReconciliationTotals } from './reconciliation-sync';

export const ORIGINAL_FIELDS = ['Sheet1_AE', 'Bank_North_AE', 'Hold_AE', 'BankExport',
  'Timesheet_Roster', 'Master_Roster', 'Q_Staff', 'Q_Salary_Scale', 'Q_Cache',
  'Q_TeacherHours', 'Q_CheckTAs', 'Q_AllowedTARules', 'SavedBal_PayrollTrial',
  'Timesheet_InputList', 'Ae_Global_Inputs'] as const;
export type OriginalField = typeof ORIGINAL_FIELDS[number];
export type TableOriginals = Partial<Pick<AppData, OriginalField>> & {
  BankExportByMonth?: Record<string, TableData>;
};

/** Explicit imports replace the baseline; edits retain the first pre-edit value. */
export function trackTableOriginals(previous: AppData, next: AppData, isEdit: boolean,
  sourceFields: readonly OriginalField[] = []): AppData {
  if (previous === next || next.TableOriginals !== previous.TableOriginals) return next;
  const tracked = trackTransactionMonthOriginal(previous, next, isEdit, sourceFields.includes('BankExport'));
  next = tracked;
  const originals = { ...next.TableOriginals };
  let changed = tracked.TableOriginals !== previous.TableOriginals;
  for (const key of ORIGINAL_FIELDS) {
    if (key === 'BankExport') continue;
    if (previous[key] === next[key]) continue;
    if (sourceFields.includes(key)) {
      Object.assign(originals, { [key]: structuredClone(next[key]) });
      changed = true;
    } else if (isEdit && !Object.prototype.hasOwnProperty.call(originals, key)) {
      Object.assign(originals, { [key]: structuredClone(previous[key]) });
      changed = true;
    }
  }
  return changed ? { ...next, TableOriginals: originals } : next;
}

export function restoreTableOriginals(data: AppData, fields: readonly OriginalField[]): AppData {
  if (!fields.length || fields.some(key => key !== 'BankExport' && !Object.prototype.hasOwnProperty.call(data.TableOriginals || {}, key))) {
    throw new Error('Chưa có bản gốc của bảng này. Hãy tải và xử lý lại file nguồn để tạo bản khôi phục.');
  }
  const restored = { ...data };
  for (const key of fields) {
    if (key !== 'BankExport') {
      Object.assign(restored, { [key]: structuredClone(data.TableOriginals![key]) });
      continue;
    }
    const original = getTransactionMonthOriginal(data);
    const period = normalizePeriod(data.globalMonth);
    const [year, month] = period.split('-');
    const reportMonth = `${month}.${year}`;
    const savedAt = new Date().toISOString();
    restored.BankExport = {...original, data: replaceTransactionPeriod(data.BankExport.data, period, original.data)};
    restored.TransactionActivity = markTransactionSaved(data, savedAt);
    restored.TransactionMonthCache = {
      ...data.TransactionMonthCache,
      activePeriod: period,
      months: {
        ...data.TransactionMonthCache?.months,
        [period]: {table: original, activity: restored.TransactionActivity,
          cloudVersionId: data.TransactionMonthCache?.cloudMonths?.[period]?.id},
      },
    };
    restored.ReconciliationByMonth = {
      ...data.ReconciliationByMonth,
      [reportMonth]: {...calculateReconciliationTotals(restored, reportMonth), generatedAt: savedAt},
    };
  }
  return restored;
}

export function grossPayRowVisible(row: Record<string, unknown>): boolean {
  return row._isNew === true || String(row['ID Number'] || row.id_number || '').trim() !== '';
}

export function sourceRowIndex(rows: Record<string, unknown>[], row: Record<string, unknown> | undefined): number {
  if (!row) return -1;
  if (row.id) return rows.findIndex(item => item.id === row.id);
  if (typeof row._originalIndex === 'number') return row._originalIndex;
  return rows.indexOf(row);
}
