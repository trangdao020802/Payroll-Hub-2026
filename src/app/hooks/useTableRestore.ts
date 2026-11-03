import { useCallback } from 'react';
import { toast } from 'sonner';
import { useAppData } from '../lib/contexts/AppDataContext';
import { restoreTableOriginals, type OriginalField } from '../lib/utils/table-originals';
import { normalizePeriod } from '../lib/utils/transaction-history';

export function useTableRestore() {
  const { appData, updateAppData } = useAppData();
  return useCallback((fields: readonly OriginalField[]) => {
    try {
      restoreTableOriginals(appData, fields);
      updateAppData(prev => restoreTableOriginals(prev, fields), true, true);
      let message = 'Đã khôi phục dữ liệu trước chỉnh sửa của bảng.';
      if (fields.includes('BankExport')) {
        const period = normalizePeriod(appData.globalMonth);
        const prefix = `Đã khôi phục Batch Payment tháng ${appData.globalMonth}`;
        if (appData.TransactionMonthCache?.months[period]?.masterDraft) message = `${prefix} theo nguồn Master mới.`;
        else if (appData.TransactionMonthCache?.cloudMonths?.[period]) message = `${prefix} theo bản tháng đã lưu gần nhất.`;
        else message = `${prefix}. Bấm Lưu tháng để cập nhật lịch sử Check STK & ID.`;
      }
      toast.success(message);
      return true;
    } catch (error) {
      toast.error(error instanceof Error ? error.message : 'Không thể khôi phục dữ liệu.');
      return false;
    }
  }, [appData, updateAppData]);
}
