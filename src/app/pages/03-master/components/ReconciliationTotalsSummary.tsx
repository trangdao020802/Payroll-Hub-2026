import { Popover, PopoverContent, PopoverTrigger } from '../../../components/ui/popover';
import { formatVNRobust } from '../../../lib/utils/format-utils';
import type { ReconciliationTotals } from '../../../lib/utils/reconciliation-sync';

export function ReconciliationTotalsSummary({month, totals}: {month: string; totals: ReconciliationTotals}) {
  const amount = (value: number) => formatVNRobust(value, 0);
  const source = totals.actualSource === 'bank-ae' ? 'Bank AE (nguồn)' : 'Batch Payment';
  return <Popover>
    <PopoverTrigger asChild>
      <button type="button" aria-label={`Chi tiết chênh lệch tháng ${month}`}
        title="Xem nguồn thực chi, Gross Pay, Deductions và công thức chênh lệch"
        className="flex flex-col items-end rounded px-1 py-0.5 hover:bg-muted/50 active:scale-[0.98]">
        <span className="text-[8px] font-bold text-muted-foreground uppercase tracking-tighter whitespace-nowrap">TOTAL VARIANCE</span>
        <span className={`text-xs font-black leading-tight tabular-nums ${!totals.hasActual ? 'text-amber-700' : Math.abs(totals.variance) < 1 ? 'text-emerald-600' : 'text-rose-600'}`}>
          {totals.hasActual ? amount(totals.variance) : 'Chưa có Batch'}
        </span>
      </button>
    </PopoverTrigger>
    <PopoverContent align="end" className="w-80 space-y-2 text-xs">
      <p className="font-semibold">Đối chiếu tháng {month}</p>
      <dl className="space-y-1.5">
        <div className="flex justify-between gap-3"><dt>{source} · thực chi</dt><dd className="tabular-nums font-semibold">{totals.hasActual ? amount(totals.actual) : 'Chưa có dữ liệu'}</dd></div>
        <div className="flex justify-between gap-3"><dt>Gross Pay</dt><dd className="tabular-nums">{amount(totals.grossPayTotal)}</dd></div>
        <div className="flex justify-between gap-3"><dt>Deductions tính lương</dt><dd className="tabular-nums">{amount(totals.deductionsTotal)}</dd></div>
        <div className="flex justify-between gap-3 border-t border-border pt-1.5"><dt>Phải chi</dt><dd className="tabular-nums font-semibold">{amount(totals.expected)}</dd></div>
      </dl>
      <p className="text-muted-foreground">Phải chi = Gross Pay + Deductions. Chênh lệch = Thực chi − Phải chi.</p>
      {!totals.hasActual && <p role="status" className="text-amber-700">Thiếu Batch Payment đúng tháng. Bấm Đồng bộ ngay để tải bản đã lưu; chưa thể kết luận chênh lệch.</p>}
    </PopoverContent>
  </Popover>;
}
