import { RefreshCw, Settings } from 'lucide-react';
import { DropdownMenu, DropdownMenuTrigger, DropdownMenuContent, DropdownMenuItem } from './ui/dropdown-menu';
import { useTableRestore } from '../hooks/useTableRestore';
import type { OriginalField } from '../lib/utils/table-originals';

export function TableRestoreButton({ fields, onRestore, placement = 'settings', title: customTitle }: { fields?: readonly OriginalField[]; onRestore?: () => void; placement?: 'settings' | 'menu' | 'plain'; title?: string }) {
  const restore = useTableRestore();
  const action = onRestore || (() => restore(fields || []));
  const title = customTitle || (onRestore ? 'Tính lại bảng từ dữ liệu nguồn, bỏ chỉnh sửa riêng của bảng.' : 'Khôi phục toàn bộ dữ liệu của bảng trước chỉnh sửa (tất cả tháng), gồm dòng đã thêm hoặc xóa. Có thể Hoàn tác.');
  const content = <><RefreshCw className="h-3.5 w-3.5 shrink-0" /><span className="truncate flex-1 text-left font-semibold">Reset to default</span></>;
  const className = 'flex w-full items-center gap-2 whitespace-nowrap rounded-lg px-3 py-2 text-xs text-primary hover:bg-muted cursor-pointer';
  if (placement === 'plain') return <button type="button" onClick={action} title={title} className={className}>{content}</button>;
  const item = <DropdownMenuItem onSelect={action} title={title} className={className}>{content}</DropdownMenuItem>;
  if (placement === 'menu') return item;
  return <DropdownMenu><DropdownMenuTrigger asChild>
    <button type="button" title="Cài đặt & Thao tác" aria-label="Cài đặt & Thao tác" className="inline-flex h-8 w-8 items-center justify-center rounded-full border border-border text-primary hover:bg-muted active:scale-[0.98]"><Settings className="h-3.5 w-3.5" /></button>
  </DropdownMenuTrigger><DropdownMenuContent align="end">{item}</DropdownMenuContent></DropdownMenu>;
}
