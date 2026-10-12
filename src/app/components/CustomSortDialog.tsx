import { useEffect, useMemo, useState } from "react";
import {
  ArrowUpDown,
  ChevronDown,
  ChevronUp,
  Plus,
  Trash2,
  X,
} from "lucide-react";

import type { Column } from "./DataTable";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "./ui/select";
import {
  normalizeSortRules,
  type SortRule,
} from "./sort-utils";

interface CustomSortDialogProps {
  open: boolean;
  columns: Column[];
  value: SortRule[];
  onOpenChange: (open: boolean) => void;
  onApply: (rules: SortRule[]) => void;
}

export function CustomSortDialog({
  open,
  columns,
  value,
  onOpenChange,
  onApply,
}: CustomSortDialogProps) {
  const sortableColumns = useMemo(
    () => columns.filter((column) => column.sortable !== false),
    [columns],
  );
  const [draft, setDraft] = useState<SortRule[]>(() => normalizeSortRules(value));

  useEffect(() => {
    if (!open) return;
    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key === "Escape") onOpenChange(false);
    };
    window.addEventListener("keydown", handleKeyDown);
    return () => window.removeEventListener("keydown", handleKeyDown);
  }, [open, onOpenChange]);

  if (!open) return null;

  const firstUnusedColumn = () =>
    sortableColumns.find((column) => !draft.some((rule) => rule.key === column.key));

  const addLevel = () => {
    const column = firstUnusedColumn();
    if (!column) return;
    setDraft((current) => [...current, { key: column.key, direction: "asc" }]);
  };

  const updateColumn = (index: number, key: string) => {
    setDraft((current) => {
      const duplicateIndex = current.findIndex((rule, ruleIndex) => ruleIndex !== index && rule.key === key);
      const next = current.filter((_, ruleIndex) => ruleIndex !== duplicateIndex);
      const targetIndex = duplicateIndex >= 0 && duplicateIndex < index ? index - 1 : index;
      next[targetIndex] = { ...next[targetIndex], key };
      return next;
    });
  };

  const moveLevel = (index: number, offset: -1 | 1) => {
    setDraft((current) => {
      const target = index + offset;
      if (target < 0 || target >= current.length) return current;
      const next = [...current];
      [next[index], next[target]] = [next[target], next[index]];
      return next;
    });
  };

  return (
    <div
      className="fixed inset-0 z-[10050] grid place-items-center bg-slate-950/35 p-4 backdrop-blur-[2px]"
      role="presentation"
      onMouseDown={() => onOpenChange(false)}
    >
      <section
        role="dialog"
        aria-modal="true"
        aria-labelledby="custom-sort-title"
        className="w-full max-w-[680px] overflow-hidden rounded-2xl border border-border bg-card text-foreground shadow-2xl"
        onMouseDown={(event) => event.stopPropagation()}
      >
        <header className="flex items-start justify-between gap-4 border-b border-border bg-[var(--table-header-bg,#FAF3E8)] px-5 py-4">
          <div className="min-w-0">
            <div className="flex items-center gap-2">
              <span className="grid h-8 w-8 shrink-0 place-items-center rounded-lg bg-primary text-primary-foreground shadow-sm">
                <ArrowUpDown className="h-4 w-4" />
              </span>
              <div>
                <h2 id="custom-sort-title" className="text-sm font-black uppercase tracking-wide">
                  Sắp xếp tùy chỉnh
                </h2>
                <p className="mt-0.5 text-[11px] font-medium text-muted-foreground">
                  Cột ở cấp trên được ưu tiên trước, tương tự Custom Sort của Excel.
                </p>
              </div>
            </div>
          </div>
          <button
            type="button"
            onClick={() => onOpenChange(false)}
            className="grid h-8 w-8 shrink-0 place-items-center rounded-full text-muted-foreground transition-colors hover:bg-primary/10 hover:text-primary"
            aria-label="Đóng"
          >
            <X className="h-4 w-4" />
          </button>
        </header>

        <div className="flex items-center justify-between gap-3 border-b border-border/70 px-5 py-2.5">
          <button
            type="button"
            onClick={addLevel}
            disabled={!firstUnusedColumn()}
            className="inline-flex h-8 items-center gap-1.5 rounded-lg border border-primary/25 bg-primary/5 px-3 text-[11px] font-black uppercase tracking-wide text-primary transition-colors hover:bg-primary/10 disabled:cursor-not-allowed disabled:opacity-40"
          >
            <Plus className="h-3.5 w-3.5" /> Thêm cấp
          </button>
          <button
            type="button"
            onClick={() => setDraft([])}
            disabled={draft.length === 0}
            className="h-8 px-2 text-[11px] font-bold text-rose-600 transition-colors hover:text-rose-700 disabled:opacity-40"
          >
            Xóa tất cả
          </button>
        </div>

        <div className="max-h-[52vh] min-h-[160px] space-y-2 overflow-y-auto bg-muted/15 p-4">
          {draft.length === 0 ? (
            <button
              type="button"
              onClick={addLevel}
              className="flex min-h-[132px] w-full flex-col items-center justify-center gap-2 rounded-xl border-2 border-dashed border-primary/20 bg-card text-center transition-colors hover:border-primary/35 hover:bg-primary/[0.025]"
            >
              <ArrowUpDown className="h-6 w-6 text-primary/35" />
              <span className="text-xs font-bold text-foreground/65">Thêm cột đầu tiên để sắp xếp</span>
            </button>
          ) : (
            draft.map((rule, index) => (
              <div
                key={`${rule.key}-${index}`}
                className="grid grid-cols-[74px_minmax(170px,1fr)_150px_84px] items-end gap-2 rounded-xl border border-border bg-card p-3 shadow-sm"
              >
                <div>
                  <span className="mb-1 block text-[9px] font-black uppercase tracking-widest text-muted-foreground">
                    Ưu tiên
                  </span>
                  <div className="flex h-9 items-center gap-2 rounded-lg bg-primary/7 px-2 text-xs font-black text-primary">
                    <span className="grid h-5 w-5 place-items-center rounded-md bg-primary text-[10px] text-primary-foreground">
                      {index + 1}
                    </span>
                    {index === 0 ? "Theo" : "Rồi"}
                  </div>
                </div>
                <label className="min-w-0">
                  <span className="mb-1 block text-[9px] font-black uppercase tracking-widest text-muted-foreground">
                    Cột
                  </span>
                  <Select value={rule.key} onValueChange={(key) => updateColumn(index, key)}>
                    <SelectTrigger className="h-9 min-w-0 text-xs font-bold">
                      <SelectValue />
                    </SelectTrigger>
                    <SelectContent className="z-[10070] max-h-72">
                      {sortableColumns.map((column) => (
                        <SelectItem key={column.key} value={column.key} className="text-xs">
                          {column.label || column.key}
                        </SelectItem>
                      ))}
                    </SelectContent>
                  </Select>
                </label>
                <label>
                  <span className="mb-1 block text-[9px] font-black uppercase tracking-widest text-muted-foreground">
                    Thứ tự
                  </span>
                  <Select
                    value={rule.direction}
                    onValueChange={(direction) => {
                      if (direction !== "asc" && direction !== "desc") return;
                      setDraft((current) =>
                        current.map((item, itemIndex) =>
                          itemIndex === index ? { ...item, direction } : item,
                        ),
                      );
                    }}
                  >
                    <SelectTrigger className="h-9 text-xs font-bold">
                      <SelectValue />
                    </SelectTrigger>
                    <SelectContent className="z-[10070]">
                      <SelectItem value="asc" className="text-xs">A → Z / Nhỏ → Lớn</SelectItem>
                      <SelectItem value="desc" className="text-xs">Z → A / Lớn → Nhỏ</SelectItem>
                    </SelectContent>
                  </Select>
                </label>
                <div className="flex h-9 items-center justify-end gap-1">
                  <button
                    type="button"
                    onClick={() => moveLevel(index, -1)}
                    disabled={index === 0}
                    className="grid h-8 w-7 place-items-center rounded-md text-muted-foreground hover:bg-primary/10 hover:text-primary disabled:opacity-25"
                    title="Tăng mức ưu tiên"
                  >
                    <ChevronUp className="h-4 w-4" />
                  </button>
                  <button
                    type="button"
                    onClick={() => moveLevel(index, 1)}
                    disabled={index === draft.length - 1}
                    className="grid h-8 w-7 place-items-center rounded-md text-muted-foreground hover:bg-primary/10 hover:text-primary disabled:opacity-25"
                    title="Giảm mức ưu tiên"
                  >
                    <ChevronDown className="h-4 w-4" />
                  </button>
                  <button
                    type="button"
                    onClick={() => setDraft((current) => current.filter((_, itemIndex) => itemIndex !== index))}
                    className="grid h-8 w-7 place-items-center rounded-md text-muted-foreground hover:bg-rose-50 hover:text-rose-600"
                    title="Xóa cấp"
                  >
                    <Trash2 className="h-3.5 w-3.5" />
                  </button>
                </div>
              </div>
            ))
          )}
        </div>

        <footer className="flex items-center justify-end gap-2 border-t border-border bg-card px-5 py-3">
          <button
            type="button"
            onClick={() => onOpenChange(false)}
            className="h-9 rounded-lg border border-border px-4 text-xs font-bold text-foreground/70 transition-colors hover:bg-muted"
          >
            Hủy
          </button>
          <button
            type="button"
            onClick={() => {
              onApply(normalizeSortRules(draft));
              onOpenChange(false);
            }}
            className="h-9 rounded-lg bg-primary px-5 text-xs font-black uppercase tracking-wide text-primary-foreground shadow-sm transition-transform active:scale-[0.98]"
          >
            Áp dụng
          </button>
        </footer>
      </section>
    </div>
  );
}
