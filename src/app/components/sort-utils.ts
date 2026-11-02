export type SortDirection = "asc" | "desc";

// localeCompare(locale, options) otherwise constructs a collator for every
// comparison: hundreds of thousands of allocations when a period is sorted.
const tableTextCollator = new Intl.Collator('vi', { numeric: true, sensitivity: 'base' });
const filterValueCollator = new Intl.Collator(undefined, { numeric: true });
export const compareTableText = (left: unknown, right: unknown) => tableTextCollator.compare(String(left), String(right));
export const compareFilterValues = (left: unknown, right: unknown) => filterValueCollator.compare(String(left), String(right));

export interface SortRule {
  key: string;
  direction: SortDirection;
}

/** Accepts both the legacy single-rule value and the current ordered list. */
export function normalizeSortRules(value: unknown): SortRule[] {
  const candidates = Array.isArray(value) ? value : value ? [value] : [];
  const seen = new Set<string>();

  return candidates.flatMap((candidate) => {
    if (!candidate || typeof candidate !== "object") return [];
    const key = String((candidate as { key?: unknown }).key || "").trim();
    const direction = (candidate as { direction?: unknown }).direction;
    if (!key || seen.has(key) || (direction !== "asc" && direction !== "desc")) {
      return [];
    }
    seen.add(key);
    return [{ key, direction } as SortRule];
  });
}
