import { getBusinessFromL07, resolveSummerBonusCenterL07 } from "./center-utils";
import { getVal, parseAnyDate, toVietnamDateString } from "./data-utils";

export interface TimesheetLinkInput {
  id: string;
  l07: string;
  url?: string;
  fileName?: string;
  sheetName?: string;
  aeCode?: string;
  date?: string;
  lastSyncedAt?: string;
  legacyRowIds?: string[];
  status?: string;
}

export interface TimesheetCoverageDateRange {
  from?: string;
  to?: string;
  preferredYear?: number;
}

export const CENTER_COVERAGE_BATCH_ROWS = 750;

const ROSTER_DATE_FIELDS = [
  "date", "ngay", "ngày", "tk_date", "session date", "sessiondate",
  "ngày học", "date of class", "scheduledate", "ngày làm việc",
  "thời gian", "kỳ", "ngày trực", "ngày tháng",
];

export function normalizeTimesheetL07(value: unknown): string {
  const text = String(value ?? "").trim().toUpperCase().replace(/\s+/g, " ");
  return text ? resolveSummerBonusCenterL07(text).l07 : "";
}

/** Check the complete table dataset before search, BU filters and pagination. */
function* scanMissingTimesheetCenters(
  inputs: readonly TimesheetLinkInput[],
  rows: readonly Record<string, unknown>[],
  mode: "l07" | "allocation" = "l07",
  source?: TimesheetCoverageDateRange & { rows: readonly Record<string, unknown>[] },
) {
  const normalizedValues = new Map<unknown, string>();
  const normalize = (value: unknown) => {
    if (!normalizedValues.has(value)) normalizedValues.set(value, normalizeTimesheetL07(value));
    return normalizedValues.get(value)!;
  };
  const configured = new Map<string, TimesheetLinkInput>();
  for (const input of inputs) {
    const l07 = normalize(input.l07);
    if (!l07 || (!input.url?.trim() && !input.fileName?.trim())) continue;
    // Pivot has one AHP allocation destination: Hai Phong. The individual
    // HP center links belong to the other Timesheet tables, not this Pivot.
    if (mode === "allocation" && getBusinessFromL07(l07) === "AHP" && l07 !== "HAI PHONG") continue;
    const existing = configured.get(l07);
    // One center counts once; prefer a directly refreshable link over a file name.
    if (!existing || (!existing.url?.trim() && input.url?.trim())) configured.set(l07, input);
  }

  // A successful import with no rows in this table does not need another fetch.
  const candidates = new Map([...configured].filter(([, input]) => input.status !== "success"));
  let batchRows = 0;
  for (const row of rows) {
    if (candidates.size === 0) break;
    if (++batchRows % CENTER_COVERAGE_BATCH_ROWS === 0) yield;
    if (row._isSubtotal || row._isTotalRow) continue;
    const l07 = mode === "allocation"
      ? row.chargeToCenterMkt || row.charge_to_center_mkt || row.l07 || row.center
      : row.l07 || row.L07 || row.center;
    const normalized = normalize(l07);
    candidates.delete(normalized);
  }

  const outsideTimeRange = new Set<string>();
  if ((source?.from || source?.to) && candidates.size > 0) {
    const eligibleOrUnknownDate = new Set<string>();
    const inputById = new Map<string, TimesheetLinkInput>();
    inputs.forEach((input) => {
      [input.id, ...(input.legacyRowIds || [])].forEach((id) => inputById.set(id, input));
    });
    const dateCache = new Map<unknown, string | null>();
    for (const row of source.rows) {
      if (eligibleOrUnknownDate.size === candidates.size) break;
      if (++batchRows % CENTER_COVERAGE_BATCH_ROWS === 0) yield;
      if (row._isSubtotal || row._isTotalRow) continue;
      const input = inputById.get(String(row._rowId || ""));
      const sourceL07 = normalize(input?.l07 || getVal(row, ["l07", "center", "maAE", "aeCode"]));
      const allocationL07 = normalize(row.chargeToCenterMkt || row.charge_to_center_mkt
        || getVal(row, ["charge to center mkt", "charge to center", "chargetocenter", "charge to center mkt name"]));
      const centers = [sourceL07, allocationL07].filter((l07) => candidates.has(l07));
      if (!centers.length) continue;
      const rawDate = getVal(row, ROSTER_DATE_FIELDS);
      if (!dateCache.has(rawDate)) {
        const date = parseAnyDate(rawDate, source.preferredYear);
        dateCache.set(rawDate, date ? toVietnamDateString(date) : null);
      }
      const date = dateCache.get(rawDate);
      const excluded = date && ((source.from && date < source.from) || (source.to && date > source.to));
      // An unreadable date does not prove that a center was removed by the
      // time filter. The calculation worker also retains rows with no date.
      const target = excluded ? outsideTimeRange : eligibleOrUnknownDate;
      centers.forEach((l07) => target.add(l07));
    }
    eligibleOrUnknownDate.forEach((l07) => outsideTimeRange.delete(l07));
  }

  const missing = [...candidates].filter(([l07]) => !outsideTimeRange.has(l07));
  return {
    expectedCount: configured.size,
    centers: missing.map(([l07]) => l07),
    inputs: missing.map(([, input]) => input),
  };
}

type CoverageArguments = Parameters<typeof scanMissingTimesheetCenters>;

export function getMissingTimesheetCenters(...args: CoverageArguments) {
  const scan = scanMissingTimesheetCenters(...args);
  let next = scan.next();
  while (!next.done) next = scan.next();
  return next.value;
}

export async function getMissingTimesheetCentersInBatches(
  signal: AbortSignal,
  ...args: CoverageArguments
) {
  const scan = scanMissingTimesheetCenters(...args);
  while (!signal.aborted) {
    const next = scan.next();
    if (next.done) return next.value;
    await new Promise<void>(resolve => setTimeout(resolve, 0));
  }
}
