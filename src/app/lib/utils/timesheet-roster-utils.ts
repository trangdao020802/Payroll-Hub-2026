import { getCenterInfoByL07, mapL07 } from "./center-utils";

type RosterRow = Record<string, unknown>;
// Imported rows are immutable between syncs. Cache their normalized identities
// so adding another center does not reparse every previously imported row.
const businessKeyCache = new WeakMap<RosterRow, string>();
const centerCache = new WeakMap<RosterRow, string>();

const INTERNAL_ROW_KEYS = new Set([
  "_uuid",
  "_rowId",
  "_sourceFile",
  "_searchStr",
  "overlap_check",
  "overlap_group",
  "overlap_position",
  "overlap_total",
  "overlap_with_ids",
]);

function normalizeText(value: unknown): string {
  return String(value ?? "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[Đđ]/g, "D")
    .trim()
    .toUpperCase()
    .replace(/\s+/g, " ");
}

function firstValue(row: RosterRow, keys: string[]): unknown {
  for (const key of keys) {
    const value = row[key];
    if (value !== undefined && value !== null && String(value).trim() !== "") {
      return value;
    }
  }

  const entries = Object.entries(row);
  for (const key of keys) {
    const normalizedKey = normalizeText(key);
    const entry = entries.find(([candidate, value]) =>
      normalizeText(candidate) === normalizedKey &&
      value !== undefined &&
      value !== null &&
      String(value).trim() !== "",
    );
    if (entry) return entry[1];
  }
  return "";
}

function normalizeDate(value: unknown): string {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return value.toISOString().slice(0, 10).replace(/-/g, "");
  }
  const source = normalizeText(value).replace(/^[A-Z]{3,4}\s+/, "");
  const ymd = source.match(/^(\d{4})[/-](\d{1,2})[/-](\d{1,2})/);
  if (ymd) {
    return `${ymd[1]}${ymd[2].padStart(2, "0")}${ymd[3].padStart(2, "0")}`;
  }
  const dmy = source.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{4})/);
  if (dmy) {
    return `${dmy[3]}${dmy[2].padStart(2, "0")}${dmy[1].padStart(2, "0")}`;
  }
  return source.replace(/[^A-Z0-9]/g, "");
}

function normalizeTime(value: unknown): string {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return `${String(value.getHours()).padStart(2, "0")}:${String(value.getMinutes()).padStart(2, "0")}`;
  }
  const source = String(value ?? "").trim();
  const match = source.match(/^(\d{1,2}):(\d{1,2})/);
  if (match) {
    return `${match[1].padStart(2, "0")}:${match[2].padStart(2, "0")}`;
  }
  return normalizeText(source).replace(/\s+/g, "");
}

function stableSerialize(row: RosterRow): string {
  return Object.entries(row)
    .filter(([key]) => !INTERNAL_ROW_KEYS.has(key))
    .sort(([left], [right]) => left.localeCompare(right))
    .map(([key, value]) => `${normalizeText(key)}=${normalizeText(value)}`)
    .join("|");
}

export function getCanonicalTimesheetCenter(row: RosterRow): string {
  const cached = centerCache.get(row);
  if (cached !== undefined) return cached;
  const values = [
    row.l07,
    row.center,
    row.maAE,
    row.aeCode,
    row["L07"],
    row["Mã AE"],
    row["Center"],
    row.charge_to_center_mkt,
    row.chargeToCenterMkt,
  ];

  for (const value of values) {
    const source = String(value ?? "").trim();
    if (!source) continue;
    const mapped = mapL07(source);
    if (getCenterInfoByL07(mapped) || normalizeText(mapped).includes("MKT LOCAL")) {
      const center = normalizeText(mapped);
      centerCache.set(row, center);
      return center;
    }
  }

  const fallback = values.find((value) => String(value ?? "").trim() !== "");
  const center = normalizeText(fallback);
  centerCache.set(row, center);
  return center;
}

function getCanonicalChargeCenter(row: RosterRow): string {
  const value = firstValue(row, [
    "charge_to_center_mkt",
    "chargeToCenterMkt",
    "Charge to Center MKT",
    "Charge to Center",
  ]);
  const source = String(value ?? "").trim();
  return source ? normalizeText(mapL07(source)) : "";
}

/**
 * Logical identity of one Timesheet session. Internal import IDs and source
 * filenames are deliberately excluded so re-syncing the same Google Sheet is
 * idempotent even when a legacy version generated new UUIDs on every import.
 */
export function getTimesheetRosterBusinessKey(row: RosterRow): string {
  const cached = businessKeyCache.get(row);
  if (cached !== undefined) return cached;
  const employee = normalizeText(firstValue(row, [
    "ma_nv", "employeeId", "ID Number", "Mã NV", "Teacher ID", "Emp ID",
  ]));
  const date = normalizeDate(firstValue(row, [
    "ngay", "date", "Date", "Ngày", "TK_Date", "Session Date",
  ]));
  const from = normalizeTime(firstValue(row, [
    "gio_vao", "from", "From", "Start", "Start Time",
  ]));
  const to = normalizeTime(firstValue(row, [
    "gio_ra", "to", "To", "End", "End Time",
  ]));
  const type = normalizeText(firstValue(row, ["type", "taskType", "Type", "Task Type"]));
  const classCode = normalizeText(firstValue(row, ["class", "classCode", "Class", "Lớp"]));
  const duration = Number(firstValue(row, ["duration", "workingHours", "Working Hours"])) || 0;
  const center = getCanonicalTimesheetCenter(row);
  const chargeCenter = getCanonicalChargeCenter(row);

  const core = [
    center,
    chargeCenter,
    employee,
    date,
    from,
    to,
    type,
    classCode,
    String(duration),
  ].join("|");
  const populatedCoreFields = [center, employee, date, from, to, type, classCode]
    .filter(Boolean).length;
  const key = populatedCoreFields >= 3 ? core : `${core}|${stableSerialize(row)}`;
  businessKeyCache.set(row, key);
  return key;
}

function hashString(value: string): string {
  let hash = 0x811c9dc5;
  for (let index = 0; index < value.length; index++) {
    hash ^= value.charCodeAt(index);
    hash = Math.imul(hash, 0x01000193);
  }
  return (hash >>> 0).toString(36);
}

export function createStableTimesheetRowId(
  row: RosterRow,
  sourceId: unknown,
  occurrence = 0,
): string {
  const sourceHash = hashString(normalizeText(sourceId) || "TIMESHEET");
  const rowHash = hashString(getTimesheetRosterBusinessKey(row));
  return `ts_${sourceHash}_${rowHash}_${occurrence}`;
}

export function dedupeTimesheetRosterRows<T extends RosterRow>(rows: T[]): T[] {
  const seen = new Set<string>();
  const uniqueReversed: T[] = [];

  // Walk backwards so the latest synchronized copy wins, while calculating
  // the expensive business key only once per row.
  for (let index = rows.length - 1; index >= 0; index -= 1) {
    const row = rows[index];
    const key = getTimesheetRosterBusinessKey(row);
    if (seen.has(key)) continue;
    seen.add(key);
    uniqueReversed.push(row);
  }

  uniqueReversed.reverse();
  return uniqueReversed;
}

export async function dedupeTimesheetRosterRowsInChunks<T extends RosterRow>(
  rows: T[],
  chunkSize = 750,
): Promise<T[]> {
  const seen = new Set<string>();
  const uniqueReversed: T[] = [];
  let processedInChunk = 0;

  for (let index = rows.length - 1; index >= 0; index -= 1) {
    const row = rows[index];
    const key = getTimesheetRosterBusinessKey(row);
    if (!seen.has(key)) {
      seen.add(key);
      uniqueReversed.push(row);
    }

    processedInChunk += 1;
    if (processedInChunk >= chunkSize) {
      processedInChunk = 0;
      await new Promise<void>((resolve) => window.setTimeout(resolve, 0));
    }
  }

  uniqueReversed.reverse();
  return uniqueReversed;
}

export function replaceTimesheetRosterRows<T extends RosterRow>(
  existingRows: T[],
  incomingRows: T[],
  options: {
    sourceRowIds: Iterable<unknown>;
    targetL07?: unknown;
    targetAeCode?: unknown;
  },
): T[] {
  const sourceRowIds = new Set(
    Array.from(options.sourceRowIds, (value) => String(value ?? "")).filter(Boolean),
  );
  const targetCenters = new Set(
    [options.targetL07, options.targetAeCode]
      .map((value) => mapL07(String(value ?? "").trim()))
      .map(normalizeText)
      .filter(Boolean),
  );

  // Incoming rows are parsed afresh on each link refresh. Deduplicate that
  // small batch once, then retain unrelated centers without rehashing the
  // entire roster on every file in a multi-link update.
  const uniqueIncoming = dedupeTimesheetRosterRows(incomingRows);
  const incomingKeys = new Set(uniqueIncoming.map(getTimesheetRosterBusinessKey));
  const retainedRows = existingRows.filter((row) => {
    if (sourceRowIds.has(String(row._rowId ?? ""))) return false;
    const rowCenter = getCanonicalTimesheetCenter(row);
    if (rowCenter && targetCenters.has(rowCenter)) return false;
    return !incomingKeys.has(getTimesheetRosterBusinessKey(row));
  });

  return [...retainedRows, ...uniqueIncoming];
}

export async function prepareTimesheetRosterReplacement<T extends RosterRow>(
  existingRows: T[],
  incomingRows: T[],
  options: Parameters<typeof replaceTimesheetRosterRows>[2],
): Promise<{ original: T[]; rows: T[] }> {
  // Normalizing dates/headers/centers for an entire roster inside a React
  // updater blocks input and rendering. Warm the immutable row caches in small
  // turns before committing; the final replacement only visits cached keys.
  let newlyPrepared = 0;
  for (const rows of [incomingRows, existingRows]) {
    for (const row of rows) {
      if (businessKeyCache.has(row)) continue;
      getTimesheetRosterBusinessKey(row);
      if (++newlyPrepared % 250 === 0) {
        await new Promise<void>(resolve => setTimeout(resolve, 0));
      }
    }
  }
  return { original: existingRows, rows: replaceTimesheetRosterRows(existingRows, incomingRows, options) };
}
