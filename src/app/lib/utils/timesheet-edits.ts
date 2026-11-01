import { getBusinessFromL07, getL07FromFileName, mapL07, resolveMktRosterCenter } from "./center-utils";
import { getVal, parseMoneyToNumber } from "./data-utils";
import { createStableTimesheetRowId } from "./timesheet-roster-utils";

type Row = Record<string, unknown>;

/** _rowId identifies an imported file, not an individual editable row. */
export function getRosterSourceKey(row: Row): string {
  return String(row._sourceKey || row._uuid || row._recordId || row.id
    || createStableTimesheetRowId(row, row._rowId || row._sourceFile || "roster"));
}

export function getTimesheetChargeToCenter(row: Row, configuredL07 = ""): string {
  let value = row.chargeToCenterMkt ?? row.charge_to_center_mkt
    ?? getVal(row, ["charge to center mkt", "charge to center", "chargetocenter", "charge to center mkt name"]);
  if (row._chargeToCenterMktOverride !== undefined) value = row._chargeToCenterMktOverride;
  const explicitBlank = row.chargeToCenterMkt === "" && row.charge_to_center_mkt === "";
  if (!String(value ?? "").trim() && row._chargeToCenterMktOverride === undefined && !explicitBlank) {
    const scope = String(configuredL07 || row._l07Override || row.l07 || row.L07 ||
      getVal(row, ["bank", "bank name", "center"]) || "").trim();
    const source = String(row._sourceFile || "");
    const sourceL07 = /SOUTH/i.test(source) ? "" : getL07FromFileName(source);
    const l07 = mapL07(scope && scope !== "UNKNOWN" ? scope : sourceL07);
    if (!/SOUTH/i.test(scope) && l07 === "MKT LOCAL NORTH") {
      const center = String(getVal(row, ['center']) || '').trim();
      value = center && mapL07(center) !== 'MKT LOCAL NORTH' ? center : getVal(row, ["class"]);
    }
  }
  return resolveMktRosterCenter(String(value ?? "")).chargeToCenterMkt;
}

/** Also normalize saved legacy rows before the calculation worker finishes. */
export function normalizeTimesheetAllocation(row: Row): Row {
  const chargeToCenterMkt = getTimesheetChargeToCenter(row);
  if (row.chargeToCenterMkt === chargeToCenterMkt) return row;
  return { ...row, _sourceKey: getRosterSourceKey(row), chargeToCenterMkt };
}

const FIELD_ALIASES = [
  ["chargeToCenterMkt", "charge_to_center_mkt", "charge to center mkt", "charge to center", "chargetocenter", "charge to center mkt name"],
  ["ngay", "date", "ngày", "tk_date", "session date", "sessiondate", "ngày làm việc", "ngày tháng"],
  ["ma_nv", "employeeId", "id number", "teacher id", "emp id", "mã nv", "manv", "staff id"],
  ["full_name", "fullName", "full name", "name", "teacher name", "tên", "họ và tên", "họ tên"],
  ["class", "classCode", "class code", "class_code", "lớp", "class name", "mã lớp"],
  ["gio_vao", "from", "start", "start time", "từ"],
  ["gio_ra", "to", "end", "end time", "đến"],
  ["type", "taskType", "task type", "type code", "task", "loại", "activity"],
  ["duration", "workingHours", "working hours", "hours", "quy ra số giờ làm", "actual hours"],
  ["notes", "note", "ghi chú", "ghi chu", "remarks"],
];
const normalizeKey = (key: string) => key.trim().toLowerCase().replace(/[\s_]+/g, "");

export function applyTimesheetRosterCellEdit(row: Row, field: string, value: unknown): Row {
  const next: Row = { ...row, _sourceKey: getRosterSourceKey(row) };
  const aliases = FIELD_ALIASES.find((group) => group.some((key) => normalizeKey(key) === normalizeKey(field)));
  let finalValue = value;
  if (aliases === FIELD_ALIASES[0]) finalValue = resolveMktRosterCenter(String(value ?? "")).chargeToCenterMkt;
  if (aliases) {
    const keys = new Set(aliases.map(normalizeKey));
    for (const key of Object.keys(row)) {
      if (keys.has(normalizeKey(key))) next[key] = finalValue;
    }
    // Keep the canonical UI/worker aliases aligned, including intentionally blank edits.
    next[aliases[0]] = finalValue;
    next[aliases[1]] = finalValue;
  }
  next[field] = finalValue;
  if (aliases === FIELD_ALIASES[0]) {
    next.chargeToCenterMkt = finalValue;
    next.charge_to_center_mkt = finalValue;
    next._chargeToCenterMktOverride = finalValue;
    next.business = getBusinessFromL07(String(finalValue));
  }
  if (field === "l07" || field === "center") {
    const l07 = mapL07(String(value ?? ""));
    Object.assign(next, { l07, center: l07, _l07Override: l07 });
  }
  if (field === "business") next._businessOverride = value;
  if (aliases === FIELD_ALIASES[8]) next._durationOverride = value;
  if (aliases === FIELD_ALIASES[5] || aliases === FIELD_ALIASES[6]) delete next._durationOverride;
  return next;
}

export function applyTimesheetPivotEdit(rows: Row[], sourceKeys: ReadonlySet<string>, field: string, value: unknown): Row[] {
  let amountApplied = false;
  const amount = Math.max(0, parseMoneyToNumber(value));
  return rows.map((row) => {
    if (!sourceKeys.has(getRosterSourceKey(row))) return row;
    if (field === "chargeToCenterMkt") return applyTimesheetRosterCellEdit(row, field, value);
    const next = { ...row, _sourceKey: getRosterSourceKey(row), _mktPivotValueOverride: amountApplied ? 0 : amount };
    amountApplied = true;
    return next;
  });
}
