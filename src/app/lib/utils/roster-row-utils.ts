import {
  getBusinessFromL07,
  getCenterInfoByAECode,
  getL07FromFileName,
  resolveMktRosterCenter,
  resolveNorthMktLocalL07,
} from "./center-utils";
import { generateUUID, getVal } from "./data-utils";

function parseDuration(value: unknown): number {
  if (typeof value === "number") return Number.isFinite(value) ? value : 0;
  if (value === null || value === undefined || value === "") return 0;

  const text = String(value).trim().replace(",", ".");
  if (text.includes(":")) {
    const [hours, minutes] = text.split(":");
    return (parseInt(hours, 10) || 0) + (parseInt(minutes, 10) || 0) / 60;
  }
  return parseFloat(text) || 0;
}

export function mapExcelRosterRow(
  row: Record<string, unknown>,
  fileName?: string,
  fileId?: string,
) {
  const rawCenter = String(
    getVal(row, [
      "cơ sở",
      "trung tâm",
      "chi nhánh",
      "center code",
      "office code",
      "center",
      "mã ae",
      "ae",
      "ae code",
    ]) || "",
  ).trim();
  const rawChargeToCenter = String(
    getVal(row, [
      "charge to center mkt",
      "charge to center",
      "chargetocenter",
      "charge to center mkt name",
    ]) || "",
  ).trim();

  const centerInfo = getCenterInfoByAECode(rawCenter);
  const fileL07 = getL07FromFileName(fileName || "");
  const groupL07 = getCenterInfoByAECode(String(getVal(row, ['l07']) || ''))?.l07;
  const isNorthSource = centerInfo?.l07 === 'MKT LOCAL NORTH' || groupL07 === 'MKT LOCAL NORTH' || fileL07 === 'MKT LOCAL NORTH';
  const regionalMktL07 = resolveNorthMktLocalL07(
    `${fileName || ""} ${rawCenter} ${rawChargeToCenter}`,
  );
  const isMktFile = /MKT|MARKETING/i.test(fileName || "");
  const classCharge = isNorthSource && (!rawCenter || centerInfo?.l07 === 'MKT LOCAL NORTH')
    ? String(getVal(row, ['class']) || '').trim() : '';
  const mktChargeSource = rawChargeToCenter || classCharge || rawCenter;
  const resolvedMktCharge = resolveMktRosterCenter(mktChargeSource);

  let l07 = centerInfo?.l07 || fileL07 || rawCenter || "UNKNOWN";
  let chargeToCenterMkt = resolveMktRosterCenter(rawChargeToCenter).chargeToCenterMkt;

  if (isMktFile || regionalMktL07 || isNorthSource) {
    l07 = regionalMktL07 || (groupL07 === 'MKT LOCAL NORTH' ? groupL07 : fileL07) || "MKT LOCAL NORTH";
    chargeToCenterMkt =
      resolvedMktCharge.chargeToCenterMkt || mktChargeSource;
  }

  const business = getBusinessFromL07(
    (isMktFile || regionalMktL07 || isNorthSource) && chargeToCenterMkt
      ? resolvedMktCharge.l07 || chargeToCenterMkt
      : l07,
  );
  const employeeId = String(
    getVal(row, [
      "id number",
      "id",
      "teacher id",
      "emp id",
      "mã nv",
      "manv",
      "staff id",
      "code",
    ]) || "",
  ).trim();
  const fullName = String(
    getVal(row, [
      "full name",
      "name",
      "teacher name",
      "tên",
      "họ và tên",
      "họ tên",
    ]) || "",
  ).trim();
  const dateValue = getVal(row, [
    "date",
    "ngay",
    "ngày",
    "tk_date",
    "session date",
    "sessiondate",
    "ngày học",
    "scheduledate",
    "ngày làm việc",
    "ngày tháng",
  ]);
  const date = dateValue === null || dateValue === undefined
    ? ""
    : String(dateValue).trim();
  const type = String(
    getVal(row, [
      "type",
      "type code",
      "type_code",
      "typecode",
      "task type",
      "task",
      "loại",
      "loại hoạt động",
      "event type",
      "activity",
      "category",
      "task type name",
    ]) || "",
  ).trim();
  const classCode = String(
    getVal(row, [
      "class",
      "class code",
      "class_code",
      "classcode",
      "lớp",
      "class name",
      "mã lớp",
      "tên lớp",
      "mã lớp học",
    ]) || "",
  ).trim();
  const from = String(
    getVal(row, ["from", "start", "start time", "từ", "giờ bắt đầu"]) || "",
  ).trim();
  const to = String(
    getVal(row, ["to", "end", "end time", "đến", "giờ kết thúc"]) || "",
  ).trim();
  const duration = parseDuration(
    getVal(row, [
      "duration",
      "quy ra số giờ làm",
      "total",
      "actual hours",
      "working hours",
      "giờ làm",
      "số giờ",
      "hours",
      "tk_duration",
      "total hours",
      "tổng giờ",
      "time",
      "thời lượng",
    ]),
  );
  const notes = String(
    getVal(row, ["notes", "note", "ghi chú", "ghi chu", "remarks"]) || "",
  )
    .trim()
    .replace(/^["']|["']$/g, "");

  return {
    ...row,
    center: rawCenter,
    l07,
    business,
    ma_nv: employeeId,
    full_name: fullName,
    ngay: date,
    type,
    class: classCode,
    gio_vao: from,
    gio_ra: to,
    chargeToCenterMkt,
    duration,
    notes,
    employeeId,
    fullName,
    maAE: rawCenter,
    date,
    taskType: type,
    classCode,
    from,
    to,
    _sourceFile: fileName || row._sourceFile || "",
    _rowId: fileId || row._rowId || "",
    _uuid: row._uuid || generateUUID(),
  };
}
