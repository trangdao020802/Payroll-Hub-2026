/* eslint-disable */
import {
  parseAnyDate,
  toVietnamDateString,
  formatTime12Hour,
  normalizeId,
  getVal,
  parseTimeStrToHours,
  generateUUID,
} from "../lib/utils/data-utils";
import {
  getBusinessFromL07,
  getCenterInfoByAECode,
  getCenterInfoByL07,
  mapL07,
} from "../lib/utils/center-utils";
import { resolveL07Logic } from "../lib/utils/l07-resolver";
import { dedupeTimesheetRosterRows } from "../lib/utils/timesheet-roster-utils";
import { getRosterSourceKey, getTimesheetChargeToCenter } from "../lib/utils/timesheet-edits";
import {
  DEFAULT_SALARY_SCALES,
  ACADEMIC_FIELDS,
} from "../constants/timesheet-logic";

export function calculateTimesheet(params: any) {
  try {
    const {
      rosterData,
      salaryScaleData,
      staffData,
      cacheData,
      fromDateStr,
      toDateStr,
      appData,
      preferredYear,
      checkTAsMap,
      classSizeMap,
      TASK_COLUMNS,
      customSalaryRates,
      customTypeRates,
    } = params;

    if (!rosterData || rosterData.length === 0) {
      return {
        processedRosterData: [],
        employeeSummary: [],
        centerSummary: [],
        isCalculating: false,
      };
    }

  // Lookups reconstructed in worker
  const staffLookup = new Map();
  (staffData || []).forEach((s: any) => {
    const sid = normalizeId(getVal(s, ["id", "id number"]));
    const sn = String(getVal(s, ["full name", "name"])).trim().toLowerCase();
    if (sid) staffLookup.set(sid, s);
    if (sn) staffLookup.set(sn, s);
  });

  const salaryScaleLookup = new Map();
  (salaryScaleData || []).forEach((s: any) => {
    const sid = normalizeId(getVal(s, ["id", "id number"]));
    const sn = String(getVal(s, ["full name", "name"])).trim().toLowerCase();
    if (sid) salaryScaleLookup.set(sid, s);
    if (sn) salaryScaleLookup.set(sn, s);
  });

  const effectiveSalaryScales: Record<string, any> = {
    ...DEFAULT_SALARY_SCALES,
    ...(appData?.customSalaryRates || {}),
    ...(customSalaryRates || {}),
  };

  const inputListLookup = new Map();
  const discardedLegacyRowIds = new Set<string>();
  (appData?.Timesheet_InputList || []).forEach((ir: any) => {
    inputListLookup.set(ir.id, ir);
    const aliases = Array.isArray(ir.legacyRowIds)
      ? ir.legacyRowIds.filter(Boolean)
      : [];
    // If the old folder-import bug generated several rows for one center,
    // only the newest imported file is authoritative.
    aliases.slice(0, -1).forEach((id: string) => discardedLegacyRowIds.add(id));
    const currentAlias = aliases.at(-1);
    if (currentAlias) inputListLookup.set(currentAlias, ir);
  });

  const cacheCodes = new Set(
    (cacheData || []).map((c: any) => String(getVal(c, ["code", "mã lớp"])).toLowerCase().trim())
  );

  const academicFieldSet = new Set(ACADEMIC_FIELDS);
  const dateCache = new Map<string, Date | null>();
  const dateStringCache = new Map<string, string>();
  const centerInfoCache = new Map<string, any>();

  const getSalaryRate = (id: string, name: string) => {
    const nid = normalizeId(id);
    const row = salaryScaleLookup.get(nid) || salaryScaleLookup.get(String(name || "").toLowerCase());
    const sCode = String(getVal(row || {}, ["s code", "scale", "salary scale"]) || "S1").trim().toUpperCase();
    const def = effectiveSalaryScales[sCode] || effectiveSalaryScales["S1"] || DEFAULT_SALARY_SCALES["S1"];
    
    // Academic Price (Đơn giá Giảng dạy): Tùy thuộc vào cột [Salary Scale] của nhân sự (S1-S7)
    // Các đơn giá cố định khác áp dụng chung
    const ac = Number(def.ac) || 33000;
    const ad = Number(def.ad) || 20000;
    const su = Number(def.summer ?? def.su) || 29474;
    const ou = Number(def.outing ?? def.ou) || 26316;
    const si = Number(def.summerInstructors ?? def.si) || 150000;

    return { ac, ad, su, ou, si, sCode };
  };

  const normalizeStr = (str: string) => String(str).replace(/\s+/g, "").toUpperCase();

  const details: any[] = [];
  const empGroup: Record<string, any> = {};
  const cenGroup: Record<string, any> = {};
  const normCenterCache = new Map();

  let checkSkipped = 0;
  let dateSkipped = 0;
  let empSkipped = 0;

  // Final safety gate: legacy sessions may already contain copies created by
  // repeated Google Sheet syncs. Exact logical sessions must be collapsed
  // before overlap detection and every downstream payroll aggregation.
  const uniqueRosterData = dedupeTimesheetRosterRows(rosterData);
  uniqueRosterData.forEach((t: any) => {
    if (discardedLegacyRowIds.has(String(t._rowId || ""))) return;
    const configuredInput = inputListLookup.get(t._rowId) as any;
    const configuredL07 = String(configuredInput?.l07 || "").trim();
    const configuredAeCode = String(configuredInput?.aeCode || "").trim();
    const rawDateVal = getVal(t, [
      "date", "ngay", "ngày", "tk_date", "session date", "sessiondate",
      "ngày học", "date of class", "scheduledate", "ngày làm việc",
      "thời gian", "kỳ", "ngày trực", "ngày tháng",
    ]);
    const dateKey = `${String(rawDateVal)}_${preferredYear}`;
    let rawDate = dateCache.get(dateKey);
    if (rawDate === undefined) {
      rawDate = parseAnyDate(rawDateVal, preferredYear);
      dateCache.set(dateKey, rawDate);
    }
    
    let rawDateStr = "";
    if (rawDate) {
      rawDateStr = dateStringCache.get(dateKey) || "";
      if (!rawDateStr) {
        rawDateStr = toVietnamDateString(rawDate);
        dateStringCache.set(dateKey, rawDateStr);
      }

      if (fromDateStr && rawDateStr < fromDateStr) {
        return;
      }
      if (toDateStr && rawDateStr > toDateStr) {
        return;
      }
    } else {
      rawDateStr = fromDateStr || "2026-01-01";
    }

    const rawEid = String(getVal(t, ["id", "id number", "teacher id", "emp id", "mã nv", "manv", "id nv", "mã nhân viên", "staff id", "staff code", "emp code", "employee code", "mã nhân sự"]) || "").trim();
    const rawName = String(getVal(t, ["full name", "name", "teacher name", "tên", "họ và tên", "họ tên", "nhân viên", "tên nhân viên", "giáo viên", "staff name"]) || "").trim();
    const kId = rawEid.toUpperCase();

    if (["ATLS", "ECP", "KDG", "PRI", "TOTAL", "TỔNG", "CLASS", "IELTS", "LỚP"].some((kw) => kId.includes(kw)) || (kId.includes(".") && !rawName)) {
      return;
    }

    let empId = normalizeId(rawEid);
    if (!empId && !rawName) return;

    let effName = rawName;
    if (!empId || !effName) {
      const sMatch = staffLookup.get(empId) || staffLookup.get(String(rawName || "").toLowerCase());
      if (sMatch) {
        if (!empId) empId = normalizeId(getVal(sMatch, ["id", "id number"]));
        if (!effName) effName = String(getVal(sMatch, ["full name", "name"]) || "");
      }
      if (!empId) empId = rawName;
      if (!effName) effName = empId;
    }

    let rawType = String(getVal(t, ["type", "code", "task code", "activity code", "type code", "type_code", "typecode", "task type", "task", "loại", "loại hoạt động", "event type", "activity", "category", "task type name", "taskType"]) || "").trim();
    let rCen = String(t.center || "").trim();
    const rawChargeToCenter = getTimesheetChargeToCenter(t, configuredL07);
    const rawAeCode = String(t.maAE || "").trim();

    const sourceFile = t._sourceFile || "";
    const fileUpper = String(sourceFile).toUpperCase();
    const bankVal = String(getVal(t, ["bank", "bank name", "bank_name", "ngân hàng", "ngan hang"]) || t.bank || "").trim().toUpperCase();
    const isMktNorthBank = bankVal === "MKT LOCAL NORTH" || bankVal.startsWith("MKT LOCAL NORTH") || bankVal.includes("MKT NORTH") || fileUpper.includes("MKT_LOCAL_NORTH") || fileUpper.includes("MKT LOCAL NORTH") || configuredL07.toUpperCase() === "MKT LOCAL NORTH";

    if (isMktNorthBank) {
      if (rawChargeToCenter) {
        rCen = rawChargeToCenter;
      } else if (t.center) {
        rCen = String(t.center).trim();
      }
    }

    let rawClassCode = String(getVal(t, ["class code", "class", "class_code", "classcode", "lớp", "class name", "mã lớp", "tên lớp", "mã lớp học", "classCode"]) || "");

    const resolvedAuth = resolveL07Logic({ rawCenter: rCen, rawChargeToCenter, sourceFile: t._sourceFile || appData?.Timesheet_RosterFileName || "", rawType, rawClassCode, rawAeCode, empId, staffLookup, normCenterCache }, TASK_COLUMNS);
    let { l07, aeCode, taskField, correctedType, correctedClass } = resolvedAuth;
    const { chargeToCenterMkt, isMktLocal } = resolvedAuth;

    // A file imported into a configured Timesheet row inherits that row's
    // center. File contents may omit Center or contain display text, but must
    // never replace the authoritative L07 selected in the input table.
    if (configuredL07) {
      l07 = mapL07(configuredL07) || configuredL07;
      const configuredCenterInfo =
        getCenterInfoByL07(l07) || getCenterInfoByAECode(configuredAeCode);
      aeCode = configuredAeCode || configuredCenterInfo?.aeCode || aeCode;
    }
    if (t._l07Override !== undefined) l07 = String(t._l07Override);

    // Use corrected values directly as they default to raw inside resolveL07Logic
    const effectiveType = correctedType;
    const effectiveClass = correctedClass;

    const rawLoai = String(getVal(t, ["loai", "loại", "loai_nv"]) || "").trim().toUpperCase();

    const [yearStr, monthStr, dayStr] = rawDateStr.split("-");
    const dateStr = `${dayStr}/${monthStr}/${yearStr}`;

    const startVal = getVal(t, ["start", "from", "start time", "từ"]);
    const endVal = getVal(t, ["end", "to", "end time", "đến"]);
    const fromStr = formatTime12Hour(startVal);
    const toStr = formatTime12Hour(endVal);

    const sH = parseTimeStrToHours(startVal);
    const eH = parseTimeStrToHours(endVal);
    let durationHours = 0;
    if (startVal !== undefined && startVal !== "" && endVal !== undefined && endVal !== "") {
      durationHours = eH >= sH ? eH - sH : eH + 24 - sH;
    } else {
      const fallbackStr = String(getVal(t, ["duration", "quy ra số giờ làm", "total", "actual hours", "working hours", "giờ làm", "số giờ", "hours", "tk_duration", "total hours", "tổng giờ", "time"]) || "0").trim();
      if (fallbackStr.includes(":")) {
        const p = fallbackStr.split(":");
        durationHours = (parseInt(p[0]) || 0) + (parseInt(p[1]) || 0) / 60;
      } else {
        durationHours = parseFloat(fallbackStr.replace(",", "."));
      }
    }
    const durationOverride = Number(t._durationOverride);
    if (
      t._durationOverride !== undefined &&
      t._durationOverride !== null &&
      t._durationOverride !== "" &&
      Number.isFinite(durationOverride) &&
      durationOverride >= 0
    ) {
      durationHours = durationOverride;
    }

    let classSize = 0;
    const classSizeVal = getVal(t, ["class size", "sĩ số", "sỹ số", "no of students", "number of student", "number of students", "students", "số hv", "số học viên", "sĩ số lớp", "total students", "số lượng học viên", "sĩ số thực tế", "sỹ số thực tế", "actual size", "size", "số lượng", "sĩ số cơ sở"]);
    if (classSizeVal) classSize = parseInt(String(classSizeVal), 10) || 0;

    if (classSize === 0) {
      const normCls = normalizeStr(rawClassCode);
      const normCenter = normalizeStr(mapL07(rCen) || aeCode);
      if (normCls && normCenter && dateStr) {
        const key = `${normCenter}_${normCls}_${dateStr}`;
        if (checkTAsMap[key]) classSize = checkTAsMap[key];
        if (classSize === 0 && classSizeMap[`${normCenter}_${normCls}`]) classSize = classSizeMap[`${normCenter}_${normCls}`];
      }
    }

    let actHours = durationHours;
    if (rawType.toLowerCase() === "tutorial" || rawType.toLowerCase().includes("tutoring")) {
      const clsLower = String(getVal(t, ["class code", "class", "lớp", "class name", "mã lớp", "tên lớp", "classcode"])).toLowerCase().trim();
      const isCached = cacheCodes.has(clsLower);
      const hasPT = String(getVal(t, ["pt name", "gvpt"])).trim() !== "";
      if (classSize > 0) {
        if (classSize === 1) actHours = 0.5;
        else if (classSize <= 4) actHours = 1;
        else if (classSize <= 8) actHours = 1.5;
        else actHours = 2;
      } else if (isCached || hasPT) { actHours = 1; } else { actHours = 1; }
    } else if (rawType.toLowerCase().includes("club")) {
      if (classSize > 0 && classSize <= 10) actHours = 1;
      else if (classSize > 10) actHours = 1.5;
      else actHours = 1.5;
    } else if (rawType.toLowerCase().includes("demo")) {
      if (classSize > 0) {
        if (classSize <= 5) actHours = Math.round((durationHours + 0.25) * 100) / 100;
        else actHours = Math.round((durationHours + 0.5) * 100) / 100;
      } else { actHours = Math.round((durationHours + 0.5) * 100) / 100; }
    }

    if (rawType.toLowerCase().includes("admin") && !actHours) actHours = 1;

    const rates = getSalaryRate(empId, effName);
    let money = 0;
    let activeRate = rates.ad;
    const customTypes = { ...(appData?.customTypeRates || {}), ...(customTypeRates || {}) };
    if (rawLoai === "KL") {
      money = 0;
      activeRate = 0;
    } else if (customTypes[taskField] !== undefined) {
      activeRate = Number(customTypes[taskField]) || 0;
      money = actHours * activeRate;
    } else if (taskField === "summer" || taskField === "discoveryCamp") {
      money = actHours * rates.su; activeRate = rates.su;
    } else if (taskField === "outing") {
      money = actHours * rates.ou; activeRate = rates.ou;
    } else if (taskField === "summerInstructors") {
      money = actHours * rates.si; activeRate = rates.si;
    } else if (academicFieldSet.has(taskField)) {
      money = actHours * rates.ac; activeRate = rates.ac;
    } else {
      money = actHours * rates.ad; activeRate = rates.ad;
    }

    const rawBUCol = String(getVal(t, ["khối", "business", "bus", "bộ phận", "bu", "khối/bu"]) || "").trim().toUpperCase();
    let centerBusiness = rawBUCol;
    if (!centerBusiness) centerBusiness = configuredInput?.bus || "";
    if (!centerBusiness) {
      let centerInfoForBus = centerInfoCache.get(l07);
      if (centerInfoForBus === undefined) { centerInfoForBus = getCenterInfoByL07(l07) || getCenterInfoByAECode(l07); centerInfoCache.set(l07, centerInfoForBus); }
      centerBusiness = (centerInfoForBus as any)?.bus || "";
    }
    if (centerBusiness === "AHN_HP") centerBusiness = "AHP";
    const allocationL07 = String(chargeToCenterMkt || l07 || "").trim();
    const canonicalAllocationL07 = mapL07(allocationL07) || allocationL07;
    const canonicalAllocationInfo = getCenterInfoByL07(canonicalAllocationL07);
    if (canonicalAllocationInfo?.bus) {
      // The allocation center is authoritative for MKT rows. In particular,
      // TN0001.LNQ must always resolve to ATN and TH0001.TPU to ATH even when
      // an imported/configured row still carries the legacy AHN business.
      centerBusiness = canonicalAllocationInfo.bus;
    } else if (isMktLocal && allocationL07) {
      centerBusiness = getBusinessFromL07(allocationL07);
    } else if (l07 && String(l07).toUpperCase().trim() === "MKT LOCAL NORTH") {
      centerBusiness = "AHN";
    }

    if (t._businessOverride !== undefined) centerBusiness = String(t._businessOverride);
    const rowSourceKey = getRosterSourceKey(t);

    const detailRow = {
      id: rowSourceKey,

      // Keep the source identity across the worker boundary so edits in Raw
      // Data and Pivot Timesheet can update the exact imported roster row.
      _uuid: t._uuid || rowSourceKey,
      _rowId: t._rowId || rowSourceKey,
      _recordId: t._recordId || rowSourceKey,
      _sourceKey: rowSourceKey,
      _sourceFile: t._sourceFile,
      _mktPivotValueOverride: t._mktPivotValueOverride,
      _durationOverride: t._durationOverride,
      
      // 12 properties required for roster-raw table
      center: rCen || rawAeCode || "",
      l07: l07,
      business: centerBusiness,
      ma_nv: empId,
      full_name: effName,
      ngay: dateStr,
      type: effectiveType,
      class: effectiveClass,
      gio_vao: fromStr,
      gio_ra: toStr,
      duration: durationHours,
      loai: rawLoai,
      notes: String(getVal(t, ["notes", "note", "ghi chú", "ghi chu", "remarks"]) || "").trim().replace(/^["']|["']$/g, ""),
      overlap_check: "Không trùng",

      // Internal for overlap check and aggregation
      _sH: sH,
      _eH: eH,
      _taskField: taskField,
      _rates: rates,
      _activeRate: activeRate,
      _money: money,

      // Backward compatibility keys
      chargeToCenterMkt,
      taskField,
      employeeId: empId,
      fullName: effName,
      maAE: aeCode,
      date: dateStr,
      // MKT LOCAL NORTH_TIMESHEET must use the source file's TYPE column.
      // Never replace a blank/missing TYPE with the internal supportMkt bucket.
      sourceType: rawType,
      taskType: isMktNorthBank ? (rawType || effectiveType) : effectiveType,
      classCode: effectiveClass,
      from: fromStr,
      to: toStr,
      workingHours: actHours,
      rate: activeRate,
      payment: money, isMktLocal,
    };
    if (isMktNorthBank && (chargeToCenterMkt === "HY0001.ECP" || String(detailRow.center).toUpperCase() === "HY0001.ECP" || String(rawChargeToCenter).toUpperCase() === "HY0001.ECP")) {
      detailRow.center = "Ecopark";
      detailRow.chargeToCenterMkt = "HY0001.ECP";
    }
    const normalizeSearchStr = (s: string) => s ? s.toLowerCase().replace(/đ/g, "d").normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/\s+/g, "") : "";
    (detailRow as any)._searchStr = normalizeSearchStr(
      `${detailRow.business} ${detailRow.l07} ${detailRow.center} ${detailRow.classCode} ` +
      `${detailRow.fullName} ${detailRow.employeeId} ${detailRow.date} ${detailRow.taskType} ${detailRow.notes}`,
    );
    details.push(detailRow);
  });

  // -------------------------------------------------------------------------
  // TIME & OVERLAP HELPERS
  // -------------------------------------------------------------------------
  function timeToMinutes(timeStr: string): number {
    if (!timeStr) return 0;
    const upper = timeStr.trim().toUpperCase();
    const isPM = upper.includes("PM");
    const isAM = upper.includes("AM");
    const clean = upper.replace(/(AM|PM)/g, "").trim();
    const parts = clean.split(':');
    let hours = parseInt(parts[0], 10) || 0;
    const minutes = parseInt(parts[1], 10) || 0;
    
    // Nếu có AM/PM thì chuyển đổi sang 24h
    if (isPM && hours < 12) hours += 12;
    else if (isAM && hours === 12) hours = 0;
    
    return hours * 60 + minutes;
  }

  function checkShiftOverlap(shiftA: any, shiftB: any): boolean {
    // 1. Khác ID nhân sự -> Không trùng
    if (shiftA.ma_nv !== shiftB.ma_nv) return false;

    // 2. Khác Ngày -> Không trùng
    if (shiftA.ngay !== shiftB.ngay) return false;

    // 3. Cùng một dòng dữ liệu -> Không tính
    if (shiftA.id === shiftB.id) return false;

    // 4. Chuyển đổi thời gian ra phút
    const startA = timeToMinutes(shiftA.gio_vao);
    let endA = timeToMinutes(shiftA.gio_ra);
    const startB = timeToMinutes(shiftB.gio_vao);
    let endB = timeToMinutes(shiftB.gio_ra);

    // Xử lý ca qua đêm (End time < Start time)
    if (endA < startA) endA += 24 * 60;
    if (endB < startB) endB += 24 * 60;

    // 5. Kiểm tra giao nhau (Dùng toán tử < tuyệt đối - hai ca liền kề 0s không giao nhau)
    return startA < endB && startB < endA;
  }

  // -------------------------------------------------------------------------
  // VALIDATION LOGIC (Policy 15-JAN-2026)
  // -------------------------------------------------------------------------
  const smsCountMap = new Map<string, number>();
  const tutoringRowsByEmpDate = new Map<string, any[]>();

  // Pass 1: Statistics for SMS and grouping for Tutoring
  details.forEach(row => {
    const typeLower = (row.type || "").toLowerCase();
    if (row.class && typeLower.includes('sms')) {
      const dateClassKey = `${row.ngay}_${row.class}`;
      smsCountMap.set(dateClassKey, (smsCountMap.get(dateClassKey) || 0) + 1);
    }
    if (typeLower.includes('tutoring') || typeLower.includes('tutorial')) {
      const empDateKey = `${row.ma_nv}_${row.ngay}`;
      if (!tutoringRowsByEmpDate.has(empDateKey)) {
        tutoringRowsByEmpDate.set(empDateKey, []);
      }
      tutoringRowsByEmpDate.get(empDateKey)!.push(row);
    }
  });

  const CLASS_REQUIRED_TYPES = ["in-class", "sms", "tutoring", "tutorial", "waiting class", "parent meeting", "report"];

  // Pass 2: Individual Row Validation
  details.forEach((row) => {
    // 1. Duration Check
    let durationError = "";
    const typeLower = (row.type || "").toLowerCase();
    const dur = row.duration || 0;

    if (typeLower.includes("pick up") || typeLower.includes("drop off")) {
      if (dur >= 0.25) durationError = "Max 0.25h";
    } else if (typeLower.includes("sms")) {
      if (Math.abs(dur - 0.25) >= 0.001) durationError = "Must be 0.25h";
    } else if (typeLower.includes("tutoring") || typeLower.includes("tutorial")) {
      if (Math.abs(dur - 0.5) >= 0.001 && Math.abs(dur - 1.0) >= 0.001) durationError = "0.5h/1.0h only";
    } else if (typeLower.includes("in-class")) {
      if (Math.abs(dur - 1.5) >= 0.001 && Math.abs(dur - 2.0) >= 0.001) durationError = "1.5h/2.0h only";
    }
    
    if (!durationError && dur > 3.0) {
      durationError = "Max 3.0h";
    }
    row.check_duration = durationError || "OK";

    // 2. Class Check
    const needsClass = CLASS_REQUIRED_TYPES.some(t => typeLower.includes(t));
    let classValid = "TRUE";
    if (needsClass) {
      if (!row.class) {
        classValid = "FALSE";
      } else {
        const classUpper = String(row.class).trim().toUpperCase();
        if (classUpper === "TRUE" || classUpper === "FALSE" || classUpper === "KHÔNG CÓ LỚP HỌC") {
          classValid = "FALSE";
        } else {
          const validClsPattern = /(KDG\s*[1-3]|PRI\s*[1-6]|PRI\s*STARTER|PRIMARY\s*STARTER|SEC\s*(STARTER|FOUND))/i;
          if (!validClsPattern.test(classUpper)) {
            classValid = "FALSE";
          }
        }
      }
    }
    row.check_class = classValid;

    // 3. SMS Check
    if (row.class && typeLower.includes('sms')) {
      const dateClassKey = `${row.ngay}_${row.class}`;
      if ((smsCountMap.get(dateClassKey) || 0) >= 2) {
        row.check_sms = "Duplicate";
      } else {
        row.check_sms = "OK";
      }
    } else {
      row.check_sms = "OK";
    }

    // 4. Tutoring Check
    // 1 ID Number in same day can have consecutive time slots (0s gap) -> OK
    // Only flag Duplicate if 2 tutoring slots for the same ID on the same date strictly overlap in time
    if (typeLower.includes('tutoring') || typeLower.includes('tutorial')) {
      const empDateKey = `${row.ma_nv}_${row.ngay}`;
      const sameDayTutoring = tutoringRowsByEmpDate.get(empDateKey) || [];
      let hasTutoringOverlap = false;

      if (sameDayTutoring.length > 1) {
        for (const other of sameDayTutoring) {
          if (other === row || other.id === row.id) continue;
          if (checkShiftOverlap(row, other)) {
            hasTutoringOverlap = true;
            break;
          }
        }
      }

      row.check_tutoring = hasTutoringOverlap ? "Duplicate" : "OK";
    } else {
      row.check_tutoring = "OK";
    }
  });

  // -------------------------------------------------------------------------
  // 3. OVERLAP DETECTION
  // -------------------------------------------------------------------------

  // Duyệt qua tất cả các dòng để kiểm tra trùng lịch (group theo NV và Ngày để tối ưu hiệu năng)
  const groups: Record<string, any[]> = {};
  details.forEach(d => {
    // Cùng ID NUMBER và CÙNG DATE
    const key = `${d.ma_nv}_${d.ngay}`;
    if (!groups[key]) groups[key] = [];
    groups[key].push(d);
  });

  let overlapGroupSequence = 0;
  Object.values(groups).forEach(group => {
    if (group.length <= 1) return;
    const neighbours = new Map<any, any[]>();
    group.forEach((row) => neighbours.set(row, []));

    for (let i = 0; i < group.length; i++) {
      for (let j = i + 1; j < group.length; j++) {
        const left = group[i];
        const right = group[j];
        if (!checkShiftOverlap(left, right)) continue;
        neighbours.get(left)!.push(right);
        neighbours.get(right)!.push(left);
      }
    }

    const visited = new Set<any>();
    group.forEach((startRow) => {
      if (visited.has(startRow) || neighbours.get(startRow)!.length === 0) return;

      const component: any[] = [];
      const stack = [startRow];
      visited.add(startRow);
      while (stack.length > 0) {
        const row = stack.pop()!;
        component.push(row);
        neighbours.get(row)!.forEach((nextRow) => {
          if (visited.has(nextRow)) return;
          visited.add(nextRow);
          stack.push(nextRow);
        });
      }

      component.sort((left, right) => {
        const timeDifference = timeToMinutes(left.gio_vao) - timeToMinutes(right.gio_vao);
        return timeDifference || String(left.class || "").localeCompare(String(right.class || ""), "vi");
      });
      const groupCode = `TL-${String(++overlapGroupSequence).padStart(3, "0")}`;

      component.forEach((row, index) => {
        const overlappingRows = neighbours.get(row)!;
        const overlapDetails = overlappingRows.map((other) => {
          const classVal = String(other.class || "").trim();
          const isNAClass =
            !classVal ||
            ["NA", "N/A", "NAN", "NULL", "NONE", "KHÔNG CÓ LỚP HỌC", "KHÔNG CÓ LỚP"].includes(classVal.toUpperCase());
          const classTag = isNAClass ? "" : ` (Lớp ${classVal})`;
          return `Ca ${other.gio_vao}-${other.gio_ra}${classTag}`;
        });
        row.overlap_check = `Trùng lịch: ${overlapDetails.join(", ")}`;
        row.overlap_group = groupCode;
        row.overlap_position = index + 1;
        row.overlap_total = component.length;
        row.overlap_with_ids = overlappingRows.map((other) => other.id);
      });
    });
  });

  // Re-sort details so overlapping shifts are placed at consecutive rows right under each other
  details.sort((a, b) => {
    // 1. Group by Employee ID
    const empCmp = String(a.ma_nv || "").localeCompare(String(b.ma_nv || ""), "vi", { numeric: true });
    if (empCmp !== 0) return empCmp;

    // 2. Group by Date
    const dateCmp = String(a.ngay || "").localeCompare(String(b.ngay || ""));
    if (dateCmp !== 0) return dateCmp;

    // 3. Group overlapping shifts together if they belong to the same overlap group
    if (a.overlap_group || b.overlap_group) {
      if (a.overlap_group && b.overlap_group) {
        if (a.overlap_group !== b.overlap_group) {
          return a.overlap_group.localeCompare(b.overlap_group);
        }
        return (a.overlap_position || 0) - (b.overlap_position || 0);
      }
      // Place overlap group in chronological position
      const timeA = timeToMinutes(a.gio_vao);
      const timeB = timeToMinutes(b.gio_vao);
      if (timeA !== timeB) return timeA - timeB;
      return a.overlap_group ? -1 : 1;
    }

    // 4. Default time sort for non-overlapping rows on the same date
    const timeA = timeToMinutes(a.gio_vao);
    const timeB = timeToMinutes(b.gio_vao);
    if (timeA !== timeB) return timeA - timeB;

    return 0;
  });

  // -------------------------------------------------------------------------
  // 2. AGGREGATION
  // -------------------------------------------------------------------------
  details.forEach(detailRow => {

    const { l07, ma_nv: empId, full_name: effName, business: centerBusiness, duration: actHours, _taskField: taskField, _rates: rates, _money: money, ngay: dateStr, isMktLocal, chargeToCenterMkt } = detailRow;

    // MKT Local is the source bucket, while CHARGE TO CENTER is the actual
    // cost-allocation destination. Roster Center must therefore group a
    // non-zero MKT Local amount under that destination L07 instead of creating
    // a misleading `MKT LOCAL NORTH` center row.
    const rawAllocationL07 = String(chargeToCenterMkt || "").trim();
    const mappedAllocationL07 = rawAllocationL07
      ? mapL07(rawAllocationL07) || rawAllocationL07
      : "";
    const effectiveL07 =
      isMktLocal && mappedAllocationL07 ? mappedAllocationL07 : l07;

    const empKey = `${effectiveL07}_${empId}_${centerBusiness}`;
    if (!empGroup[empKey]) {
      const nid = normalizeId(empId);
      const staffRow = staffLookup.get(nid) || staffLookup.get(String(effName || "").toLowerCase()) || {};
      const bankAccount = String(getVal(staffRow, ["bank account number", "account number", "stk", "số tài khoản"]) || "").trim();

      empGroup[empKey] = {
        id: generateUUID(), business: centerBusiness, center: effectiveL07, employeeId: empId, fullName: effName, bankAccountNumber: bankAccount, salaryScale: rates.sCode, from: fromDateStr, to: toDateStr, className: detailRow.classCode,
        noteDays: new Set(), inClass: 0, inClassAtls: 0, demo: 0, tutoring: 0, waitingClass: 0, clubActivity: 0, parentMeeting: 0, pickUpDropOff: 0, pickUpDropOffAtls: 0, sms: 0, smsAtls: 0, progressReport: 0, progressReportAtls: 0, prepareLessonTutoring: 0, prepareLessonClubs: 0, meetingTraining: 0, pt: 0, discoveryCamp: 0, outing: 0, summer: 0, summerInstructors: 0, conductTest: 0, renewalProjects: 0, supportLxo: 0, supportEc: 0, supportMkt: 0, lpar01: 0, lret01: 0, ldem01: 0, ldec01: 0, moth01: 0, other: 0, totalHours: 0, academicHours: 0, adminHours: 0, klHours: 0, baseSalary: 0, totalSalary: 0, deductionHours: 0, _rates: rates, siRate: rates.si,
      };
    }
    const eRow = empGroup[empKey];
    if (detailRow.loai === "KL") {
      eRow.klHours += actHours;
    } else {
      if (academicFieldSet.has(taskField)) eRow.academicHours += actHours; else eRow.adminHours += actHours;
      eRow.totalHours += actHours;
      if (eRow[taskField] === undefined) eRow[taskField] = 0;
      eRow[taskField] += actHours;
      eRow.baseSalary += money; eRow.totalSalary += money;
    }
    if (detailRow.notes) eRow.noteDays.add(`${dateStr}: ${detailRow.notes}`);

    const cenId = `${effectiveL07}|${rates.sCode}`;
    if (!cenGroup[cenId]) {
      cenGroup[cenId] = {
        id: generateUUID(), l07: effectiveL07, business: centerBusiness, salaryScale: rates.sCode, acRate: rates.ac, adRate: rates.ad, suRate: rates.su, ouRate: rates.ou, siRate: rates.si, from: fromDateStr, to: toDateStr, inClass: 0, inClassAtls: 0, demo: 0, tutoring: 0, waitingClass: 0, clubActivity: 0, parentMeeting: 0, pickUpDropOff: 0, pickUpDropOffAtls: 0, sms: 0, smsAtls: 0, progressReport: 0, progressReportAtls: 0, prepareLessonTutoring: 0, prepareLessonClubs: 0, meetingTraining: 0, pt: 0, discoveryCamp: 0, outing: 0, summer: 0, summerInstructors: 0, conductTest: 0, renewalProjects: 0, supportLxo: 0, supportEc: 0, supportMkt: 0, lpar01: 0, lret01: 0, ldem01: 0, ldec01: 0, moth01: 0, other: 0, totalHours: 0, academicHours: 0, adminHours: 0, klHours: 0,
      };
    }
    const cRow = cenGroup[cenId];
    if (detailRow.loai === "KL") {
      cRow.klHours += actHours;
    } else {
      if (taskField) { if (cRow[taskField] === undefined) cRow[taskField] = 0; cRow[taskField] += actHours; }
      if (academicFieldSet.has(taskField)) cRow.academicHours += actHours; else cRow.adminHours += actHours;
      cRow.totalHours += actHours;
    }
  });

  const finalize = (groupObj: Record<string, any>) => Object.values(groupObj).map((row: any, index) => {
    // 2. Tổng Hợp Công Thức Tính Lương (Khoản 2 deduction)
    const deductionHours = ((row.inClassAtls || 0) + (row.clubActivity || 0) + (row.parentMeeting || 0)) / 2;

    const isMktLocal = row.center === "MKT LOCAL NORTH" || row.center === "MKT LOCAL SOUTH" || (row.center && typeof row.center === "string" && row.center.startsWith("MKT LOCAL NORTH_")) || row.l07 === "MKT LOCAL NORTH" || row.l07 === "MKT LOCAL SOUTH" || (row.l07 && typeof row.l07 === "string" && row.l07.startsWith("MKT LOCAL NORTH_"));
    
    // Resolve rates safely from row._rates or defaults
    const rAc = row._rates?.ac !== undefined ? row._rates.ac : 33000;
    const rAd = 20000; // Fixed Admin Price
    const rSu = 29474; // Fixed Summer Price
    const rOu = 26316; // Fixed Outing Price
    const rSi = 150000; // Fixed Summer Instructors Price

    let baseSalary = 0, totalSalary = 0, cMktLocal = 0, chargeLxo = 0, cEc = 0, cPtDemo = 0, cRenewal = 0, cDiscovery = 0, cSummerOuting = 0, cSummerInstructors = 0, chargeOther = 0;
    let chargeLdem01 = 0, chargeLdec01 = 0, chargeLpar01 = 0, chargeLret01 = 0, chargeMoth01 = 0, chargeExtraSummerInstructors = 0;

    if (isMktLocal) { 
      const rawTotalSalary = (row.totalHours || 0) * 20000; 
      totalSalary = Math.round(rawTotalSalary); 
      baseSalary = totalSalary; 
      cMktLocal = totalSalary; 
      chargeLxo = 0;
    } else {
      // 8 components for Total Salary
      const p1 = (row.academicHours || 0) * rAc;
      
      // Admin Hours excluding those with special rates
      const specialAdminHours = (row.summer || 0) + (row.outing || 0) + (row.discoveryCamp || 0) + (row.supportEc || 0) + (row.summerInstructors || 0) + (row.other || 0) + (row.moth01 || 0);
      const generalAdminHours = Math.max(0, (row.adminHours || 0) - specialAdminHours);
      const p2 = (generalAdminHours - deductionHours) * 20000;
      
      const p3 = (row.summer || 0) * 29474;
      const p4 = (row.outing || 0) * 26316;
      const p5 = (row.discoveryCamp || 0) * 29474;
      const p6 = (row.supportEc || 0) * 20000;
      const p7 = (row.summerInstructors || 0) * 150000;
      const p8 = ((row.other || 0) + (row.moth01 || 0)) * 20000;

      totalSalary = Math.round(p1 + p2 + p3 + p4 + p5 + p6 + p7 + p8);
      
      // Base Salary definition: components 1, 2, 6, 8 (Fixed monthly parts)
      baseSalary = Math.round(p1 + p2 + p6 + p8);

      // Charges for Summary Sheet (Part 3)
      cEc = Math.round((row.supportEc || 0) * 20000);
      cPtDemo = Math.round((row.demo || 0) * rAc) + Math.round((row.pt || 0) * 20000);
      chargeOther = Math.round((row.other || 0) * 20000);
      
      // Charge Renewal: ROUND(([Prepare lesson - Clubs] + [Renewal Projects]) * 20,000, 0) + ([Club activity] * [Academic Price]) - ([Club activity] / 2 * 20,000)
      cRenewal = Math.round(((row.prepareLessonClubs || 0) + (row.renewalProjects || 0)) * 20000) + ((row.clubActivity || 0) * rAc) - (((row.clubActivity || 0) / 2) * 20000);
      
      cDiscovery = Math.round((row.discoveryCamp || 0) * 29474);
      cSummerOuting = Math.round(((row.outing || 0) * 26316) + ((row.summer || 0) * 29474));
      cSummerInstructors = Math.round((row.summerInstructors || 0) * 150000);

      chargeLdem01 = Math.round((row.ldem01 || 0) * 20000);
      chargeLdec01 = Math.round((row.ldec01 || 0) * 20000);
      chargeLpar01 = Math.round((row.lpar01 || 0) * 20000);
      chargeLret01 = Math.round((row.lret01 || 0) * 20000);
      chargeMoth01 = Math.round((row.moth01 || 0) * 20000);
      chargeExtraSummerInstructors = Math.round((row.extraSummerInstructors || 0) * 150000);

      cMktLocal = chargeLdem01 + chargeLdec01 + chargeLpar01 + chargeLret01 + chargeMoth01 + Math.round((row.supportMkt || 0) * 20000);

      // Charge LXO (Trừ lùi)
      chargeLxo = totalSalary - cEc - cPtDemo - chargeOther - cRenewal - cDiscovery - cSummerOuting - cSummerInstructors - chargeExtraSummerInstructors - cMktLocal;
      if (chargeLxo < 0) chargeLxo = 0;
    }
    return { ...row, id: index + 1, deductionHours, baseSalary, totalSalary, chargeLxo, chargeEc: cEc, chargePtDemo: cPtDemo, chargeLdem01, chargeLdec01, chargeLpar01, chargeLret01, chargeMoth01, chargeMktLocal: cMktLocal, chargeOther, chargeRenewalProjects: cRenewal, chargeDiscoveryCamp: cDiscovery, chargeSummerOuting: cSummerOuting, chargeSummerInstructors: cSummerInstructors, chargeExtraSummerInstructors };
  });

  const normalizeSearchStr = (s: string) => s ? s.toLowerCase().replace(/đ/g, "d").normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/\s+/g, "") : "";
  const empResult = finalize(empGroup).map((r: any) => ({ ...r, noteDays: r.noteDays ? Array.from(r.noteDays).join(" | ") : "", _searchStr: normalizeSearchStr(`${r.className} ${r.fullName} ${r.employeeId}`) }));
  const cenAggregate: Record<string, any> = {};
  empResult.forEach((c) => {
    const key = c.l07 || c.center;
    if (!cenAggregate[key]) {
      cenAggregate[key] = {
        id: c.id, l07: key, business: c.business, from: c.from, to: c.to,
        chargeLxo: c.chargeLxo || 0,
        chargeEc: c.chargeEc || 0,
        chargePtDemo: c.chargePtDemo || 0,
        chargeLdem01: c.chargeLdem01 || 0,
        chargeLdec01: c.chargeLdec01 || 0,
        chargeLpar01: c.chargeLpar01 || 0,
        chargeLret01: c.chargeLret01 || 0,
        chargeMoth01: c.chargeMoth01 || 0,
        chargeMktLocal: c.chargeMktLocal || 0,
        chargeOther: c.chargeOther || 0,
        chargeRenewalProjects: c.chargeRenewalProjects || 0,
        chargeDiscoveryCamp: c.chargeDiscoveryCamp || 0,
        chargeSummerOuting: c.chargeSummerOuting || 0,
        chargeSummerInstructors: c.chargeSummerInstructors || 0,
        chargeExtraSummerInstructors: c.chargeExtraSummerInstructors || 0,
        totalSalary: c.totalSalary || 0,
        totalHours: c.totalHours || 0,
        salaryScale: c.salaryScale,
      };
    } else {
      cenAggregate[key].totalSalary += c.totalSalary || 0;
      cenAggregate[key].chargeLxo += c.chargeLxo || 0;
      cenAggregate[key].chargeEc += c.chargeEc || 0;
      cenAggregate[key].chargePtDemo += c.chargePtDemo || 0;
      cenAggregate[key].chargeLdem01 += c.chargeLdem01 || 0;
      cenAggregate[key].chargeLdec01 += c.chargeLdec01 || 0;
      cenAggregate[key].chargeLpar01 += c.chargeLpar01 || 0;
      cenAggregate[key].chargeLret01 += c.chargeLret01 || 0;
      cenAggregate[key].chargeMoth01 += c.chargeMoth01 || 0;
      cenAggregate[key].chargeMktLocal += c.chargeMktLocal || 0;
      cenAggregate[key].chargeOther += c.chargeOther || 0;
      cenAggregate[key].chargeRenewalProjects += c.chargeRenewalProjects || 0;
      cenAggregate[key].chargeDiscoveryCamp += c.chargeDiscoveryCamp || 0;
      cenAggregate[key].chargeSummerOuting += c.chargeSummerOuting || 0;
      cenAggregate[key].chargeSummerInstructors += c.chargeSummerInstructors || 0;
      cenAggregate[key].chargeExtraSummerInstructors += c.chargeExtraSummerInstructors || 0;
      cenAggregate[key].totalHours += c.totalHours || 0;
    }
  });
  const cenResult = Object.values(cenAggregate);

  return { processedRosterData: details, employeeSummary: empResult, centerSummary: cenResult, isCalculating: false };

  } catch (error: any) {
    return {
      processedRosterData: [],
      employeeSummary: [],
      centerSummary: [],
      isCalculating: false,
      error: error.message || String(error)
    };
  }
}

if (typeof window === "undefined" && typeof self !== "undefined") {
  const inputFields = [
    "rosterData",
    "salaryScaleData",
    "staffData",
    "cacheData",
  ] as const;
  let pendingInput: {
    requestId?: string;
    params: any;
    data: Record<(typeof inputFields)[number], any[]>;
  } | null = null;

  const resultFields = ["processedRosterData", "employeeSummary", "centerSummary"] as const;
  let pendingOutput: {
    result: any; requestId?: string; fieldIndex: number; offset: number; chunkId: number;
  } | null = null;
  const sendNextResultChunk = () => {
    const output = pendingOutput;
    if (!output) return;
    while (output.fieldIndex < resultFields.length) {
      const field = resultFields[output.fieldIndex];
      const rows = output.result[field] || [];
      if (output.offset >= rows.length) {
        output.result[field] = [];
        output.fieldIndex++;
        output.offset = 0;
        continue;
      }
      const end = Math.min(output.offset + 2_000, rows.length);
      self.postMessage({ type: "timesheet-result-chunk", requestId: output.requestId,
        field, chunkId: ++output.chunkId, rows: rows.slice(output.offset, end) });
      // postMessage has cloned these rows. Release them in this worker before
      // waiting for the UI, rather than retaining both complete datasets.
      for (let i = output.offset; i < end; i++) rows[i] = undefined;
      output.offset = end;
      return;
    }
    self.postMessage({ type: "timesheet-result-complete", requestId: output.requestId });
    pendingOutput = null;
  };

  const postChunkedResult = (result: any, requestId?: string, backpressure = false) => {
    if (backpressure) {
      pendingOutput = { result, requestId, fieldIndex: 0, offset: 0, chunkId: 0 };
      self.postMessage({ type: "timesheet-result-start", requestId, error: result.error });
      sendNextResultChunk();
      return;
    }
    const resultFields = [
      "processedRosterData",
      "employeeSummary",
      "centerSummary",
    ] as const;
    // Larger batches materially reduce postMessage/structured-clone overhead
    // on 50k-200k row Timesheet files while still keeping the UI responsive.
    const chunkSize = 10_000;

    self.postMessage({
      type: "timesheet-result-start",
      requestId,
      error: result.error,
    });
    resultFields.forEach((field) => {
      const rows = Array.isArray(result[field]) ? result[field] : [];
      for (let offset = 0; offset < rows.length; offset += chunkSize) {
        self.postMessage({
          type: "timesheet-result-chunk",
          requestId,
          field,
          rows: rows.slice(offset, offset + chunkSize),
        });
      }
    });
    self.postMessage({ type: "timesheet-result-complete", requestId });
  };

  self.onmessage = (e: MessageEvent) => {
    const message = e.data || {};

    if (message.type === "timesheet-result-ack") {
      if (pendingOutput?.requestId === message.requestId && pendingOutput?.chunkId === message.chunkId) sendNextResultChunk();
      return;
    }

    if (message.type === "timesheet-input-start") {
      pendingOutput = null;
      pendingInput = {
        requestId: message.requestId,
        params: message.params || {},
        data: {
          rosterData: [],
          salaryScaleData: [],
          staffData: [],
          cacheData: [],
        },
      };
      return;
    }

    if (message.type === "timesheet-input-chunk") {
      if (!pendingInput || pendingInput.requestId !== message.requestId) return;
      const target = pendingInput.data[message.field as (typeof inputFields)[number]];
      if (target && Array.isArray(message.rows)) target.push(...message.rows);
      return;
    }

    if (message.type === "timesheet-input-complete") {
      if (!pendingInput || pendingInput.requestId !== message.requestId) return;
      const request = pendingInput;
      pendingInput = null;
      postChunkedResult(
        calculateTimesheet({ ...request.params, ...request.data }),
        request.requestId,
        request.params.resultBackpressure === true,
      );
      return;
    }

    // Backwards-compatible direct calculation path.
    postChunkedResult(calculateTimesheet(message), message.requestId);
  };
}
