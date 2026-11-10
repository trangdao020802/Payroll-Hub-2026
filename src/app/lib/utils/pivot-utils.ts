/* eslint-disable @typescript-eslint/no-explicit-any */
import {
  getBusinessFromL07,
  resolveMktRosterCenter,
  resolveSummerBonusCenterL07,
} from "./center-utils";
import { parsePayrollMonth } from "./hold-carryover";
import { parseDurationToHours } from "../schemas/excel-schema";
import { normalizeGrossPaySpecialCenter, resolveMasterSpecialCenter } from "./master-special-centers";

export const PIVOT_CACHE_VERSION = 13;
export const PIVOT_MKT_TYPE_CACHE_KEY = "pivot_master_mkt_type_data";
export const PIVOT_MKT_TYPE_CACHE_VERSION = 2;
export const PIVOT_SOURCE_MARKER_PREFIX = "__PIVOT_SOURCE__";
export const ZHN_SHARED_L07 = "ZHN0000.GY";

export type PivotGroupedData = Record<
  string,
  Record<string, Record<string, Record<string, number>>>
>;

export interface PivotMktTypeCache {
  cacheVersion: number;
  groupedData: PivotGroupedData;
  typeColumns: string[];
  months: string[];
  updatedAt: number;
  manualMonths?: string[];
}

export function normalizePivotL07(l07Raw: string): string {
  const l07 = String(l07Raw || "").trim();
  return resolveMasterSpecialCenter(l07)?.l07 || l07;
}

export function getPivotZhnSourceLabel(...sourceValues: unknown[]): string {
  for (const value of sourceValues) {
    const normalized = String(value || "").trim().toUpperCase();
    if (!normalized) continue;
    if (normalized.includes("CONTEST")) return "CONTEST";
    if (normalized.includes("CAMBRIDGE")) return "CAMBRIDGE";
  }
  return "";
}

export function markPivotZhnSource(
  values: Record<string, number>,
  l07Raw: string,
  ...sourceValues: unknown[]
): void {
  if (normalizePivotL07(l07Raw).toUpperCase() !== ZHN_SHARED_L07) return;
  const sourceLabel = getPivotZhnSourceLabel(l07Raw, ...sourceValues);
  if (!sourceLabel) return;
  values[`${PIVOT_SOURCE_MARKER_PREFIX}${sourceLabel}`] = 1;
}

export function getPivotSourceLabels(values: Record<string, number> = {}): string[] {
  const preferredOrder = ["CAMBRIDGE", "CONTEST"];
  return Object.keys(values || {})
    .filter((key) => key.startsWith(PIVOT_SOURCE_MARKER_PREFIX) && Number(values[key]) !== 0)
    .map((key) => key.slice(PIVOT_SOURCE_MARKER_PREFIX.length))
    .filter(Boolean)
    .sort((a, b) => {
      const aIndex = preferredOrder.indexOf(a);
      const bIndex = preferredOrder.indexOf(b);
      if (aIndex !== -1 || bIndex !== -1) {
        if (aIndex === -1) return 1;
        if (bIndex === -1) return -1;
        return aIndex - bIndex;
      }
      return a.localeCompare(b);
    });
}

export function formatPivotTypeHeader(typeRaw: string): string {
  if (!typeRaw) return "UNSPECIFIED";
  let t = String(typeRaw).trim();
  if (!t || t.toUpperCase() === "N/A" || t.toUpperCase() === "NAN") return "UNSPECIFIED";

  // Strip prefix "CHARGE TO " or "CHARGE " (case insensitive)
  if (/^CHARGE\s+TO\s+/i.test(t)) {
    t = t.replace(/^CHARGE\s+TO\s+/i, "").trim();
  } else if (/^CHARGE\s+/i.test(t)) {
    t = t.replace(/^CHARGE\s+/i, "").trim();
  }

  const cleanUpper = t.toUpperCase();
  if (cleanUpper === "ADD" || cleanUpper === "CANCEL") return "EXCLUDE";
  if (cleanUpper === "CENTER MKT" || cleanUpper === "MKT LOCAL NORTH" || cleanUpper === "MKT LOCAL") return "MKT LOCAL";
  if (
    cleanUpper === "BONUS" ||
    cleanUpper === "EXTRA INSTRUCTORS" ||
    cleanUpper === "EXTRA INSTRUCTOR" ||
    cleanUpper === "EXTRA SUMMER INSTRUCTOR" ||
    cleanUpper === "EXTRA SUMMER INSTRUCTORS" ||
    cleanUpper === "SUMMER INSTRUCTORS BONUS" ||
    cleanUpper === "CHARGE TO EXTRA SUMMER INSTRUCTORS" ||
    cleanUpper === "CHARGE TO EXTRA INSTRUCTORS"
  ) {
    return "EXTRA SUMMER INSTRUCTORS";
  }
  if (!t) return "UNSPECIFIED";

  // The reference Master logic groups every Roster task into the five
  // standard payroll columns. Some source files contain a suffix/description
  // after the code, while older files only contain the four-letter prefix.
  // Canonicalize both variants so amounts cannot be split into hidden columns.
  if (cleanUpper.startsWith("LPAR")) return "LPAR01";
  if (cleanUpper.startsWith("LRET")) return "LRET01";
  if (cleanUpper.startsWith("LDEM")) return "LDEM01";
  if (cleanUpper.startsWith("LDEC")) return "LDEC01";
  if (cleanUpper.startsWith("MOTH")) return "MOTH01";

  return cleanUpper;
}

export function isMktRosterTypeColumn(typeRaw: string): boolean {
  if (!typeRaw) return false;
  if (String(typeRaw).startsWith(PIVOT_SOURCE_MARKER_PREFIX)) return false;
  const type = formatPivotTypeHeader(typeRaw);
  if (!type || type === "EXCLUDE" || type === "ADD" || type === "CANCEL" || type === "UNSPECIFIED") return false;
  const standardGross = new Set([
    "LXO", "EC", "PT-DEMO", "OTHER", "RENEWAL PROJECTS", "DISCOVERY CAMP",
    "SUMMER OUTING", "SUMMER INSTRUCTORS", "EXTRA SUMMER INSTRUCTORS", "EXTRA INSTRUCTORS"
  ]);
  if (standardGross.has(type.toUpperCase())) return false;
  return true;
}

const clonePivotGroupedData = (data: PivotGroupedData = {}): PivotGroupedData =>
  JSON.parse(JSON.stringify(data || {}));

export function normalizePivotMonth(value: unknown): string {
  const month = String(value || "").trim();
  return parsePayrollMonth(month.replace(/(\d)\s+((?:19|20)\d{2})$/, "$1.$2"))?.dot || month;
}

export function omitPivotMonths(data: PivotGroupedData, months: readonly string[]): PivotGroupedData {
  if (months.length === 0) return data;
  const omitted = new Set(months.map(normalizePivotMonth));
  const result: PivotGroupedData = {};
  for (const [bu, centers] of Object.entries(data)) {
    for (const [l07, periods] of Object.entries(centers)) {
      const kept = Object.entries(periods).filter(([month]) => !omitted.has(normalizePivotMonth(month)));
      if (kept.length > 0) (result[bu] ??= {})[l07] = Object.fromEntries(kept);
    }
  }
  return result;
}

export function replacePivotCacheMonth(current: PivotGroupedData, rebuilt: PivotGroupedData, month: string): PivotGroupedData {
  const restored = getPivotDataMonths(rebuilt);
  if (!restored.includes(month)) return current;
  const result = omitPivotMonths(current, [month]);
  const selected = omitPivotMonths(rebuilt, restored.filter(value => value !== month));
  for (const [bu, centers] of Object.entries(selected)) {
    const target = result[bu] ??= {};
    for (const [l07, periods] of Object.entries(centers)) {
      target[l07] = {...target[l07], ...periods};
    }
  }
  return result;
}

export function requestPivotCacheMonthReset(month: string): void {
  try {
    const raw = localStorage.getItem('pivot_master_processed_data');
    if (!raw) return;
    const cache = JSON.parse(raw);
    if (cache?.cacheVersion === PIVOT_CACHE_VERSION) {
      const pending = Array.isArray(cache.resetMonths) ? cache.resetMonths : [];
      localStorage.setItem('pivot_master_processed_data', JSON.stringify({...cache, resetMonths: [...new Set([...pending, month])]}));
    }
  } catch {
    // A missing/invalid cache is rebuilt from Master when Pivot opens.
  }
}

export function getMasterRosterMonth(row: Record<string, unknown>, reportingMonth: string): string {
  return normalizePivotMonth(row._fileMonth || row.month || row["Tháng báo cáo"] || row["Tháng"] || reportingMonth);
}

// Legacy Master imports appended sources. Source rows are stored in import
// order, so keep the last identifiable file for each month as a whole.
function latestPivotRosterRows<T extends Record<string, unknown>>(rows: T[], reportingMonth: string): T[] {
  const latestSource = new Map<string, unknown>();
  const sourceKey = (row: Record<string, unknown>) => row._masterInputId || row._sourceFile;
  for (const row of rows) {
    if (row && sourceKey(row)) latestSource.set(getMasterRosterMonth(row, reportingMonth), sourceKey(row));
  }
  return rows.filter(row => {
    if (!row) return false;
    const latest = latestSource.get(getMasterRosterMonth(row, reportingMonth));
    return !latest || sourceKey(row) === latest;
  });
}

export function getPivotDataMonths(data: PivotGroupedData = {}): string[] {
  const months = new Set<string>();
  Object.values(data || {}).forEach((l07Rows) => {
    Object.values(l07Rows || {}).forEach((monthRows) => {
      Object.keys(monthRows || {}).forEach((month) => {
        const normalized = normalizePivotMonth(month);
        if (normalized) months.add(normalized);
      });
    });
  });
  return Array.from(months);
}

export function extractPivotMktTypeData(
  data: PivotGroupedData = {},
  knownTypeColumns: string[] = [],
): PivotMktTypeCache {
  const managedTypes = new Set(
    (knownTypeColumns || [])
      .map(formatPivotTypeHeader)
      .filter(isMktRosterTypeColumn),
  );
  const groupedData: PivotGroupedData = {};
  const months = new Set<string>();

  Object.entries(data || {}).forEach(([bu, l07Rows]) => {
    Object.entries(l07Rows || {}).forEach(([l07, monthRows]) => {
      const uL07 = String(l07).trim().toUpperCase();
      if (
        uL07.includes("MKT LOCAL NORTH") ||
        uL07.startsWith("MKT LOCAL") ||
        uL07.includes("MKT_LOCAL")
      ) {
        return;
      }
      Object.entries(monthRows || {}).forEach(([month, typeRows]) => {
        Object.entries(typeRows || {}).forEach(([rawType, rawAmount]) => {
          const type = formatPivotTypeHeader(rawType);
          if (!managedTypes.has(type) && !isMktRosterTypeColumn(type)) return;
          const amount = Number(rawAmount);
          if (!Number.isFinite(amount)) return;

          managedTypes.add(type);
          months.add(month);
          if (!groupedData[bu]) groupedData[bu] = {};
          if (!groupedData[bu][l07]) groupedData[bu][l07] = {};
          if (!groupedData[bu][l07][month]) groupedData[bu][l07][month] = {};
          groupedData[bu][l07][month][type] =
            (groupedData[bu][l07][month][type] || 0) + amount;
        });
      });
    });
  });

  return {
    cacheVersion: PIVOT_MKT_TYPE_CACHE_VERSION,
    groupedData,
    typeColumns: Array.from(managedTypes).sort(),
    months: Array.from(months).sort(),
    updatedAt: Date.now(),
  };
}

export function readPivotMktTypeCache(
  fallbackData: PivotGroupedData = {},
  fallbackTypeColumns: string[] = [],
): PivotMktTypeCache {
  try {
    if (typeof localStorage !== "undefined") {
      const cached = localStorage.getItem(PIVOT_MKT_TYPE_CACHE_KEY);
      if (cached) {
        const parsed = JSON.parse(cached);
        if (
          parsed?.cacheVersion === PIVOT_MKT_TYPE_CACHE_VERSION &&
          parsed.groupedData &&
          Array.isArray(parsed.typeColumns) &&
          Array.isArray(parsed.months)
        ) {
          // Validate that the cache contains only MKT Roster type columns and clean L07 rows
          const sanitizedGroupedData: PivotGroupedData = {};
          Object.entries(parsed.groupedData || {}).forEach(([bu, l07Rows]: [string, any]) => {
            Object.entries(l07Rows || {}).forEach(([l07, monthRows]: [string, any]) => {
              const uL07 = String(l07).trim().toUpperCase();
              if (
                uL07.includes("MKT LOCAL NORTH") ||
                uL07.startsWith("MKT LOCAL") ||
                uL07.includes("MKT_LOCAL")
              ) {
                return;
              }
              if (!sanitizedGroupedData[bu]) sanitizedGroupedData[bu] = {};
              sanitizedGroupedData[bu][l07] = monthRows;
            });
          });

          const sanitizedTypeColumns = parsed.typeColumns
            .map(formatPivotTypeHeader)
            .filter(isMktRosterTypeColumn);
          return {
            ...parsed,
            groupedData: sanitizedGroupedData,
            typeColumns: sanitizedTypeColumns,
          };
        }
      }
    }
  } catch (error) {
    console.warn("Không thể đọc cache TYPE MKT Local của Pivot Master", error);
  }

  // Migration cho dữ liệu đã có trước khi tách cache TYPE MKT Local.
  return extractPivotMktTypeData(fallbackData, fallbackTypeColumns);
}

export function writePivotMktTypeCache(cache: PivotMktTypeCache): void {
  try {
    if (typeof localStorage !== "undefined") {
      localStorage.setItem(PIVOT_MKT_TYPE_CACHE_KEY, JSON.stringify(cache));
    }
  } catch (error) {
    console.warn("Không thể lưu cache TYPE MKT Local của Pivot Master", error);
  }
}

export function updatePivotMktTypeCache(
  currentCache: PivotMktTypeCache,
  incomingData: PivotGroupedData,
  incomingTypeColumns: string[],
  affectedMonths: string[],
  manualUpload = false,
): PivotMktTypeCache {
  const monthsToReplace = new Set(affectedMonths.map(normalizePivotMonth).filter(Boolean));
  if (monthsToReplace.size === 0) return currentCache;
  const groupedData = omitPivotMonths(currentCache.groupedData, [...monthsToReplace]);

  // CHỈ quản lý các cột TYPE của MKT Roster. Tuyệt đối không bao gồm LXO hay các cột Gross Pay khác!
  const managedTypes = new Set(
    [...(currentCache?.typeColumns || []), ...(incomingTypeColumns || [])]
      .map(formatPivotTypeHeader)
      .filter((type) => type && isMktRosterTypeColumn(type)),
  );

  Object.entries(incomingData || {}).forEach(([bu, l07Rows]) => {
    Object.entries(l07Rows || {}).forEach(([l07, monthRows]) => {
      const uL07 = String(l07).trim().toUpperCase();
      if (
        uL07.includes("MKT LOCAL NORTH") ||
        uL07.startsWith("MKT LOCAL") ||
        uL07.includes("MKT_LOCAL")
      ) {
        return;
      }
      Object.entries(monthRows || {}).forEach(([rawMonth, typeRows]) => {
        const month = normalizePivotMonth(rawMonth);
        if (!monthsToReplace.has(month)) return;
        Object.entries(typeRows || {}).forEach(([rawType, rawAmount]) => {
          const type = formatPivotTypeHeader(rawType);
          if (!managedTypes.has(type) || !isMktRosterTypeColumn(type)) return;
          const amount = Number(rawAmount);
          if (!Number.isFinite(amount)) return;

          if (!groupedData[bu]) groupedData[bu] = {};
          if (!groupedData[bu][l07]) groupedData[bu][l07] = {};
          if (!groupedData[bu][l07][month]) groupedData[bu][l07][month] = {};
          groupedData[bu][l07][month][type] =
            (groupedData[bu][l07][month][type] || 0) + amount;
        });
      });
    });
  });

  return {
    cacheVersion: PIVOT_MKT_TYPE_CACHE_VERSION,
    groupedData,
    typeColumns: Array.from(managedTypes).sort(),
    months: Array.from(
      new Set([...(currentCache?.months || []), ...monthsToReplace]),
    ).sort(),
    updatedAt: Date.now(),
    manualMonths: [
      ...(currentCache.manualMonths || []).filter(month => !monthsToReplace.has(normalizePivotMonth(month))),
      ...(manualUpload ? monthsToReplace : []),
    ],
  };
}

export function applyPivotMktTypeCache(
  baseData: PivotGroupedData,
  cache: PivotMktTypeCache,
  affectedMonths?: readonly string[],
): PivotGroupedData {
  const result = clonePivotGroupedData(baseData || {});
  const managedTypes = new Set(
    (cache?.typeColumns || [])
      .map(formatPivotTypeHeader)
      .filter(isMktRosterTypeColumn),
  );
  const managedMonths = new Set((affectedMonths ?? cache?.months ?? []).map(normalizePivotMonth));

  // Chỉ xóa/cập nhật các cột TYPE MKT trong các tháng đã được quản lý.
  // Toàn bộ cột Gross Pay/Pivot khác (như LXO, EC, PT-DEMO,...) TUYỆT ĐỐI giữ nguyên 100%!
  Object.keys(result).forEach((bu) => {
    Object.keys(result[bu] || {}).forEach((l07) => {
      Object.entries(result[bu][l07] || {}).forEach(([month, typeRows]) => {
        if (!managedMonths.has(normalizePivotMonth(month))) return;
        Object.keys(typeRows || {}).forEach((rawType) => {
          const formatted = formatPivotTypeHeader(rawType);
          if (managedTypes.has(formatted) && isMktRosterTypeColumn(formatted)) {
            delete typeRows[rawType];
          }
        });
        if (Object.keys(typeRows || {}).length === 0) {
          delete result[bu][l07][month];
        }
      });
      if (Object.keys(result[bu][l07] || {}).length === 0) delete result[bu][l07];
    });
    if (Object.keys(result[bu] || {}).length === 0) delete result[bu];
  });

  Object.entries(cache?.groupedData || {}).forEach(([bu, l07Rows]) => {
    Object.entries(l07Rows || {}).forEach(([l07, monthRows]) => {
      const uL07 = String(l07).trim().toUpperCase();
      if (
        uL07.includes("MKT LOCAL NORTH") ||
        uL07.startsWith("MKT LOCAL") ||
        uL07.includes("MKT_LOCAL")
      ) {
        return;
      }
      Object.entries(monthRows || {}).forEach(([month, typeRows]) => {
        const normalizedMonth = normalizePivotMonth(month);
        if (affectedMonths && !managedMonths.has(normalizedMonth)) return;
        const resultMonths = result[bu]?.[l07];
        const matchingMonth = Object.keys(resultMonths || {}).find(key => normalizePivotMonth(key) === normalizedMonth);
        const targetMonth = matchingMonth || normalizedMonth;
        Object.entries(typeRows || {}).forEach(([rawType, rawAmount]) => {
          const type = formatPivotTypeHeader(rawType);
          if (!managedTypes.has(type) || !isMktRosterTypeColumn(type)) return;
          const amount = Number(rawAmount);
          if (!Number.isFinite(amount)) return;

          const center = ((result[bu] ??= {})[l07] ??= {});
          (center[targetMonth] ??= {})[type] = amount;
        });
      });
    });
  });

  return result;
}

export function sanitizePivotData(
  groupedData: Record<string, Record<string, Record<string, number>>>,
  typeColumns: string[] = []
) {
  const newGroupedData: Record<string, Record<string, Record<string, number>>> = {};
  const uniqueTypes = new Set<string>();

  if (groupedData) {
    Object.keys(groupedData).forEach(bu => {
      const buObj = groupedData[bu];
      if (!buObj) return;
      if (!newGroupedData[bu]) newGroupedData[bu] = {};

      Object.keys(buObj).forEach(l07 => {
        const l07Obj = buObj[l07];
        if (!l07Obj) return;
        if (!newGroupedData[bu][l07]) newGroupedData[bu][l07] = {};

        Object.keys(l07Obj).forEach(rawType => {
          const amount = l07Obj[rawType];
          if (!amount || isNaN(amount)) return;
          if (rawType.startsWith(PIVOT_SOURCE_MARKER_PREFIX)) {
            newGroupedData[bu][l07][rawType] = 1;
            return;
          }
          const cleanType = formatPivotTypeHeader(rawType);
          if (cleanType === "EXCLUDE" || cleanType === "ADD" || cleanType === "CANCEL") return;

          uniqueTypes.add(cleanType);

          if (!newGroupedData[bu][l07][cleanType]) {
            newGroupedData[bu][l07][cleanType] = 0;
          }
          newGroupedData[bu][l07][cleanType] += amount;
        });
      });
    });
  }

  if (typeColumns && typeColumns.length > 0) {
    typeColumns.forEach(t => {
      const clean = formatPivotTypeHeader(t);
      if (clean !== "EXCLUDE" && clean !== "ADD" && clean !== "CANCEL") {
        uniqueTypes.add(clean);
      }
    });
  }

  const sortedTypes = Array.from(uniqueTypes).sort((a, b) => {
    if (a === "MKT LOCAL") return -1;
    if (b === "MKT LOCAL") return 1;
    if (a === "UNSPECIFIED") return 1;
    if (b === "UNSPECIFIED") return -1;
    return a.localeCompare(b);
  });

  return {
    groupedData: newGroupedData,
    typeColumns: sortedTypes
  };
}

const KNOWN_NON_CHARGE_KEYS = new Set([
  "NO", "NO.", "ID NUMBER", "ID", "MÃ NV", "FULL NAME", "HỌ VÀ TÊN", "BANK ACCOUNT NUMBER", "BANK NAME",
  "CITAD CODE", "TAX CODE", "CONTRACT NO", "TOTAL PAYMENT", "TOTAL", "THỰC NHẬN", "CENTER", "MÃ TT", "TRUNG TÂM",
  "BUSINESS", "BU", "L07", "_RAWAE", "THÁNG", "THÁNG BÁO CÁO", "SALARY SCALE", "FROM", "TO", "TYPE", "TÊN FILE",
  "_FILEBANK", "_FILEMONTH", "_ROWID", "STATUS", "NOTE", "GHI CHÚ"
]);

const STANDARD_GROSS_PAY_CHARGES: { key: string; type: string }[] = [
  { key: "CHARGE TO LXO", type: "LXO" },
  { key: "CHARGE TO EC", type: "EC" },
  { key: "CHARGE TO PT-DEMO", type: "PT-DEMO" },
  { key: "Charge MKT Local", type: "MKT LOCAL" },
  { key: "CHARGE TO OTHER", type: "OTHER" },
  { key: "Charge Renewal Projects", type: "RENEWAL PROJECTS" },
  { key: "Charge Discovery Camp", type: "DISCOVERY CAMP" },
  { key: "Charge Summer Outing", type: "SUMMER OUTING" },
  { key: "Charge Summer Instructors", type: "SUMMER INSTRUCTORS" },
  { key: "Extra Summer Instructors", type: "EXTRA SUMMER INSTRUCTORS" },
  { key: "CHARGE TO EXTRA SUMMER INSTRUCTORS", type: "EXTRA SUMMER INSTRUCTORS" },
  { key: "Charge Extra Summer Instructors", type: "EXTRA SUMMER INSTRUCTORS" },
  { key: "Extra Instructors", type: "EXTRA SUMMER INSTRUCTORS" },
  { key: "CHARGE TO EXTRA INSTRUCTORS", type: "EXTRA SUMMER INSTRUCTORS" },
  { key: "BONUS", type: "EXTRA SUMMER INSTRUCTORS" },
];

export function buildPivotFromAppData(
  sheet1Rows: any[] = [],
  _holdRows: any[] = [],
  rosterRows: any[] = [],
  reportingMonth?: string,
) {
  void _holdRows;
  const newGroupedData: Record<string, Record<string, Record<string, Record<string, number>>>> = {};
  const uniqueTypes = new Set<string>();
  const activeReportingMonth = String(reportingMonth || "")
    .trim()
    .replace("/", ".");

  const parseMoney = (val: any): number => {
    if (typeof val === "number") return val;
    if (!val) return 0;
    const str = String(val).replace(/,/g, "").trim();
    const num = parseFloat(str);
    return isNaN(num) ? 0 : num;
  };

  const addAmount = (
    buRaw: string,
    l07Raw: string,
    monthRaw: string,
    typeRaw: string,
    amount: number,
    sourceRaw?: unknown,
  ) => {
    if (!amount || isNaN(amount)) return;
    let bu = (buRaw || "").trim().toUpperCase();
    const sourceMonth = (monthRaw || "").trim();
    let month = sourceMonth || activeReportingMonth || "03.2026";

    // Sanity check: swap if BU is a month format (e.g. 03.2026, 03/2026, THÁNG 3) or month is a known BU
    const isMonthStr = (s: string) => /^\d{1,2}[./-]\d{2,4}$/.test(s) || /^(THÁNG|THANG|MONTH)\b/i.test(s);
    const isKnownBU = (s: string) => ["AHN", "EC", "LXO", "OTHER", "AFL", "AEC", "KINDY", "PRIMARY", "SECONDARY", "MKT"].includes(s.toUpperCase());

    if (isMonthStr(bu) || (isKnownBU(sourceMonth) && !isKnownBU(bu))) {
      const temp = bu;
      bu = sourceMonth.toUpperCase();
      month = temp;
    }

    if (!bu || bu === "UNKNOWN" || isMonthStr(bu)) {
      bu = "OTHER";
    }

    const specialCenter = resolveMasterSpecialCenter(l07Raw, bu);
    const l07 = specialCenter?.l07 || normalizePivotL07((l07Raw || "UNKNOWN").trim());
    if (specialCenter) bu = specialCenter.business;
    month = normalizePivotMonth(month);
    const type = formatPivotTypeHeader(typeRaw);

    if (type === "EXCLUDE" || type === "ADD" || type === "CANCEL") return;

    uniqueTypes.add(type);

    if (!newGroupedData[bu]) newGroupedData[bu] = {};
    if (!newGroupedData[bu][l07]) newGroupedData[bu][l07] = {};
    if (!newGroupedData[bu][l07][month]) newGroupedData[bu][l07][month] = {};
    markPivotZhnSource(newGroupedData[bu][l07][month], l07, sourceRaw);
    if (!newGroupedData[bu][l07][month][type]) newGroupedData[bu][l07][month][type] = 0;
    newGroupedData[bu][l07][month][type] += amount;
  };

  sheet1Rows.forEach((sourceRow) => {
    if (!sourceRow) return;
    const row = normalizeGrossPaySpecialCenter(sourceRow);
    const rawL07 = row["L07"] || row["Center"] || row["CHARGE TO CENTER"] || "";
    const isSummerBonusRow = [
      "Extra Summer Instructors",
      "CHARGE TO EXTRA SUMMER INSTRUCTORS",
      "Charge Extra Summer Instructors",
      "Extra Instructors",
      "CHARGE TO EXTRA INSTRUCTORS",
      "BONUS",
    ].some((key) => parseMoney(row[key]) !== 0) ||
      String(row["Note"] || row["Sheet Source"] || "").toUpperCase().includes("SUMMER BONUS");
    const summerCenter = isSummerBonusRow && !resolveMasterSpecialCenter(rawL07)
      ? resolveSummerBonusCenterL07(rawL07)
      : null;
    const l07 = summerCenter?.l07 || normalizePivotL07(rawL07);
    const bu = summerCenter?.business || row["Business"] || row["BU"] || "";
    const sourceCenter =
      row["_rawAE"] ||
      row["Center"] ||
      row["CENTER"] ||
      row["Mã ae"] ||
      row["MÃ AE"] ||
      row["Note"] ||
      rawL07;
    const month = row["_fileMonth"] || row["Tháng báo cáo"] || row["Tháng"] || row["month"] || activeReportingMonth || "03.2026";
    if (!l07) return;

    // MKT LOCAL NORTH (and all regional variants like MKT LOCAL NORTH_HP, MKT LOCAL NORTH_TN,
    // MKT LOCAL NORTH_TH, MKT LOCAL NORTH_PT, MKT HP, etc.) is only the temporary Gross Pay bucket in Sheet1_AE.
    // Its breakdown amounts are distributed to the real L07 rows via MKT Local Roster,
    // therefore NO MKT LOCAL NORTH rows from Sheet1_AE should ever appear or add amounts (especially TYPE = OTHER) in Pivot Master!
    const uL07 = String(l07).trim().toUpperCase();
    if (
      uL07.includes("MKT LOCAL NORTH") ||
      uL07.startsWith("MKT LOCAL") ||
      uL07.includes("MKT_LOCAL_NORTH") ||
      uL07 === "MKT" ||
      uL07 === "NTW"
    ) {
      return;
    }

    const seenTypesInRow = new Set<string>();

    // 1. First process standard Gross Pay charge columns
    STANDARD_GROSS_PAY_CHARGES.forEach(({ key, type }) => {
      if (seenTypesInRow.has(type)) return;
      if (row[key] !== undefined && row[key] !== null && row[key] !== "") {
        const amt = parseMoney(row[key]);
        if (amt !== 0) {
          seenTypesInRow.add(type);
          addAmount(bu, l07, month, type, amt, sourceCenter);
        }
      }
    });

    // 2. Check if row contains any other non-standard charge columns
    Object.keys(row).forEach((key) => {
      const uKey = key.toUpperCase().trim();
      if (KNOWN_NON_CHARGE_KEYS.has(uKey)) return;
      if (
        uKey.includes("CENTER") ||
        uKey.includes("TRUNG TÂM") ||
        uKey.includes("NOTE") ||
        uKey.includes("STATUS") ||
        uKey.includes("ACCOUNT") ||
        uKey.includes("NAME") ||
        uKey.includes("CODE")
      ) return;

      if (
        uKey.includes("CHARGE") ||
        uKey.startsWith("LDEC") ||
        uKey.startsWith("LDEM") ||
        uKey.startsWith("LPAR") ||
        uKey.startsWith("LRET") ||
        uKey.startsWith("MOTH")
      ) {
        const cleanType = formatPivotTypeHeader(key);
        if (!seenTypesInRow.has(cleanType) && cleanType !== "EXCLUDE" && cleanType !== "ADD" && cleanType !== "CANCEL") {
          const amt = parseMoney(row[key]);
          if (amt !== 0) {
            seenTypesInRow.add(cleanType);
            addAmount(bu, l07, month, cleanType, amt, sourceCenter);
          }
        }
      }
    });

    // 3. If no individual charge columns were found in this row, fall back to TOTAL PAYMENT
    if (seenTypesInRow.size === 0) {
      const totalPay = parseMoney(row["TOTAL PAYMENT"] || row["TOTAL"] || 0);
      const type = row["Type"] || row["LOẠI"] || row["Phân loại"] || row["Nghiệp vụ"] || "UNSPECIFIED";
      const cleanType = formatPivotTypeHeader(type);
      if (totalPay !== 0 && cleanType !== "EXCLUDE" && cleanType !== "ADD" && cleanType !== "CANCEL") {
        addAmount(bu, l07, month, cleanType, totalPay, sourceCenter);
      }
    }
  });

  // holdRows removed as per user request ("xóa hold đi, ko lấy dữ liệu sheet hold ae_master")

  latestPivotRosterRows(rosterRows, activeReportingMonth || "03.2026").forEach((row) => {
    if (!row) return;
    const center = String(
      row["chargeToCenterCode"] ||
        row["chargeToCenterMkt"] ||
        row["CHARGE TO CENTER"] ||
        row["Charge To Center MKT"] ||
        row["Center"] ||
        "",
    ).trim();
    const resolvedCenter = resolveMktRosterCenter(center);
    const l07 = resolvedCenter.l07 || row["l07"] || row["L07"] || "";
    const uL07 = String(l07).trim().toUpperCase();
    if (
      !l07 ||
      uL07.includes("MKT LOCAL NORTH") ||
      uL07.startsWith("MKT LOCAL") ||
      uL07.includes("MKT_LOCAL_NORTH") ||
      uL07 === "MKT"
    ) {
      return;
    }

    const rawDuration =
      row["durationHours"] ??
      row["duration"] ??
      row["DURATION"] ??
      row["HOURS"] ??
      0;
    const durationHours = row["durationHours"] !== undefined
      ? parseMoney(row["durationHours"])
      : parseDurationToHours(rawDuration);
    const salary = durationHours > 0
      ? durationHours * 20000
      : parseMoney(
          row["calculatedSalary"] ||
            row["TOTAL PAYMENT"] ||
            row["TOTAL"] ||
            row["totalPayment"] ||
            0,
        );
    const bu =
      resolvedCenter.business ||
      row["bu"] ||
      row["Business"] ||
      row["business"] ||
      getBusinessFromL07(l07);
    const month = getMasterRosterMonth(row, activeReportingMonth || "03.2026");
    const rawRowType = row["type"] || row["Type"] || row["LOẠI"] || row["Phân loại"] || row["Nghiệp vụ"] || "MKT LOCAL";
    let rowType = formatPivotTypeHeader(rawRowType);

    // Prevent Roster rows from accidentally writing into Gross Pay columns (e.g. LXO, EC)
    if (rowType === "LXO" || rowType === "EC" || rowType === "PT-DEMO" || rowType === "OTHER") {
      rowType = "MKT LOCAL";
    }

    if (salary > 0 && l07 && rowType !== "EXCLUDE") {
      addAmount(bu, l07, month, rowType, salary, center);
    }
  });

  const priorityOrder = [
    "MKT LOCAL",
    "LXO",
    "EC",
    "PT-DEMO",
    "LDEC01",
    "LDEM01",
    "LPAR01",
    "LRET01",
    "MOTH01",
    "RENEWAL PROJECTS",
    "DISCOVERY CAMP",
    "SUMMER OUTING",
    "SUMMER INSTRUCTORS",
    "EXTRA SUMMER INSTRUCTORS",
    "OTHER",
  ];

  const sortedTypes = Array.from(uniqueTypes).sort((a, b) => {
    if (a === "UNSPECIFIED") return 1;
    if (b === "UNSPECIFIED") return -1;
    const idxA = priorityOrder.indexOf(a);
    const idxB = priorityOrder.indexOf(b);
    if (idxA !== -1 && idxB !== -1) return idxA - idxB;
    if (idxA !== -1) return -1;
    if (idxB !== -1) return 1;
    return a.localeCompare(b);
  });

  return { groupedData: newGroupedData, typeColumns: sortedTypes };
}
