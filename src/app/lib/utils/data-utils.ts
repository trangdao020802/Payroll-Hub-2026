import { withDeadline } from "../cloud/deadline";

export function parseMoneyToNumber(val: unknown): number {
  if (typeof val === "number") return Number.isFinite(val) ? val : 0;
  if (typeof val === "bigint") return Number(val);
  if (val instanceof Date) {
    // A calendar value is never a payroll amount. Converting it back to an
    // Excel serial leaks dates such as 20/07/2026 into charge columns as 46223.
    return 0;
  }
  if (val === null || val === undefined || val === "") return 0;

  const source = String(val).trim();
  if (!source) return 0;

  // Dates returned by the Excel worker are serialized to ISO strings. They
  // must never be stripped down to digits and treated as payroll amounts.
  if (
    /^\d{4}-\d{2}-\d{2}(?:T\d{2}:\d{2}:\d{2}(?:\.\d+)?Z)?$/i.test(
      source,
    ) ||
    /^\d{1,2}[/-]\d{1,2}[/-]\d{4}$/.test(source)
  ) {
    return 0;
  }

  // Preserve scientific notation when Excel/browser already supplied it.
  if (/^[+-]?\d+(?:\.\d+)?[eE][+-]?\d+$/.test(source)) {
    const scientificValue = Number(source);
    return Number.isFinite(scientificValue) ? scientificValue : 0;
  }

  const isNegativeByParentheses = /^\(.*\)$/.test(source);
  let numericText = source
    .replace(/[\s\u00a0]/g, "")
    .replace(/[^\d.,+-]/g, "");

  const isNegative = isNegativeByParentheses || numericText.includes("-");
  numericText = numericText.replace(/[+-]/g, "");
  if (!numericText || !/\d/.test(numericText)) return 0;

  const dotCount = (numericText.match(/\./g) || []).length;
  const commaCount = (numericText.match(/,/g) || []).length;

  if (dotCount > 0 && commaCount > 0) {
    // When both separators exist, the last one is the decimal separator and
    // every earlier separator is a thousands separator.
    const decimalIndex = Math.max(
      numericText.lastIndexOf("."),
      numericText.lastIndexOf(","),
    );
    const integerPart = numericText.slice(0, decimalIndex).replace(/[.,]/g, "");
    const decimalPart = numericText.slice(decimalIndex + 1).replace(/[.,]/g, "");
    numericText = decimalPart
      ? `${integerPart}.${decimalPart}`
      : integerPart;
  } else if (dotCount > 0 || commaCount > 0) {
    const separator = dotCount > 0 ? "." : ",";
    const parts = numericText.split(separator);
    const lastPart = parts[parts.length - 1];
    const separatorIsDecimal =
      lastPart.length > 0 &&
      (lastPart.length <= 2 ||
        (parts.length === 2 &&
          lastPart.length === 3 &&
          parts[0].length > 3));

    numericText = separatorIsDecimal
      ? `${parts.slice(0, -1).join("")}.${lastPart}`
      : parts.join("");
  }

  const parsedValue = Number(numericText);
  if (!Number.isFinite(parsedValue)) return 0;
  return isNegative ? -Math.abs(parsedValue) : parsedValue;
}
export function formatNumber(
  val: unknown,
  type: "string" | "number" | "money" | "date" = "number",
): string {
  if (val === null || val === undefined || val === "") return "";
  if (type === "string") return String(val);
  if (type === "date") {
    const dateValue = val instanceof Date ? val : parseAnyDate(String(val));
    if (!dateValue || isNaN(dateValue.getTime())) return String(val);
    return dateValue.toLocaleDateString("vi-VN", {
      day: "2-digit",
      month: "2-digit",
      year: "numeric",
    });
  }

  const numericValue = parseMoneyToNumber(val);
  const rounded = Math.round(numericValue);
  return new Intl.NumberFormat("vi-VN", {
    minimumFractionDigits: 0,
    maximumFractionDigits: 0,
  }).format(rounded);
}
export function formatMoneyVND(val: unknown): string {
  const n = parseMoneyToNumber(val);
  const rounded = Math.round(n);
  return rounded.toLocaleString("vi-VN", {
    minimumFractionDigits: 0,
    maximumFractionDigits: 0,
  });
}
export function removeVietnameseTones(str: string): string {
  if (!str) return '';
  return str.normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/đ/g, "d").replace(/Đ/g, "D");
}

function normalizeColumnRuleKey(columnKey: string): string {
  return removeVietnameseTones(String(columnKey || ""))
    .replace(/([a-z0-9])([A-Z])/g, "$1 $2")
    .toUpperCase()
    .replace(/[^A-Z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

export function isNonSummableTextColumn(columnKey: string): boolean {
  const normalized = normalizeColumnRuleKey(columnKey);
  const compact = normalized.replace(/\s+/g, "");
  return (
    compact.includes("CHARGETOCENTER") ||
    compact.includes("CHARGECENTER") ||
    compact.includes("CENTER") ||
    compact === "L07" ||
    normalized.includes("CHARGE TO CENTER") ||
    normalized.includes("CHARGE CENTER") ||
    normalized.includes("CENTER") ||
    normalized.includes("L07") ||
    normalized === "THANG" ||
    normalized.includes("THANG PHAT SINH") ||
    normalized.includes("THANG BAO CAO") ||
    normalized.includes("THANG HOLD") ||
    normalized.includes("THANG THANH TOAN") ||
    normalized.includes("THANH TOAN VAO THANG") ||
    normalized.includes("CAC THANG") ||
    normalized.includes("KY BAO CAO") ||
    normalized.includes("KY DANG THEO DOI") ||
    normalized.includes("REPORTING MONTH") ||
    normalized.includes("ARISING MONTH") ||
    normalized.includes("SALARY SCALE")
  );
}

export function isChargeAmountColumn(columnKey: string): boolean {
  const normalized = normalizeColumnRuleKey(columnKey);
  const compact = normalized.replace(/\s+/g, "");

  if (
    compact.includes("CHARGETOCENTER") ||
    compact.includes("CHARGECENTER") ||
    compact.includes("CHARGETYPE") ||
    compact.includes("CHARGECODE") ||
    compact.includes("CENTER") ||
    compact === "L07" ||
    normalized.includes("CHARGE TYPE") ||
    normalized.includes("CHARGE TO CENTER") ||
    normalized.includes("CHARGE CENTER") ||
    normalized.includes("CHARGE CODE") ||
    normalized.includes("CENTER") ||
    normalized.includes("L07")
  ) {
    return false;
  }

  return (
    normalized.includes("CHARGE") ||
    /^(LDEC|LDEM|LPAR|LRET|MOTH)\d*/.test(compact) ||
    normalized.includes("EXTRA SUMMER INSTRUCTORS")
  );
}
export function formatIdNumber(id: unknown): string {
  if (id === undefined || id === null || id === "") return "";

  let normalized = String(id).trim().replace(/[\s\u00a0]/g, "");

  // Excel can expose an ID as scientific notation or as a decimal ending in
  // zero. Convert those forms back to their integer representation first.
  if (/^[+]?\d+(?:\.\d+)?[eE][+-]?\d+$/.test(normalized)) {
    const numericId = Number(normalized);
    if (Number.isSafeInteger(numericId) && numericId >= 0) {
      normalized = numericId.toLocaleString("fullwide", {
        useGrouping: false,
        maximumFractionDigits: 0,
      });
    }
  } else if (/^\d+\.0+$/.test(normalized)) {
    normalized = normalized.replace(/\.0+$/, "");
  } else if (/^\+\d+$/.test(normalized)) {
    normalized = normalized.slice(1);
  }

  return /^\d+$/.test(normalized) && normalized.length < 12
    ? normalized.padStart(12, "0")
    : normalized;
}
export function prepareDataForExport<T>(data: T[]): T[] {
  return data;
}
export function parseAnyDate(value: unknown, preferredYear?: number): Date | null {
  if (value === null || value === undefined || value === "") return null;
  if (value instanceof Date) {
    return isNaN(value.getTime()) ? null : new Date(value.getTime());
  }

  if (typeof value === "number" && Number.isFinite(value)) {
    // Excel serial date (1900 date system). Day-only headers are resolved by
    // their owning import flow instead of being treated as serial dates.
    if (value > 59 && value < 100000) {
      const excelEpoch = Date.UTC(1899, 11, 30);
      const parsed = new Date(excelEpoch + value * 86400000);
      return isNaN(parsed.getTime()) ? null : parsed;
    }
    return null;
  }

  const raw = String(value).trim();
  if (!raw) return null;

  // Recover Date objects serialized by the legacy roster mapper as
  // "undefined/undefined/Mon Jun 01 2026 ...". Existing browser data may
  // retain that shape even after the importer is upgraded.
  const legacyDateTail = raw.match(/^undefined\/undefined\/(.+)$/i);
  if (legacyDateTail) {
    const recovered = new Date(legacyDateTail[1]);
    return isNaN(recovered.getTime()) ? null : recovered;
  }

  if (/(?:^|[/\s])undefined(?:$|[/\s])/i.test(raw)) return null;

  const dmy = raw.match(/^(\d{1,2})[./-](\d{1,2})(?:[./-](\d{2}|\d{4}))$/);
  if (dmy) {
    const day = Number(dmy[1]);
    const month = Number(dmy[2]);
    let year = Number(dmy[3]);
    if (year < 100) year += 2000;
    const parsed = new Date(year, month - 1, day);
    return parsed.getFullYear() === year &&
      parsed.getMonth() === month - 1 &&
      parsed.getDate() === day
      ? parsed
      : null;
  }

  // Support MM.YYYY / MM/YYYY / MM-YYYY (e.g. 04.2026, 12.2025)
  const my = raw.match(/^(\d{1,2})[./-](\d{4})$/);
  if (my) {
    const month = Number(my[1]);
    const year = Number(my[2]);
    if (month >= 1 && month <= 12 && year >= 1900 && year <= 2200) {
      return new Date(year, month - 1, 1);
    }
  }

  // Support YYYY.MM / YYYY/MM / YYYY-MM (e.g. 2026-04, 2026.04)
  const ym = raw.match(/^(\d{4})[./-](\d{1,2})$/);
  if (ym) {
    const year = Number(ym[1]);
    const month = Number(ym[2]);
    if (month >= 1 && month <= 12 && year >= 1900 && year <= 2200) {
      return new Date(year, month - 1, 1);
    }
  }

  const dayOnly = raw.match(/^(\d{1,2})$/);
  if (dayOnly && preferredYear) return null;

  const parsed = new Date(raw);
  return isNaN(parsed.getTime()) ? null : parsed;
}

export function isDateColumn(columnKey?: unknown, columnLabel?: unknown, type?: unknown): boolean {
  if (type === "date") return true;
  const normalized = removeVietnameseTones(
    `${String(columnKey || "")} ${String(columnLabel || "")}`,
  )
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
  return /(^| )(date|ngay|full date|session date|thang|month|period|ky bao cao)( |$)/.test(normalized);
}

export function normalizeDateFilterValue(value: unknown): string {
  if (value === null || value === undefined || value === "") return "";
  const raw = String(value).trim();
  if (!raw) return "";

  const isRecoverableLegacyDate = /^undefined\/undefined\//i.test(raw);
  if (!isRecoverableLegacyDate && /undefined|null|invalid date/i.test(raw)) {
    return "";
  }

  // Preserve the calendar day encoded by Excel/worker ISO strings. Formatting
  // an ISO midnight through the runtime timezone can otherwise shift it to the
  // previous day in the filter menu.
  const isoDate = raw.match(/^(\d{4})-(\d{2})-(\d{2})(?:T|$)/);
  if (isoDate) return `${isoDate[3]}/${isoDate[2]}/${isoDate[1]}`;

  const parsed = parseAnyDate(value);
  if (!parsed) return raw;
  return parsed.toLocaleDateString("vi-VN", {
    day: "2-digit",
    month: "2-digit",
    year: "numeric",
    timeZone: "Asia/Bangkok",
  });
}
const normalizedRowKeyCache = new WeakMap<object, Map<string, string>>();

function normalizeLookupKey(value: unknown): string {
  return String(value ?? "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[Đđ]/g, "D")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

/**
 * Read the first matching value from an imported row.
 *
 * Excel/Google Sheet headers are user-controlled, so callers can provide a
 * list of aliases. Matching is accent/case/spacing insensitive and the
 * normalized key map is cached per row to avoid repeatedly scanning every
 * header while processing large workbooks.
 */
export function getVal(
  row: Record<string, unknown> | null | undefined,
  keyOrAliases: string | readonly string[],
): unknown {
  if (!row || typeof row !== "object") return null;

  const aliases = Array.isArray(keyOrAliases)
    ? keyOrAliases
    : [keyOrAliases];
  let firstDefinedValue: unknown = null;

  for (const alias of aliases) {
    if (Object.prototype.hasOwnProperty.call(row, alias)) {
      const value = row[alias];
      if (value !== undefined && value !== null && value !== "") return value;
      if (firstDefinedValue === null && value !== undefined) {
        firstDefinedValue = value;
      }
    }
  }

  let normalizedKeys = normalizedRowKeyCache.get(row);
  if (!normalizedKeys) {
    normalizedKeys = new Map<string, string>();
    Object.keys(row).forEach((rowKey) => {
      const normalized = normalizeLookupKey(rowKey);
      if (normalized && !normalizedKeys!.has(normalized)) {
        normalizedKeys!.set(normalized, rowKey);
      }
    });
    normalizedRowKeyCache.set(row, normalizedKeys);
  }

  for (const alias of aliases) {
    const actualKey = normalizedKeys.get(normalizeLookupKey(alias));
    if (!actualKey) continue;
    const value = row[actualKey];
    if (value !== undefined && value !== null && value !== "") return value;
    if (firstDefinedValue === null && value !== undefined) {
      firstDefinedValue = value;
    }
  }

  return firstDefinedValue;
}

export const AUDIT_RAW_TYPE_ALIASES = [
  "type",
  "TYPE",
  "task type",
  "task code",
  "activity code",
  "event type",
  "tk_type",
  "code",
  "loại",
  "loại công việc",
] as const;

export function normalizeAuditRawType(value: unknown): string {
  return String(value || "")
    .trim()
    .toUpperCase()
    .replace(/[‐‑‒–—−]/g, "-")
    .replace(/\s*-\s*/g, "-")
    .replace(/\s+/g, " ");
}

export function getAuditRawType(
  row: Record<string, unknown> | null | undefined,
): string {
  return String(getVal(row, AUDIT_RAW_TYPE_ALIASES) || "").trim();
}

export function isAuditInClassType(value: unknown): boolean {
  const normalized = normalizeAuditRawType(value);
  return normalized === "IN-CLASS" || normalized === "IN-CLASS ATLS";
}

export function parseTimeStrToHours(timeValue: unknown): number {
  if (timeValue === null || timeValue === undefined || timeValue === "") return 0;

  if (timeValue instanceof Date) {
    if (Number.isNaN(timeValue.getTime())) return 0;
    return (
      timeValue.getHours() +
      timeValue.getMinutes() / 60 +
      timeValue.getSeconds() / 3600
    );
  }

  if (typeof timeValue === "number" && Number.isFinite(timeValue)) {
    // XLSX can expose a clock value either as an Excel day fraction (0.375)
    // or as an hour value (9). Normalise both to hours-of-day.
    const dayFraction = Math.abs(timeValue) < 1
      ? timeValue
      : timeValue >= 24
        ? timeValue - Math.floor(timeValue)
        : null;
    return dayFraction === null ? timeValue : dayFraction * 24;
  }

  const source = String(timeValue).trim();
  if (!source) return 0;

  const numeric = Number(source.replace(",", "."));
  if (Number.isFinite(numeric) && !source.includes(":")) {
    return parseTimeStrToHours(numeric);
  }

  const match = source.match(/^(\d{1,2}):([0-5]?\d)(?::([0-5]?\d))?\s*(AM|PM)?$/i);
  if (!match) return 0;

  let hours = Number(match[1]);
  const minutes = Number(match[2]);
  const seconds = Number(match[3] || 0);
  const meridiem = match[4]?.toUpperCase();
  if (meridiem === "PM" && hours < 12) hours += 12;
  if (meridiem === "AM" && hours === 12) hours = 0;
  return hours + minutes / 60 + seconds / 3600;
}
export async function getExcelFileBuffer(
  file: File,
): Promise<{ buffer: ArrayBuffer; name: string }> {
  if (!file) {
    throw new Error("Không tìm thấy thông tin file để đọc.");
  }

  const buffer = await file.arrayBuffer();
  if (file.name.toLowerCase().endsWith(".gsheet")) {
    const pointerText = new TextDecoder("utf-8").decode(buffer).trim();
    if (!pointerText) {
      throw new Error(`File Google Sheet đang trống (${file.name}).`);
    }
    if (looksLikeHtmlPayload(pointerText)) {
      throw new Error(
        `File ${file.name} chứa trang HTML thay vì dữ liệu Google Sheet. Vui lòng tải lại file từ Google Drive hoặc kiểm tra quyền chia sẻ.`,
      );
    }

    let source = "";
    if (pointerText.startsWith("{")) {
      let parsedPointer: unknown;
      try {
        parsedPointer = JSON.parse(pointerText);
      } catch {
        throw new Error(`File liên kết Google Sheet không hợp lệ (${file.name}).`);
      }

      if (parsedPointer && typeof parsedPointer === "object" && !Array.isArray(parsedPointer)) {
        const pointer = parsedPointer as Record<string, unknown>;
        const url = String(pointer.url || "").trim();
        const documentId = String(pointer.doc_id || "").trim();
        const resourceId = String(pointer.resource_id || "")
          .replace(/^spreadsheet:/i, "")
          .trim();
        source =
          url ||
          (documentId
            ? `https://docs.google.com/spreadsheets/d/${documentId}`
            : resourceId
              ? `https://docs.google.com/spreadsheets/d/${resourceId}`
              : "");
      }
    } else if (/^https:\/\/docs\.google\.com\/spreadsheets\//i.test(pointerText)) {
      source = pointerText;
    }

    if (source) {
      try {
        const downloadedFile = await fetchGoogleSheetAsFile(source, file.name);
        return {
          buffer: await downloadedFile.arrayBuffer(),
          name: downloadedFile.name,
        };
      } catch (downloadErr: unknown) {
        const msg =
          downloadErr instanceof Error
            ? downloadErr.message
            : "Không thể tải Google Sheet";
        throw new Error(`${msg} (${file.name})`);
      }
    }

    if (pointerText.startsWith("{")) {
      throw new Error(
        `File ${file.name} không chứa đường dẫn Google Sheet hợp lệ.`,
      );
    }

    return {
      buffer,
      name: file.name.replace(/\.gsheet$/i, ".csv"),
    };
  }

  return {
    buffer,
    name: file.name,
  };
}

export function looksLikeHtmlPayload(value: string): boolean {
  const preview = String(value || "")
    .replace(/^\uFEFF/, "")
    .trimStart()
    .slice(0, 2048);
  return /^(?:<!--[\s\S]{0,512}?-->\s*)?<\s*(?:!doctype\s+html|html|head|body|script|title|meta|link)\b/i.test(
    preview,
  );
}
export function formatTime12Hour(timeStr: unknown): string {
  return String(timeStr);
}
export const COMMON_FIELD_ALIASES: Record<string, string[]> = {
  No: ["STT", "NO", "NUMBER", "SỐ THỨ TỰ"],
  "ID Number": [
    "ID",
    "MÃ NV",
    "CMND",
    "CCCD",
    "MÃ NHÂN VIÊN",
    "EMPLOYEE ID",
    "MÃ SỐ",
    "ID NUMBER",
  ],
  "Full Name": [
    "NAME",
    "TÊN",
    "HỌ VÀ TÊN",
    "TÊN NHÂN VIÊN",
    "FULL NAME",
    "TEACHER",
    "GIÁO VIÊN",
  ],
  "Full name": [
    "NAME",
    "TÊN",
    "HỌ VÀ TÊN",
    "TÊN NHÂN VIÊN",
    "FULL NAME",
  ],
  "Salary Scale": [
    "SCALE",
    "MỨC LƯƠNG",
    "RANK",
    "BẬC LƯƠNG",
    "SALARY RANK",
  ],
  From: [
    "FROM",
    "TỪ",
    "TỪ NGÀY",
    "START DATE",
    "NGÀY BẮT ĐẦU",
    "START",
    "DATE FROM",
    "FROM DATE",
  ],
  To: [
    "TO",
    "ĐẾN",
    "ĐẾN NGÀY",
    "END DATE",
    "NGÀY KẾT THÚC",
    "END",
    "DATE TO",
    "TO DATE",
  ],
  "Bank Account Number": [
    "ACCOUNT",
    "TÀI KHOẢN",
    "STK",
    "SỐ TÀI KHOẢN",
    "BANK ACCOUNT",
  ],
  "Bank Name": [
    "BANK NAME",
    "NGÂN HÀNG",
    "TÊN NGÂN HÀNG",
    "TEN NGAN HANG",
  ],
  "CITAD code": ["CITAD", "MÃ CITAD", "CITAD CODE"],
  "TAX CODE": ["TAX", "MST", "MÃ SỐ THUẾ", "TAX CODE"],
  "Contract No": [
    "CONTRACT",
    "HỢP ĐỒNG",
    "SỐ HỢP ĐỒNG",
    "CONTRACT NO",
  ],
  "CHARGE TO LXO": ["LXO", "CHARGE LXO", "CHARGE TO LXO"],
  "CHARGE TO EC": ["EC", "CHARGE EC", "CHARGE TO EC"],
  "CHARGE TO PT-DEMO": [
    "PT-DEMO",
    "CHARGE PT-DEMO",
    "CHARGE TO PT-DEMO",
  ],
  "Charge MKT Local": [
    "MKT",
    "MKT LOCAL",
    "CHARGE MKT LOCAL",
    "CHARGE TO MKT LOCAL",
    "CHARGE MKT",
    "CHARGE TO CENTER MKT",
  ],
  "CHARGE TO OTHER": ["CHARGE OTHER", "CHARGE TO OTHER", "OTHER"],
  "Charge Renewal Projects": [
    "RENEWAL",
    "RENEWAL PROJECTS",
    "CHARGE TO RENEWAL PROJECTS",
  ],
  "Charge Discovery Camp": [
    "DISCOVERY",
    "DISCOVERY CAMP",
    "CHARGE TO DISCOVERY CAMP",
  ],
  "Charge Summer Outing": [
    "SUMMER OUTING",
    "CHARGE TO SUMMER OUTING",
  ],
  "Charge Summer Instructors": [
    "SUMMER INSTRUCTORS",
    "CHARGE TO SUMMER INSTRUCTORS",
  ],
  "Extra Summer Instructors": [
    "EXTRA SUMMER INSTRUCTORS",
    "CHARGE TO EXTRA SUMMER INSTRUCTORS",
    "EXTRA INSTRUCTOR",
    "EXTRA INSTRUCTORS",
    "EXTRA SUMMER INSTRUCTOR",
    "EXTRA INSTRUCTOR BONUS",
    "SUMMER INSTRUCTORS BONUS",
    "BONUS",
  ],
  "TOTAL PAYMENT": [
    "TOTAL",
    "TỔNG",
    "THỰC NHẬN",
    "TỔNG THANH TOÁN",
    "TOTAL PAYMENT",
    "NET PAY",
    "AMOUNT",
  ],
  Center: [
    "CENTER",
    "COST CENTER",
    "TRUNG TÂM",
    "AE CODE",
    "AE",
    "MÃ AE",
    "MÃ TT",
    "MÃ TRUNG TÂM",
    "L07",
  ],
  Business: ["BUSINESS", "KHỐI", "BUS", "BỘ PHẬN", "BU"],
  Type: ["TYPE", "EVENT TYPE", "CLASS TYPE", "LOẠI LỚP", "LOẠI"],
  Class: [
    "CLASS",
    "CLASS NAME",
    "TÊN LỚP",
    "LỚP",
    "MÃ LỚP",
    "CLASS CODE",
  ],
  Date: ["DATE", "NGÀY", "DATE OF CLASS", "NGÀY DẠY", "SCHEDULE DATE"],
  Duration: ["DURATION", "HOURS", "SỐ GIỜ", "GIỜ", "TOTAL HOURS"],
};

const normalizeHeaderForMatching = (value: string): string =>
  removeVietnameseTones(String(value || ""))
    .toUpperCase()
    .replace(/[^A-Z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();

const getHeaderTokenScore = (left: string, right: string): number => {
  const leftTokens = new Set(left.split(" ").filter(Boolean));
  const rightTokens = new Set(right.split(" ").filter(Boolean));
  if (leftTokens.size === 0 || rightTokens.size === 0) return 0;
  let intersectionSize = 0;
  leftTokens.forEach((token) => {
    if (rightTokens.has(token)) intersectionSize++;
  });
  return (
    (intersectionSize / (leftTokens.size + rightTokens.size - intersectionSize)) *
    100
  );
};

export function scoreMatch(
  header: string,
  target: string,
  aliases: string[] = [],
): number {
  const normalizedHeader = normalizeHeaderForMatching(header);
  const normalizedTarget = normalizeHeaderForMatching(target);
  if (!normalizedHeader || !normalizedTarget) return 0;
  if (normalizedHeader === normalizedTarget) return 100;

  const normalizedAliases = aliases.map(normalizeHeaderForMatching);
  if (normalizedAliases.includes(normalizedHeader)) return 95;

  if (normalizedHeader.includes(normalizedTarget)) {
    return 85;
  }

  if (
    normalizedAliases.some(
      (alias) =>
        alias &&
        normalizedHeader.includes(alias),
    )
  ) {
    return 80;
  }

  const candidateScores = [
    getHeaderTokenScore(normalizedHeader, normalizedTarget),
    ...normalizedAliases.map((alias) =>
      getHeaderTokenScore(normalizedHeader, alias),
    ),
  ];
  const bestTokenScore = Math.max(...candidateScores);
  return bestTokenScore >= 60 ? Math.min(79, bestTokenScore) : 0;
}
export function normalizeId(id: unknown): string { return String(id); }
export function toVietnamDateString(date: Date): string {
  if (!(date instanceof Date) || Number.isNaN(date.getTime())) return "";

  // Timesheet filters compare this value with the ISO date keys produced by
  // the month selector (YYYY-MM-DD). Returning Date#toString here previously
  // produced values such as "Mon Jun 01 2026...", so valid rows were removed
  // by the lexical range comparison and monthly results were incomplete.
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}
export function generateUUID(): string { return Math.random().toString(); }
export async function fetchGoogleSheetAsFile(
  url: string,
  name: string,
  timeoutMs = 90_000,
): Promise<File> {
  return withDeadline(signal => downloadGoogleSheet(url, name, signal),
    "Tải Google Sheet quá lâu. Dữ liệu cũ được giữ nguyên; vui lòng thử cập nhật lại link.", timeoutMs);
}

async function downloadGoogleSheet(
  url: string,
  name: string,
  signal: AbortSignal,
): Promise<File> {
  const sourceUrl = String(url || "").trim();
  if (!sourceUrl) {
    throw new Error("URL Google Sheet đang để trống.");
  }

  let response: Response | null = null;
  let fetchError: Error | null = null;

  try {
    response = await fetch("/api/gs-export", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ url: sourceUrl }),
      signal,
    });
  } catch (err: unknown) {
    fetchError = err instanceof Error ? err : new Error(String(err));
  }

  // If server-side export succeeded:
  if (response && response.ok) {
    const contentType = (
      response.headers.get("content-type") ||
      "text/csv; charset=utf-8"
    ).toLowerCase();
    const encodedSourceName = response.headers.get("x-spreadsheet-name") || "";
    let sourceName = "";
    if (encodedSourceName) {
      try {
        sourceName = decodeURIComponent(encodedSourceName);
      } catch {
        sourceName = encodedSourceName;
      }
    }

    const fileBlob = await response.blob();
    const payloadPreview = await fileBlob.slice(0, 2048).text();
    if (
      contentType.includes("text/html") ||
      looksLikeHtmlPayload(payloadPreview)
    ) {
      throw new Error(
        "Dịch vụ Google Sheet trả về trang HTML thay vì dữ liệu. Vui lòng tải lại trang và kiểm tra quyền chia sẻ file.",
      );
    }

    const isExcel =
      contentType.includes("spreadsheetml") ||
      contentType.includes("application/vnd.ms-excel");
    const extension = isExcel ? ".xlsx" : ".csv";
    const requestedName = String(sourceName || name || "GoogleSheet").trim();
    const finalName = /\.(xlsx?|xls|csv|txt|gsheet)$/i.test(requestedName)
      ? isExcel
        ? requestedName.replace(/\.(csv|txt|gsheet)$/i, extension)
        : requestedName.replace(/\.(xlsx?|xls)$/i, extension)
      : `${requestedName}${extension}`;
    if (fileBlob.size === 0) {
      throw new Error("Google Sheet trả về file rỗng.");
    }

    return new File([fileBlob], finalName, {
      type: isExcel
        ? "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        : "text/csv;charset=utf-8",
    });
  }

  // Fallback: try direct client-side fetch for publicly accessible Google Sheets
  if (signal.aborted) throw fetchError || new Error("Đã dừng tải Google Sheet.");
  const dMatch = sourceUrl.match(/\/d\/([a-zA-Z0-9-_]{15,})/);
  const pubMatch = sourceUrl.match(/\/d\/e\/([a-zA-Z0-9-_]{20,})/);
  let directUrl = "";

  if (pubMatch) {
    directUrl = `https://docs.google.com/spreadsheets/d/e/${pubMatch[1]}/pub?output=csv`;
  } else if (dMatch) {
    const gidMatch = sourceUrl.match(/[#&?]gid=([0-9]+)/);
    const gid = gidMatch ? gidMatch[1] : "0";
    directUrl = `https://docs.google.com/spreadsheets/d/${dMatch[1]}/export?format=csv&gid=${gid}`;
  }

  if (directUrl) {
    try {
      const directRes = await fetch(directUrl, { signal });
      if (directRes.ok) {
        const fileBlob = await directRes.blob();
        if (fileBlob.size && !looksLikeHtmlPayload(await fileBlob.slice(0, 2048).text())) {
          const requestedName = String(name || "GoogleSheet").trim();
          const finalName = requestedName.toLowerCase().endsWith(".csv")
            ? requestedName
            : requestedName.replace(/\.(xlsx?|xls|gsheet|txt)$/i, "") + ".csv";
          return new File([fileBlob], finalName, {
            type: "text/csv;charset=utf-8",
          });
        }
      }
    } catch {
      // Direct client fetch failed (e.g. CORS on private sheet)
    }
  }

  if (response && !response.ok) {
    let message = "";
    try {
      const responseText = await response.text();
      try {
        const parsed = JSON.parse(responseText) as { error?: unknown };
        message = String(parsed.error || responseText);
      } catch {
        message = looksLikeHtmlPayload(responseText)
          ? `Dịch vụ đọc Google Sheet tạm thời không khả dụng (HTTP ${response.status}).`
          : responseText;
      }
    } catch {
      message = `Máy chủ trả về mã lỗi ${response.status}`;
    }
    throw new Error(message || "Không thể tải dữ liệu Google Sheet.");
  }

  if (fetchError) {
    throw new Error(
      `Không thể kết nối đến dịch vụ đọc Google Sheet (${fetchError.message || "Failed to fetch"}). Vui lòng kiểm tra quyền chia sẻ file hoặc cấu hình tài khoản Google.`,
    );
  }

  throw new Error("Không thể tải dữ liệu Google Sheet.");
}
export function isMoneyColumn(col: string): boolean {
  const normalized = normalizeColumnRuleKey(col);
  if (!normalized || isNonSummableTextColumn(col)) return false;

  if (
    normalized.includes("DATE") ||
    normalized.includes("NGAY") ||
    normalized === "FROM" ||
    normalized === "TO" ||
    normalized.includes("ID NUMBER") ||
    normalized.includes("ACCOUNT NUMBER") ||
    normalized.includes("BANK ACCOUNT") ||
    normalized.includes("CENTER") ||
    normalized.includes("CITAD") ||
    normalized.includes("TAX CODE") ||
    normalized.includes("CONTRACT")
  ) {
    return false;
  }

  return (
    isChargeAmountColumn(col) ||
    normalized.includes("TOTAL") ||
    normalized.includes("PAYMENT") ||
    normalized.includes("AMOUNT") ||
    normalized.includes("SALARY") ||
    normalized.includes("LUONG") ||
    normalized.includes("SO TIEN") ||
    normalized.includes("TIEN THUONG") ||
    normalized.includes("BONUS")
  );
}
export async function fetchWithBackoff(
  input: RequestInfo | URL | (() => Promise<Response>),
  init?: RequestInit,
  retries = 3,
): Promise<Response> {
  let attempt = 0;
  while (true) {
    const response =
      typeof input === "function" ? await input() : await fetch(input, init);
    if (
      attempt >= retries ||
      (response.status !== 429 && response.status < 500)
    ) {
      return response;
    }
    await new Promise((resolve) =>
      setTimeout(resolve, Math.min(500 * 2 ** attempt, 4000)),
    );
    attempt += 1;
  }
}

export function getHoldRowAmount(r: Record<string, unknown> | null | undefined): number {
  if (!r || typeof r !== "object") return 0;

  const remainingKeys = [
    "Số dư HOLD còn lại",
    "Số tiền HOLD còn lại",
    "SỐ TIỀN HOLD CÒN LẠI",
    "SỐ DƯ HOLD CÒN LẠI",
    "Số dư còn lại",
    "SỐ DƯ CÒN LẠI",
    "Remaining Balance",
    "Remaining Hold",
    "HOLD còn lại",
    "Hold còn lại",
  ];

  for (const key of remainingKeys) {
    if (r[key] !== undefined && r[key] !== null && String(r[key]).trim() !== "") {
      const val = parseMoneyToNumber(r[key]);
      if (!isNaN(val) && val !== 0) {
        return val;
      }
    }
  }

  return parseMoneyToNumber(
    r["TOTAL PAYMENT"] ||
      r["Payment Amount"] ||
      r["Grand Total"] ||
      r["GRAND TOTAL"] ||
      r["Total Payment"] ||
      0
  );
}
