/* eslint-disable @typescript-eslint/no-explicit-any */
import * as XLSX from "xlsx";
import { mapExcelRosterRow } from "../lib/utils/roster-row-utils";
import { mapL07 } from '../lib/utils/center-utils';
import {
  createStableTimesheetRowId,
  dedupeTimesheetRosterRows,
} from "../lib/utils/timesheet-roster-utils";

export type ExcelDataKind = "roster" | "salary" | "staff" | "cache" | "raw";
export type ExcelParseMode = "auto" | "roster" | "raw";

export interface ExcelParseResult {
  kind: ExcelDataKind;
  rows: Record<string, unknown>[];
}

interface ExcelWorkerRequest {
  requestId: string;
  fileBuffer: ArrayBuffer;
  fileName: string;
  fileId?: string;
  mode?: ExcelParseMode;
}

function findHeaderRow(rows: any[][]) {
  const isHeaderRow = (row: any[]) =>
    row.some((cell) => {
      const value = String(cell ?? "").toLowerCase().trim();
      if (!value) return false;
      return (
        value === "no" ||
        value === "no." ||
        value === "stt" ||
        value === "id" ||
        value === "nv" ||
        value === "mã nv" ||
        value === "tên" ||
        value === "họ tên" ||
        value === "họ và tên" ||
        value === "nhân viên" ||
        value === "employee" ||
        value === "time" ||
        value === "hours" ||
        value === "duration" ||
        value.includes("giờ làm") ||
        value.includes("số giờ") ||
        value.includes("mã ae") ||
        value.includes("account") ||
        value === "center" ||
        value === "charge to center" ||
        value === "charge to center mkt" ||
        value === "s code" ||
        value === "class" ||
        value.includes("class name") ||
        value === "lớp" ||
        value === "ngày" ||
        value === "date" ||
        value === "cơ sở" ||
        value === "location" ||
        value === "id number" ||
        value === "from" ||
        value === "to" ||
        value === "start" ||
        value === "end" ||
        value === "task" ||
        value === "activity" ||
        value === "session" ||
        value.includes("phòng") ||
        value === "name" ||
        value === "full name" ||
        value === "code" ||
        value === "ma nv" ||
        value === "teacher"
      );
    });

  for (let index = 0; index < Math.min(50, rows.length); index++) {
    if (isHeaderRow(rows[index] || [])) return index;
  }
  return -1;
}

function normalizeSheetName(sheetName: string): string {
  return String(sheetName || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toUpperCase()
    .replace(/[^A-Z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function isRosterSheetName(sheetName: string): boolean {
  const normalized = normalizeSheetName(sheetName);
  return (
    normalized === "ROSTER" ||
    normalized === "Q ROSTER" ||
    normalized.startsWith("Q ROSTER ") ||
    normalized.startsWith("ROSTER ")
  );
}

function isGenericDataSheetName(sheetName: string): boolean {
  const normalized = normalizeSheetName(sheetName);
  return normalized.includes("DU LIEU") || normalized.includes("DATA");
}

function isMktLocalNorthFileName(fileName: string): boolean {
  const normalized = normalizeSheetName(fileName);
  return (
    normalized.includes("MKT LOCAL NORTH") ||
    normalized.includes("NORTH MKT")
  );
}

export function parseExcelData(
  fileBuffer: ArrayBuffer,
  fileName: string,
): Record<string, unknown>[] {
  const name = fileName || "unknown.xlsx";
  if (!fileBuffer) throw new Error("No file data provided to worker.");

  const lowerName = name.toLowerCase();
  const isCsv =
    lowerName.endsWith(".csv") ||
    lowerName.endsWith(".gsheet") ||
    lowerName.endsWith(".txt");
  const workbook = isCsv
    ? XLSX.read(new TextDecoder("utf-8").decode(fileBuffer), {
        type: "string",
        cellDates: true,
        raw: true,
      })
    : XLSX.read(fileBuffer, {
        type: "array",
        cellDates: true,
        raw: true,
        dense: true,
      });

  const rosterSheets = workbook.SheetNames.filter(isRosterSheetName);
  const hasRosterTab = rosterSheets.length > 0;
  const hasGenericDataTab = workbook.SheetNames.some(isGenericDataSheetName);
  const preferredRosterSheet = rosterSheets.find(
    (sheetName) => normalizeSheetName(sheetName) === "Q ROSTER",
  ) || rosterSheets.find(
    (sheetName) => normalizeSheetName(sheetName) === "ROSTER",
  ) || rosterSheets[0];
  const sheetsToProcess = hasRosterTab
    ? [preferredRosterSheet]
    : hasGenericDataTab
      ? workbook.SheetNames.filter(isGenericDataSheetName)
      : workbook.SheetNames;

  const allRows: Record<string, unknown>[] = [];
  for (const sheetName of sheetsToProcess) {
    const worksheet = workbook.Sheets[sheetName];
    if (!worksheet) continue;

    const reference = worksheet["!ref"];
    const decodedRange = reference ? XLSX.utils.decode_range(reference) : null;
    const previewRange = decodedRange
      ? {
          s: decodedRange.s,
          e: {
            r: Math.min(decodedRange.e.r, decodedRange.s.r + 49),
            c: decodedRange.e.c,
          },
        }
      : undefined;
    const previewRows = XLSX.utils.sheet_to_json<any[]>(worksheet, {
      header: 1,
      defval: "",
      blankrows: false,
      raw: false,
      range: previewRange,
    });

    let headerRowIndex = findHeaderRow(previewRows);
    if (headerRowIndex === -1) {
      const normalizedSheetName = sheetName.toLowerCase();
      const looksLikeDataSheet =
        normalizedSheetName.includes("roster") ||
        normalizedSheetName.includes("lịch") ||
        normalizedSheetName.includes("data") ||
        normalizedSheetName.includes("sheet") ||
        normalizedSheetName.includes("thống kê") ||
        normalizedSheetName.includes("salary") ||
        normalizedSheetName.includes("staff") ||
        normalizedSheetName.includes("báo cáo") ||
        normalizedSheetName.includes("danh sách") ||
        workbook.SheetNames.length === 1;

      if (looksLikeDataSheet) {
        headerRowIndex = previewRows.findIndex(
          (row, index) =>
            index < 20 &&
            (row || []).filter((cell) => String(cell ?? "").trim() !== "")
              .length >= 3,
        );
      }
    }

    if (headerRowIndex === -1 && workbook.SheetNames.length > 1) continue;

    const parsedRows = XLSX.utils.sheet_to_json<Record<string, unknown>>(
      worksheet,
      {
        defval: "",
        blankrows: false,
        raw: false,
        range:
          headerRowIndex === -1
            ? undefined
            : (decodedRange?.s.r || 0) + headerRowIndex,
      },
    );
    const isMktLocalNorthRoster =
      isMktLocalNorthFileName(fileName) && isRosterSheetName(sheetName);

    parsedRows.forEach((row) => {
      const cleaned: Record<string, unknown> = {};
      for (const [key, value] of Object.entries(row)) {
        if (!key.toLowerCase().includes("__empty") && String(value ?? "").trim() !== "") {
          cleaned[key] = value;
        }
      }

      if (Object.keys(cleaned).length > 0) {
        if (isMktLocalNorthRoster) {
          const keys = Object.keys(cleaned);
          
          // Rename 'code' to 'Type' if 'type' doesn't exist
          const hasType = keys.some(k => k.toLowerCase().trim() === "type");
          if (!hasType) {
            const codeKey = keys.find(k => k.toLowerCase().trim() === "code");
            if (codeKey) {
              cleaned["Type"] = cleaned[codeKey];
              delete cleaned[codeKey];
            }
          }

          // In MKT Local North ROSTER/Q_ROSTER, CENTER and CHARGE TO
          // CENTER are aliases for the allocation code that resolves to L07.
          const hasChargeToCenter = keys.some(k => k.toLowerCase().trim().includes("charge to center") ||
            ['chargetocenter', 'chargetocentermkt', 'chargetocentermktname'].includes(k.toLowerCase().replace(/[\s_]+/g, '')));
          if (!hasChargeToCenter) {
            const centerKey = keys.find(k => k.toLowerCase().trim() === "center");
            const classKey = keys.find(k => k.toLowerCase().trim() === "class");
            if (centerKey && mapL07(String(cleaned[centerKey])) === 'MKT LOCAL NORTH' && classKey) {
              // Center identifies the source group; Class is the destination.
              cleaned['Charge to Center MKT'] = cleaned[classKey];
            } else if (centerKey) {
              cleaned["Charge to Center"] = cleaned[centerKey];
              delete cleaned[centerKey];
            } else {
              // Updated North MKT Roster uses Class for the charge destination.
              // Keep the original Class and all older allocation aliases intact.
              if (classKey) cleaned["Charge to Center MKT"] = cleaned[classKey];
            }
          }
        }
        allRows.push(cleaned);
      }
    });
  }

  if (allRows.length === 0) {
    throw new Error("File trống hoặc không tìm thấy dòng Tiêu đề hợp lệ.");
  }
  return allRows;
}

function classifyRows(
  rows: Record<string, unknown>[],
  fileName: string,
): Exclude<ExcelDataKind, "raw"> {
  const fileNameLower = fileName.toLowerCase();
  if (fileNameLower.includes("salary")) return "salary";
  if (fileNameLower.includes("staff")) return "staff";
  if (fileNameLower.includes("cache")) return "cache";

  const headers = Object.keys(rows[0] || {}).map((key) =>
    key.toLowerCase().trim(),
  );
  if (headers.includes("academic price") || headers.includes("s code")) {
    return "salary";
  }
  if (headers.includes("bank account number")) return "staff";
  if (headers.includes("today")) return "cache";
  return "roster";
}

export function prepareExcelResult(
  rawRows: Record<string, unknown>[],
  fileName: string,
  fileId?: string,
  mode: ExcelParseMode = "auto",
): ExcelParseResult {
  if (mode === "raw") return { kind: "raw", rows: rawRows };
  const kind = mode === "roster" ? "roster" : classifyRows(rawRows, fileName);
  if (kind === "roster") {
    const mappedRows = dedupeTimesheetRosterRows(
      rawRows.map((row) => mapExcelRosterRow(row, fileName, fileId)),
    );
    return {
      kind,
      rows: mappedRows.map((row) => ({
        ...row,
        _uuid: createStableTimesheetRowId(row, fileId || fileName),
      })),
    };
  }

  return {
    kind,
    rows: rawRows.map((row) => ({
      ...row,
      _sourceFile: fileName,
      _rowId: fileId || row._rowId || "",
    })),
  };
}

if (typeof window === "undefined" && typeof self !== "undefined") {
  self.onmessage = async (event: MessageEvent<ExcelWorkerRequest>) => {
    const {
      requestId,
      fileBuffer,
      fileName,
      fileId,
      mode = "auto",
    } = event.data;
    try {
      const rawRows = parseExcelData(fileBuffer, fileName);
      const result = prepareExcelResult(rawRows, fileName, fileId, mode);
      self.postMessage({ requestId, success: true, result });
    } catch (error: any) {
      self.postMessage({
        requestId,
        success: false,
        error: error?.message || String(error),
      });
    }
  };
}
