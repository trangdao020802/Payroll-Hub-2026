/* eslint-disable @typescript-eslint/no-explicit-any */
import * as XLSX from "xlsx";
import { csvParseRows } from "d3";
import { mapExcelRosterRow } from "../lib/utils/roster-row-utils";
import { mapL07 } from '../lib/utils/center-utils';
import {
  createStableTimesheetRowId,
  dedupeTimesheetRosterRows,
  getTimesheetRosterBusinessKey,
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
  resultBackpressure?: boolean;
  rosterCsvStreaming?: boolean;
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
        dense: true,
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
        blankrows: false,
        raw: false,
        range:
          headerRowIndex === -1
            ? undefined
            : (decodedRange?.s.r || 0) + headerRowIndex,
      },
    );
    // The formatted rows own their values. Let the workbook cells go before
    // constructing the roster, particularly for large Google Sheet exports.
    delete workbook.Sheets[sheetName];
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

function finishRosterRows(mappedRows: Record<string, unknown>[], sourceId: string): ExcelParseResult {
  const rows = dedupeTimesheetRosterRows(mappedRows);
  for (const row of rows) {
    row._uuid = createStableTimesheetRowId(row, sourceId);
  }
  return { kind: "roster", rows };
}

// Split only between CSV records. Quoted commas, CRLF and multiline notes
// stay intact for the same CSV parser used by the small-file path.
function* csvRows(text: string): Generator<string[]> {
  let start = 0;
  let recordCount = 0;
  let quoted = false;
  let fieldStart = true;
  for (let index = 0; index < text.length; index++) {
    const char = text[index];
    if (quoted) {
      if (char === '"') {
        if (text[index + 1] === '"') index++;
        else quoted = false;
      }
    } else if (char === '"' && fieldStart) {
      quoted = true;
      fieldStart = false;
    } else if (char === '\r' || char === '\n') {
      if (char === '\r' && text[index + 1] === '\n') index++;
      fieldStart = true;
      if (++recordCount === 500) {
        yield* csvParseRows(text.slice(start, index + 1), row => row);
        start = index + 1;
        recordCount = 0;
      }
    } else {
      fieldStart = char === ',';
    }
  }
  if (start < text.length) yield* csvParseRows(text.slice(start), row => row);
}

function csvHeaders(cells: string[]): string[] {
  const used = new Set<string>();
  return cells.map(name => {
    if (!name || name.toLowerCase().includes("__empty")) return "";
    let key = name;
    for (let suffix = 1; used.has(key); suffix++) key = `${name}_${suffix}`;
    used.add(key);
    return key;
  });
}

function* parseRosterCsv(fileBuffer: ArrayBuffer, fileName: string, fileId?: string): Generator<Record<string, unknown>> {
  const csv = new TextDecoder("utf-8").decode(fileBuffer);
  if (/^sep=[^,]/i.test(csv)) return;
  let headers: string[] | undefined;
  let previewRemaining = 50;

  for (const cells of csvRows(csv)) {
    if (!headers) {
      if (!cells.some(cell => cell.trim())) continue;
      if (findHeaderRow([cells]) === 0) {
        const keys = csvHeaders(cells);
        if (keys.filter(Boolean).length >= 2) headers = keys;
      }
      if (!headers && --previewRemaining === 0) return;
      continue;
    }

    const row: Record<string, unknown> = Object.create(null);
    for (let index = 0; index < Math.min(cells.length, headers.length); index++) {
      const key = headers[index];
      const value = cells[index];
      if (key && value.trim()) row[key] = value.replace(/\r\n/g, "\n");
    }
    if (Object.keys(row).length) yield mapExcelRosterRow(row, fileName, fileId);
  }
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
    return finishRosterRows(
      rawRows.map((row) => mapExcelRosterRow(row, fileName, fileId)),
      fileId || fileName,
    );
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
  type Chunk = { rows: ExcelParseResult['rows']; keys?: string[] };
  const outputs = new Map<string, { chunks: Iterator<Chunk>; chunkId: number }>();

  function* resultChunks(rows: ExcelParseResult['rows']): Generator<Chunk> {
    for (let offset = 0; offset < rows.length; offset += 1000) {
      const end = Math.min(offset + 1000, rows.length);
      const batch = rows.slice(offset, end);
      for (let index = offset; index < end; index++) delete rows[index];
      yield { rows: batch };
    }
  }

  function* csvChunks(rows: Iterable<Record<string, unknown>>, sourceId: string): Generator<Chunk> {
    let chunk: Required<Chunk> = { rows: [], keys: [] };
    for (const row of rows) {
      row._uuid = createStableTimesheetRowId(row, sourceId);
      chunk.keys.push(getTimesheetRosterBusinessKey(row));
      chunk.rows.push(row);
      if (chunk.rows.length === 500) {
        yield chunk;
        chunk = { rows: [], keys: [] };
      }
    }
    if (chunk.rows.length) yield chunk;
  }

  const fail = (requestId: string, error: unknown) => {
    outputs.delete(requestId);
    self.postMessage({ requestId, success: false, error: error instanceof Error ? error.message : String(error) });
  };
  const sendNext = (requestId: string) => {
    const output = outputs.get(requestId);
    if (!output) return;
    try {
      const chunk = output.chunks.next();
      if (chunk.done) {
        outputs.delete(requestId);
        self.postMessage({ type: "excel-result-complete", requestId, success: true });
      } else {
        self.postMessage({ type: "excel-result-chunk", requestId, chunkId: ++output.chunkId, ...chunk.value });
      }
    } catch (error) {
      fail(requestId, error);
    }
  };
  const startOutput = (requestId: string, kind: ExcelDataKind, chunks: Iterator<Chunk>) => {
    outputs.set(requestId, { chunks, chunkId: 0 });
    self.postMessage({ type: "excel-result-start", requestId, kind });
    sendNext(requestId);
  };
  self.onmessage = (event: MessageEvent<ExcelWorkerRequest & { type?: string; chunkId?: number }>) => {
    if (event.data.type === "excel-result-ack") {
      const output = outputs.get(event.data.requestId);
      if (output?.chunkId === event.data.chunkId) sendNext(event.data.requestId);
      return;
    }
    const {
      requestId,
      fileBuffer,
      fileName,
      fileId,
      mode = "auto",
    } = event.data;
    try {
      let result: ExcelParseResult | undefined;
      if (mode === "roster" && /\.(csv|txt|gsheet)$/i.test(fileName)) {
        const rest = parseRosterCsv(fileBuffer, fileName, fileId);
        const first = rest.next();
        if (!first.done) {
          const sourceId = fileId || fileName;
          if (event.data.resultBackpressure && event.data.rosterCsvStreaming) {
            function* allRows() { yield first.value; yield* rest; }
            startOutput(requestId, "roster", csvChunks(allRows(), sourceId));
            return;
          }
          result = finishRosterRows([first.value, ...rest], sourceId);
        }
      }
      result ??= prepareExcelResult(parseExcelData(fileBuffer, fileName), fileName, fileId, mode);
      if (event.data.resultBackpressure) {
        startOutput(requestId, result.kind, resultChunks(result.rows));
      } else {
        self.postMessage({ requestId, success: true, result });
      }
    } catch (error) {
      fail(requestId, error);
    }
  };
}
