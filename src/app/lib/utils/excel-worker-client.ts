import { generateUUID, getExcelFileBuffer } from "./data-utils";
import ExcelWorker from "../../workers/excelParser.worker?worker";
import type {
  ExcelParseMode,
  ExcelParseResult,
} from "../../workers/excelParser.worker";

type PendingExcelRequest = {
  resolve: (result: ExcelParseResult) => void;
  reject: (error: Error) => void;
  timer: ReturnType<typeof setTimeout>;
  timeoutMs: number;
  result?: ExcelParseResult;
  chunkId?: number;
};

let excelWorker: Worker | null = null;
const pendingExcelRequests = new Map<string, PendingExcelRequest>();

function releaseExcelWorker() {
  if (!excelWorker) return;
  excelWorker.onmessage = null;
  excelWorker.onerror = null;
  excelWorker.onmessageerror = null;
  excelWorker.terminate();
  excelWorker = null;
}

function failExcelWorker(message: string) {
  const error = new Error(message);
  pendingExcelRequests.forEach(({ reject, timer }) => {
    clearTimeout(timer);
    reject(error);
  });
  pendingExcelRequests.clear();
  releaseExcelWorker();
}

function resetDeadline(request: PendingExcelRequest) {
  clearTimeout(request.timer);
  request.timer = setTimeout(() => failExcelWorker(
    "Đọc file quá lâu hoặc worker đã dừng. Dữ liệu cũ được giữ nguyên; hãy thử cập nhật lại link.",
  ), request.timeoutMs);
}

function getExcelWorker() {
  if (excelWorker) return excelWorker;
  excelWorker = new ExcelWorker();
  excelWorker.onmessage = (event: MessageEvent) => {
    const requestId = String(event.data?.requestId || "");
    const pending = pendingExcelRequests.get(requestId);
    if (!pending) return;
    resetDeadline(pending);
    if (event.data?.type === "excel-result-start") {
      pending.result = { kind: event.data.kind, rows: [] };
      pending.chunkId = 0;
      return;
    }
    if (event.data?.type === "excel-result-chunk") {
      if (!pending.result || event.data.chunkId !== (pending.chunkId || 0) + 1 || !Array.isArray(event.data.rows)) {
        failExcelWorker("Dữ liệu file truyền không đầy đủ. Vui lòng thử đồng bộ lại.");
        return;
      }
      for (const row of event.data.rows) pending.result.rows.push(row);
      pending.chunkId = event.data.chunkId;
      const worker = excelWorker;
      // Let input, paint and other queued work run before accepting more rows.
      setTimeout(() => {
        if (worker && excelWorker === worker && pendingExcelRequests.has(requestId)) {
          worker.postMessage({ type: "excel-result-ack", requestId, chunkId: event.data.chunkId });
        }
      }, 0);
      return;
    }
    clearTimeout(pending.timer);
    pendingExcelRequests.delete(requestId);
    // Parsing expands a workbook substantially. Release that worker heap
    // before the calculation worker receives the imported rows.
    if (pendingExcelRequests.size === 0) releaseExcelWorker();
    if (event.data?.success) {
      pending.resolve(pending.result || event.data.result as ExcelParseResult);
    } else {
      pending.reject(
        new Error(event.data?.error || "Không thể đọc dữ liệu Excel."),
      );
    }
  };
  excelWorker.onerror = (event) => {
    failExcelWorker(event.message || "Excel Worker đã dừng bất thường.");
  };
  excelWorker.onmessageerror = () => failExcelWorker("Không thể nhận dữ liệu từ Excel Worker. Vui lòng thử lại.");
  return excelWorker;
}

export const parseExcelInWorker = async (
  file: File,
  options: { fileId?: string; mode?: ExcelParseMode; timeoutMs?: number } = {},
): Promise<ExcelParseResult> => {
  const { buffer, name } = await getExcelFileBuffer(file);
  const requestId = generateUUID();
  const worker = getExcelWorker();

  return new Promise((resolve, reject) => {
    const pending = { resolve, reject, timeoutMs: options.timeoutMs ?? 120_000 } as PendingExcelRequest;
    resetDeadline(pending);
    pendingExcelRequests.set(requestId, pending);
    try {
      worker.postMessage(
        {
          requestId,
          fileBuffer: buffer,
          fileName: name,
          fileId: options.fileId,
          mode: options.mode || "auto",
          resultBackpressure: true,
        },
        [buffer],
      );
    } catch (error) {
      clearTimeout(pending.timer);
      pendingExcelRequests.delete(requestId);
      if (pendingExcelRequests.size === 0) releaseExcelWorker();
      reject(error instanceof Error ? error : new Error(String(error)));
    }
  });
};
