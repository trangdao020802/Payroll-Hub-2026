/* eslint-disable @typescript-eslint/no-explicit-any */
import { useMemo, useState, useEffect, useRef, useSyncExternalStore, useCallback } from "react";
import { generateUUID, parseAnyDate } from "../lib/utils/data-utils";
import { mapL07 } from "../lib/utils/center-utils";
import { TASK_COLUMNS } from "../constants/timesheet-logic";
import { useAppData } from "../lib/contexts/AppDataContext";
import { useUiSettings } from "../lib/ui-settings";
import { isTimesheetImporting, subscribeTimesheetImports } from "../lib/utils/timesheet-import-queue";
import TimesheetWorker from "../workers/timesheet.worker?worker";

const emptyResult = (isCalculating: boolean) => ({
  processedRosterData: [], employeeSummary: [], centerSummary: [], isCalculating,
});

export function useTimesheetCalculations(
  rosterData: any[],
  salaryScaleData: any[],
  staffData: any[],
  cacheData: any[],
  fromDateStr: string,
  toDateStr: string,
  enabled = true,
  includeProcessedRosterData = true,
) {
  const { appData } = useAppData();
  const uiSettings = useUiSettings();
  const importing = useSyncExternalStore(subscribeTimesheetImports, isTimesheetImporting, () => false);

  const [result, setResult] = useState<any>({
    processedRosterData: [],
    employeeSummary: [],
    centerSummary: [],
    isCalculating: true,
  });

  const workerRef = useRef<Worker | null>(null);
  const sourceRef = useRef<{ rosterData: any[]; salaryScaleData: any[]; staffData: any[]; cacheData: any[]; ready: boolean } | null>(null);
  const [retryKey, setRetryKey] = useState(0);
  const retry = useCallback(() => setRetryKey(key => key + 1), []);
  useEffect(() => () => {
    workerRef.current?.terminate(); workerRef.current = null; sourceRef.current = null;
  }, []);

  const fromDateVal = fromDateStr;
  let tempFDate: Date | null = null;
  if (fromDateVal) {
    const parts = fromDateVal.split("-");
    if (parts.length === 3) {
      tempFDate = new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10), 0, 0, 0);
    }
  }
  const preferredYear = uiSettings.defaultAuditYear || (tempFDate ? tempFDate.getFullYear() : new Date().getFullYear());

  const checkTAsData = appData?.Q_CheckTAs;
  // Status/progress/URL changes do not change payroll calculations. Use a
  // primitive key so they cannot terminate/restart a worker with the same data.
  const timesheetInputMetadataKey = JSON.stringify((appData?.Timesheet_InputList || []).map((row: any) => ({
    id: row.id, l07: row.l07, aeCode: row.aeCode, bus: row.bus, legacyRowIds: row.legacyRowIds || [],
  })));
  const timesheetInputMetadata = useMemo(
    () => JSON.parse(timesheetInputMetadataKey), [timesheetInputMetadataKey],
  );

  const { classSizeMap, checkTAsMap } = useMemo(() => {
    // Class-size inputs are Timesheet-owned. Payroll Master data must never
    // participate in a Timesheet calculation.
    const csMap: Record<string, number> = {};
    const ctaMap: Record<string, number> = {};
    if (!enabled || importing) return { classSizeMap: csMap, checkTAsMap: ctaMap };

    const safeCheckTAsData = checkTAsData || [];
    safeCheckTAsData.forEach((row: any) => {
      const clsName = String(row["Class Name"] || row["Lớp"] || row["Class"] || row["Mã lớp"] || "").trim();
      const centerRaw = String(row["Center Name"] || row["Mã AE"] || row["Center"] || row["Center Code"] || row["L07"] || row["Trung tâm"] || "");
      const sessionDate = row["Session Date"] || row["Ngày"] || row["Date"] || row["Ngày học"] || row["Session"] || row["SessionDate"];
      const numStudents = parseInt(String(row["Number of Student"] || row["Number of Students"] || row["No of Student"] || row["Sĩ số"] || row["Sỹ số"] || row["Students"] || row["Số HV"] || row["Số học viên"] || row["Sĩ số lớp"] || row["Total Students"] || row["Số lượng học viên"] || row["Sĩ số thực tế"] || row["Sỹ số thực tế"] || row["Actual Size"] || row["Class Size"] || row["Size"] || row["Số lượng"] || row["Sĩ số cơ sở"] || ""), 10) || 0;

      const parsedDate = parseAnyDate(sessionDate, preferredYear);
      const normCls = clsName.replace(/\s+/g, "").toUpperCase();
      const centerL07 = mapL07(centerRaw);
      const normCenter = centerL07.replace(/\s+/g, "").toUpperCase();

      if (numStudents > 0 && normCls) {
        const classKey = `${normCenter}_${normCls}`;
        if (!csMap[classKey] || csMap[classKey] < numStudents) csMap[classKey] = numStudents;
      }
      if (parsedDate && clsName) {
        const dateStr = `${String(parsedDate.getDate()).padStart(2, "0")}/${String(parsedDate.getMonth() + 1).padStart(2, "0")}/${parsedDate.getFullYear()}`;
        const key = `${normCenter}_${normCls}_${dateStr}`;
        ctaMap[key] = numStudents;
      }
    });
    return { classSizeMap: csMap, checkTAsMap: ctaMap };
  }, [enabled, importing, checkTAsData, preferredYear]);

  useEffect(() => {
    let inputSendTimer: ReturnType<typeof setTimeout> | undefined;
    let calculationTimer: ReturnType<typeof setTimeout> | undefined;

    if (!enabled || importing || rosterData.length === 0) {
      workerRef.current?.terminate(); workerRef.current = null; sourceRef.current = null;
      // Release detailed rows while a source workbook is fetched/parsed, and
      // when this consumer is hidden. No global cache pins the old dataset.
      let cancelled = false;
      Promise.resolve().then(() => {
        if (!cancelled) setResult(emptyResult(enabled && (importing || rosterData.length > 0)));
      });
      return () => { cancelled = true; };
    }

    let cancelled = false;
    Promise.resolve().then(() => {
      if (!cancelled) setResult(emptyResult(true));
    });
    const customSalaryRates = appData?.customSalaryRates;
    const customTypeRates = appData?.customTypeRates;

    const params = {
      rosterData,
      salaryScaleData,
      staffData,
      cacheData,
      fromDateStr,
      toDateStr,
      appData: {
        Timesheet_InputList: timesheetInputMetadata,
        Timesheet_RosterFileName: appData?.Timesheet_RosterFileName,
        customSalaryRates,
        customTypeRates,
      },
      preferredYear,
      includeProcessedRosterData,
      checkTAsMap,
      classSizeMap,
      TASK_COLUMNS,
      customSalaryRates,
      customTypeRates,
    };

    const sameSource = sourceRef.current?.ready && sourceRef.current.rosterData === rosterData
      && sourceRef.current.salaryScaleData === salaryScaleData && sourceRef.current.staffData === staffData
      && sourceRef.current.cacheData === cacheData;
    if (!sameSource) {
      workerRef.current?.terminate(); workerRef.current = null; sourceRef.current = null;
    }

    const startWorker = () => {
      try {
        const reused = !!workerRef.current && !!sourceRef.current?.ready;
        const worker = workerRef.current || new TimesheetWorker();
        workerRef.current = worker;
        const requestId = generateUUID();
        const chunkedResult: Record<string, any[]> = {
          processedRosterData: [],
          employeeSummary: [],
          centerSummary: [],
        };
        let chunkedError: string | undefined;

        const commitWorkerResult = (resultData: any) => {
          const finalResult = {
            processedRosterData: resultData.processedRosterData || [],
            employeeSummary: resultData.employeeSummary || [],
            centerSummary: resultData.centerSummary || [],
            isCalculating: false,
            error: resultData.error,
          };
          worker.onmessage = null;
          clearTimeout(calculationTimer);
          if (!cancelled) setResult(finalResult);
        };

        worker.onmessage = (e) => {
          if (cancelled || workerRef.current !== worker) return;
          const resultData = e.data || {};
          if (resultData.requestId && resultData.requestId !== requestId) return;
          armCalculationDeadline();

          if (resultData.type === "timesheet-result-start") {
            chunkedError = resultData.error;
            return;
          }

          if (resultData.type === "timesheet-result-chunk") {
            const target = chunkedResult[resultData.field];
            if (target && Array.isArray(resultData.rows)) {
              target.push(...resultData.rows);
            }
            if (typeof resultData.chunkId === "number") {
              setTimeout(() => {
                if (!cancelled && workerRef.current === worker) worker.postMessage({ type: "timesheet-result-ack", requestId, chunkId: resultData.chunkId });
              }, 0);
            }
            return;
          }

          if (resultData.type === "timesheet-result-complete") {
            commitWorkerResult({ ...chunkedResult, error: chunkedError });
            return;
          }

          // Backwards-compatible fallback for an older cached worker bundle.
          if (resultData.error) {
            console.error("Timesheet worker error string:", resultData.error);
          }
          commitWorkerResult(resultData);
        };
        worker.onerror = (err: ErrorEvent | any) => {
          if (cancelled || workerRef.current !== worker) return;
          const errorDetails = err?.message || err?.error?.message;
          console.warn("Timesheet worker warning:", errorDetails);
          worker.onmessage = null;
          worker.onerror = null;
          worker.onmessageerror = null;
          worker.terminate();
          if (workerRef.current === worker) workerRef.current = null;
          sourceRef.current = null;
          clearTimeout(calculationTimer);
          setResult((prev: any) => ({
            ...prev,
            isCalculating: false,
            error:
              errorDetails ||
              "Không thể tính Timesheet trong Worker. Vui lòng tải lại trang và thử lại.",
          }));
        };
        worker.onmessageerror = () => worker.onerror?.call(worker, { message: 'Không thể nhận kết quả Timesheet. Bấm Thử lại để tiếp tục.' } as ErrorEvent);
        const inputFields = [
          "rosterData",
          "salaryScaleData",
          "staffData",
          "cacheData",
        ] as const;
        const workerParams = {
          ...params,
          resultBackpressure: true,
          rosterData: undefined,
          salaryScaleData: undefined,
          staffData: undefined,
          cacheData: undefined,
        };
        const armCalculationDeadline = () => {
          clearTimeout(calculationTimer);
          calculationTimer = setTimeout(() => {
            if (cancelled || workerRef.current !== worker) return;
            worker.terminate(); workerRef.current = null; sourceRef.current = null;
            setResult({ ...emptyResult(false), error: 'Lọc Timesheet quá lâu. Bấm Thử lại để tiếp tục.' });
          }, 120000);
        };
        armCalculationDeadline();
        if (reused) {
          worker.postMessage({ type: 'timesheet-recalculate', requestId, params: workerParams });
          return;
        }
        sourceRef.current = { rosterData, salaryScaleData, staffData, cacheData, ready: false };
        // Source rows are sent once per dataset, in bounded clone turns.
        // Period/rate changes reuse the worker's source and only send parameters.
        const inputChunkSize = 2_000;
        let fieldIndex = 0;
        let rowOffset = 0;

        worker.postMessage({
          type: "timesheet-input-start",
          requestId,
          params: workerParams,
        });

        const sendNextInputChunk = () => {
          if (workerRef.current !== worker) return;
          if (fieldIndex >= inputFields.length) {
            worker.postMessage({ type: "timesheet-input-complete", requestId });
            if (sourceRef.current) sourceRef.current.ready = true;
            return;
          }

          const field = inputFields[fieldIndex];
          const rows = params[field] || [];
          if (rowOffset >= rows.length) {
            fieldIndex += 1;
            rowOffset = 0;
            inputSendTimer = setTimeout(sendNextInputChunk, 0);
            return;
          }

          worker.postMessage({
            type: "timesheet-input-chunk",
            requestId,
            field,
            rows: rows.slice(rowOffset, rowOffset + inputChunkSize),
          });
          rowOffset += inputChunkSize;
          inputSendTimer = setTimeout(sendNextInputChunk, 0);
        };
        inputSendTimer = setTimeout(sendNextInputChunk, 0);
      } catch (workerError) {
        workerRef.current?.terminate(); workerRef.current = null; sourceRef.current = null;
        clearTimeout(calculationTimer);
        console.error("Failed to instantiate TimesheetWorker:", workerError);
        setResult((prev: any) => ({
          ...prev,
          isCalculating: false,
          error:
            workerError instanceof Error
              ? workerError.message
              : "Không thể khởi tạo Timesheet Worker.",
        }));
      }

    };
    // Coalesce successive source updates; a folder sync must not calculate
    // every intermediate version of the complete roster.
    const startTimer = setTimeout(startWorker, 250);
    return () => {
      cancelled = true;
      clearTimeout(startTimer);
      if (inputSendTimer) clearTimeout(inputSendTimer);
      clearTimeout(calculationTimer);
      if (workerRef.current) {
        workerRef.current.onmessage = null;
        workerRef.current.onerror = null;
        workerRef.current.onmessageerror = null;
        workerRef.current.postMessage({ type: 'timesheet-cancel-result' });
      }
    };
  }, [
    enabled,
    importing,
    retryKey,
    includeProcessedRosterData,
    rosterData,
    salaryScaleData,
    staffData,
    cacheData,
    fromDateStr,
    toDateStr,
    classSizeMap,
    checkTAsMap,
    checkTAsData,
    appData?.Timesheet_RosterFileName,
    appData?.customSalaryRates,
    appData?.customTypeRates,
    timesheetInputMetadata,
    timesheetInputMetadataKey,
    preferredYear,
  ]);

  return { ...result, retry };
}
