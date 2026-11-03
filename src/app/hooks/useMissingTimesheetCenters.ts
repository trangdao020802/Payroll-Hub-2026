import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { toast } from "sonner";
import { useAppData } from "../lib/contexts/AppDataContext";
import {
  CENTER_COVERAGE_BATCH_ROWS,
  getMissingTimesheetCenters,
  getMissingTimesheetCentersInBatches,
  type TimesheetCoverageDateRange,
} from "../lib/utils/timesheet-center-coverage";
import { parseTimesheetSyncDate } from "../lib/utils/timesheet-sync-date";
import { useTimesheetLinkSync } from "./useTimesheetLinkSync";
import { withTimesheetImportBatch } from "../lib/utils/timesheet-import-queue";

type CenterCoverage = ReturnType<typeof getMissingTimesheetCenters>;
const emptyCoverage: CenterCoverage = { centers: [], inputs: [], expectedCount: 0 };

export function useMissingTimesheetCenters(
  rows: readonly Record<string, unknown>[],
  mode: "l07" | "allocation" = "l07",
  { from, to, preferredYear }: TimesheetCoverageDateRange = {},
  enabled = true,
) {
  const { appData } = useAppData();
  const { syncRow } = useTimesheetLinkSync();
  const [refreshing, setRefreshing] = useState(false);
  const running = useRef(false);
  const inputs = appData.Timesheet_InputList;
  const roster = appData.Timesheet_Roster;
  const query = useMemo(() => {
    if (!enabled) return null;
    const source = { rows: roster || [], from, to, preferredYear };
    const scanSize = rows.length + (from || to ? source.rows.length : 0);
    const immediate = scanSize < CENTER_COVERAGE_BATCH_ROWS
      ? getMissingTimesheetCenters(inputs, rows, mode, source)
      : undefined;
    return { key: Symbol(), inputs, rows, mode, source, immediate };
  }, [enabled, inputs, roster, rows, mode, from, to, preferredYear]);
  const [completed, setCompleted] = useState<{ key: symbol; coverage: CenterCoverage }>();
  const isChecking = !!query && !query.immediate && completed?.key !== query.key;
  const coverage = query?.immediate ?? (completed && completed.key === query?.key ? completed.coverage : emptyCoverage);

  useEffect(() => {
    if (!query || query.immediate) return;
    const controller = new AbortController();
    void getMissingTimesheetCentersInBatches(controller.signal, query.inputs, query.rows, query.mode, query.source)
      .then(coverage => {
        if (coverage && !controller.signal.aborted) setCompleted({ key: query.key, coverage });
      });
    return () => controller.abort();
  }, [query]);

  const latestSyncAt = useMemo(() => {
    let latestLabel = "";
    let latestTime = Number.NEGATIVE_INFINITY;
    for (const input of appData.Timesheet_InputList || []) {
      const label = String(input.lastSyncedAt || "").trim();
      if (!label) continue;
      const parsed = parseTimesheetSyncDate(label);
      const time = parsed?.getTime();
      if (time !== undefined && time !== null && Number.isFinite(time) && time > latestTime) {
        latestTime = time;
        latestLabel = label;
      }
    }
    return latestLabel;
  }, [appData.Timesheet_InputList]);

  // Persisted Settings statuses may outlive a request and do not indicate that
  // this table's refresh button is running. Only an explicit click starts it.
  const isRefreshing = refreshing;

  const refreshMissing = useCallback(async () => {
    if (running.current || isRefreshing || coverage.inputs.length === 0) return;
    running.current = true;
    setRefreshing(true);
    let successCount = 0;
    let emptyCount = 0;
    try {
      // Capture just this table's missing sources and preserve each successful
      // update if another source fails. Sequential fetches avoid Google throttling.
      await withTimesheetImportBatch(async () => {
        for (const input of coverage.inputs) {
          const count = await syncRow(input.id, undefined, undefined, true);
          if (count !== null) successCount++;
          if (count === 0) emptyCount++;
        }
      });
      const message = `Đã làm mới ${successCount}/${coverage.inputs.length} link Center còn thiếu.`;
      if (successCount === coverage.inputs.length) toast.success(message);
      else toast.warning(message);
      if (emptyCount) toast.info(`Đã bỏ qua ${emptyCount} link đồng bộ thành công nhưng không có dòng dữ liệu.`);
    } finally {
      running.current = false;
      setRefreshing(false);
    }
  }, [coverage.inputs, isRefreshing, syncRow]);

  return { ...coverage, isChecking, isRefreshing, refreshMissing, latestSyncAt };
}

export type MissingTimesheetCenters = ReturnType<typeof useMissingTimesheetCenters>;
