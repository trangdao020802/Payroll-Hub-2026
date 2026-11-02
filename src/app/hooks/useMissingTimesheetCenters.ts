import { useCallback, useMemo, useRef, useState } from "react";
import { toast } from "sonner";
import { useAppData } from "../lib/contexts/AppDataContext";
import { getMissingTimesheetCenters, type TimesheetCoverageDateRange } from "../lib/utils/timesheet-center-coverage";
import { parseTimesheetSyncDate } from "../lib/utils/timesheet-sync-date";
import { useTimesheetLinkSync } from "./useTimesheetLinkSync";
import { withTimesheetImportBatch } from "../lib/utils/timesheet-import-queue";

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
  const coverage = useMemo(
    () => enabled ? getMissingTimesheetCenters(appData.Timesheet_InputList, rows, mode, {
      rows: appData.Timesheet_Roster || [], from, to, preferredYear,
    }) : { centers: [], inputs: [], expectedCount: 0 },
    [enabled, appData.Timesheet_InputList, appData.Timesheet_Roster, rows, mode, from, to, preferredYear],
  );
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

  return { ...coverage, isRefreshing, refreshMissing, latestSyncAt };
}

export type MissingTimesheetCenters = ReturnType<typeof useMissingTimesheetCenters>;
