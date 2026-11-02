import { runTimesheetImport } from "../lib/utils/timesheet-import-queue";
import { useCallback, useEffect, useLayoutEffect, useRef } from "react";
import { toast } from "sonner";
import { useAppData } from "../lib/contexts/AppDataContext";
import { fetchGoogleSheetAsFile } from "../lib/utils/data-utils";
import { parseExcelInWorker } from "../lib/utils/excel-worker-client";
import { applyTimesheetLinkResult, resolveTimesheetLink } from "../lib/utils/timesheet-link-sync";
import { formatTimesheetSyncDate } from "../lib/utils/timesheet-sync-date";
import { getCenterInfoByL07 } from "../lib/utils/center-utils";
import { prepareTimesheetRosterReplacement } from "../lib/utils/timesheet-roster-utils";

// Shared by Settings Actions and the missing-center refresh controls.
const pendingSources = new Set<string>();
const useCommittedDataEffect = typeof window === "undefined" ? useEffect : useLayoutEffect;

export function useTimesheetLinkSync() {
  const { appData, updateAppData } = useAppData();
  const { Timesheet_InputList } = appData;
  const latestData = useRef(appData);
  useCommittedDataEffect(() => { latestData.current = appData; }, [appData]);
  const syncRow = useCallback(async (
    id: string,
    urlOverride?: string,
    customUploadDate?: string,
    silentSuccess = false,
  ): Promise<number | null> => {
    const originalInput = Timesheet_InputList.find((row) => row.id === id);
    const input = originalInput && { ...originalInput, url: urlOverride || originalInput.url };
    if (!input) {
      toast.error("Không tìm thấy dòng tương ứng.");
      return null;
    }
    if (pendingSources.has(id)) return null;
    pendingSources.add(id);
    return runTimesheetImport(async () => {
      if (!silentSuccess) updateAppData((previous) => ({
        ...previous,
        Timesheet_InputList: previous.Timesheet_InputList.map((row) => row.id === id ? {
          ...row, status: "processing", ...(urlOverride ? { url: urlOverride } : {}),
        } : row),
      }), false);
      try {
        const source = await resolveTimesheetLink({ ...input, url: urlOverride || input.url });
        const file = await fetchGoogleSheetAsFile(source.url, input.sheetName || "Sheet1");
        if (!file) throw new Error("Không lấy được nội dung file.");
        const parsed = await parseExcelInWorker(file, { fileId: id, mode: "roster" });
        const current = latestData.current;
        const currentInput = current.Timesheet_InputList.find(row => row.id === id);
        const prepared = await prepareTimesheetRosterReplacement(current.Timesheet_Roster || [], parsed.rows, {
          sourceRowIds: [id, ...(currentInput?.legacyRowIds || [])],
          targetL07: input.l07,
          targetAeCode: input.aeCode || getCenterInfoByL07(input.l07)?.aeCode || "",
        });
        const syncedAt = formatTimesheetSyncDate(new Date());
        updateAppData((previous) => applyTimesheetLinkResult(previous, input, parsed.rows, {
          url: source.url,
          fileName: file.name,
          date: customUploadDate || source.date || syncedAt,
          lastSyncedAt: syncedAt,
        }, prepared), false);
        if (!silentSuccess) toast.success(`Đã đồng bộ ${input.l07}: ${parsed.rows.length} dòng (Đã ghi đè dữ liệu cũ).`);
        return parsed.rows.length;
      } catch (error) {
        updateAppData((previous) => ({
          ...previous,
          Timesheet_InputList: previous.Timesheet_InputList.map((row) => row.id === id ? { ...row, status: "error" } : row),
        }), false);
        toast.error(`${input.l07}: ${error instanceof Error ? error.message : String(error)}`);
        return null;
      } finally {
        pendingSources.delete(id);
      }
    });
  }, [Timesheet_InputList, updateAppData]);
  return { syncRow };
}
