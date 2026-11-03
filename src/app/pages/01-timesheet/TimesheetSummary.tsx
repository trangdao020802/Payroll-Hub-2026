import { SupabaseSchemaError } from '../../lib/utils/supabase-sync-errors';
import { useTableRestore } from '../../hooks/useTableRestore';
import { TableRestoreButton } from '../../components/TableRestoreButton';
import { chooseExcelExport } from "../../components/ExportScopeDialog";
import React, { useState, useMemo, useEffect, useRef } from "react";
import {
  FileSpreadsheet,
  Download,
  Settings,
  RefreshCw,
  Trash2,
  Copy,
  Plus,
  ArrowLeft,
  ChevronDown,
} from "lucide-react";
import * as XLSX from "xlsx";
import { toast } from "sonner";
import { useAppData } from "../../lib/contexts/AppDataContext";
import { isSupabaseConfigured } from "../../lib/supabase";
import {
  clearSupabaseRosterData,
  preflightTimesheetSync,
  syncRosterToSupabase,
  SQL_SETUP_SCRIPT,
} from "../../lib/supabase-sync-utils";
import { runTimesheetImport, withTimesheetImportBatch } from "../../lib/utils/timesheet-import-queue";
import { useTimesheetCalculations } from "../../hooks/useTimesheetCalculations";
import { getDynamicEmployeeColumns, CENTER_COLUMNS } from "../../constants/timesheet-columns";
import { TimesheetInputTable } from "./components/TimesheetInputTable";
import type { TimesheetInputRow } from "./components/TimesheetInputTable";
import { AppData } from "../../types";
import {
  getL07FromFileName,
  getCenterInfoByL07,
  getBusinessFromL07,
} from "../../lib/utils/center-utils";
import {
  isFileNameStoredAsL07,
  resolveTimesheetCenterFromFileName,
  shouldSkipTimesheetSource,
} from "../../lib/utils/timesheet-input-resolver";
import {
  replaceTimesheetRosterRows,
} from "../../lib/utils/timesheet-roster-utils";
import { formatTimesheetSyncDate } from "../../lib/utils/timesheet-sync-date";
import { 
  generateUUID, 
} from "../../lib/utils/data-utils";
import {
  DropdownMenu,
  DropdownMenuContent,
  DropdownMenuItem,
  DropdownMenuLabel,
  DropdownMenuSeparator,
  DropdownMenuTrigger,
} from "../../components/ui/dropdown-menu";
import { ConfirmDialog } from "../../components/shared/ConfirmDialog";
import {
  Dialog,
  DialogContent,
  DialogHeader,
  DialogTitle,
  DialogDescription,
  DialogFooter,
} from "../../components/ui/dialog";
import { Button } from "../../components/ui/button";
import {
  TableInitialMark,
  TableTitleRemainder,
} from "../../components/TableInitialMark";

import { parseExcelInWorker } from "../../lib/utils/excel-worker-client";
import type { ExcelParseResult } from "../../workers/excelParser.worker";
import { useTimesheetLinkSync } from "../../hooks/useTimesheetLinkSync";
import { DEFAULT_TIMESHEET_FOLDER_URL as DEFAULT_FOLDER_URL, MKT_LOCAL_NORTH_URL, fetchTimesheetFolderFiles } from "../../lib/utils/timesheet-link-sync";

interface TimesheetSummaryPageProps {
  onBack?: () => void;
}

export default function TimesheetSummaryPage({ onBack }: TimesheetSummaryPageProps = {}) {
  const { appData, updateAppData } = useAppData();
  const restoreTable = useTableRestore();

  const [activeTab] = useState<"files">("files");
  const [fromDate] = useState("");
  const [toDate] = useState("");
  const [debouncedFromDate, setDebouncedFromDate] = useState("");
  const [debouncedToDate, setDebouncedToDate] = useState("");

  const [isSyncing, setIsSyncing] = useState(false);
  const [syncProgress, setSyncProgress] = useState(0);
  const [totalSyncRows, setTotalSyncRows] = useState(0);
  const [syncedRowsCount, setSyncedRowsCount] = useState(0);
  const [showSqlDialog, setShowSqlDialog] = useState(false);
  const [syncSchemaError, setSyncSchemaError] = useState('');

  const [isFetchingGgSheet, setIsFetchingGgSheet] = useState(false);
  const urlInputBusy = useRef(false);
  const [bulkUploadProgress, setBulkUploadProgress] = useState<{
    current: number;
    total: number;
  } | null>(null);
  const [, setRefreshKey] = useState(0);
  const [showClearAllDialog, setShowClearAllDialog] = useState(false);
  const [showClearEmptyL07Dialog, setShowClearEmptyL07Dialog] = useState(false);
  const [clearRowTargetId, setClearRowTargetId] = useState<string | null>(null);

  const handleUrlInput = async (id: string, url: string) => {
    if (!url.trim() || urlInputBusy.current) return;
    if (shouldSkipTimesheetSource(url)) {
      toast.info("Hệ thống tự động bỏ qua link MKT HP.");
      return;
    }
    const isFolder = url.includes("folders/") || url.includes("drive/folders/") || url.includes("?id=");

    urlInputBusy.current = true;
    setIsFetchingGgSheet(true);
    try {
      if (isFolder) {
        let folderId = url.trim();
        const match = url.match(/folders\/([a-zA-Z0-9-_]+)/);
        if (match) {
          folderId = match[1];
        } else {
          try {
            const urlObj = new URL(url);
            if (urlObj.searchParams.has("id")) {
              folderId = urlObj.searchParams.get("id") || folderId;
            }
          } catch { /* ignore */ }
        }

        const files = await fetchTimesheetFolderFiles(folderId);
        if (files.length === 0) {
          throw new Error("Không tìm thấy file nào trong thư mục này.");
        }

        const driveFiles = files.filter(f => {
          const name = String(f.name || "").toLowerCase();
          return !name.includes("copy") && !shouldSkipTimesheetSource(
            f.name,
            f.webViewLink,
            f.url,
          );
        });

        if (driveFiles.length === 0 && files.length > 0) {
          throw new Error("Tất cả các file trong thư mục đều là file 'copy' nên hệ thống tự động bỏ qua.");
        }

        toast.info(`Tìm thấy ${driveFiles.length} file hợp lệ. Đang tự động đối chiếu và nạp dữ liệu...`);

        const currentInputs = [...(appData.Timesheet_InputList || [])];
        const configuredInputs = currentInputs.filter(
          (row) => !isFileNameStoredAsL07(row),
        );
        const queuedByRowId = new Map<
          string,
          {
            id: string;
            file: File;
            url: string;
            uploadDate: string;
            modifiedAt: number;
          }
        >();
        let skipCount = 0;

        for (const f of driveFiles) {
          const fileName = String(f.name || "");
          const configuredCenter = resolveTimesheetCenterFromFileName(
            fileName,
            configuredInputs,
          );
          if (!configuredCenter) {
            skipCount++;
            continue;
          }
          const rowId = configuredCenter.id;

          const sheetUrl = `https://docs.google.com/spreadsheets/d/${f.id}`;
          const fileContent = JSON.stringify({ url: sheetUrl });
          const blob = new Blob([fileContent], { type: 'application/json' });
          let name = fileName;
          if (!name.toLowerCase().endsWith(".gsheet")) {
            name = name.replace(/\.(xlsx|xls|csv)$/i, "") + ".gsheet";
          }
          const fileObj = new File([blob], name, { type: 'application/json' });
          const modifiedAt = Date.parse(String(f.modifiedTime || f.createdTime || "")) || 0;
          const modifiedDate = modifiedAt ? new Date(modifiedAt) : new Date();
          const uploadDate = formatTimesheetSyncDate(modifiedDate);
          const queued = queuedByRowId.get(rowId);
          if (!queued || modifiedAt >= queued.modifiedAt) {
            if (queued) skipCount++;
            queuedByRowId.set(rowId, {
              id: rowId,
              file: fileObj,
              url: `${sheetUrl}/edit`,
              uploadDate,
              modifiedAt,
            });
          } else {
            skipCount++;
          }
        }

        const toProcess = Array.from(queuedByRowId.values());
        const successCount = toProcess.length;

        if (successCount > 0) {
          // Set matched rows to a "ready" status first, but don't start processing yet
          updateAppData(prev => ({ ...prev, Timesheet_InputList: prev.Timesheet_InputList.map(r => {
            const match = queuedByRowId.get(r.id);
            if (match) {
              return {
                ...r,
                status: "ready" as const,
                url: match.url,
                fileName: match.file.name,
                date: match.uploadDate,
              };
            }
            return r;
          }) }), false);
          
          let completed = 0;
          await withTimesheetImportBatch(async () => {
            for (let i = 0; i < toProcess.length; i++) {
              const item = toProcess[i];
              try {
                const count = await handleSyncRow(item.id, item.url, item.uploadDate, true);
                if (count !== null) completed++;
              } catch (err: unknown) {
                const msg = err instanceof Error ? err.message : String(err);
                console.error(`Failed to process ${item.file.name}:`, err);
                handleUpdateRow(item.id, "status", "error");
                toast.error(`Lỗi xử lý ${item.file.name}: ${msg}`);
              }

              if (i < toProcess.length - 1) {
                await new Promise(resolve => setTimeout(resolve, 150));
              }
            }
          });
          
          toast.success(`Đã nạp xong từ thư mục! Thành công: ${completed}/${successCount} trung tâm${skipCount > 0 ? `, Bỏ qua: ${skipCount}` : ""}.`);
        } else {
          toast.warning(`Không tìm thấy trung tâm nào khớp với các file trong thư mục.`);
        }
      } else {
        await handleSyncRow(id, url);
      }
    } catch (error: unknown) {
      console.error(error);
      const msg = error instanceof Error ? error.message : "Lỗi xử lý link";
      toast.error(msg);
    } finally {
      urlInputBusy.current = false;
      setIsFetchingGgSheet(false);
    }
  };

  useEffect(() => {
    const timer = setTimeout(() => {
      setDebouncedFromDate(fromDate);
      setDebouncedToDate(toDate);
    }, 500);
    return () => clearTimeout(timer);
  }, [fromDate, toDate]);

  useEffect(() => {
    if (!debouncedFromDate || !debouncedToDate) return;

    updateAppData((prev) => {
      if (
        prev.Timesheet_Dates?.from === debouncedFromDate &&
        prev.Timesheet_Dates?.to === debouncedToDate
      ) {
        return prev;
      }

      return {
        ...prev,
        Timesheet_Dates: { from: debouncedFromDate, to: debouncedToDate },
      };
    }, false);
  }, [debouncedFromDate, debouncedToDate, updateAppData]);



  const rosterData = useMemo(
    () => appData.Timesheet_Roster || [],
    [appData.Timesheet_Roster],
  );
  const salaryScaleData = useMemo(() => appData.Q_Salary_Scale || [], [appData.Q_Salary_Scale]);
  const staffData = useMemo(() => appData.Q_Staff || [], [appData.Q_Staff]);
  const cacheData = useMemo(() => appData.Q_Cache || [], [appData.Q_Cache]);
  const inputRows = useMemo(() => appData.Timesheet_InputList || [
    { id: "1", l07: "", aeCode: "", bus: "", url: "", status: "pending" },
  ], [appData.Timesheet_InputList]);

  // The North.MKT Roster source is fixed by payroll. Repair older persisted
  // rows as well, so users never need to paste the link again.
  useEffect(() => {
    const needsRepair = inputRows.some(
      (row) => String(row.l07 || "").trim().toUpperCase() === "MKT LOCAL NORTH" && row.url !== MKT_LOCAL_NORTH_URL,
    );
    if (!needsRepair) return;
    updateAppData((prev) => ({
      ...prev,
      Timesheet_InputList: (prev.Timesheet_InputList || []).map((row) =>
        String(row.l07 || "").trim().toUpperCase() === "MKT LOCAL NORTH"
          ? { ...row, url: MKT_LOCAL_NORTH_URL }
          : row,
      ),
    }), false);
  }, [inputRows, updateAppData]);

  // Repair rows created by the previous folder-import bug. Their L07 was the
  // filename itself; move the imported file/data back onto the already
  // configured center row and remove only that generated duplicate.
  useEffect(() => {
    updateAppData((prev) => {
      const rows = prev.Timesheet_InputList || [];
      const malformedRows = rows.filter(isFileNameStoredAsL07);
      if (malformedRows.length === 0) return prev;

      const validRows = rows.filter((row) => !isFileNameStoredAsL07(row));
      const replacementByTargetId = new Map<string, (typeof rows)[number]>();
      const removableIds = new Set<string>();

      malformedRows.forEach((row) => {
        if (!row.fileName) return;
        const target = resolveTimesheetCenterFromFileName(row.fileName, validRows);
        if (!target) return;
        const currentTarget = replacementByTargetId.get(target.id) || target;
        replacementByTargetId.set(target.id, {
          ...currentTarget,
          url: row.url || target.url,
          fileName: row.fileName,
          sheetName: row.sheetName || target.sheetName,
          status: row.status,
          count: row.count,
          date: row.date,
          lastSyncedAt: row.lastSyncedAt || target.lastSyncedAt,
          columnMapping: row.columnMapping || target.columnMapping,
          legacyRowIds: Array.from(
            new Set([
              ...(currentTarget.legacyRowIds || []),
              row.id,
            ]),
          ),
        });
        removableIds.add(row.id);
      });

      if (removableIds.size === 0) return prev;
      const nextInputs = rows
        .filter((row) => !removableIds.has(row.id))
        .map((row) => replacementByTargetId.get(row.id) || row);

      return {
        ...prev,
        Timesheet_InputList: nextInputs,
      };
    }, false, true, ["Timesheet_Roster", "Q_Staff", "Q_Salary_Scale", "Q_Cache", "Timesheet_InputList"]);
  }, [updateAppData]);

  const handleAddRow = () => {
    updateAppData((prev) => ({
      ...prev,
      Timesheet_InputList: [
        ...inputRows,
        {
          id: generateUUID(),
          l07: "",
          aeCode: "",
          bus: "",
          url: "",
          status: "pending",
        },
      ],
    }));
  };
  const handleUpdateRow = (
    id: string,
    field: keyof TimesheetInputRow,
    val: string | number | boolean | Record<string, unknown> | undefined,
  ) => {
    updateAppData(
      (prev) => ({
        ...prev,
        Timesheet_InputList: (prev.Timesheet_InputList || []).map((r) => {
          if (r.id === id) {
            const updated = { ...r, [field]: val };
            if (
              (field === "l07" && val === "MKT LOCAL NORTH") ||
              (field === "aeCode" && (val === "MKT LOCAL NORTH" || val === "NTW"))
            ) {
              updated.url = MKT_LOCAL_NORTH_URL;
            }
            if (["l07", "url", "fileName", "sheetName"].some((key) => updated[key] !== r[key])) {
              updated.status = "pending";
              updated.lastSyncedAt = undefined;
            }
            return updated;
          }
          return r;
        }),
      }),
      false,
    );
  };
  const executeClearRow = (id: string) => {
    updateAppData((prev) => {
      const targetRow = (prev.Timesheet_InputList || []).find((r) => r.id === id);
      const ownedRowIds = new Set([id, ...(targetRow?.legacyRowIds || [])]);
      return {
        ...prev,
        Timesheet_InputList: (prev.Timesheet_InputList || []).map((r) =>
        r.id === id
          ? {
              ...r,
              url: "",
              fileName: undefined,
              sheetName: undefined,
              status: "pending",
              count: undefined,
              date: undefined,
              lastSyncedAt: undefined,
              columnMapping: undefined,
              legacyRowIds: [],
            }
          : r,
        ),
        Timesheet_Roster: replaceTimesheetRosterRows(
          prev.Timesheet_Roster || [],
          [],
          {
            sourceRowIds: ownedRowIds,
            targetL07: targetRow?.l07,
            targetAeCode: targetRow?.aeCode,
          },
        ),
        Q_Salary_Scale: (prev.Q_Salary_Scale || []).filter(
          (r) => !ownedRowIds.has(String(r._rowId || "")),
        ),
        Q_Staff: (prev.Q_Staff || []).filter(
          (r) => !ownedRowIds.has(String(r._rowId || "")),
        ),
        Q_Cache: (prev.Q_Cache || []).filter(
          (r) => !ownedRowIds.has(String(r._rowId || "")),
        ),
      };
    });
    toast.success("Đã xóa dữ liệu dòng");
  };

  const handleClearRow = (id: string) => {
    setClearRowTargetId(id);
  };

  const executeClearAll = async () => {
    updateAppData((prev) => ({
      ...prev,
      Timesheet_InputList: (prev.Timesheet_InputList || []).map((r) => ({
        ...r,
        url: "",
        fileName: undefined,
        sheetName: undefined,
        status: "pending",
        count: undefined,
        date: undefined,
        lastSyncedAt: undefined,
        columnMapping: undefined,
        legacyRowIds: [],
      })),
      Timesheet_Roster: [],
      Q_Salary_Scale: [],
      Q_Staff: [],
      Q_Cache: [],
      Timesheets: [],
      TA_Employee_Summary: { headers: [], data: [] },
      TA_Center_Summary: { headers: [], data: [] },
      Timesheet_Dates: { from: "", to: "" },
      Timesheet_RosterFileName: "",
      Timesheet_RosterEditHistory: [],
      Timesheet_SkipSupabaseRestore: true,
      Timesheet_LocalClearedAt: new Date().toISOString(),
    }), false);

    if (isSupabaseConfigured()) {
      try {
        await clearSupabaseRosterData();
        toast.success("Đã xóa dữ liệu Timesheet local và Roster trên Supabase; vẫn giữ danh sách center.");
      } catch (error: unknown) {
        const message = error instanceof Error ? error.message : String(error);
        toast.error(`Đã xóa dữ liệu local nhưng chưa xóa được Supabase: ${message}`);
      }
    } else {
      toast.success("Đã xóa toàn bộ dữ liệu Timesheet local; vẫn giữ danh sách center.");
    }
  };

  const handleClearAll = () => {
    setShowClearAllDialog(true);
  };

  const executeClearEmptyL07 = () => {
    updateAppData((prev) => ({
      ...prev,
      Timesheet_InputList: (prev.Timesheet_InputList || []).filter(
        (r) => r.l07 && r.l07.trim() !== "",
      ),
    }));
    toast?.success("Đã xóa các dòng chưa có mã L07.");
  };

  const handleClearEmptyL07 = () => {
    setShowClearEmptyL07Dialog(true);
  };

  const handleRecalculate = () => {
    restoreTable(["Timesheet_InputList"]);
    setRefreshKey((prev) => prev + 1);
  };

  const handleSaveData = async () => {
    updateAppData(prev => ({
      ...prev,
      updatedAt: new Date().toISOString()
    }), true);
    
    if (isSupabaseConfigured()) {
      toast.info("Đang tự động đồng bộ dữ liệu hiện tại lên Supabase...");
      await handleSyncToSupabase();
    } else {
      toast.success("Đã lưu dữ liệu hiện tại offline thành công!");
    }
  };

  const handleSyncToSupabase = async () => {
    if (!isSupabaseConfigured()) {
      toast.error("Supabase chưa được cấu hình! Vui lòng cài đặt URL và Anon Key trong phần cấu hình.");
      return;
    }

    if (!rosterData || rosterData.length === 0) {
      toast.warning("Không có dữ liệu Roster để đồng bộ.");
      return;
    }

    setIsSyncing(true);
    setTotalSyncRows(rosterData.length);
    setSyncedRowsCount(0);
    setSyncProgress(0);

    try {
      await preflightTimesheetSync(['roster_cham_cong']);
      const dataToSync = (computedData.processedRosterData && computedData.processedRosterData.length > 0) 
        ? computedData.processedRosterData 
        : rosterData;

      const { successCount, totalRows } = await syncRosterToSupabase(
        dataToSync as Record<string, unknown>[],
        (current, total) => {
          setSyncedRowsCount(current);
          setTotalSyncRows(total);
          setSyncProgress(Math.round((current / total) * 100));
        }
      );

      toast.success(`Đồng bộ thành công ${successCount.toLocaleString()}/${totalRows.toLocaleString()} dòng lên Supabase.`);
      
      updateAppData((prev: AppData) => ({
        ...prev,
        updatedAt: new Date().toISOString(),
        lastSupabaseSyncAt: new Date().toISOString(),
        Timesheet_SkipSupabaseRestore: false,
      }), true);
      toast.success("Đã tự động lưu cứng dữ liệu trên web.");
    } catch (err: unknown) {
      console.error("Supabase Sync Error:", err);
      let errMsg = err instanceof Error ? err.message : String(err);
      
      if (errMsg.includes("Failed to fetch") || errMsg.includes("fetch")) {
        errMsg = "Không thể kết nối tới Supabase (Failed to fetch). Vui lòng kiểm tra lại URL Supabase trong phần Settings và đảm bảo Project của bạn đang hoạt động (không bị tạm dừng).";
      }

      // Detailed alert as requested for debugging RLS and column issues
      alert('Lỗi Supabase: ' + errMsg);
      toast.error(`Đồng bộ thất bại: ${errMsg}`);
      if (err instanceof SupabaseSchemaError) {
        setSyncSchemaError(err.message);
        setShowSqlDialog(true);
      }
    } finally {
      setIsSyncing(false);
    }
  };

  const { syncRow: handleSyncRow } = useTimesheetLinkSync();

  const handleReloadFromFolder = async (id: string) => {
    setIsFetchingGgSheet(true);
    try {
      await handleSyncRow(id);
    } finally {
      setIsFetchingGgSheet(false);
    }
  };

  const handleUploadFiles = (files: File[]) => runTimesheetImport(async () => {
    const currentInputs = appData.Timesheet_InputList || [];
    const updatedInputs = [...currentInputs];
    const toProcess: { id: string; file: File }[] = [];
    let unmatchedCount = 0;

    const filteredFiles = files.filter(
      (file) => !file.name.toLowerCase().includes("copy") && !shouldSkipTimesheetSource(file.name),
    );
    if (filteredFiles.length === 0 && files.length > 0) {
      toast.info("Tất cả các file đã chọn đều là file copy nên hệ thống tự động bỏ qua.");
      return;
    }

    for (const file of filteredFiles) {
      const configuredCenter = resolveTimesheetCenterFromFileName(
        file.name,
        updatedInputs.filter((row) => !isFileNameStoredAsL07(row)),
      );
      if (!configuredCenter) {
        unmatchedCount++;
        continue;
      }
      const matchIndex = updatedInputs.findIndex(
        (row) => row.id === configuredCenter.id,
      );

      if (matchIndex !== -1) {
        updatedInputs[matchIndex] = {
          ...updatedInputs[matchIndex],
          status: "processing",
        };
        toProcess.push({ id: updatedInputs[matchIndex].id, file });
      }
    }

    const latestFileByRow = new Map<string, { id: string; file: File }>();
    toProcess.forEach((item) => latestFileByRow.set(item.id, item));
    const queue = Array.from(latestFileByRow.values());
    if (queue.length === 0) {
      if (unmatchedCount > 0) {
        toast.warning(
          `${unmatchedCount} file không khớp danh sách L07 đã cấu hình nên không được thêm vào bảng.`,
        );
      }
      return;
    }

    type BatchResult = {
      id: string;
      file: File;
      parsed?: ExcelParseResult;
      error?: Error;
    };
    const results: BatchResult[] = [];
    setBulkUploadProgress({ current: 0, total: queue.length });

    for (let index = 0; index < queue.length; index++) {
      const item = queue[index];
      try {
        const parsed = await parseExcelInWorker(item.file, {
          fileId: item.id,
          mode: "auto",
        });
        results.push({ ...item, parsed });
      } catch (error: unknown) {
        results.push({
          ...item,
          error: error instanceof Error ? error : new Error(String(error)),
        });
      }
      setBulkUploadProgress({ current: index + 1, total: queue.length });
    }

    updateAppData((prev) => {
      let nextRoster = prev.Timesheet_Roster || [];
      let nextSalary = prev.Q_Salary_Scale || [];
      let nextStaff = prev.Q_Staff || [];
      let nextCache = prev.Q_Cache || [];
      const statusById = new Map(
        results.map((result) => [result.id, result] as const),
      );

      results.forEach((result) => {
        const input = updatedInputs.find((row) => row.id === result.id);
        if (!result.parsed) return;
        const ownedRowIds = new Set([
          result.id,
          ...(input?.legacyRowIds || []),
        ]);
        const targetL07 = String(input?.l07 || "").trim().toLowerCase();
        const targetAe = String(input?.aeCode || "").trim().toLowerCase();

        nextRoster = replaceTimesheetRosterRows(
          nextRoster,
          result.parsed.kind === "roster" ? result.parsed.rows : [],
          {
            sourceRowIds: ownedRowIds,
            targetL07,
            targetAeCode: targetAe,
          },
        );
        nextSalary = nextSalary.filter(
          (row: Record<string, unknown>) =>
            !ownedRowIds.has(String(row._rowId || "")),
        );
        nextStaff = nextStaff.filter(
          (row: Record<string, unknown>) =>
            !ownedRowIds.has(String(row._rowId || "")),
        );
        nextCache = nextCache.filter(
          (row: Record<string, unknown>) =>
            !ownedRowIds.has(String(row._rowId || "")),
        );

        if (result.parsed.kind === "salary") {
          nextSalary = nextSalary.concat(result.parsed.rows);
        } else if (result.parsed.kind === "staff") {
          nextStaff = nextStaff.concat(result.parsed.rows);
        } else if (result.parsed.kind === "cache") {
          nextCache = nextCache.concat(result.parsed.rows);
        }
      });

      const dateLabel = formatTimesheetSyncDate(new Date());
      const nextInputs = updatedInputs.map((input) => {
        const result = statusById.get(input.id);
        if (!result) return input;
        const detectedL07 = getL07FromFileName(result.file.name) || "";
        const centerInfo = detectedL07
          ? getCenterInfoByL07(detectedL07)
          : null;
        return {
          ...input,
          l07: input.l07 || detectedL07,
          aeCode: input.aeCode || centerInfo?.aeCode || "",
          bus:
            input.bus ||
            centerInfo?.bus ||
            (detectedL07 ? getBusinessFromL07(detectedL07) : ""),
          status: result.error ? ("error" as const) : ("success" as const),
          count: result.parsed?.rows.length || 0,
          fileName: result.file.name,
          date: dateLabel,
          lastSyncedAt: result.error ? input.lastSyncedAt : dateLabel,
          legacyRowIds: result.parsed ? [] : input.legacyRowIds,
        };
      });

      return {
        ...prev,
        Timesheet_Roster: nextRoster,
        Q_Salary_Scale: nextSalary,
        Q_Staff: nextStaff,
        Q_Cache: nextCache,
        Timesheet_InputList: nextInputs,
      };
    }, false);

    setBulkUploadProgress(null);
    const successCount = results.filter((result) => result.parsed).length;
    const errorCount = results.length - successCount;
    if (successCount > 0) {
      toast.success(`Đã xử lý ${successCount}/${results.length} file Excel.`);
    }
    if (errorCount > 0) {
      const firstError = results.find((result) => result.error)?.error?.message;
      toast.error(`${errorCount} file bị lỗi${firstError ? `: ${firstError}` : "."}`);
    }
    if (unmatchedCount > 0) {
      toast.warning(
        `Đã bỏ qua ${unmatchedCount} file không khớp danh sách L07; không tạo thêm dòng mới.`,
      );
    }
  });

  const handleUploadFile = (
    rowId: string,
    file: File,
    sourceMetadata?: { url?: string; uploadDate?: string },
  ) => runTimesheetImport(async () => {
    if (file.name.toLowerCase().includes("copy")) {
      toast?.info(`Hệ thống tự động bỏ qua file có tên 'copy': ${file.name}`);
      return;
    }
    if (shouldSkipTimesheetSource(file.name, sourceMetadata?.url)) {
      toast?.info(`Hệ thống tự động bỏ qua nguồn MKT HP: ${file.name}`);
      return;
    }

    handleUpdateRow(rowId, "status", "processing");
    try {
      const parsed = await parseExcelInWorker(file, {
        fileId: rowId,
        mode: "auto",
      });
      const allRows = parsed.rows;

      if (allRows.length > 0) {
        updateAppData((prev) => {
          const next = { ...prev };
          
          const targetRow = (prev.Timesheet_InputList || []).find(r => r.id === rowId);
          const ownedRowIds = new Set([
            rowId,
            ...(targetRow?.legacyRowIds || []),
          ]);
          const detectedL07 = getL07FromFileName(file.name);
          const finalL07 = targetRow?.l07 || detectedL07 || "";
          const centerInfo = finalL07 ? getCenterInfoByL07(finalL07) : null;
          const targetL07Lower = finalL07.trim().toLowerCase();
          const aeCodeLower = (targetRow?.aeCode || centerInfo?.aeCode || "").trim().toLowerCase();

          next.Timesheet_Roster = replaceTimesheetRosterRows(
            next.Timesheet_Roster || [],
            parsed.kind === "roster" ? allRows : [],
            {
              sourceRowIds: ownedRowIds,
              targetL07: targetL07Lower,
              targetAeCode: aeCodeLower,
            },
          );
          next.Q_Salary_Scale = (next.Q_Salary_Scale || []).filter(
            (r: Record<string, unknown>) =>
              !ownedRowIds.has(String(r._rowId || "")),
          );
          next.Q_Staff = (next.Q_Staff || []).filter(
            (r: Record<string, unknown>) =>
              !ownedRowIds.has(String(r._rowId || "")),
          );
          next.Q_Cache = (next.Q_Cache || []).filter(
            (r: Record<string, unknown>) =>
              !ownedRowIds.has(String(r._rowId || "")),
          );

          if (parsed.kind === "salary")
            next.Q_Salary_Scale = next.Q_Salary_Scale.concat(allRows);
          else if (parsed.kind === "staff")
            next.Q_Staff = next.Q_Staff.concat(allRows);
          else if (parsed.kind === "cache")
            next.Q_Cache = next.Q_Cache.concat(allRows);

          const d = new Date();
          const syncedAt = formatTimesheetSyncDate(d);
          const bu =
            targetRow?.bus ||
            centerInfo?.bus ||
            (detectedL07 ? getBusinessFromL07(detectedL07) : "");

          next.Timesheet_InputList = (next.Timesheet_InputList || []).map((input) =>
            input.id === rowId
              ? {
                  ...input,
                  l07: input.l07 || detectedL07 || "",
                  aeCode: input.aeCode || centerInfo?.aeCode || "",
                  bus: input.bus || bu || "",
                  status: "success",
                  fileName: file.name,
                  url: sourceMetadata?.url || input.url,
                  date:
                    sourceMetadata?.uploadDate || syncedAt,
                  lastSyncedAt: syncedAt,
                  legacyRowIds: [],
                }
              : input
          );

          return next;
        }, false, true, ["Timesheet_Roster", "Q_Staff", "Q_Salary_Scale", "Q_Cache", "Timesheet_InputList"]);

        toast?.success(`Đọc thành công ${file.name}`);
      }
    } catch (err: unknown) {
      const errMsg = err instanceof Error ? err.message : "Unknown error";
      const errName = file.name;
      console.error(`[TimesheetSummary] Error reading ${errName}:`, err);
      handleUpdateRow(rowId, "status", "error");
      toast?.error(
        `Lỗi đọc ${errName}: ${errMsg}`,
      );
    }
  });

  const computedData = useTimesheetCalculations(
    rosterData,
    salaryScaleData,
    staffData,
    cacheData,
    debouncedFromDate,
    debouncedToDate,
    true,
    false,
  );

  useEffect(() => {
    if (computedData.isCalculating || computedData.error) return;
    updateAppData(
      (prev: AppData) => ({
        ...prev,
        TA_Employee_Summary: {
          headers: getDynamicEmployeeColumns(rosterData).map((c) =>
            String(c.label),
          ),
          data: computedData.employeeSummary,
        },
        TA_Center_Summary: {
          headers: CENTER_COLUMNS.map((c) => c.label),
          data: computedData.centerSummary,
        },
      }),
      false,
    );
  }, [computedData.employeeSummary, computedData.centerSummary, computedData.isCalculating, computedData.error, rosterData, updateAppData]);

  const activeData = inputRows;

  const handleUploadFileA = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    try {
      const { rows: allRows } = await parseExcelInWorker(file, { mode: "raw" });

      console.log("Parsed File A:", allRows.slice(0, 5));
      updateAppData((prev) => ({ ...prev, Q_TeacherHours: allRows }));
      toast?.success(`Tải lên File A thành công (${allRows.length} dòng)`);
      if (e.target) e.target.value = "";
    } catch (err: unknown) {
      console.error(err);
      const msg = err instanceof Error ? err.message : "Lỗi khi đọc File A";
      toast?.error(msg);
      if (e.target) e.target.value = "";
    }
  };

  const handleExport = () => {
    if (activeData.length === 0) {
      toast?.error("Không có dữ liệu");
      return;
    }
    const ws = XLSX.utils.json_to_sheet(activeData.map((row, index) => ({"No.": index + 1, "L07": row.l07 || "", "AE Code": row.aeCode || "", "Business": row.bus || "", "File / Link": row.name || row.url || "", "Upload Date": row.lastSyncedAt || row.date || "", "Status": row.status || "", "Actions": ""})));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, activeTab);
    XLSX.writeFile(wb, `Timesheet_Export_${activeTab}.xlsx`);
  };

  return (
    <div 
      className="page-timesheet-summary flex-1 flex flex-col min-h-0 bg-transparent m-0 gap-4 w-full h-full overflow-hidden"
      style={{
        paddingLeft: "0px",
        paddingTop: "0px",
        paddingBottom: "0px",
        paddingRight: "0px",
        borderWidth: "0px",
      }}
    >
      <button data-action="save-data" className="hidden" onClick={handleSaveData} />
      
      <input
        type="file"
        id="fileA"
        accept=".xlsx,.xls,.csv"
        className="hidden"
        onChange={handleUploadFileA}
      />

      <div 
        className="bg-card flex-1 flex flex-col min-h-0 w-full relative overflow-hidden border border-border rounded-xl shadow-sm"
        style={{ paddingLeft: "0px", paddingTop: "0px", paddingBottom: "0px", paddingRight: "0px", borderWidth: "1px" }}
      >
        <div className="absolute inset-0 bg-accent/5 opacity-[0.05] pointer-events-none hidden" />

        <div
          id="timesheet-summary-header"
          className="unified-table-frame-header relative z-10 flex min-h-[56px] w-full min-w-0 shrink-0 flex-col items-stretch justify-between gap-2 bg-[var(--table-header-bg,#FAF3E8)] px-4 py-2.5 md:flex-row md:items-center border-b border-border"
        >
          {computedData?.error && (
            <div role="alert" className="absolute top-0 left-0 right-0 z-50 flex items-center justify-center gap-3 bg-red-100 p-2 text-xs text-red-800">
              <span>Không thể tính lại Timesheet. {computedData.error}</span>
              <button type="button" className="shrink-0 rounded-full border border-red-300 px-3 py-1 font-medium" onClick={computedData.retry}>Thử lại</button>
            </div>
          )}
          {isSyncing && (
            <div className="absolute top-0 left-0 right-0 bg-secondary text-primary-foreground border-b border-none px-8 py-3 flex flex-col sm:flex-row items-center justify-between gap-3 z-50 animate-in fade-in slide-in-from-top duration-300">
              <div className="flex items-center gap-3">
                <RefreshCw className="w-5 h-5 text-primary animate-spin" />
                <div>
                  <p className="text-xs font-black text-primary uppercase tracking-wider">
                    Đang đồng bộ dữ liệu lên Supabase...
                  </p>
                  <p className="text-[10px] font-bold text-foreground uppercase mt-0.5">
                    Đã lưu thành công: {syncedRowsCount.toLocaleString()} / {totalSyncRows.toLocaleString()} dòng ({syncProgress}%)
                  </p>
                </div>
              </div>
              <div className="w-full sm:w-64 bg-accent/20 rounded-full h-2.5 overflow-hidden relative">
                <div 
                  className="bg-primary h-full rounded-full transition-all duration-300"
                  style={{ width: `${syncProgress}%` }}
                />
              </div>
            </div>
          )}
          <div className="relative z-10 flex min-w-0 flex-1 items-center gap-0.5">
            {onBack && (
              <button
                onClick={onBack}
                className="mr-2 flex h-8 w-8 shrink-0 cursor-pointer items-center justify-center rounded-lg border border-border bg-card text-foreground shadow-2xs transition-all hover:bg-muted active:scale-95"
                title="Quay lại"
                aria-label="Quay lại bảng Timesheet"
              >
                <ArrowLeft className="h-4 w-4" />
              </button>
            )}
            <div className="app-table-title-lockup min-w-0 flex-1">
              <div className="app-table-title-line">
                <TableInitialMark
                  label="Setting Timesheet"
                  className="shrink-0 text-primary"
                />
                <h1 className="truncate text-base leading-5 font-bold tracking-tight text-foreground">
                  <TableTitleRemainder label="Setting Timesheet" className="font-bold text-[16px]" style={{ fontWeight: "bold", fontSize: "16px" }} />
                </h1>
              </div>
              <div className="app-table-title-meta flex min-w-0 flex-wrap items-center gap-x-2 gap-y-0.5 text-xs font-medium text-muted-foreground">
                <span className="inline-flex items-center gap-1 whitespace-nowrap">
                  <strong className="font-bold text-foreground">{inputRows.length || 0}</strong>
                  centers
                </span>
                <span aria-hidden="true" className="text-border">•</span>
                <span className="inline-flex items-center gap-1 whitespace-nowrap">
                  <strong className="font-bold text-foreground">{computedData?.employeeSummary?.length || 0}</strong>
                  employees
                </span>
                <span aria-hidden="true" className="text-border">•</span>
                <span className="inline-flex items-center gap-1 whitespace-nowrap">
                  <strong className="font-bold text-foreground">{(computedData?.processedRosterData?.length || 0).toLocaleString()}</strong>
                  records
                </span>
              </div>
            </div>
          </div>

          <div className="relative z-10 flex shrink-0 items-center justify-end">
            <div className="flex items-center">
              <DropdownMenu>
                <DropdownMenuTrigger asChild>
                  <button 
                    id="summary-settings-btn"
                    className="group relative z-10 flex h-8 cursor-pointer items-center gap-1.5 rounded-lg border border-border bg-card px-3 text-foreground shadow-2xs transition-all hover:bg-muted active:scale-95"
                    aria-label="Mở cài đặt Timesheet"
                  >
                    <Settings className="h-3.5 w-3.5 shrink-0 text-primary transition-transform duration-300 group-hover:rotate-45" />
                    <span className="hidden select-none text-[10px] font-bold uppercase tracking-wide sm:inline">
                      CÀI ĐẶT
                    </span>
                    <ChevronDown className="h-3 w-3 shrink-0 text-muted-foreground" />
                  </button>
                </DropdownMenuTrigger>
                <DropdownMenuContent
                  align="end"
                  className="w-64 border border-border shadow-xl p-2 bg-card rounded-xl z-[999999]"
                >
                  <TableRestoreButton placement="menu" fields={["Timesheet_InputList"]} />
                  <DropdownMenuLabel className="text-[10px] font-bold uppercase tracking-wider text-muted-foreground px-3 py-1.5">
                    CÀI ĐẶT & TIỆN ÍCH
                  </DropdownMenuLabel>
                  <DropdownMenuSeparator className="bg-border/60 mx-1" />

                  <DropdownMenuItem
                    onSelect={() => window.dispatchEvent(new Event("open-ui-settings"))}
                    className="flex items-center gap-2.5 px-3 py-2 rounded-lg cursor-pointer hover:bg-muted text-foreground transition-colors font-medium text-xs"
                  >
                    <Settings className="w-3.5 h-3.5 text-primary" />
                    <span>Cấu hình Giao diện</span>
                  </DropdownMenuItem>

                  <DropdownMenuSeparator className="bg-border/60 mx-1" />
                  
                  <DropdownMenuItem
                    onSelect={handleAddRow}
                    className="flex items-center gap-2.5 px-3 py-2 rounded-lg cursor-pointer hover:bg-muted text-foreground transition-colors font-medium text-xs"
                  >
                    <Plus className="w-3.5 h-3.5 text-primary" />
                    <span>Thêm dòng trung tâm</span>
                  </DropdownMenuItem>

                  <DropdownMenuSeparator className="bg-border/60 mx-1" />

                  <DropdownMenuItem
                    onSelect={handleClearAll}
                    className="flex items-center gap-2.5 px-3 py-2 rounded-lg cursor-pointer hover:bg-rose-50 text-rose-600 transition-colors font-medium text-xs"
                  >
                    <Trash2 className="w-3.5 h-3.5" />
                    <span>Xóa dữ liệu trang Timesheet</span>
                  </DropdownMenuItem>
                  <DropdownMenuSeparator className="bg-border/60 mx-1" />
                  <DropdownMenuItem
                    onSelect={() => handleUrlInput(inputRows[0].id, DEFAULT_FOLDER_URL)}
                    disabled={isFetchingGgSheet}
                    className="flex items-center gap-2.5 px-3 py-2 rounded-lg cursor-pointer hover:bg-muted text-foreground transition-colors font-medium text-xs disabled:opacity-50"
                  >
                    <FileSpreadsheet className={`w-3.5 h-3.5 text-amber-600 ${isFetchingGgSheet ? "animate-spin" : ""}`} />
                    <span>Đồng bộ Google Sheet (Folder)</span>
                  </DropdownMenuItem>
                  <DropdownMenuSeparator className="bg-border/60 mx-1" />
                  <DropdownMenuItem
                    onSelect={handleSyncToSupabase}
                    disabled={isSyncing}
                    className="flex items-center gap-2.5 px-3 py-2 rounded-lg cursor-pointer hover:bg-muted text-primary transition-colors font-bold text-xs disabled:opacity-50 disabled:cursor-not-allowed"
                  >
                    <RefreshCw className={`w-3.5 h-3.5 ${isSyncing ? "animate-spin" : ""}`} />
                    <span>ĐỒNG BỘ LÊN SUPABASE</span>
                  </DropdownMenuItem>
                  <DropdownMenuSeparator className="bg-border/60 mx-1" />

                  <DropdownMenuItem
                    onSelect={() => chooseExcelExport(handleExport, handleExport)}
                    className="flex items-center gap-2.5 px-3 py-2 rounded-lg cursor-pointer hover:bg-muted text-foreground transition-colors font-medium text-xs"
                  >
                    <Download className="w-3.5 h-3.5 text-primary" />
                    <span>Xuất Excel</span>
                  </DropdownMenuItem>
                </DropdownMenuContent>
              </DropdownMenu>
            </div>
          </div>
        </div>
        
        {/* Service Account Info Card removed as requested */}


        <div className="flex-1 flex flex-col min-h-0 relative overflow-hidden p-0">
          <TimesheetInputTable
            rows={inputRows}
            onAddRow={handleAddRow}
            onUpdateRow={handleUpdateRow}
            onClearRow={handleClearRow}
            onClearAll={handleClearAll}
            onClearEmptyL07={handleClearEmptyL07}
            onUploadFile={handleUploadFile}
            onUploadFiles={handleUploadFiles}
            onUrlInput={handleUrlInput}
            onRefresh={handleRecalculate}
            onSyncRow={handleSyncRow}
            onReloadFromFolder={handleReloadFromFolder}
            isProcessing={isFetchingGgSheet || bulkUploadProgress !== null}
          />
        </div>
      </div>

      <Dialog open={showSqlDialog} onOpenChange={setShowSqlDialog}>
        <DialogContent className="max-w-2xl bg-card rounded-3xl border-none shadow-2xl p-0 overflow-hidden">
          <div className="bg-primary p-8 text-primary-foreground">
            <DialogHeader>
              <DialogTitle className="text-2xl font-black uppercase tracking-wider">Thiết lập Bảng Supabase</DialogTitle>
              <DialogDescription className="text-primary-foreground/80 font-medium">
                {syncSchemaError || 'Cấu trúc dữ liệu Supabase cần cập nhật.'} Chạy script trong đúng dự án Supabase, sau đó thử đồng bộ lại.
              </DialogDescription>
            </DialogHeader>
          </div>
          <div className="p-8">
            <div className="relative group">
              <pre className="bg-foreground text-secondary p-6 rounded-2xl text-[10px] tabular-nums leading-relaxed overflow-x-auto max-h-[300px] border border-primary/20 shadow-inner custom-scrollbar">
                {SQL_SETUP_SCRIPT}
              </pre>
              <Button
                variant="outline"
                size="sm"
                className="absolute top-4 right-4 bg-card/10 hover:bg-card/20 border-white/20 text-primary-foreground gap-2 backdrop-blur-md opacity-0 group-hover:opacity-100 transition-all"
                onClick={() => {
                  navigator.clipboard.writeText(SQL_SETUP_SCRIPT);
                  toast.success("Đã copy script SQL!");
                }}
              >
                <Copy className="w-3.5 h-3.5" />
                SAO CHÉP
              </Button>
            </div>
            <div className="mt-6 space-y-4">
              <h4 className="text-xs font-black uppercase tracking-widest text-foreground/50">Các bước thực hiện:</h4>
              <ol className="text-[11px] font-bold text-foreground/80 space-y-2 list-decimal pl-4">
                <li>Truy cập vào Dashboard Supabase của bạn.</li>
                <li>Chọn dự án và vào phần <span className="text-primary">SQL Editor</span>.</li>
                <li>Bấm <span className="text-primary">New Query</span> và dán nội dung script trên vào.</li>
                <li>Bấm <span className="text-primary">Run</span> để tạo bảng và cấu hình quyền truy cập (RLS).</li>
                <li>Quay lại đây và thử Đồng bộ lại.</li>
              </ol>
            </div>
          </div>
          <DialogFooter className="p-6 bg-background border-t border-border/50">
            <Button 
              onClick={() => setShowSqlDialog(false)}
              className="bg-primary hover:opacity-90 text-primary-foreground rounded-xl px-8 font-black uppercase tracking-widest text-[10px]"
            >
              Tôi đã hiểu
            </Button>
          </DialogFooter>
        </DialogContent>
      </Dialog>

      <ConfirmDialog
        isOpen={showClearAllDialog}
        onClose={() => setShowClearAllDialog(false)}
        onConfirm={() => {
          setShowClearAllDialog(false);
          executeClearAll();
        }}
        title="Xóa dữ liệu trang Timesheet?"
        description="Toàn bộ dữ liệu nhập file, ca làm Roster và bảng tổng hợp sẽ bị xóa. Dữ liệu Roster trên Supabase cũng sẽ được dọn sạch để tránh tự động tải lại. Dữ liệu Audit, Balance và Master được giữ nguyên."
        confirmText="XÓA TRANG TIMESHEET"
        variant="destructive"
      />

      <ConfirmDialog
        isOpen={showClearEmptyL07Dialog}
        onClose={() => setShowClearEmptyL07Dialog(false)}
        onConfirm={() => {
          setShowClearEmptyL07Dialog(false);
          executeClearEmptyL07();
        }}
        title="Xóa các dòng chưa có mã L07?"
        description="Tất cả các dòng cấu hình file chưa được gán mã cơ sở L07 sẽ bị xóa khỏi danh sách nhập liệu."
        confirmText="XÓA DÒNG TRỐNG"
        variant="destructive"
      />

      <ConfirmDialog
        isOpen={!!clearRowTargetId}
        onClose={() => setClearRowTargetId(null)}
        onConfirm={() => {
          if (clearRowTargetId) {
            executeClearRow(clearRowTargetId);
            setClearRowTargetId(null);
          }
        }}
        title="Xác nhận xóa dòng này?"
        description={`Bạn có chắc chắn muốn xóa file và dữ liệu ca làm của dòng "${(appData.Timesheet_InputList || []).find((r) => r.id === clearRowTargetId)?.l07 || (appData.Timesheet_InputList || []).find((r) => r.id === clearRowTargetId)?.center || "đã chọn"}" khỏi Timesheet?`}
        confirmText="XÓA DÒNG NÀY"
        variant="destructive"
      />
    </div>
  );
}
