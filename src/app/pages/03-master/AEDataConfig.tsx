import { processMasterAEData, parseMonthFromFileName, type AERow } from "../../lib/utils/master-ae-processing";
import "./master-config.css";
import { TableRestoreButton } from '../../components/TableRestoreButton';
import { chooseExcelExport } from "../../components/ExportScopeDialog";
import { createMasterExportDefinition } from "../../lib/utils/master-excel-export";
import { downloadHierarchicalWorkbook } from "../../lib/utils/excel-export";
/* eslint-disable @typescript-eslint/no-explicit-any */
import { useState, useRef, useEffect } from "react";
import { useNavigate } from "react-router";
import {
  Plus,
  Link as LinkIcon,
  UploadCloud,
  Layers,
  Trash2,
  FileSpreadsheet,
  Loader2,
  AlertTriangle,
  Check,
  RefreshCw,
  ChevronLeft,
  ChevronRight,
  ChevronDown,
  Wrench,
  Settings,
  Search,
  Folder,
  Download,
} from "lucide-react";
import { motion, AnimatePresence } from "motion/react";
import * as XLSX from "xlsx";
import { useAppData } from "../../lib/contexts/AppDataContext";
import { PIVOT_MKT_TYPE_CACHE_KEY } from "../../lib/utils/pivot-utils";
import { fetchWithBackoff, generateUUID } from "../../lib/utils/master-data-utils";

import MasterImportWorker from "../../workers/masterImport.worker?worker";
import type { MasterWorkbookPayload } from "../../workers/masterImport.worker";
import { clearMasterPageData } from "../../lib/utils/data-clear-scopes";

import {
  TableInitialMark,
  TableTitleRemainder,
} from "../../components/TableInitialMark";

import { toast } from "sonner";
import {
  Dialog,
  DialogContent,
  DialogHeader,
  DialogTitle,
  DialogDescription,
  DialogFooter,
} from "../../components/ui/dialog";
import { Button } from "../../components/ui/button";
import { Input } from "../../components/ui/input";
import { ColumnMappingDialog } from "./components/ColumnMappingDialog";
import { ConfirmDialog } from "../../components/shared/ConfirmDialog";
import {

  DropdownMenu,
  DropdownMenuContent,
  DropdownMenuItem,
  DropdownMenuLabel,
  DropdownMenuSeparator,
  DropdownMenuTrigger,
} from "../../components/ui/dropdown-menu";

interface PendingUpload {
  file: File;
  existingRowId?: string;
}

type PendingMasterRequest = {
  resolve: (result: MasterWorkbookPayload) => void;
  reject: (error: Error) => void;
};

let masterImportWorker: Worker | null = null;
const pendingMasterRequests = new Map<string, PendingMasterRequest>();

function getMasterImportWorker() {
  if (masterImportWorker) return masterImportWorker;
  masterImportWorker = new MasterImportWorker();
  masterImportWorker.onmessage = (event: MessageEvent) => {
    const requestId = String(event.data?.requestId || "");
    const pending = pendingMasterRequests.get(requestId);
    if (!pending) return;
    pendingMasterRequests.delete(requestId);
    if (event.data?.success) {
      pending.resolve(event.data.result as MasterWorkbookPayload);
    } else {
      pending.reject(
        new Error(event.data?.error || "Không thể xử lý file Master."),
      );
    }
  };
  masterImportWorker.onerror = (event) => {
    const error = new Error(
      event.message || "Master Import Worker đã dừng bất thường.",
    );
    pendingMasterRequests.forEach(({ reject }) => reject(error));
    pendingMasterRequests.clear();
    masterImportWorker?.terminate();
    masterImportWorker = null;
  };
  return masterImportWorker;
}

function parseMasterFileInWorker(
  file: File,
  isMktFile: boolean,
  targetFields: string[],
) {
  const requestId = generateUUID();
  const worker = getMasterImportWorker();
  return new Promise<MasterWorkbookPayload>((resolve, reject) => {
    pendingMasterRequests.set(requestId, { resolve, reject });
    worker.postMessage({ requestId, file, isMktFile, targetFields });
  });
}

const containerVariants = {
  hidden: { opacity: 0 },
  visible: {
    opacity: 1,
    transition: {
      staggerChildren: 0.1,
    },
  },
} as const;

export function AEDataConfig({
  onSwitchToFinal,
}: {
  onSwitchToFinal?: () => void;
}) {
  const navigate = useNavigate();
  const { appData, updateAppData } = useAppData();
  const [isProcessing, setIsProcessing] = useState(false);
  const [processingMessage, setProcessingMessage] = useState("");
  const [progress, setProgress] = useState(0);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const preparedMasterFilesRef = useRef(
    new Map<string, MasterWorkbookPayload>(),
  );
  const [pendingUploads, setPendingUploads] = useState<PendingUpload[]>([]);
  const [choices, setChoices] = useState<
    { file: File; action: "update" | "new" | "skip"; targetId?: string }[]
  >([]);
  const [showDialog, setShowDialog] = useState(false);
  
  const [linkDialogOpen, setLinkDialogOpen] = useState(false);
  const [linkInput, setLinkInput] = useState("");
  const [isFetchingLink, setIsFetchingLink] = useState(false);
  const [activeLinkRowId, setActiveLinkRowId] = useState<string | null>(null);

  const [folderLinkDialogOpen, setFolderLinkDialogOpen] = useState(false);
  const [folderLinkInput, setFolderLinkInput] = useState("");
  const [isFetchingFolder, setIsFetchingFolder] = useState(false);

  const prepareUploadChoices = (uploads: PendingUpload[]) => {
    setPendingUploads(uploads);
    setChoices(uploads.map((upload) => ({
      file: upload.file,
      action: upload.existingRowId ? "update" : "new",
      targetId: upload.existingRowId,
    })));
  };

  const [showClearDialog, setShowClearDialog] = useState(false);
  const [deleteTargetId, setDeleteTargetId] = useState<string | null>(null);
  const [searchTerm, setSearchTerm] = useState("");
  const [debouncedSearchTerm, setDebouncedSearchTerm] = useState("");
  const [showSearch, setShowSearch] = useState(false);
  const [currentPage, setCurrentPage] = useState(1);
  const itemsPerPage = 50;

  useEffect(() => {
    const timer = setTimeout(() => {
      setDebouncedSearchTerm(searchTerm);
      setCurrentPage(1);
    }, 300);
    return () => clearTimeout(timer);
  }, [searchTerm]);

  const [mappingDialog, setMappingDialog] = useState<{
    isOpen: boolean;
    rowId: string | null;
  }>({
    isOpen: false,
    rowId: null,
  });

  const masterAeFields = [
    "No",
    "ID Number",
    "Full name",
    "Full Name",
    "HỌ VÀ TÊN",
    "Salary Scale",
    "From",
    "To",
    "Bank Account Number",
    "Bank Name",
    "CITAD code",
    "TAX CODE",
    "Contract No",
    "CHARGE TO LXO",
    "CHARGE TO EC",
    "CHARGE TO PT-DEMO",
    "Charge MKT Local",
    "CHARGE TO OTHER",
    "Charge Renewal Projects",
    "Charge Discovery Camp",
    "Charge Summer Outing",
    "Charge Summer Instructors",
    "TOTAL PAYMENT",
    "Center",
  ];

  const exportConfigListToExcel = () => {
    if (!appData.Ae_Global_Inputs || appData.Ae_Global_Inputs.length === 0) {
      toast.error("Không có file/bản ghi cấu hình nào để tải xuống!");
      return;
    }
    const exportData = appData.Ae_Global_Inputs.map((row, idx) => ({
      "STT": idx + 1,
      "Tên File": row.name,
      "Region / Bank": row.bank || "",
      "Tháng": row.month || "",
      "Trạng Thái": row.status || "ready",
      "Tên File Gốc": row.fileObj?.name || "",
    }));
    const ws = XLSX.utils.json_to_sheet(exportData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Master_Config_List");
    XLSX.writeFile(wb, `Master_AE_Files_List_${Date.now()}.xlsx`);
    toast.success("Đã tải xuống danh sách cấu hình file Excel!");
  };

  const filteredData = appData.Ae_Global_Inputs.filter(
    (row) =>
      row.name.toLowerCase().includes(debouncedSearchTerm.toLowerCase()) ||
      (row.bank || "").toLowerCase().includes(debouncedSearchTerm.toLowerCase()) ||
      (row.month || "").toLowerCase().includes(debouncedSearchTerm.toLowerCase()),
  );

  const totalPages = Math.max(1, Math.ceil(filteredData.length / itemsPerPage));
  const paginatedData = filteredData.slice(
    (currentPage - 1) * itemsPerPage,
    currentPage * itemsPerPage,
  );

  // Keep the page valid when deleting or filtering the last rows on a page.
  if (currentPage > totalPages) {
    setCurrentPage(totalPages);
  }

  const clearPageData = () => {
    updateAppData(clearMasterPageData);
    localStorage.removeItem("pivot_master_processed_data");
    localStorage.removeItem(PIVOT_MKT_TYPE_CACHE_KEY);
    setShowClearDialog(false);
    toast.success("Đã xóa dữ liệu trang Master; Timesheet, Audit và Balance được giữ nguyên.");
  };

  const addRow = () => {
    const newRow: AERow = {
      id: Date.now().toString(),
      name: "",
      status: "ready",
      bank: "",
      month: "",
    };
    updateAppData((prev) => ({
      ...prev,
      Ae_Global_Inputs: [...prev.Ae_Global_Inputs, newRow],
    }));
  };

  const deleteRow = (id: string | undefined) => {
    if (!id) return;
    setDeleteTargetId(id);
  };

  const confirmDeleteRow = () => {
    if (!deleteTargetId) return;
    updateAppData((prev) => ({
      ...prev,
      Ae_Global_Inputs: prev.Ae_Global_Inputs.filter((row) => row.id !== deleteTargetId),
    }));
    toast.success("Đã xóa dòng cấu hình");
    setDeleteTargetId(null);
  };

  const updateRow = (id: string, field: keyof AERow, value: any) => {
    updateAppData((prev) => ({
      ...prev,
      Ae_Global_Inputs: prev.Ae_Global_Inputs.map((row) => {
        if (row.id === id) {
          const updated = { ...row, [field]: value };
          const uName = String(updated.name || "").toUpperCase();
          const uVal = String(value || "").toUpperCase();
          const uBank = String(updated.bank || "").toUpperCase();
          if (
            uName.includes("MKT") ||
            uName.includes("MARKETING") ||
            uVal.includes("MKT") ||
            uVal.includes("MARKETING") ||
            uBank.includes("MKT") ||
            uBank.includes("MARKETING")
          ) {
            updated.bank = "MKT LOCAL NORTH";
          }
          if (field === "name") {
            const guessedMonth = parseMonthFromFileName(value);
            if (guessedMonth) {
              updated.month = guessedMonth;
            }
          }
          return updated;
        }
        return row;
      }),
    }));
  };

  const handleFileUpload = async (id: string, file: File) => {
    const allowedExtensions = [".xlsx", ".xls", ".csv", ".gsheet"];
    const maxSize = 100 * 1024 * 1024; // 100MB

    const guessBank = (name: string) => {
      const u = name.toUpperCase();
      if (u.includes("MKT") || u.includes("MARKETING")) return "MKT LOCAL NORTH";
      if (u.includes("NORTH")) return "NORTH";
      if (u.includes(" TN") || u.includes(" THAI NGUYEN") || u.includes("_TN_") || u.includes("TN.")) return "THAI NGUYEN";
      if (u.includes(" TH") || u.includes(" THANH HOA") || u.includes("_TH_") || u.includes("TH.")) return "THANH HOA";
      if (u.includes(" PT") || u.includes(" PHU THO") || u.includes("_PT_") || u.includes("PT.")) return "PHU THO";
      return "";
    };

    const fileExtension = file.name
      .substring(file.name.lastIndexOf("."))
      .toLowerCase();
    if (!allowedExtensions.includes(fileExtension)) {
      toast.error(
        `Định dạng file không hợp lệ: ${file.name}. Vui lòng tải lên file Excel (.xlsx, .xls).`,
      );
      return;
    }

    if (file.size > maxSize) {
      toast.error(
        `File quá lớn: ${file.name}. Vui lòng tải lên file nhỏ hơn 100MB.`,
      );
      return;
    }

    setIsProcessing(true);
    setProcessingMessage("Đang tự động map cột...");
    try {
      const guessedBank = guessBank(file.name);
      const parsed = await parseMasterFileInWorker(
        file,
        guessedBank === "MKT LOCAL NORTH",
        masterAeFields,
      );
      preparedMasterFilesRef.current.set(id, parsed);

      updateAppData((prev) => ({
        ...prev,
        Ae_Global_Inputs: prev.Ae_Global_Inputs.map((row) =>
          row.id === id
            ? {
                ...row,
                fileObj: file,
                previousSourceName: row.previousSourceName || row.name,
                name: file.name,
                status: "Uploaded",
                bank: guessedBank || row.bank,
                month:
                  parseMonthFromFileName(file.name) ||
                  row.month ||
                  prev.globalMonth,
                columnMapping: parsed.mapping,
              }
            : row,
        ),
      }), false);
      toast.success(`Đã tải lên và tự động map cột cho file: ${file.name}`);
    } catch (error: any) {
      toast.error(`Lỗi đọc ${file.name}: ${error?.message || String(error)}`);
    } finally {
      setIsProcessing(false);
    }
  };

  const handleLinkSubmit = async () => {
    if (!activeLinkRowId || !linkInput.trim()) return;
    setIsFetchingLink(true);
    try {
      const currentRow = appData.Ae_Global_Inputs.find(r => r.id === activeLinkRowId);
      const response = await fetchWithBackoff(`/api/fetch-google-sheet?url=${encodeURIComponent(linkInput)}`);
      if (!response.ok) {
        const errorData = await response.json().catch(() => ({}));
        throw new Error(errorData.error || "Không thể tải dữ liệu. Hãy đảm bảo link đã được share.");
      }
      
      const arrayBuffer = await response.arrayBuffer();
      const blob = new Blob([arrayBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      
      let baseName = currentRow?.name || `GoogleSheet_Export_${Date.now()}`;
      if (baseName.endsWith(".csv") || baseName.endsWith(".xlsx") || baseName.endsWith(".xls") || baseName.endsWith(".gsheet")) {
        baseName = baseName.substring(0, baseName.lastIndexOf("."));
      }
      const fileName = `${baseName}.xlsx`;
      const file = new File([blob], fileName, { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      
      // Map file and upload to the specific row
      await handleFileUpload(activeLinkRowId, file);
      
      toast.success("Đã tải dữ liệu từ Google Sheet!");
      setLinkDialogOpen(false);
      setLinkInput("");
    } catch (error: any) {
      toast.error(error.message || "Lỗi tải Google Sheet");
    } finally {
      setIsFetchingLink(false);
      setActiveLinkRowId(null);
    }
  };

  const handleFolderLinkSubmit = async () => {
    if (!folderLinkInput.trim()) return;
    setIsFetchingFolder(true);
    try {
      let folderId = folderLinkInput.trim();
      const match = folderLinkInput.match(/folders\/([a-zA-Z0-9-_]+)/);
      if (match) {
        folderId = match[1];
      } else {
        try {
          const url = new URL(folderLinkInput);
          if (url.searchParams.has("id")) {
            folderId = url.searchParams.get("id") || folderId;
          }
        } catch {
          // Ignore invalid URL
        }
      }

      const response = await fetchWithBackoff(`/api/drive-folder-files?folderId=${encodeURIComponent(folderId)}`);
      if (!response.ok) {
        const errData = await response.json().catch(() => ({}));
        throw new Error(errData.error || "Không thể lấy danh sách file từ thư mục. Vui lòng kiểm tra lại link hoặc file credentials.json.");
      }

      const data = await response.json();
      if (!data.success || !data.files || data.files.length === 0) {
        throw new Error("Không tìm thấy file nào trong thư mục này.");
      }

      const newPending: PendingUpload[] = [];
      data.files.forEach((f: any) => {
        const sheetUrl = `https://docs.google.com/spreadsheets/d/${f.id}`;
        const fileContent = JSON.stringify({ url: sheetUrl });
        const blob = new Blob([fileContent], { type: 'application/json' });
        
        let name = f.name || `GoogleSheet_${f.id}`;
        if (!name.endsWith(".gsheet") && !name.endsWith(".xlsx") && !name.endsWith(".xls")) {
          name += ".gsheet";
        } else if (name.endsWith(".xlsx") || name.endsWith(".xls")) {
          // If it already has an extension but is fetched via Drive, just append .gsheet so our parser treats it as URL
          // Wait, actually Google Drive might return name with .xlsx, but we want it to act like .gsheet 
          name = name.replace(/\.(xlsx|xls)$/i, ".gsheet");
        }

        const file = new File([blob], name, { type: 'application/json' });

        const existingRow = appData.Ae_Global_Inputs.find(
          (row) => row.name === name || row.name === name.replace(".gsheet", ".csv"),
        );

        if (existingRow) {
          newPending.push({ file, existingRowId: existingRow.id });
        } else {
          newPending.push({ file });
        }
      });

      prepareUploadChoices(newPending);
      setShowDialog(true);
      setFolderLinkDialogOpen(false);
      setFolderLinkInput("");
      toast.success(`Đã tìm thấy ${data.files.length} file trong thư mục.`);

    } catch (error: any) {
      toast.error(error.message || "Lỗi tải dữ liệu từ Google Drive Folder");
    } finally {
      setIsFetchingFolder(false);
    }
  };

  const handleMultiUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (!files) return;

    const newPending: PendingUpload[] = [];
    Array.from(files).forEach((file) => {
      const existingRow = appData.Ae_Global_Inputs.find(
        (row) => row.name === file.name,
      );
      if (existingRow) {
        newPending.push({ file, existingRowId: existingRow.id });
      } else {
        newPending.push({ file });
      }
    });

    prepareUploadChoices(newPending);
    setShowDialog(true);
    e.target.value = ""; // Reset input
  };

  const confirmUploads = (
    choices: {
      file: File;
      action: "update" | "new" | "skip";
      targetId?: string;
    }[],
  ) => {
    const newRows: AERow[] = [];
    const updates: {
      id: string;
      file: File;
      bank?: string;
      columnMapping?: Record<string, string>;
      status: string;
    }[] = [];

    const guessBank = (name: string) => {
      const u = name.toUpperCase();
      if (u.includes("MKT") || u.includes("MARKETING")) return "MKT LOCAL NORTH";
      if (u.includes("NORTH")) return "NORTH";
      if (u.includes(" TN") || u.includes(" THAI NGUYEN") || u.includes("_TN_") || u.includes("TN.")) return "THAI NGUYEN";
      if (u.includes(" TH") || u.includes(" THANH HOA") || u.includes("_TH_") || u.includes("TH.")) return "THANH HOA";
      if (u.includes(" PT") || u.includes(" PHU THO") || u.includes("_PT_") || u.includes("PT.")) return "PHU THO";
      return "";
    };

    const activeChoices = choices.filter((choice) => choice.action !== "skip");
    for (let index = 0; index < activeChoices.length; index++) {
      const choice = activeChoices[index];

      const guessedBank = guessBank(choice.file.name);
      const id =
        choice.action === "update" && choice.targetId
          ? choice.targetId
          : `${Date.now()}-${index}-${generateUUID()}`;
      preparedMasterFilesRef.current.delete(id);

      if (choice.action === "update" && choice.targetId) {
        updates.push({
          id: choice.targetId,
          file: choice.file,
          bank: guessedBank,
          columnMapping: {},
          status: "ready",
        });
      } else if (choice.action === "new") {
        newRows.push({
          id,
          name: choice.file.name,
          status: "ready",
          fileObj: choice.file,
          bank: guessedBank,
          month: parseMonthFromFileName(choice.file.name) || appData.globalMonth || "",
          columnMapping: {},
        });
      }
    }

    const nextInputs = appData.Ae_Global_Inputs.map((row) => {
        const update = updates.find((u) => u.id === row.id);
      if (update) preparedMasterFilesRef.current.delete(row.id);
        return update
          ? {
              ...row,
              fileObj: update.file,
              status: update.status,
              bank: update.bank || row.bank,
              month:
                parseMonthFromFileName(update.file.name) ||
                row.month ||
                appData.globalMonth,
              columnMapping: update.columnMapping,
            }
          : row;
      }).concat(newRows);

    updateAppData(
      (prev) => ({
        ...prev,
        Ae_Global_Inputs: nextInputs,
      }),
      false,
    );

    setShowDialog(false);
    prepareUploadChoices([]);
    const uploadedCount = newRows.length + updates.length;
    if (uploadedCount === 0) {
      toast.info("Không có file nào được thêm vào danh sách.");
      return;
    }

    toast.success(
      `Đã tải ${uploadedCount} file lên danh sách. Bấm “Xử lý dữ liệu” khi bạn muốn bắt đầu xử lý.`,
    );
  };

  const processAEData = () => processMasterAEData({
    appData,
    updateAppData,
    preparedMasterFiles: preparedMasterFilesRef.current,
    parseMasterFileInWorker,
    masterAeFields,
    setIsProcessing,
    setProgress,
    setProcessingMessage,
    onComplete: () => {
      if (onSwitchToFinal) onSwitchToFinal();
      else navigate("/master-ae");
    },
  });

  return (
    <motion.div
      id="master-config-workspace"
      variants={containerVariants}
      initial="hidden"
      animate="visible"
      className="page-master-config flex-1 flex flex-col min-h-0 bg-transparent m-0 gap-0 w-full h-full overflow-hidden"
      style={{ padding: "10px 12px 12px 12px", boxSizing: "border-box" }}
    >
      {/* One shared frame for title, data area and pagination. */}
      <div 
        className="unified-table-frame table-container bg-card text-card-foreground flex-1 flex flex-col min-h-0 w-full max-w-full relative overflow-hidden border border-border shadow-xs"
        style={{
          borderRadius: "var(--table-radius, 12px)",
          borderWidth: "1px",
          borderColor: "var(--table-frame-border, var(--border))",
          padding: 0,
        }}
      >

        {/* Integrated Header & Controls */}
        <div 
          className="master-config-header unified-table-frame-header relative z-10 flex min-h-[56px] w-full min-w-0 shrink-0 flex-col items-stretch justify-between gap-2 px-3 py-2 md:flex-row md:items-center border-b border-border bg-[var(--table-header-bg,#FAF3E8)]"
          style={{ backgroundColor: "var(--table-header-bg, #FAF3E8)", minHeight: "56px" }}
        >
          <div className="relative z-10 flex min-w-0 flex-1 items-center gap-0.5">
            <button
              type="button"
              onClick={() => {
                if (onSwitchToFinal) onSwitchToFinal();
                else navigate("/master-ae");
              }}
              className="master-config-icon mr-2 flex shrink-0 cursor-pointer items-center justify-center rounded-full border border-border bg-card text-foreground shadow-sm transition-all hover:bg-muted active:scale-[0.98]"
              title="Quay lại Gross Pay"
              aria-label="Quay lại bảng Gross Pay"
            >
              <ChevronLeft className="h-4 w-4" />
            </button>
            <div className="app-table-title-lockup min-w-0 flex-1">
              <div className="app-table-title-line">
                <TableInitialMark
                  label="Setting Master"
                  className="shrink-0 text-primary"
                />
                <h1 className="truncate text-base font-extrabold leading-5 tracking-tight text-foreground">
                  <TableTitleRemainder label="Setting Master" />
                </h1>
              </div>
              <div className="app-table-title-meta flex min-w-0 flex-wrap items-center gap-x-2 gap-y-0.5 text-[11px] font-medium leading-4 text-muted-foreground">
                <span className="inline-flex items-center gap-1 whitespace-nowrap">
                  <strong className="font-bold text-foreground">{appData.Ae_Global_Inputs.length || 0}</strong>
                  file cấu hình
                </span>
                <span aria-hidden="true" className="text-border">•</span>
                <span className="inline-flex items-center gap-1 whitespace-nowrap">
                  <strong className="font-bold text-foreground">{appData.Sheet1_AE.data.length || 0}</strong>
                  bản ghi Gross Pay
                </span>
                <span aria-hidden="true" className="text-border">•</span>
                <span className="inline-flex items-center gap-1 whitespace-nowrap">
                  <strong className="font-bold text-foreground">{appData.Master_Roster?.length || 0}</strong>
                  bản ghi MKT
                </span>
              </div>
            </div>
          </div>

          <div className="relative z-10 flex shrink-0 items-center justify-end gap-2">
            <AnimatePresence>
              {showSearch && (
                <motion.div
                  initial={{ opacity: 0, x: 20, scale: 0.95 }}
                  animate={{ opacity: 1, x: 0, scale: 1 }}
                  exit={{ opacity: 0, x: 20, scale: 0.95 }}
                  className="relative group"
                >
                  <input
                    id="search-input"
                    name="search-input"
                    type="text"
                    placeholder="TÌM KIẾM..."
                    value={searchTerm}
                    onChange={(e) => setSearchTerm(e.target.value)}
                    className="h-8 w-56 rounded-full border border-border bg-card pl-9 pr-3 text-[11px] font-medium text-foreground shadow-sm outline-none transition-all placeholder:text-muted-foreground focus:border-primary focus:ring-2 focus:ring-primary/15"
                    autoFocus
                  />
                  <Search className="w-4 h-4 text-primary/30 absolute left-3.5 top-1/2 -translate-y-1/2 group-focus-within:text-primary transition-colors" />
                </motion.div>
              )}
            </AnimatePresence>

            <div className="flex items-center gap-2">
              {/* Hidden multi-file upload input */}
              <input
                type="file"
                ref={fileInputRef}
                className="hidden"
                accept=".xlsx,.xls,.csv"
                multiple
                onChange={handleMultiUpload}
              />

              <button
                onClick={() => processAEData()}
                disabled={isProcessing}
                className="master-header-action flex items-center gap-1.5 rounded-full bg-primary px-4 text-[11px] font-bold uppercase tracking-wide text-primary-foreground shadow-sm transition-all hover:brightness-95 active:scale-[0.98] disabled:opacity-50"
              >
                {isProcessing ? (
                  <Loader2 className="w-4 h-4 animate-spin shrink-0" />
                ) : (
                  <Layers className="w-4 h-4 shrink-0" />
                )}
                <span>Xử lý</span>
              </button>

              <DropdownMenu>
                <DropdownMenuTrigger asChild>
                  <button
                    className="master-header-action group relative z-10 flex cursor-pointer items-center gap-1.5 rounded-full border border-border bg-card px-3 text-foreground shadow-sm transition-all hover:bg-muted active:scale-[0.98]"
                    aria-label="Mở cài đặt Master"
                  >
                    <Settings className="h-3.5 w-3.5 shrink-0 text-primary transition-transform duration-300 group-hover:rotate-45" />
                    <ChevronDown className="h-3 w-3 shrink-0 text-muted-foreground" />
                  </button>
                </DropdownMenuTrigger>
                <DropdownMenuContent
                  align="end"
                  className="w-64 rounded-2xl border border-border/50 bg-card p-2 text-card-foreground shadow-2xl z-[999999]"
                >
                  <TableRestoreButton placement="menu" fields={["Ae_Global_Inputs"]} />
                  <DropdownMenuLabel className="px-3 py-2 text-[0.625rem] font-bold uppercase tracking-widest text-muted-foreground">
                    Cài đặt &amp; tiện ích
                  </DropdownMenuLabel>
                  <DropdownMenuSeparator className="bg-primary/10 mx-1.5" />

                  <DropdownMenuItem
                    onClick={addRow}
                    className="cursor-pointer font-bold uppercase text-[0.6875rem] gap-3 p-3 rounded-xl transition-all hover:bg-primary/5"
                  >
                    <Plus className="w-4 h-4" />
                    <span>Thêm dòng mới</span>
                  </DropdownMenuItem>

                  <DropdownMenuItem
                    onClick={() => setShowSearch(!showSearch)}
                    className={`cursor-pointer font-bold uppercase text-[0.6875rem] gap-3 p-3 rounded-xl transition-all ${showSearch ? "bg-primary text-primary-foreground shadow-lg shadow-primary/20" : "hover:bg-primary/5"}`}
                  >
                    <Search className="w-4 h-4" />
                    <span>{showSearch ? "Ẩn tìm kiếm" : "Hiện tìm kiếm"}</span>
                  </DropdownMenuItem>

                  <DropdownMenuItem
                    onClick={() => fileInputRef.current?.click()}
                    className="cursor-pointer font-bold uppercase text-[0.6875rem] gap-3 hover:bg-primary/5 text-primary p-3 rounded-xl transition-all"
                  >
                    <UploadCloud className="w-4 h-4" />
                    <span>Upload nhiều File</span>
                  </DropdownMenuItem>
                  
                  <DropdownMenuSeparator className="bg-primary/10 mx-1.5" />

                  <DropdownMenuItem
                    onClick={() => setFolderLinkDialogOpen(true)}
                    className="cursor-pointer font-bold uppercase text-[0.6875rem] gap-3 hover:bg-blue-50 text-blue-600 p-3 rounded-xl transition-all"
                  >
                    <Folder className="w-4 h-4" />
                    <span>Upload thư mục Google Drive</span>
                  </DropdownMenuItem>

                  <DropdownMenuSeparator className="bg-primary/10 mx-1.5" />

                  <DropdownMenuItem
                    onClick={() => chooseExcelExport(exportConfigListToExcel, () => { void downloadHierarchicalWorkbook(createMasterExportDefinition(appData)); })}
                    className="cursor-pointer font-bold uppercase text-[0.6875rem] gap-3 hover:bg-teal-50 text-teal-600 p-3 rounded-xl transition-all"
                  >
                    <Download className="w-4 h-4" />
                    <span>Xuất danh sách cấu hình Excel</span>
                  </DropdownMenuItem>

                  <DropdownMenuSeparator className="bg-primary/10 mx-1.5" />

                  <DropdownMenuItem
                    onClick={() => setShowClearDialog(true)}
                    className="cursor-pointer font-bold uppercase text-[0.6875rem] gap-3 hover:bg-rose-50 text-rose-500 p-3 rounded-xl transition-all"
                  >
                    <Trash2 className="w-4 h-4" />
                    <span>Xóa dữ liệu trang Master</span>
                  </DropdownMenuItem>
                </DropdownMenuContent>
              </DropdownMenu>
            </div>
          </div>
        </div>

        {isProcessing && (
          <div className="relative z-10 flex shrink-0 flex-col gap-1.5 border-b border-border bg-primary/[0.025] px-3 py-2 text-primary">
            <div className="flex items-center justify-between">
              <div className="flex items-center gap-2">
                <Loader2 className="w-4 h-4 animate-spin" />
                <span className="font-bold uppercase text-[0.625rem] tracking-widest">
                  {processingMessage}
                </span>
              </div>
              <span className="text-xs font-bold">{Math.round(progress)}%</span>
            </div>
            <div className="h-1.5 w-full overflow-hidden rounded-full bg-muted">
              <div
                className="bg-primary h-full transition-all duration-300"
                style={{ width: `${progress}%` }}
              ></div>
            </div>
          </div>
        )}

        <div
          ref={(element) => {
            element?.style.setProperty("padding", "0px", "important");
          }}
          className="data-table-wrapper master-config-data-table flex-1 min-h-0 flex flex-col w-full max-w-full p-0 font-[family-name:var(--font-table,var(--font-main))] overflow-hidden"
          style={{ padding: "0px" }}
        >
          <div className="table-body-region master-config-table-region relative flex-1 min-h-0 w-full max-w-full overflow-auto custom-scrollbar bg-card shadow-none">
            <table className="master-config-table relative z-10 w-full border-separate border-spacing-0 text-left" style={{ borderWidth: "0px" }}>
              <colgroup>
                <col className="master-config-col-number" />
                <col className="master-config-col-name" />
                <col className="master-config-col-region" />
                <col className="master-config-col-month" />
                <col />
                <col className="master-config-col-status" />
                <col className="master-config-col-delete" />
              </colgroup>
              <thead>
                <tr className="bg-muted/20">
                  <th
                    style={{ padding: "12px 16px" }}
                    className="sticky top-0 z-20 text-[0.7rem] font-bold text-muted-foreground uppercase tracking-[0.15em] text-center border-b border-r border-border whitespace-nowrap min-w-[60px] bg-muted/30 backdrop-blur-xs shadow-[0_1px_0_rgba(0,0,0,0.02)]"
                  >
                    No
                  </th>
                  <th
                    style={{ padding: "12px 16px" }}
                    className="sticky top-0 z-20 text-[0.7rem] font-bold text-muted-foreground uppercase tracking-[0.15em] border-b border-r border-border whitespace-nowrap text-center min-w-[280px] bg-muted/30 backdrop-blur-xs shadow-[0_1px_0_rgba(0,0,0,0.02)]"
                  >
                    TÊN FILE
                  </th>
                  <th
                    style={{ padding: "12px 16px" }}
                    className="sticky top-0 z-20 text-[0.7rem] font-bold text-muted-foreground uppercase tracking-[0.15em] border-b border-r border-border whitespace-nowrap text-center min-w-[180px] bg-muted/30 backdrop-blur-xs shadow-[0_1px_0_rgba(0,0,0,0.02)]"
                  >
                    REGION / BANK
                  </th>
                  <th
                    style={{ padding: "12px 16px" }}
                    className="sticky top-0 z-20 text-[0.7rem] font-bold text-muted-foreground uppercase tracking-[0.15em] border-b border-r border-border whitespace-nowrap text-center min-w-[120px] bg-muted/30 backdrop-blur-xs shadow-[0_1px_0_rgba(0,0,0,0.02)]"
                  >
                    THÁNG
                  </th>
                  <th
                    style={{ padding: "12px 16px" }}
                    className="sticky top-0 z-20 text-[0.7rem] font-bold text-muted-foreground uppercase tracking-[0.15em] border-b border-r border-border whitespace-nowrap text-center min-w-[320px] bg-muted/30 backdrop-blur-xs shadow-[0_1px_0_rgba(0,0,0,0.02)]"
                  >
                    NGUỒN FILE
                  </th>
                  <th
                    style={{ padding: "12px 16px" }}
                    className="sticky top-0 z-20 text-[0.7rem] font-bold text-muted-foreground uppercase tracking-[0.15em] border-b border-r border-border whitespace-nowrap text-center min-w-[150px] bg-muted/30 backdrop-blur-xs shadow-[0_1px_0_rgba(0,0,0,0.02)]"
                  >
                    TRẠNG THÁI
                  </th>
                  <th
                    style={{ padding: "12px 16px" }}
                    className="sticky top-0 z-20 text-[0.7rem] font-bold text-muted-foreground uppercase tracking-[0.15em] text-center border-b border-border whitespace-nowrap min-w-[70px] bg-muted/30 backdrop-blur-xs shadow-[0_1px_0_rgba(0,0,0,0.02)]"
                  >
                    XÓA
                  </th>
                </tr>
              </thead>
              <tbody className="bg-card text-card-foreground">
                {paginatedData.map((row, idx) => (
                    <tr
                      key={row.id}
                      className="group transition-colors hover:bg-muted/20"
                    >
                      <td
                        style={{
                          padding: "8px 12px",
                          fontFamily: "var(--font-table, var(--font-main))",
                        }}
                        className="text-center border-b border-r border-border min-w-[50px]"
                      >
                        <span className="text-[0.875rem] font-medium text-foreground/40">
                          {(currentPage - 1) * itemsPerPage + idx + 1}
                        </span>
                      </td>
                      <td
                        style={{
                          padding: "8px 12px",
                          fontFamily: "var(--font-table, var(--font-main))",
                        }}
                        className="border-b border-r border-border min-w-[220px]"
                      >
                        <input
                          id={`name-${row.id}`}
                          name={`name-${row.id}`}
                          type="text"
                          value={row.name}
                          title={row.name}
                          onChange={(e) =>
                            updateRow(row.id, "name", e.target.value)
                          }
                          placeholder="Tên file..."
                          className="w-full bg-transparent border-none focus:ring-0 text-[0.875rem] font-semibold text-foreground placeholder:text-foreground/20 p-0 uppercase tracking-tight"
                        />
                      </td>
                      <td
                        style={{
                          padding: "8px 12px",
                          fontFamily: "var(--font-table, var(--font-main))",
                        }}
                        className="border-b border-r border-border min-w-[170px]"
                      >
                        <select
                          id={`bank-${row.id}`}
                          name={`bank-${row.id}`}
                          value={row.bank || ""}
                          onChange={(e) =>
                            updateRow(row.id, "bank", e.target.value)
                          }
                          className="w-full bg-transparent border-none focus:ring-0 text-[0.875rem] font-bold text-foreground/70 p-0 uppercase cursor-pointer appearance-none tracking-wider pr-4"
                        >
                          <option value="" className="text-foreground/40">
                            Chọn Region...
                          </option>
                          <option value="NORTH">NORTH</option>
                          <option value="THANH HOA">THANH HOA</option>
                          <option value="PHU THO">PHU THO</option>
                          <option value="THAI NGUYEN">THAI NGUYEN</option>
                          <option value="MKT LOCAL NORTH">MKT LOCAL NORTH</option>
                        </select>
                      </td>
                      <td
                        style={{
                          padding: "8px 12px",
                          fontFamily: "var(--font-table, var(--font-main))",
                        }}
                        className="border-b border-r border-border text-center min-w-[110px]"
                      >
                        <select
                          id={`month-${row.id}`}
                          name={`month-${row.id}`}
                          value={row.month || parseMonthFromFileName(row.name) || appData.globalMonth || "03.2026"}
                          onChange={(e) =>
                            updateRow(row.id, "month", e.target.value)
                          }
                          className="w-full bg-transparent border-none focus:ring-0 text-[0.875rem] font-bold text-foreground/70 p-0 uppercase cursor-pointer appearance-none tracking-wider text-center"
                        >
                          {[
                            "01.2026", "02.2026", "03.2026", "04.2026", "05.2026", "06.2026",
                            "07.2026", "08.2026", "09.2026", "10.2026", "11.2026", "12.2026"
                          ].map((m) => (
                            <option key={m} value={m} className="text-foreground font-normal">{m}</option>
                          ))}
                        </select>
                      </td>
                      <td
                        style={{
                          padding: "8px 12px",
                          fontFamily: "var(--font-table, var(--font-main))",
                        }}
                        className="border-b border-r border-border min-w-[280px]"
                      >
                        <div className="master-config-source flex items-center gap-2">
                          <input
                            type="file"
                            id={`file-${row.id}`}
                            name={`file-${row.id}`}
                            className="hidden"
                            accept=".xlsx,.xls"
                            onChange={(e) =>
                              e.target.files?.[0] &&
                              handleFileUpload(row.id, e.target.files[0])
                            }
                          />
                          <button
                            onClick={() =>
                              document.getElementById(`file-${row.id}`)?.click()
                            }
                            className={`flex items-center gap-2 px-3.5 py-1.5 rounded-xl border font-bold text-[0.6875rem] tracking-wider uppercase transition-all flex-1 justify-center ${row.fileObj ? "bg-emerald-500 text-primary-foreground border-emerald-500 shadow-sm" : "bg-card text-card-foreground text-primary border-primary/20 hover:bg-primary/5"}`}
                          >
                            <FileSpreadsheet className="w-3.5 h-3.5 shrink-0" />
                            <span className="whitespace-nowrap">{row.fileObj ? "ĐÃ CHỌN" : "CHỌN FILE"}</span>
                          </button>
                          <button
                            onClick={() => {
                              setActiveLinkRowId(row.id);
                              setLinkInput("");
                              setLinkDialogOpen(true);
                            }}
                            className="p-1.5 rounded-full bg-blue-50 border border-blue-100 text-blue-500 hover:bg-blue-100 transition-colors shrink-0"
                            title="Dán link Google Sheet"
                          >
                            <LinkIcon className="w-4 h-4" />
                          </button>

                          {row.fileObj && (
                            <button
                              onClick={() =>
                                setMappingDialog({
                                  isOpen: true,
                                  rowId: row.id,
                                })
                              }
                              className="p-1.5 border border-primary/10 rounded-full bg-card text-card-foreground text-primary hover:bg-primary/5 transition-all shadow-sm shrink-0"
                              title="Cấu hình Mapping Cột"
                            >
                              <Wrench className="w-4 h-4" />
                            </button>
                          )}
                          {row.fileObj && (
                            <span
                              title={row.fileObj.name}
                              className="master-config-source-name text-[0.625rem] font-bold text-foreground/50 truncate uppercase tracking-wider"
                            >
                              {row.fileObj.name}
                            </span>
                          )}
                        </div>
                      </td>
                      <td
                        style={{
                          padding: "8px 12px",
                          fontFamily: "var(--font-table, var(--font-main))",
                        }}
                        className="border-b border-r border-border min-w-[140px]"
                      >
                        <div className="flex items-center justify-center">
                          {row.status === "Success" ? (
                            <div className="flex items-center gap-1.5 text-emerald-600 bg-emerald-50 px-3 py-1 rounded-full text-[0.6875rem] font-bold uppercase tracking-wider border border-emerald-200">
                              <Check className="w-3.5 h-3.5 shrink-0" />
                              <span className="whitespace-nowrap">Thành công</span>
                            </div>
                          ) : row.status === "ready" ? (
                            <div className="flex items-center gap-1.5 text-foreground/50 bg-foreground/5 px-3 py-1 rounded-full text-[0.6875rem] font-bold uppercase tracking-wider border border-border">
                              <div className="w-1.5 h-1.5 rounded-full bg-foreground/30 shrink-0" />
                              <span className="whitespace-nowrap">Sẵn sàng</span>
                            </div>
                          ) : row.status.includes("Error") ? (
                            <div className="flex items-center gap-1.5 text-rose-600 bg-rose-50 px-3 py-1 rounded-full text-[0.6875rem] font-bold uppercase tracking-wider border border-rose-200">
                              <AlertTriangle className="w-3.5 h-3.5 shrink-0" />
                              <span className="whitespace-nowrap">Lỗi</span>
                            </div>
                          ) : (
                            <div className="flex items-center gap-1.5 text-amber-600 bg-amber-50 px-3 py-1 rounded-full text-[0.6875rem] font-bold uppercase tracking-wider border border-amber-200">
                              <RefreshCw className="w-3.5 h-3.5 animate-spin shrink-0" />
                              <span className="whitespace-nowrap">Xử lý...</span>
                            </div>
                          )}
                        </div>
                      </td>
                      <td className="px-3 py-2 text-center border-b border-border min-w-[60px]">
                        <button
                          onClick={() => deleteRow(row.id)}
                          className="p-1.5 text-rose-500 hover:bg-rose-500/10 rounded-full transition-all"
                        >
                          <Trash2 className="w-4 h-4" />
                        </button>
                      </td>
                    </tr>
                  ))}
              </tbody>
            </table>

            {paginatedData.length === 0 && (
              <div
                className="master-config-empty-state absolute inset-x-0 bottom-0 top-[40px] z-0 flex min-h-[220px] items-center justify-center border-t border-border bg-[var(--table-data-bg,var(--card,#fff))] px-6 py-8 text-center text-muted-foreground"
                role="status"
                aria-live="polite"
              >
                <div className="flex max-w-md flex-col items-center gap-4">
                  <div className="flex h-20 w-20 items-center justify-center rounded-2xl border border-primary/10 bg-primary/5">
                    <FileSpreadsheet className="h-9 w-9 text-primary/25" />
                  </div>
                  <div className="space-y-1.5">
                    <p className="font-display text-base font-bold uppercase tracking-tight text-primary/55">
                      {appData.Ae_Global_Inputs.length === 0
                        ? "Chưa có file From AE"
                        : "Không tìm thấy file phù hợp"}
                    </p>
                    <p className="text-[0.625rem] font-bold uppercase tracking-widest text-muted-foreground/65">
                      {appData.Ae_Global_Inputs.length === 0
                        ? "Thêm dòng hoặc upload file để bắt đầu"
                        : "Hãy thay đổi hoặc xóa nội dung tìm kiếm"}
                    </p>
                  </div>
                  {appData.Ae_Global_Inputs.length === 0 && (
                    <div className="flex flex-wrap items-center justify-center gap-2 pt-1">
                      <button
                        type="button"
                        onClick={addRow}
                        className="flex h-8 items-center gap-1.5 rounded-full border border-primary/20 bg-card px-3 text-[0.625rem] font-bold uppercase tracking-wide text-primary shadow-sm transition-colors hover:bg-primary/5"
                      >
                        <Plus className="h-3.5 w-3.5" />
                        Thêm dòng
                      </button>
                      <button
                        type="button"
                        onClick={() => fileInputRef.current?.click()}
                        className="flex h-8 items-center gap-1.5 rounded-full bg-primary px-3 text-[0.625rem] font-bold uppercase tracking-wide text-primary-foreground shadow-sm transition-opacity hover:opacity-90"
                      >
                        <UploadCloud className="h-3.5 w-3.5" />
                        Upload file
                      </button>
                    </div>
                  )}
                </div>
              </div>
            )}
        </div>

        <div 
          className="master-config-footer table-footer-pagination unified-table-frame-footer flex h-[52px] min-h-[52px] max-h-[52px] shrink-0 items-center justify-between border-t border-border bg-card/90 px-4 py-2"
        >
          <div className="flex items-center gap-3">
            <button
              type="button"
              onClick={() => processAEData()}
              disabled={isProcessing}
              className="group rounded-full p-1.5 text-primary/40 transition-all hover:bg-primary/10 hover:text-primary disabled:opacity-50"
              title="Xử lý dữ liệu các file đã tải"
              aria-label="Xử lý dữ liệu các file đã tải"
            >
              <RefreshCw className={`h-4 w-4 ${isProcessing ? "animate-spin" : ""}`} />
            </button>
            <p className="text-[12px] font-bold uppercase tracking-widest text-foreground/40">
              Hiển thị{" "}
              <span className="text-foreground">
                {filteredData.length === 0 ? 0 : (currentPage - 1) * itemsPerPage + 1}
              </span>{" "}
              -{" "}
              <span className="text-foreground">
                {Math.min(currentPage * itemsPerPage, filteredData.length)}
              </span>{" "}
              / <span className="text-foreground">{filteredData.length}</span> file
            </p>
          </div>
          <div className="flex items-center gap-2">
            <button
              type="button"
              onClick={() => setCurrentPage((page) => Math.max(page - 1, 1))}
              disabled={currentPage === 1}
              className="flex h-7 w-7 items-center justify-center rounded-full border border-primary/10 text-primary/60 transition-all hover:bg-primary/10 disabled:opacity-30"
              aria-label="Trang trước"
            >
              <ChevronLeft className="h-4 w-4" />
            </button>
            <span className="min-w-[72px] text-center text-[12px] font-bold uppercase tracking-widest text-foreground/60">
              Trang {currentPage} / {totalPages}
            </span>
            <button
              type="button"
              onClick={() => setCurrentPage((page) => Math.min(page + 1, totalPages))}
              disabled={currentPage === totalPages}
              className="flex h-7 w-7 items-center justify-center rounded-full border border-primary/10 text-primary/60 transition-all hover:bg-primary/10 disabled:opacity-30"
              aria-label="Trang sau"
            >
              <ChevronRight className="h-4 w-4" />
            </button>
          </div>
        </div>
      </div>

      </div>

      <input
        type="file"
        id="file-upload"
        name="file-upload"
        ref={fileInputRef}
        className="hidden"
        multiple
        accept=".xlsx,.xls,.gsheet"
        onChange={handleMultiUpload}
      />

      {/* Confirmation Dialog for Multi-Upload */}
      <Dialog open={folderLinkDialogOpen} onOpenChange={setFolderLinkDialogOpen}>
        <DialogContent className="sm:max-w-md bg-card text-card-foreground border border-primary/10 rounded-2xl shadow-2xl">
          <DialogHeader>
            <DialogTitle className="font-bold uppercase tracking-widest text-primary text-sm">Nhập link Google Drive Folder</DialogTitle>
          </DialogHeader>
          <div className="flex items-center space-x-2 py-4">
            <div className="grid flex-1 gap-2">
              <Input
                id="folder-link"
                placeholder="https://drive.google.com/drive/folders/..."
                value={folderLinkInput}
                onChange={(e) => setFolderLinkInput(e.target.value)}
                disabled={isFetchingFolder}
              />
              <p className="text-[0.65rem] text-muted-foreground mt-1 font-medium">
                <span className="font-bold text-red-500">Lưu ý quan trọng:</span> File <span className="text-primary font-bold">credentials.json</span> (của Google Cloud) phải được tạo/upload vào <span className="font-bold">thư mục gốc của phần mềm này</span> (ở cột cây thư mục bên trái màn hình), KHÔNG phải tạo trên Google Drive. Thư mục Drive chỉ cần được cấp quyền View cho email Service Account.
              </p>
            </div>
          </div>
          <DialogFooter className="sm:justify-end gap-2">
            <Button
              type="button"
              variant="secondary"
              onClick={() => setFolderLinkDialogOpen(false)}
              disabled={isFetchingFolder}
              className="font-bold uppercase text-[0.625rem] tracking-widest px-6 py-2.5 rounded-xl hover:bg-primary/5 transition-all bg-card text-card-foreground border-primary/10"
            >
              Hủy
            </Button>
            <Button
              type="button"
              onClick={handleFolderLinkSubmit}
              disabled={isFetchingFolder || !folderLinkInput.trim()}
              className="bg-primary text-primary-foreground font-bold uppercase text-[0.625rem] tracking-widest px-6 py-2.5 rounded-xl hover:bg-primary/90 shadow-lg shadow-primary/20 transition-all"
            >
              {isFetchingFolder ? "Đang tải..." : "Lấy danh sách file"}
            </Button>
          </DialogFooter>
        </DialogContent>
      </Dialog>

      <Dialog open={linkDialogOpen} onOpenChange={setLinkDialogOpen}>
        <DialogContent className="sm:max-w-md bg-card text-card-foreground border border-primary/10 rounded-2xl shadow-2xl">
          <DialogHeader>
            <DialogTitle className="font-bold uppercase tracking-widest text-primary text-sm">Nhập link Google Sheet</DialogTitle>
          </DialogHeader>
          <div className="flex items-center space-x-2 py-4">
            <div className="grid flex-1 gap-2">
              <Input
                id="link"
                placeholder="https://docs.google.com/spreadsheets/d/..."
                value={linkInput}
                onChange={(e) => setLinkInput(e.target.value)}
                disabled={isFetchingLink}
              />
              <p className="text-[0.65rem] text-muted-foreground mt-1 uppercase tracking-wider font-bold">
                Lưu ý: Chia sẻ dưới dạng <span className="text-primary">"Bất kỳ ai có liên kết"</span>
              </p>
            </div>
          </div>
          <DialogFooter className="sm:justify-end gap-2">
            <Button
              type="button"
              variant="secondary"
              onClick={() => setLinkDialogOpen(false)}
              disabled={isFetchingLink}
              className="font-bold uppercase text-[0.625rem] tracking-widest px-6 py-2.5 rounded-xl hover:bg-primary/5 transition-all bg-card text-card-foreground border-primary/10"
            >
              Hủy
            </Button>
            <Button
              type="button"
              onClick={handleLinkSubmit}
              disabled={isFetchingLink || !linkInput.trim()}
              className="bg-primary text-primary-foreground font-bold uppercase text-[0.625rem] tracking-widest px-6 py-2.5 rounded-xl hover:bg-primary/90 shadow-lg shadow-primary/20 transition-all"
            >
              {isFetchingLink ? "Đang tải..." : "Tải dữ liệu"}
            </Button>
          </DialogFooter>
        </DialogContent>
      </Dialog>

      <Dialog open={showDialog} onOpenChange={setShowDialog}>
        <DialogContent className="max-w-2xl border border-primary/10 shadow-2xl bg-card text-card-foreground rounded-2xl p-6 max-h-[80vh] flex flex-col">
          <DialogHeader>
            <DialogTitle className="font-bold uppercase tracking-widest text-primary text-sm">
              Xác nhận tải lên danh sách file
            </DialogTitle>
            <DialogDescription className="font-bold text-foreground/40 text-[0.625rem] uppercase tracking-widest mt-2">
              Phát hiện {pendingUploads.length} file. Vui lòng chọn hành động
              cho từng file.
            </DialogDescription>
          </DialogHeader>

          <div className="flex-1 overflow-auto custom-scrollbar my-6 border border-primary/10 rounded-xl bg-primary/5 font-[family-name:var(--font-table,var(--font-main))]">
            <table className="w-full text-left border-separate border-spacing-0">
              <thead className="sticky top-0 bg-card text-card-foreground z-10">
                <tr className="bg-primary/10">
                  <th className="px-4 py-3 text-[0.625rem] font-bold uppercase tracking-widest text-primary/60 border-b border-primary/10">
                    Tên File
                  </th>
                  <th className="px-4 py-3 text-[0.625rem] font-bold uppercase tracking-widest text-primary/60 border-b border-primary/10">
                    Hành động
                  </th>
                </tr>
              </thead>
              <tbody className="divide-y divide-primary/5">
                {choices.map((choice, idx) => (
                  <tr
                    key={idx}
                    className=""
                  >
                    <td className="px-4 py-3 text-[0.6875rem] font-bold text-foreground truncate max-w-[300px] uppercase tracking-tight">
                      {choice.file.name}
                    </td>
                    <td className="px-4 py-3">
                      <div className="flex items-center gap-3">
                        <select
                          id={`action-${idx}`}
                          name={`action-${idx}`}
                          value={choice.action}
                          onChange={(e) => {
                            const newChoices = [...choices];
                            newChoices[idx].action = e.target.value as any;
                            setChoices(newChoices);
                          }}
                          className="bg-card text-card-foreground border border-primary/10 rounded-lg px-3 py-1.5 text-[0.625rem] font-bold uppercase tracking-widest focus:outline-none focus:ring-2 focus:ring-primary/20"
                        >
                          <option value="new">Tạo mới</option>
                          {choice.targetId && (
                            <option value="update">Ghi đè</option>
                          )}
                          <option value="skip">Bỏ qua</option>
                        </select>
                        {choice.action === "update" && (
                          <RefreshCw className="w-3.5 h-3.5 text-blue-500 animate-spin" />
                        )}
                        {choice.action === "new" && (
                          <Plus className="w-3.5 h-3.5 text-emerald-500" />
                        )}
                        {choice.action === "skip" && (
                          <AlertTriangle className="w-3.5 h-3.5 text-amber-500" />
                        )}
                      </div>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          <DialogFooter className="gap-3">
            <Button
              variant="outline"
              onClick={() => setShowDialog(false)}
              className="border-primary/10 bg-card text-card-foreground font-bold uppercase text-[0.625rem] tracking-widest px-6 py-2.5 rounded-xl hover:bg-primary/5 transition-all"
            >
              Hủy bỏ
            </Button>
            <Button
              onClick={() => confirmUploads(choices)}
              className="bg-primary text-primary-foreground font-bold uppercase text-[0.625rem] tracking-widest px-6 py-2.5 rounded-xl hover:bg-primary/90 shadow-lg shadow-primary/20 transition-all"
            >
              Xác nhận tải lên
            </Button>
          </DialogFooter>
        </DialogContent>
      </Dialog>

      <Dialog open={showClearDialog} onOpenChange={setShowClearDialog}>
        <DialogContent className="sm:max-w-md border border-primary/10 shadow-2xl bg-card text-card-foreground rounded-2xl p-6">
          <DialogHeader>
            <DialogTitle className="font-bold uppercase tracking-widest text-primary text-sm">
              Xóa dữ liệu trang Master
            </DialogTitle>
            <DialogDescription className="font-bold text-foreground/40 text-[0.625rem] uppercase tracking-widest mt-2">
              Thao tác này xóa toàn bộ file tải lên, dữ liệu và kết quả thuộc
              trang Master. Dữ liệu Timesheet, Audit và Balance không bị ảnh hưởng.
            </DialogDescription>
          </DialogHeader>
          <DialogFooter className="gap-3 mt-6">
            <Button
              variant="outline"
              onClick={() => setShowClearDialog(false)}
              className="border-primary/10 bg-card text-card-foreground font-bold uppercase text-[0.625rem] tracking-widest px-6 py-2.5 rounded-xl hover:bg-primary/5 transition-all"
            >
              Hủy
            </Button>
            <Button
              variant="destructive"
              onClick={clearPageData}
              className="bg-rose-500 text-primary-foreground font-bold uppercase text-[0.625rem] tracking-widest px-6 py-2.5 rounded-xl hover:bg-rose-600 shadow-lg shadow-rose-500/20 transition-all"
            >
              Xóa trang Master
            </Button>
          </DialogFooter>
        </DialogContent>
      </Dialog>

      <ColumnMappingDialog
        isOpen={mappingDialog.isOpen}
        onClose={() => setMappingDialog({ isOpen: false, rowId: null })}
        file={
          appData.Ae_Global_Inputs.find((r) => r.id === mappingDialog.rowId)
            ?.fileObj || null
        }
        targetFields={masterAeFields}
        initialMapping={
          appData.Ae_Global_Inputs.find((r) => r.id === mappingDialog.rowId)
            ?.columnMapping || {}
        }
        onSave={(mapping) => {
          if (mappingDialog.rowId) {
            updateRow(mappingDialog.rowId, "columnMapping", mapping);
            toast.success("Đã lưu cấu hình mapping cột");
          }
        }}
      />

      <ConfirmDialog
        isOpen={!!deleteTargetId}
        onClose={() => setDeleteTargetId(null)}
        onConfirm={confirmDeleteRow}
        title="Xác nhận xóa dòng cấu hình"
        description={`Bạn có chắc chắn muốn xóa dòng cấu hình "${appData.Ae_Global_Inputs.find((r) => r.id === deleteTargetId)?.name || appData.Ae_Global_Inputs.find((r) => r.id === deleteTargetId)?.bank || "này"}" khỏi trang Master?`}
        confirmText="XÓA DÒNG CẤU HÌNH"
        variant="destructive"
      />
    </motion.div>
  );
}

