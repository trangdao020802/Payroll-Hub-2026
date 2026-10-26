/* eslint-disable @typescript-eslint/no-explicit-any, @typescript-eslint/no-unused-vars */
import React, { useMemo, useRef, useState, useCallback, useEffect } from "react";
import { useTableRestore } from '../../hooks/useTableRestore';
import { TableRestoreButton } from '../../components/TableRestoreButton';
import { grossPayRowVisible, sourceRowIndex } from '../../lib/utils/table-originals';
import { chooseExcelExport } from "../../components/ExportScopeDialog";
import { downloadTableExcel } from "../../lib/utils/table-excel";
import { useAppData } from "../../lib/contexts/AppDataContext";
import {
  FileText,
  Landmark,
  PauseCircle,
  Trash2,
  Settings,
  Download,
  Search,
  Users,
  ChevronDown,
  RefreshCw,
  UploadCloud,
  CreditCard,
  PanelLeftClose,
  PanelLeftOpen,
  Wallet,
  CornerDownRight,
  Ban,
  XCircle,
  Plus,
  Table,
  Eye,
  EyeOff,
  X,
  ArrowLeft,
} from "lucide-react";
import { DataTable } from "../../components/DataTable";
import {
  TableInitialMark,
  TableTitleRemainder,
} from "../../components/TableInitialMark";
import {
  Tooltip,
  TooltipTrigger,
  TooltipContent,
} from "../../components/ui/tooltip";
import {
  Popover,
  PopoverTrigger,
  PopoverContent,
} from "../../components/ui/popover";
import {
  DropdownMenu,
  DropdownMenuContent,
  DropdownMenuItem,
  DropdownMenuLabel,
  DropdownMenuSeparator,
  DropdownMenuTrigger,
} from "../../components/ui/dropdown-menu";
import { ConfirmDialog } from "../../components/shared/ConfirmDialog";
import { toast } from "sonner";
import { motion, AnimatePresence } from "motion/react";
import * as XLSX from "xlsx";
import {
  parseMoneyToNumber,
  prepareDataForExport,
  isChargeAmountColumn,
  isNonSummableTextColumn,
} from "../../lib/utils/data-utils";
import { formatVNRobust } from "../../lib/utils/format-utils";
import { Button } from "../../components/ui/button";
import { useMasterAELogic, MasterAETab } from "../../hooks/useMasterAELogic";
import { BulkPayment } from "../04-balance/BulkPayment";
import { HoldAETable } from "./components/HoldAETable";

import { AEDataConfig } from "./AEDataConfig";
import { PivotSheet } from "../04-balance/PivotSheet";
import { Table2 } from "lucide-react";
import { useUiSettings, UI_SETTINGS_KEY } from "../../lib/ui-settings";
import localforage from "localforage";
import {
  downloadHierarchicalWorkbook,
  downloadTransactionBankExport,
} from "../../lib/utils/excel-export";
import { createMasterExportDefinition } from "../../lib/utils/master-excel-export";
import { TransactionReferenceCell } from "../../components/TransactionReferenceCell";
import {
  getTransactionReferenceField,
  type TransactionReferenceAuditEntry,
} from "../../lib/utils/transaction-reference-sync";

const containerVariants = {
  hidden: { opacity: 0 },
  visible: {
    opacity: 1,
    transition: {
      staggerChildren: 0.1,
    },
  },
} as const;

function getMasterTableTitle(tab: string, period: string): string {
  switch (tab) {
    case "Sheet1_AE":
      return `GROSS PAY - THÁNG ${period}`;
    case "Deductions":
      return "BẢNG CHI TIẾT KHẤU TRỪ VÀ THU KHÁC";
    case "NetPay":
      return "BẢNG LƯƠNG NET THỰC CHUYỂN";
    case "Mkt_Local_North":
      return "BẢNG PHÂN PHỐI LƯƠNG AE LOCAL (NORTH)";
    default:
      return tab;
  }
}

export function MasterAE() {
  const { appData, updateAppData } = useAppData();
  const uiSettings = useUiSettings();

  const handleUpdateUiSettings = async (newPartial: any) => {
    const newSettings = { ...uiSettings, ...newPartial };
    await localforage.setItem(UI_SETTINGS_KEY, newSettings);
    window.dispatchEvent(new Event("ui-settings-changed"));
  };

  const [view, setView] = useState<"list" | "upload">("list");
  const isRefreshing = false;
  const [showLeftCard, setShowLeftCard] = useState(true);
  const [showClearBankExportDialog, setShowClearBankExportDialog] =
    useState(false);


  const {
    activeTab,
    setActiveTab,
    searchTerm,
    setSearchTerm,
    showSearch,
    setShowSearch,
    processAEData,
    reMapAECodes,
    handleCellChange,
    handleDeleteRow,
    clearCurrentTableData,
  } = useMasterAELogic();

  const restoreTable = useTableRestore();
  const handleRefreshData = useCallback(() => {
    if (activeTab === "Sheet1_AE" || activeTab === "Bank_North_AE") restoreTable([activeTab]);
  }, [activeTab, restoreTable]);

  const [cameFromBulkPayment, setCameFromBulkPayment] = useState(false);

  const handleOpenTransactionReference = useCallback(
    (
      audit: TransactionReferenceAuditEntry,
      row: Record<string, unknown>,
    ) => {
      const sourceSearch = String(
        audit.transactionId ||
          row["ID Number"] ||
          row["Mã AE"] ||
          audit.newValue ||
          "",
      ).trim();
      const transactionSearch = String(
        audit.transactionId ||
          audit.transactionAccount ||
          audit.transactionName ||
          "",
      ).trim();
      sessionStorage.setItem(
        "transaction_reference_return",
        JSON.stringify({
          targetTable: audit.targetTable,
          targetLabel:
            audit.targetTable === "Sheet1_AE" ? "Gross Pay" : "Deductions",
          sourceSearch,
          transactionSearch,
          transactionKey: audit.transactionKey,
        }),
      );
      localStorage.setItem("bulk_payment_right_tab", "table");
      localStorage.setItem("master_ae_active_tab", "BulkPayment");
      setSearchTerm(transactionSearch);
      setShowSearch(true);
      setCameFromBulkPayment(false);
      setActiveTab("BulkPayment");
      toast.info("Đã mở giao dịch tham chiếu trong bảng Transaction.");
    },
    [setActiveTab, setSearchTerm, setShowSearch],
  );

  useEffect(() => {
    const handleFilter = (e: any) => {
      if (e.detail && e.detail.search) {
        setSearchTerm(e.detail.search);
        setShowSearch(true);
        if (e.detail.from === "BulkPayment") {
          setCameFromBulkPayment(true);
        }
      }
    };
    window.addEventListener("master-ae-filter", handleFilter);
    return () => window.removeEventListener("master-ae-filter", handleFilter);
  }, [setSearchTerm, setShowSearch]);

  useEffect(() => {
    const handleGlobalKeyDown = (e: KeyboardEvent) => {
      const activeEl = document.activeElement;
      if (activeEl?.tagName === "INPUT" || activeEl?.tagName === "TEXTAREA") return;
      
      const key = e.key.toUpperCase();
      if (!["A", "H", "C"].includes(key)) return;
      
      const row = (window as any).__hoveredHoldAERow;
      if (row && activeTab === "Hold_AE") {
        let state = "Add";
        if (key === "H") state = "Hold";
        if (key === "C") state = "Cancel";
        
        if (row["Nghiệp vụ"] !== state) {
          const applyHoldOperation = (window as any).__applyHoldAEOperation;
          if (typeof applyHoldOperation === "function") {
            applyHoldOperation(row, state);
          }
        }
      }
    };
    
    window.addEventListener("keydown", handleGlobalKeyDown);
    return () => window.removeEventListener("keydown", handleGlobalKeyDown);
  }, [activeTab]);

  useEffect(() => {
    const handleRequestTabChange = (e: any) => {
      if (e.detail && e.detail.tab) {
        if (e.detail.tab === "upload") {
          setView("upload");
        } else {
          setView("list");
          setActiveTab(e.detail.tab as any);
        }
      }
    };
    window.addEventListener("master-ae-request-tab-change", handleRequestTabChange);
    return () => window.removeEventListener("master-ae-request-tab-change", handleRequestTabChange);
  }, [setActiveTab]);

  const handleAddRow = (idx?: number) => {
    if (activeTab === "BulkPayment") return;

    // Auto increment primary key or just generate a unique row
    updateAppData((prev) => {
      const tabDataKey = activeTab as keyof typeof prev;
      const targetTab = prev[tabDataKey];
      if (!targetTab || !("data" in targetTab)) return prev;

      const data = [...targetTab.data];
      const headers = targetTab.headers;

      const newRow: Record<string, any> = {
        id: `custom_${Date.now()}`, // fallback id
        _isNew: true,
      };

      headers.forEach((h: string) => {
        newRow[h] = "";
      });
      newRow.id = `custom_${crypto.randomUUID()}`;
      newRow["Tháng báo cáo"] = currentPeriodVal;

      let insertIdx = idx === undefined ? undefined : sourceRowIndex(data, currentData.data[idx]);
      if (insertIdx === undefined && tableRef.current) {
        const activeCell = tableRef.current.getActiveCell?.();
        const filteredAndSorted = tableRef.current.getFilteredAndSortedData?.();
        if (activeCell && filteredAndSorted) {
          const targetRow = filteredAndSorted[activeCell.r];
          if (targetRow) {
            const actualIdx = sourceRowIndex(data, targetRow);
            if (actualIdx >= 0) {
              insertIdx = actualIdx;
            }
          }
        }
      }

      if (
        insertIdx !== undefined &&
        insertIdx >= 0 &&
        insertIdx < data.length
      ) {
        data.splice(insertIdx + 1, 0, newRow);
      } else {
        data.push(newRow);
      }

      return {
        ...prev,
        [tabDataKey]: {
          ...targetTab,
          data,
        },
      };
    });
    toast.success("Đã thêm dòng mới");
  };

  const handleAddColumn = () => {
    const colName = prompt("Nhập tên cột mới muốn thêm:");
    if (!colName || !colName.trim()) return;
    const trimmed = colName.trim();

    updateAppData((prev) => {
      const tabDataKey = activeTab as keyof typeof prev;
      const targetTab = prev[tabDataKey];
      if (!targetTab || !("headers" in targetTab)) return prev;

      const headers = [...targetTab.headers];
      if (!headers.includes(trimmed)) {
        headers.push(trimmed);
      }

      return {
        ...prev,
        [tabDataKey]: {
          ...targetTab,
          headers,
        },
      };
    });
    toast.success(`Đã thêm cột "${trimmed}"`);
  };

  const [showClearDialog, setShowClearDialog] = useState(false);

  const tabs = useMemo(
    () =>
      [
        { id: "Sheet1_AE", label: "Gross Pay", icon: FileText },
        { id: "Hold_AE", label: "Deductions", icon: PauseCircle },
        { id: "BulkPayment", label: "Bulk Payment", icon: CreditCard },
        { id: "Pivot", label: "Pivot Master", icon: Table2 },
      ] as const,
    [],
  );

  const normalizeMonthLabel = useCallback((raw: any, fallbackYear = "2026"): string => {
    const str = String(raw || "").toUpperCase().trim();
    if (!str) return "";
    
    // Match "01.2026", "THÁNG 01/2026", "01-2026", "1.2026"
    const fullMatch = str.match(/(?:THÁNG|THANG|T)?\s*(\d{1,2})[./\- ]\s*(\d{4})/i);
    if (fullMatch) {
      const mm = fullMatch[1].padStart(2, "0");
      const yyyy = fullMatch[2];
      return `${mm}.${yyyy}`;
    }

    // Match "THÁNG 01", "T01", "01"
    const mOnly = str.match(/(?:THÁNG|THANG|T)?\s*(\d{1,2})\b/i);
    if (mOnly) {
      const mm = mOnly[1].padStart(2, "0");
      return `${mm}.${fallbackYear}`;
    }

    return str;
  }, []);

  const currentPeriodVal = useMemo(() => {
    const firstRowMonth = 
      appData.Sheet1_AE?.data?.[0]?.["Tháng báo cáo"] || 
      appData.Sheet1_AE?.data?.[0]?.["_fileMonth"] || 
      appData.Hold_AE?.data?.[0]?.["Tháng báo cáo"] ||
      appData.Hold_AE?.data?.[0]?.["Tháng"];
    const rawGlobal = appData.globalMonth || firstRowMonth || "03.2026";
    const parts = String(rawGlobal).split(".");
    const year = parts.length === 2 ? parts[1] : "2026";
    return normalizeMonthLabel(rawGlobal, year) || "03.2026";
  }, [appData.globalMonth, appData.Sheet1_AE?.data, appData.Hold_AE?.data, normalizeMonthLabel]);

  const currentData = useMemo(() => {
    const raw =
      activeTab === "BulkPayment"
        ? appData.BankExport
        : appData[activeTab as keyof typeof appData] || appData.Sheet1_AE;
    
    if (raw && Array.isArray(raw.data)) {
      const globalYear = currentPeriodVal.split(".")[1] || "2026";

      // Map data to ensure "Tháng báo cáo" is correctly populated
      const mappedData = raw.data.map((r: any, sourceIndex: number) => {
        const mappedRow = { ...r, _originalIndex: sourceIndex };
        const rawM = mappedRow["Tháng báo cáo"] || mappedRow["_fileMonth"] || mappedRow["Tháng"];
        if (rawM) {
          mappedRow["Tháng báo cáo"] = normalizeMonthLabel(rawM, globalYear);
        }
        return mappedRow;
      });

      // Filter all tabs by ONLY the selected reporting month
      const filteredRows = mappedData.filter((r: any) => {
        if (activeTab === "Hold_AE") {
          const nv = String(r["Nghiệp vụ"] || "").toUpperCase().trim();
          const ss = String(r["Sheet Source"] || "").toUpperCase().trim();
          if (nv === "BONUS" || nv === "B" || nv.includes("BONUS") || nv === "⏩" || nv === "⏯" || ss.includes("BONUS")) {
            return false;
          }
        }
        if (activeTab === "Sheet1_AE") {
          if (!grossPayRowVisible(r)) return false;
        }
        const rawM = r["Tháng báo cáo"] || r["_fileMonth"] || r["Tháng"];
        if (!rawM) return true;
        const rowM = normalizeMonthLabel(rawM, globalYear);
        return rowM === currentPeriodVal;
      });
      return { ...raw, data: filteredRows };
    }
    
    return raw;
  }, [activeTab, appData, currentPeriodVal, normalizeMonthLabel]);

  const [tableSummaryState, setTableSummaryState] = useState<{
    tab: string;
    source: any[] | null;
    rows: any[];
  }>({ tab: "", source: null, rows: [] });

  const handleFilteredTableDataChange = useCallback(
    (rows: any[]) => {
      setTableSummaryState({
        tab: activeTab,
        source: currentData.data,
        rows,
      });
    },
    [activeTab, currentData.data],
  );

  const currentSummaryRows =
    tableSummaryState.tab === activeTab &&
    tableSummaryState.source === currentData.data
      ? tableSummaryState.rows
      : currentData.data;

  const filteredSheet1Data = useMemo(() => {
    const raw = appData.Sheet1_AE?.data || [];
    const globalYear = currentPeriodVal.split(".")[1] || "2026";
    
    return raw.filter((r: any) => {
      const rawM = r["Tháng báo cáo"] || r["_fileMonth"] || r["Tháng"];
      if (!rawM) return true;
      const rowM = normalizeMonthLabel(rawM, globalYear);
      return rowM === currentPeriodVal;
    });
  }, [appData.Sheet1_AE?.data, currentPeriodVal, normalizeMonthLabel]);

  const totalSheet1Filtered = useMemo(() => {
    return filteredSheet1Data.reduce((acc, row: any) => acc + parseMoneyToNumber(row["TOTAL PAYMENT"] || row["Total Payment"] || row["Số tiền"] || row["Sale Incentive Amount"] || row["Grand Total"] || row["Payment Amount"] || 0), 0);
  }, [filteredSheet1Data]);

  const currentTabTotalColumn = useMemo(() => {
    const availableColumns = new Map<string, string>();
    (currentData.headers || []).forEach((header: string) => {
      availableColumns.set(String(header).toUpperCase(), header);
    });
    (currentData.data || []).slice(0, 20).forEach((row: any) => {
      Object.keys(row || {}).forEach((key) => {
        availableColumns.set(String(key).toUpperCase(), key);
      });
    });

    const preferredColumns = [
      "TOTAL PAYMENT",
      "PAYMENT AMOUNT",
      "SALE INCENTIVE AMOUNT",
      "SALES/REHIRING AE GP AMOUNT (FINAL)",
      "SỐ TIỀN",
      "GRAND TOTAL",
      "AMOUNT",
      "PHÁT SINH TĂNG/GIẢM",
      "SỐ TIỀN PHẠT",
    ];

    for (const preferredColumn of preferredColumns) {
      const actualColumn = availableColumns.get(preferredColumn.toUpperCase());
      if (actualColumn) return actualColumn;
    }
    return "";
  }, [currentData.headers, currentData.data]);

  const currentTabTotalSum = useMemo(() => {
    if (!currentTabTotalColumn || !Array.isArray(currentSummaryRows)) return 0;
    return currentSummaryRows.reduce((sum: number, row: any) => {
      if (!row || row._isTotalRow) return sum;
      return sum + parseMoneyToNumber(row[currentTabTotalColumn]);
    }, 0);
  }, [currentSummaryRows, currentTabTotalColumn]);

  const currentTableSubtitle = useMemo(() => {
    switch (activeTab) {
      case "Sheet1_AE":
        return `Gross Pay Details (${currentPeriodVal}) · ${currentData.data.length} rows`;
      case "Hold_AE":
        return `Deductions & Adjustments (${currentPeriodVal}) · ${currentData.data.length} rows`;
      default:
        return `Master AE Data (${currentPeriodVal}) · ${currentData.data.length} rows`;
    }
  }, [activeTab, currentData.data.length, currentPeriodVal]);

  const recordsByCategory = useMemo(() => {
    const counts: Record<string, number> = {};
    if (!currentData || !Array.isArray(currentData.data)) return [];
    currentData.data.forEach((r: any) => {
      let cat = String(r.business || r.Business || r.BU || r.l07 || r.L07 || "Unknown").toUpperCase();
      
      // Simplify labels
      if (cat.includes("MKT LOCAL NORTH") || cat === "MKT" || cat === "NTW" || cat.startsWith("MKT LOCAL NORTH_")) cat = "MKT North";
      
      counts[cat] = (counts[cat] || 0) + 1;
    });
    return Object.entries(counts).sort((a, b) => b[1] - a[1]);
  }, [currentData]);

  const columns = useMemo(() => {
    const headers = currentData.headers && currentData.headers.length > 0 
      ? [...currentData.headers] 
      : activeTab === "Sheet1_AE" 
        ? [
            "No.",
            "Tháng báo cáo",
            "L07",
            "Business",
            "ID Number",
            "Full name",
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
            "Extra Summer Instructors",
            "TOTAL PAYMENT",
          ]
        : activeTab === "Bank_North_AE"
          ? ["STT", "L07", "Tháng báo cáo", "Mã AE", "STK AE", "Beneficiary Name", "Business", "Sale Incentive Amount", "Bank", "Note"]
          : activeTab === "Hold_AE"
            ? ["Sheet Source", "STT", "L07", "Tháng báo cáo", "Phân quyền", "Mã AE", "STK AE", "Beneficiary Name", "Business", "Sales/Rehiring AE GP Amount (Final)", "TOTAL PAYMENT", "Bank", "Note", "Tháng phát sinh", "Nghiệp vụ", "Tình trạng thanh toán", "Trạng thái"]
            : activeTab === "BulkPayment"
              ? ["Payment Serial Number", "Tháng báo cáo", "Transaction Type Code", "Payment Type", "Customer Reference No", "Beneficiary Account No.", "Beneficiary Name", "Document ID", "Place of Issue", "ID Issuance Date", "Beneficiary Bank Swift Code / IFSC Code", "Transaction Currency", "Payment Amount", "Charge Type", "Payment details"]
              : [];
    
    // Merge additional keys from data to ensure nothing is hidden
    if (currentData.data && currentData.data.length > 0) {
      const allKeys = Object.keys(currentData.data[0]);
      allKeys.forEach(key => {
        const kUp = key.toUpperCase();
        if (
          !key.startsWith("_") &&
          kUp !== "ID" &&
          kUp !== "_ID" &&
          kUp !== "UUID" &&
          kUp !== "ROWID" &&
          kUp !== "RECORDID" &&
          !headers.some(h => String(h).toUpperCase() === kUp)
        ) {
          headers.push(key);
        }
      });
    }

    const sheet1HiddenCols = [
      "SALARY SCALE",
      "FROM",
      "TO",
      "BANK NAME",
      "CITAD CODE",
      "TAX CODE",
      "CONTRACT NO",
      "TÊN FILE",
      "CENTER",
      "THÁNG",
      "MONTH"
    ];

    let cleanHeaders = headers.filter(h => {
      const u = String(h).trim().toUpperCase();
      if (u === "ID" || u === "_ID" || u === "UUID" || u === "ROWID" || u === "RECORDID" || u.startsWith("_")) {
        return false;
      }
      if (activeTab === "Sheet1_AE" && sheet1HiddenCols.includes(u)) {
        return false;
      }
      return true;
    });

    // Ensure "Tháng báo cáo" exists and is visible for relevant tabs
    const hUp = cleanHeaders.map(h => String(h).toUpperCase());
    if (!hUp.includes("THÁNG BÁO CÁO") && activeTab !== "Bank_North_AE") {
      cleanHeaders.push("Tháng báo cáo");
    }

    const isNoCol = (h: string) => {
      const u = String(h).trim().toUpperCase();
      return u === "NO." || u === "NO" || u === "STT";
    };
    const isBusCol = (h: string) => {
      const u = String(h).trim().toUpperCase();
      return u === "BUSINESS" || u === "BU";
    };
    const isL07Col = (h: string) => {
      const u = String(h).trim().toUpperCase();
      return u === "L07";
    };
    const isNoteCol = (h: string) => {
      const u = String(h).trim().toUpperCase();
      return u === "NOTE" || u === "DIỄN GIẢI";
    };

    const noCol = cleanHeaders.find(isNoCol);
    const busCol = cleanHeaders.find(isBusCol);
    const l07Col = cleanHeaders.find(isL07Col);
    const noteCol = cleanHeaders.find(isNoteCol);

    const remaining = cleanHeaders.filter(h => !isNoCol(h) && !isBusCol(h) && !isL07Col(h) && !isNoteCol(h));

    const finalHeaders: string[] = [];
    if (noCol) finalHeaders.push(noCol);
    else finalHeaders.push("No.");

    if (busCol) finalHeaders.push(busCol);
    if (l07Col) finalHeaders.push(l07Col);

    finalHeaders.push(...remaining);

    if (noteCol && activeTab === "Hold_AE") {
      finalHeaders.push(noteCol);
    }

    cleanHeaders = finalHeaders;

    return cleanHeaders
      .map((header: string) => {
        const h = header.toUpperCase();
        const isLabel = h === "LABEL";
        const isTextOnlyColumn = isNonSummableTextColumn(header);
        const isChargeAmount = isChargeAmountColumn(header);
        let type: "text" | "number" | "currency" | "label" = "text";
        if (isChargeAmount) {
          type = "currency";
        } else if (
          h.includes("TOTAL") ||
          h.includes("PAYMENT") ||
          h.includes("AE") ||
          h.includes("LỆCH") ||
          h.includes("TIỀN") ||
          h.includes("AMOUNT") ||
          h.includes("INCENTIVE") ||
          h.includes("PHẠT") ||
          h.includes("THƯỞNG") ||
          h.includes("CHI") ||
          h.includes("THU") ||
          h.includes("SALARY") ||
          h.includes("LƯƠNG") ||
          h.includes("CỘNG") ||
          h.includes("GP") ||
          h.includes("VALUE")
        ) {
          if (
            !(
              h.includes("ID") ||
              h.includes("ACCOUNT") ||
              h.includes("NUMBER") ||
              h.includes("CODE") ||
              h.includes("STK") ||
              h.includes("MÃ") ||
              h.includes("CENTER") ||
              h.includes("KHÁCH HÀNG") ||
              h.includes("SCALE") ||
              h.includes("HỆ SỐ") ||
              h.includes("RATE")
            )
          ) {
            type = "currency";
          }
        }
        
        if (
          h.includes("SCALE") ||
          h.includes("HỆ SỐ") ||
          h.includes("RATE") ||
          h.includes("DAY") ||
          h.includes("NGÀY")
        ) {
          type = "number";
        }

        if (isTextOnlyColumn) type = "text";
        
        if (isLabel) type = "label";

        const isReadOnly =
          activeTab === "Hold_AE" &&
          [
            "Tháng báo cáo",
            "Nghiệp vụ",
            "Trạng thái",
            "Tháng phát sinh",
            "Tình trạng thanh toán",
          ].includes(header);

        let renderOption:
          | ((value: any, row: any) => React.ReactNode)
          | undefined;

        let hidden = false;
        if (activeTab === "Sheet1_AE") {
          hidden = sheet1HiddenCols.includes(h);
        }

        // Custom render for "Tháng báo cáo" to ensure it's always populated
        if (h === "THÁNG BÁO CÁO") {
          renderOption = (value: any, row: any) => {
            return value || row["_fileMonth"] || row["Tháng"] || "";
          };
        }
        if (activeTab === "Hold_AE" && header === "Nghiệp vụ") {
          renderOption = (value: any, row: any) => {
            const nghiepVu = String(row["Nghiệp vụ"] || row["NGHIỆP VỤ"] || "").toUpperCase();
            
            const isHold = nghiepVu.includes("HOLD") || nghiepVu === "H";
            const isCancel = nghiepVu.includes("CANCEL") || nghiepVu === "C";
            const isAdd = !isHold && !isCancel;
            const currentPeriodVal = appData.globalMonth || "03.2026";
            const currentPeriodParts = currentPeriodVal.split(".");
            const currentMonthNum = parseInt(currentPeriodParts[0], 10) || 3;
            const currentYearNum = parseInt(currentPeriodParts[1], 10) || 2026;
            const currentPeriod = `${String(currentMonthNum).padStart(2, "0")}.${currentYearNum}`;

            const rowReportingMonth = String(row["Tháng báo cáo"] || "").trim();
            const isPeriodMatch =
              rowReportingMonth === currentPeriod ||
              rowReportingMonth.endsWith(".2025") ||
              rowReportingMonth.endsWith("/2025");

            let activeChar = "A";
            let activeColor = "bg-emerald-600 border-emerald-600 text-primary-foreground hover:bg-emerald-700 hover:border-emerald-700";
            let activeLabel = "Add";

            if (isHold) {
              activeChar = "H";
              activeColor = "bg-amber-500 border-amber-500 text-primary-foreground hover:bg-amber-600 hover:border-amber-600";
              activeLabel = "Hold";
            } else if (isCancel) {
              activeChar = "C";
              activeColor = "bg-rose-500 border-rose-500 text-primary-foreground hover:bg-rose-600 hover:border-rose-600";
              activeLabel = "Cancel";
            }
            
            return (
              <div
                onMouseEnter={() => {
                  (window as any).__hoveredHoldAERow = row;
                }}
                onMouseLeave={() => {
                  if ((window as any).__hoveredHoldAERow === row) {
                    (window as any).__hoveredHoldAERow = null;
                  }
                }}
                className="flex items-center justify-center gap-2 w-full py-1 outline-none"
              >
                <button
                  tabIndex={-1}
                  onClick={(e) => {
                    e.stopPropagation();
                    if (!isPeriodMatch) return;
                    const states = ["Add", "Hold", "Cancel"];
                    const currentState = isHold ? "Hold" : isCancel ? "Cancel" : "Add";
                    let nextIdx = states.indexOf(currentState) + 1;
                    if (nextIdx >= states.length) nextIdx = 0;
                    handleCellChange(activeTab, row, "Nghiệp vụ", states[nextIdx]);
                  }}
                  className={`flex items-center justify-center gap-1.5 px-2.5 py-1 text-[11px] font-bold uppercase tracking-wider rounded-lg border shadow-sm transition-all select-none h-7 w-24 hover:brightness-95 active:scale-95 ${
                    !isPeriodMatch
                      ? "bg-secondary/30 border-[#e7dbdc] text-foreground/40 opacity-40 cursor-not-allowed pointer-events-none shadow-none"
                      : activeColor
                  }`}
                  title={!isPeriodMatch ? `Chỉ sửa đổi được tại card tháng chọn: ${rowReportingMonth}` : `Bấm để đổi nghiệp vụ (hoặc dùng phím A/H/C khi di chuột)`}
                  disabled={!isPeriodMatch}
                >
                  <div className="flex items-center justify-center gap-1.5">
                    <span className="inline-flex items-center justify-center w-4 h-4 rounded-full bg-black/10 text-white text-[10px] font-extrabold">
                      {activeChar}
                    </span>
                    <span>{activeLabel}</span>
                  </div>
                </button>
              </div>
            );
          };
        }

        const transactionReferenceField =
          activeTab === "Sheet1_AE"
            ? getTransactionReferenceField(header)
            : null;
        if (transactionReferenceField) {
          renderOption = (value: any, row: any) => (
            <TransactionReferenceCell
              value={value}
              row={row}
              field={transactionReferenceField}
              onOpenTransaction={handleOpenTransactionReference}
            />
          );
        }

        return {
          key: header,
          label: header,
          type,
          hidden,
          sortable: header === "Nghiệp vụ" ? false : true,
          filterable: true,
          readOnly: isReadOnly,
          render: renderOption,
          width: header === "Nghiệp vụ" ? 140 : undefined,
          showGrandTotal:
            !isTextOnlyColumn &&
            (isChargeAmount || type === "currency" || type === "number"),
        };
      });
  }, [
    currentData.headers,
    currentData.data,
    activeTab,
    handleCellChange,
    handleOpenTransactionReference,
    appData.globalMonth,
  ]);

  const handleExportExcel = useCallback(() => {
    if (activeTab === "BulkPayment") {
      if (appData.BankExport.data.length === 0) return;
      downloadTransactionBankExport(appData.BankExport.data);
      return;
    }

    if (downloadTableExcel(activeTab === "Hold_AE" ? "master_ae_Hold_AE" : `master-ae-${activeTab}`)) return;
    if (currentData.data.length === 0) return;

    const ws = XLSX.utils.json_to_sheet(prepareDataForExport(currentData.data));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, activeTab);
    XLSX.writeFile(wb, `Master_AE_${activeTab}.xlsx`);
  }, [currentData.data, activeTab, appData.BankExport.data]);

  const handleExportAllExcel = useCallback(() => {
    void downloadHierarchicalWorkbook(createMasterExportDefinition(appData))
      .then(() =>
        toast.success("Đã xuất toàn bộ 4 trang Master và các bảng phụ."),
      )
      .catch((error) => {
        console.error("Master workbook export failed:", error);
        toast.error("Không thể xuất workbook Master.");
      });
  }, [appData]);

  const tableRef = useRef<any>(null);

  useEffect(() => {
    const tabName = view === "upload" ? "upload" : activeTab;
    window.dispatchEvent(new CustomEvent("master-ae-tab-changed", { detail: { tab: tabName } }));
  }, [activeTab, view]);

  useEffect(() => {
    const handleUploadRequest = () => setView("upload");
    const handleRefreshRequest = () => updateAppData(prev => ({ ...prev }), false);
    const handleExportRequest = () => handleExportExcel();
    const handleSectionExportRequest = () => handleExportAllExcel();
    const handleClearRequest = () => setShowClearDialog(true);

    window.addEventListener("master-ae-request-upload", handleUploadRequest);
    window.addEventListener("master-ae-request-refresh", handleRefreshRequest);
    window.addEventListener("master-ae-request-export", handleExportRequest);
    window.addEventListener("app-export-section-excel", handleSectionExportRequest);
    window.addEventListener("master-ae-request-clear", handleClearRequest);

    return () => {
      window.removeEventListener("master-ae-request-upload", handleUploadRequest);
      window.removeEventListener("master-ae-request-refresh", handleRefreshRequest);
      window.removeEventListener("master-ae-request-export", handleExportRequest);
      window.removeEventListener("app-export-section-excel", handleSectionExportRequest);
      window.removeEventListener("master-ae-request-clear", handleClearRequest);
    };
  }, [setActiveTab, updateAppData, handleExportExcel, handleExportAllExcel]);

  return (
    <div className="page-master-ae flex-1 flex flex-col min-h-0 relative overflow-hidden bg-transparent">
      <AnimatePresence initial={false}>
        {view === "list" && (
          <motion.div
            key="list-main"
            initial={{ opacity: 0, y: 20 }}
            animate={{ opacity: 1, y: 0 }}
            exit={{ opacity: 0, y: "100%" }}
            className="flex-1 flex flex-col min-h-0 gap-0 relative overflow-hidden bg-transparent w-full p-0"
            style={{
              paddingLeft: "0px",
              paddingRight: "0px",
              paddingTop: "0px",
              paddingBottom: "0px",
            }}
          >
            {/* Inner Content Area holding Table */}
            <div className="flex-1 min-h-0 relative overflow-hidden w-full h-full">
              {/* Right Panel: Content Grid */}
              <div 
                className="flex-1 flex flex-col min-h-0 h-full overflow-hidden relative w-full content-area"
                style={{ paddingTop: "0px", paddingBottom: "0px", borderWidth: "0px", paddingLeft: "0px", paddingRight: "0px", borderColor: "#ccd5ef" }}
              >
                <div 
                  className="table-container flex-1 flex flex-col min-h-0 relative bg-transparent rounded-none shadow-none overflow-hidden master-ae-table-wrapper"
                  data-split-panels={activeTab === "BulkPayment"}
                  style={{ paddingTop: "0px", paddingLeft: "0px", paddingRight: "0px", paddingBottom: "0px", borderWidth: "0px" }}
                >
                  {activeTab === "BulkPayment" && (
                    <BulkPayment
                      showLeftCard={showLeftCard}
                      setShowLeftCard={setShowLeftCard}
                      searchTerm={searchTerm}
                      onSearchTermChange={setSearchTerm}
                      onTabChange={(target) => {
                        if (target === "upload") {
                          setView("upload");
                          return;
                        }
                        setActiveTab(target);
                        localStorage.setItem("master_ae_active_tab", target);
                        window.dispatchEvent(new CustomEvent("master-ae-tab-changed", { detail: { tab: target } }));
                      }}
                    />
                  )}
                  <div 
                    className="flex-1 flex flex-col min-h-0 w-full h-full overflow-hidden" 
                    style={{ 
                      display: activeTab === "Pivot" ? "flex" : "none",
                      padding: "0px",
                    }}
                  >
                    <PivotSheet />
                  </div>
                  {activeTab !== "BulkPayment" && activeTab !== "Pivot" && (
                    <div className="flex-1 flex flex-col min-h-0 w-full h-full overflow-hidden relative">
                      <div className="absolute inset-0 striped-pattern opacity-0 pointer-events-none overflow-hidden" />
                      
                      {activeTab === "Hold_AE" ? (
                        <HoldAETable
                          ref={tableRef}
                          searchTerm={searchTerm}
                          onSearchTermChange={setSearchTerm}
                          onAddRow={handleAddRow}
                          cameFromBulkPayment={cameFromBulkPayment}
                          onBackToBulkPayment={() => {
                            setSearchTerm("");
                            localStorage.removeItem("master_ae_search");
                            localStorage.setItem("bulk_payment_right_tab", "reconcile");
                            setActiveTab("BulkPayment");
                            setCameFromBulkPayment(false);
                            window.dispatchEvent(new CustomEvent("bulk-payment-set-right-tab", { detail: { tab: "reconcile" } }));
                          }}
                          onOpenTransactionReference={handleOpenTransactionReference}
                        />
                      ) : currentData.data.length === 0 ? (
                        <div className="flex-1 flex flex-col items-center justify-center text-primary/10 p-12 relative z-10">
                          <div className="w-20 h-20 bg-secondary/20 rounded-full flex items-center justify-center mb-6 border border-primary/5">
                            <Table className="w-10 h-10 text-primary/20" />
                          </div>
                          <p className="font-bold uppercase text-xl tracking-tight text-primary/40">
                            Chưa có dữ liệu {activeTab}
                          </p>
                          <div className="flex gap-2 my-3 text-primary">
                            <TableRestoreButton onRestore={handleRefreshData} />
                            <button type="button" onClick={() => handleAddRow()} className="rounded-full border border-border px-3 py-1.5 text-xs">Thêm dòng</button>
                          </div>
                          <p className="text-[0.625rem] font-bold uppercase opacity-40 tracking-widest mt-2 text-center max-w-md">
                            Vui lòng vào phần Setting Master để chọn file AE Final, hệ thống sẽ tự động cập nhật dữ liệu.
                          </p>
                        </div>
                      ) : (
                        <div 
                          className="unified-table-frame table-container flex-1 flex flex-col min-h-0 w-full h-full px-0 py-0 m-0 relative overflow-hidden gap-0 bg-card border border-border shadow-xs z-10"
                          style={{
                            borderRadius: "var(--table-radius, 12px)",
                            borderWidth: "1px",
                            borderColor: "var(--table-frame-border, var(--border))",
                            padding: 0,
                          }}
                        >
                          {/* Top Toolbar Header with Settings Button */}
                          <div 
                            className="unified-table-frame-header flex min-h-[56px] items-center justify-between gap-3 bg-[var(--table-header-bg,#FAF3E8)] px-3 py-2 shrink-0 select-none border-b border-border"
                            style={{ backgroundColor: "var(--table-header-bg, #FAF3E8)", minHeight: "56px" }}
                          >
                            <div className="app-table-title-lockup min-w-0">
                              <div className="app-table-title-line">
                                <TableInitialMark
                                  label={getMasterTableTitle(activeTab, currentPeriodVal)}
                                  className="shrink-0 text-primary"
                                />
                                <h3 
                                  className="truncate font-bold tracking-tight text-foreground"
                                  style={{ fontSize: "13px", lineHeight: "23px", height: "18.0012px" }}
                                >
                                  <TableTitleRemainder
                                    label={getMasterTableTitle(activeTab, currentPeriodVal)}
                                  />
                                </h3>
                              </div>
                              <p
                                className="app-table-title-meta truncate font-medium text-muted-foreground"
                                style={{ fontSize: "10px", lineHeight: "14.375px" }}
                              >
                                {currentTableSubtitle}
                              </p>
                            </div>

                            <div className="flex shrink-0 items-center gap-3">
                              {cameFromBulkPayment && (
                                <button
                                  onClick={() => {
                                    setSearchTerm("");
                                    localStorage.removeItem("master_ae_search");
                                    localStorage.setItem("bulk_payment_right_tab", "reconcile");
                                    setActiveTab("BulkPayment");
                                    setCameFromBulkPayment(false);
                                    window.dispatchEvent(new CustomEvent("bulk-payment-set-right-tab", { detail: { tab: "reconcile" } }));
                                  }}
                                  className="px-3 py-1 bg-indigo-50 hover:bg-indigo-100 text-indigo-700 border border-indigo-200 transition-all shadow-3xs rounded-full active:scale-95 cursor-pointer flex items-center justify-center gap-1.5 font-bold text-[10px] h-7 uppercase tracking-wider"
                                  title="Quay lại Bảng Reconciliation"
                                >
                                  <ArrowLeft className="w-3 h-3" />
                                  <span>Về Reconciliation</span>
                                </button>
                              )}
                              
                              {/* Search Input shown dynamically based on showSearch state */}
                              {showSearch && (
                                <div 
                                  className="flex h-8 items-center gap-2 rounded-full border border-border bg-card px-3 text-xs shadow-sm transition-all focus-within:border-primary focus-within:ring-2 focus-within:ring-primary/15 animate-in fade-in slide-in-from-right duration-250"
                                  style={{ width: "220px" }}
                                >
                                  <Search className="w-3.5 h-3.5 text-slate-400 shrink-0" />
                                  <input
                                    type="text"
                                    placeholder="Tìm kiếm..."
                                    value={searchTerm}
                                    onChange={(e) => setSearchTerm(e.target.value)}
                                    className="w-full bg-transparent border-none outline-none text-xs text-foreground placeholder:text-muted-foreground"
                                  />
                                  {searchTerm && (
                                    <button
                                      onClick={() => setSearchTerm("")}
                                      className="text-slate-400 hover:text-slate-700 text-xs p-0.5 transition-colors cursor-pointer"
                                      title="Xóa tìm kiếm"
                                    >
                                      <X className="w-3 h-3" />
                                    </button>
                                  )}
                                  <button
                                    onClick={() => {
                                      setShowSearch(false);
                                      setSearchTerm("");
                                    }}
                                    className="text-slate-400 hover:text-slate-600 hover:bg-slate-100 p-0.5 rounded-full transition-colors cursor-pointer"
                                    title="Đóng tìm kiếm"
                                  >
                                    <X className="w-3 h-3" />
                                  </button>
                                </div>
                              )}

                              {/* Key statistics block perfectly matching the second image */}
                              <div className="flex items-center gap-4 mr-1">
                                <div className="flex flex-col items-end">
                                  <span className="text-[8px] font-bold text-muted-foreground uppercase tracking-tighter whitespace-nowrap">NHÂN VIÊN</span>
                                  <span className="text-xs font-black text-foreground leading-tight">{currentData.data.length}</span>
                                </div>
                                <div className="flex flex-col items-end border-l border-border pl-4">
                                  <span className="text-[8px] font-bold text-muted-foreground uppercase tracking-tighter whitespace-nowrap">TỔNG TIỀN</span>
                                  <div className="bg-muted/60 px-2 py-0.5 rounded border border-border/60 mt-0.5">
                                    <span className="text-xs font-black text-foreground tracking-tight tabular-nums">{formatVNRobust(currentTabTotalSum, 0)}</span>
                                  </div>
                                </div>
                              </div>
 
                                 {/* Nút Cài đặt (Settings Button) */}
                                 <DropdownMenu>
                                   <DropdownMenuTrigger asChild>
                                     <button
                                       className="group flex h-8 cursor-pointer items-center gap-1.5 rounded-full border border-border bg-card px-3 text-foreground shadow-sm transition-all hover:bg-muted active:scale-95"
                                       title="Cài đặt & Thao tác"
                                       aria-label="Mở cài đặt và thao tác bảng Master"
                                     >
                                       <Settings className="h-3.5 w-3.5 text-primary transition-transform duration-300 group-hover:rotate-45" />
                                       <ChevronDown className="h-3 w-3 text-muted-foreground" />
                                     </button>
                                   </DropdownMenuTrigger>
                                 <DropdownMenuContent align="end" className="w-56 p-2 rounded-2xl shadow-2xl border-slate-100 z-[99999]">
                                   <DropdownMenuLabel className="text-[10px] font-black text-slate-400 uppercase tracking-widest px-3 py-2">
                                     Action Center
                                   </DropdownMenuLabel>
                                   <DropdownMenuSeparator className="bg-slate-50" />
                                   
                                   {/* Toggle Search Menu Item */}
                                   <DropdownMenuItem
                                     onClick={() => setShowSearch(!showSearch)}
                                     className="flex items-center gap-3 px-3 py-2.5 rounded-xl cursor-pointer hover:bg-slate-50 transition-colors"
                                   >
                                     <Search className="w-4 h-4 text-primary" />
                                     <span className="text-xs font-bold text-slate-700">
                                       {showSearch ? "Ẩn công cụ tìm kiếm" : "Tìm kiếm..."}
                                     </span>
                                   </DropdownMenuItem>

                                   <DropdownMenuItem
                                     onClick={() => handleAddRow()}
                                     className="flex items-center gap-3 px-3 py-2.5 rounded-xl cursor-pointer hover:bg-slate-50 transition-colors"
                                   >
                                    <Plus className="w-4 h-4 text-primary" />
                                    <span className="text-xs font-bold text-slate-700">Thêm dòng mới</span>
                                  </DropdownMenuItem>
                                  <DropdownMenuItem
                                    onClick={handleRefreshData}
                                    className="flex items-center gap-3 px-3 py-2.5 rounded-xl cursor-pointer hover:bg-slate-50 transition-colors"
                                  >
                                    <RefreshCw className="w-4 h-4 text-primary" />
                                    <span className="text-xs font-bold text-slate-700">Reset to default</span>
                                  </DropdownMenuItem>
                                  {activeTab !== "Sheet1_AE" && (
                                    <DropdownMenuItem
                                      onClick={() => {
                                        if (tableRef?.current?.resetTableConfig) {
                                          tableRef.current.resetTableConfig();
                                        } else {
                                          toast.error("Không tìm thấy cấu hình bảng");
                                        }
                                      }}
                                      className="flex items-center gap-3 px-3 py-2.5 rounded-xl cursor-pointer hover:bg-slate-50 transition-colors"
                                    >
                                      <RefreshCw className="w-4 h-4 text-amber-600 animate-pulse" />
                                      <span className="text-xs font-bold text-slate-700">Khôi phục bố cục bảng</span>
                                    </DropdownMenuItem>
                                  )}
                                  <DropdownMenuItem
                                    onClick={() => window.dispatchEvent(new Event("open-ui-settings"))}
                                    className="flex items-center gap-3 px-3 py-2.5 rounded-xl cursor-pointer hover:bg-slate-50 transition-colors"
                                  >
                                    <Settings className="w-4 h-4 text-slate-500" />
                                    <span className="text-xs font-bold text-slate-700">Cài đặt Giao diện</span>
                                  </DropdownMenuItem>
                                  <DropdownMenuItem
                                    onClick={() => chooseExcelExport(handleExportExcel, handleExportAllExcel)}
                                    className="flex items-center gap-3 px-3 py-2.5 rounded-xl cursor-pointer hover:bg-slate-50 transition-colors"
                                  >
                                    <Download className="w-4 h-4 text-emerald-600" />
                                    <span className="text-xs font-bold text-slate-700">Xuất Excel</span>
                                  </DropdownMenuItem>
                                  <DropdownMenuSeparator className="bg-slate-50" />
                                  <DropdownMenuItem
                                    onClick={() => setShowClearDialog(true)}
                                    className="flex items-center gap-3 px-3 py-2.5 rounded-xl cursor-pointer hover:bg-rose-50 text-rose-600 transition-colors"
                                  >
                                    <Trash2 className="w-4 h-4" />
                                    <span className="text-xs font-bold">
                                      {activeTab === "Sheet1_AE"
                                        ? "Xóa dữ liệu bảng Gross Pay"
                                        : "Xóa dữ liệu bảng Bank North"}
                                    </span>
                                  </DropdownMenuItem>
                                </DropdownMenuContent>
                              </DropdownMenu>
                            </div>
                          </div>

                          <DataTable
                            className="!overflow-visible"
                            hideColumnVisibilityToggle={false}
                            showRowNumber={true}
                            scrollContainerStyle={{ borderRadius: "0", border: "none" }}
                            storageKey={`master-ae-${activeTab}`}
                            ignoreSavedHiddenColumns={true}
                            selectable={false}
                            defaultItemsPerPage={50}
                            ignoreSavedPagination={true}
                            ref={tableRef}
                            columns={columns}
                            data={currentData.data}
                            onFilteredDataChange={handleFilteredTableDataChange}
                            onCellChange={(row, col, val) => handleCellChange(activeTab, row, col, val)}
                            onDeleteRow={(row, idx) => handleDeleteRow(activeTab, row)}
                            onAddRow={handleAddRow}
                            isEditable={true}
                            stickyFirstColumn={false}
                            externalSearchTerm={searchTerm}
                            onExternalSearchChange={setSearchTerm}
                            hideSearch={true}
                            showPagination={true}
                            showFooter={true}
                            footerClassName="bg-card text-foreground border-t border-border font-bold"
                            headerClassName="bg-slate-100 text-accent border-[#E2E8F0] font-bold"
                          />
                        </div>
                      )}
                    </div>
                  )}
                </div>
              </div>
            </div>
            <ConfirmDialog
              isOpen={showClearDialog}
              onClose={() => setShowClearDialog(false)}
              onConfirm={() => {
                clearCurrentTableData();
                setShowClearDialog(false);
              }}
              title={`Xóa dữ liệu bảng ${activeTab === "Sheet1_AE" ? "Gross Pay" : "Bank North"}?`}
              description="Thao tác này chỉ xóa dữ liệu trong bảng hiện tại. Các bảng Master, Timesheet, Audit và Balance khác được giữ nguyên."
              confirmText="XÓA BẢNG NÀY"
              variant="destructive"
            />
            <ConfirmDialog
              isOpen={showClearBankExportDialog}
              onClose={() => setShowClearBankExportDialog(false)}
              onConfirm={() => {
                updateAppData((prev) => ({
                  ...prev,
                  BankExport: { ...prev.BankExport, data: [] },
                  TransactionMonthCache: undefined,
                }));
                setShowClearBankExportDialog(false);
                toast.success("Đã xóa dữ liệu bảng kê");
              }}
              title="Xóa dữ liệu bảng kê?"
              description="Hành động này sẽ xóa sạch dữ liệu trong bảng kê Bulk Payment. Bạn có chắc chắn muốn tiếp tục?"
              confirmText="Xác nhận xoá"
              variant="destructive"
            />
          </motion.div>
        )}
        {view === "upload" && (
          <motion.div
            key="upload"
            initial={{ y: "-100%", opacity: 0 }}
            animate={{ y: 0, opacity: 1 }}
            exit={{ y: "-100%", opacity: 0 }}
            transition={{ type: "spring", damping: 25, stiffness: 200 }}
            className="absolute inset-0 flex w-full flex-col"
            style={{
              paddingLeft: "18px",
              paddingRight: "12px",
              paddingTop: "12px",
              paddingBottom: "12px",
              boxSizing: "border-box",
            }}
          >
            <AEDataConfig onSwitchToFinal={() => { setActiveTab("Sheet1_AE"); setView("list"); }} />
          </motion.div>
        )}
      </AnimatePresence>
    </div>
  );
}
