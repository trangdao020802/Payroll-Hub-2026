import { TableRestoreButton } from '../../components/TableRestoreButton';
import { chooseExcelExport } from "../../components/ExportScopeDialog";
import { TransactionHistoryPanel } from "./components/TransactionHistoryPanel";
import { registerTableExport } from "../../lib/utils/table-excel";
import { getDeductionContribution } from "../../lib/utils/reconciliation-sync";
import { aggregateAdjustmentTotals } from "../../lib/utils/adjustment-summary";
import { downloadTransactionBankExport } from "../../lib/utils/excel-export";
import {
  canonicalTransactionHeaders,
  withCanonicalTransactionDocumentId,
} from "../../lib/utils/transaction-history";
/* eslint-disable @typescript-eslint/no-explicit-any, @typescript-eslint/no-unused-vars */
import React, { useState, useCallback, useMemo, useEffect } from "react";
import { useAppData } from "../../lib/contexts/AppDataContext";
import {
  useBulkPaymentLogic,
  isPastMonthHold,
} from "../../hooks/useBulkPaymentLogic";
import { DEFAULT_CENTERS } from "../../constants";
import {
  CreditCard,
  PlayCircle,
  Trash2,
  Save,
  Download,
  CheckCircle2,
  AlertCircle,
  FileText,
  Settings,
  Search,
  RefreshCw,
  FileSpreadsheet,
  AlertTriangle,
  PanelLeftClose,
  PanelLeftOpen,
  Wrench,
  Plus,
  Eye,
  Menu,
  Filter,
  Check,
  ArrowLeft,
  ArrowRight,
  Coins,
  TrendingDown,
  Calendar,
  Copy,
  Sparkles,
  Layers,
  Table,
  Table2,
  BarChart2,
  Info,
  X,
  Scale,
  ChevronDown,
  ChevronLeft,
  ChevronRight,
  ChevronsLeft,
  ChevronsRight,
  ExternalLink,
  Zap,
  Maximize2,
} from "lucide-react";
import {
  parseMoneyToNumber,
  formatMoneyVND,
  formatIdNumber,
} from "../../lib/utils/data-utils";
import {
  buildBankAccountIndex,
  getBankAccount,
  normalizePayrollId,
} from "../../lib/utils/bank-account-resolver";
import {
  processBulkPaymentTotals,
  isBUOfBankType,
  BankType,
} from "../../lib/utils/payment-processor";
import * as XLSX from "xlsx";
import { Button } from "../../components/ui/button";
import {
  TableInitialMark,
  TableTitleRemainder,
} from "../../components/TableInitialMark";
import {
  Dialog,
  DialogContent,
  DialogDescription,
  DialogFooter,
  DialogHeader,
  DialogTitle,
} from "../../components/ui/dialog";
import { toast } from "sonner";
import {
  Popover,
  PopoverTrigger,
  PopoverContent,
} from "../../components/ui/popover";
import {
  Tooltip,
  TooltipTrigger,
  TooltipContent,
} from "../../components/ui/tooltip";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "../../components/ui/select";
import { DataTable, isAhpBuValue } from "../../components/DataTable";
import { BulkPaymentAnalytics } from "./components/BulkPaymentAnalytics";
import { buildBulkPaymentAnalytics } from "../../lib/utils/bulk-payment-analytics";
import { markTransactionSaved } from "../../lib/utils/transaction-activity";
import { applyLocalTransactionVersions } from "../../lib/utils/transaction-local-history";
import type { TransactionVersion } from "../../lib/transaction-history-store";
import {
  applyBulkReconciliationReferenceSync,
  applyTransactionReferenceSync,
  buildTransactionReferenceSyncPlan,
  hasPendingReconciliationReferenceSync,
  type TransactionReferenceCorrection,
  type TransactionRawTimesheetCorrection,
} from "../../lib/utils/transaction-reference-sync";
import { protectSavedTransactionIdentity } from "../../lib/utils/transaction-saved-fields";
import { motion, AnimatePresence } from "motion/react";
import {
  DropdownMenu,
  DropdownMenuContent,
  DropdownMenuItem,
  DropdownMenuLabel,
  DropdownMenuSeparator,
  DropdownMenuTrigger,
} from "../../components/ui/dropdown-menu";

const containerVariants = {
  hidden: { opacity: 0 },
  visible: {
    opacity: 1,
    transition: {
      staggerChildren: 0.1,
    },
  },
} as const;

const itemVariants = {
  hidden: { opacity: 0, y: 20 },
  visible: {
    opacity: 1,
    y: 0,
    transition: {
      type: "spring",
      stiffness: 100,
      damping: 15,
    },
  },
} as const;

function PayrollBowIcon({ className = "" }: { className?: string }) {
  return (
    <svg
      viewBox="0 0 24 24"
      className={className}
      aria-hidden="true"
      focusable="false"
    >
      <path
        d="M10.8 9.6C8.7 5.9 5.5 4.5 3.7 6.1c-1.8 1.6-.5 4.8 2.1 5.2 1.7.3 3.6-.3 5.2-1.1"
        fill="currentColor"
        opacity=".88"
      />
      <path
        d="M13.2 9.6c2.1-3.7 5.3-5.1 7.1-3.5 1.8 1.6.5 4.8-2.1 5.2-1.7.3-3.6-.3-5.2-1.1"
        fill="currentColor"
        opacity=".88"
      />
      <path
        d="m9.6 11.1-2.4 7.1 4.8-2.4 4.8 2.4-2.4-7.1"
        fill="currentColor"
        opacity=".68"
      />
      <rect x="9.2" y="8.3" width="5.6" height="4.8" rx="2.1" fill="currentColor" />
    </svg>
  );
}

const ALL_ANALYS_BUSINESS_UNITS = "__ALL_BUSINESS_UNITS__";

interface TransactionReferenceReturnContext {
  targetTable: "Sheet1_AE" | "Hold_AE";
  targetLabel: "Gross Pay" | "Deductions";
  sourceSearch: string;
  transactionSearch: string;
  transactionKey: string;
}

function readTransactionReferenceReturn(): TransactionReferenceReturnContext | null {
  try {
    const raw = sessionStorage.getItem("transaction_reference_return");
    if (!raw) return null;
    const parsed = JSON.parse(raw) as TransactionReferenceReturnContext;
    if (parsed.targetTable !== "Sheet1_AE" && parsed.targetTable !== "Hold_AE") {
      return null;
    }
    return parsed;
  } catch {
    return null;
  }
}

// Reconcile exposes only comparison and action columns. The internal
// sheet1Amount and holdAmount fields remain available to calculations, export,
// and sync logic but are intentionally excluded from display/column selection.
const RECONCILE_DISPLAY_COLUMN_KEYS = [
  "serialNo",
  "docId",
  "name",
  "accountNo",
  "benefitsAccountNo",
  "actualAmount",
  "expectedAmount",
  "variance",
  "processSync",
  "problems",
] as const;

const isIdColumnKey = (k: string): boolean => {
  if (!k) return false;
  const lower = String(k).trim().toLowerCase();

  return (
    lower === "id" ||
    lower === "_id" ||
    lower === "document id" ||
    lower === "id issuance date" ||
    lower === "id issuance" ||
    lower === "id number" ||
    lower === "place of issue" ||
    lower.includes("document id") ||
    lower.includes("id issuance") ||
    lower.includes("id number") ||
    lower.includes("place of issue") ||
    lower.includes("national id") ||
    lower.includes("citizen id") ||
    lower.includes("cmnd") ||
    lower.includes("cccd") ||
    lower.includes("mã id") ||
    lower.startsWith("id_") ||
    lower.startsWith("id ") ||
    lower.endsWith("_id") ||
    lower.endsWith(" id")
  );
};

export function BulkPayment({
  showLeftCard: propShowLeftCard,
  setShowLeftCard: propSetShowLeftCard,
  searchTerm: externalSearchTerm,
  onSearchTermChange: externalOnSearchTermChange,
  onTabChange,
}: {
  showLeftCard?: boolean;
  setShowLeftCard?: React.Dispatch<React.SetStateAction<boolean>>;
  searchTerm?: string;
  onSearchTermChange?: (val: string) => void;
  onTabChange?: (
    tab: "Sheet1_AE" | "Hold_AE" | "BulkPayment" | "Pivot" | "upload",
  ) => void;
} = {}) {
  const { appData, updateAppData } = useAppData();

  const {
    globalMonth,
    monthPeriod,
    holdPaymentDetails,
    calculationSummary,
    dynamicReportStats,
    remainingHoldByMonth,
    bankExportData,
    isGenerating,
    isSuccess,
    reportStats,
    isRefreshing,
    hasPendingTransactionEdits,
    handleGenerateReport,
    handleClearReport,
    handleExportExcel,
    handleCellChange,
    handleSaveTransactionEdits,
    handleDeleteRow,
    handleDeleteRows,
    handleRefresh,
    generateAllSummaryText,
    monMatchComp,
    isMonthInStrComp,
  } = useBulkPaymentLogic();

  const [activeBalanceSection, setActiveBalanceSection] = useState<string>("I");
  const [internalSearchTerm, setInternalSearchTerm] = useState("");
  const searchTerm =
    externalSearchTerm !== undefined ? externalSearchTerm : internalSearchTerm;
  const setSearchTerm =
    externalOnSearchTermChange !== undefined
      ? externalOnSearchTermChange
      : setInternalSearchTerm;
  const [debouncedSearch, setDebouncedSearch] = useState("");
  const [internalShowLeftCard, setInternalShowLeftCard] = useState(true);
  const [showClearDialog, setShowClearDialog] = useState(false);
  const [deleteConfirmTarget, setDeleteConfirmTarget] = useState<{
    row?: any;
    rows?: any[];
  } | null>(null);

  const [editingCell, setEditingCell] = useState<{
    id: string;
    field: "docId" | "accountNo" | "sourceDocId" | "sourceAccountNo";
  } | null>(null);
  const [editingValue, setEditingValue] = useState("");

  // Expose these as variables for the UI
  const {
    targetMonthLabelComp,
    monthShortStrComp,
    monthDashStrComp,
    currentMonthNumComp,
    currentYearNumComp,
  } = monthPeriod;

  const [activeLeftTab, setActiveLeftTab] = useState<
    "summary" | "adjustments" | "reconcile"
  >("summary");
  const [transactionReferenceReturn, setTransactionReferenceReturn] =
    useState<TransactionReferenceReturnContext | null>(() =>
      readTransactionReferenceReturn(),
    );
  const [rightPanelTab, setRightPanelTab] = useState<
    "table" | "reconcile" | "visuals"
  >(() => {
    if (readTransactionReferenceReturn()) return "table";
    const saved = localStorage.getItem("bulk_payment_right_tab");
    return saved === "reconcile" || saved === "table" || saved === "visuals"
      ? saved
      : "table";
  });
  const [analysSelectedBusiness, setAnalysSelectedBusiness] = useState(
    ALL_ANALYS_BUSINESS_UNITS,
  );
  const [analysSearchTerm, setAnalysSearchTerm] = useState("");
  const [analysSearchVisible, setAnalysSearchVisible] = useState(false);

  useEffect(() => {
    const handleSetRightTab = (e: any) => {
      if (e.detail && e.detail.tab) {
        setRightPanelTab(e.detail.tab);
        localStorage.setItem("bulk_payment_right_tab", e.detail.tab);
      }
    };
    window.addEventListener("bulk-payment-set-right-tab", handleSetRightTab);
    return () =>
      window.removeEventListener(
        "bulk-payment-set-right-tab",
        handleSetRightTab,
      );
  }, []);

  const handleBackFromTransactionReference = useCallback(() => {
    if (!transactionReferenceReturn || !onTabChange) return;
    const context = transactionReferenceReturn;
    sessionStorage.removeItem("transaction_reference_return");
    localStorage.setItem("master_ae_active_tab", context.targetTable);
    localStorage.setItem("master_ae_search", context.sourceSearch);
    setTransactionReferenceReturn(null);
    onTabChange(context.targetTable);
    window.dispatchEvent(
      new CustomEvent("master-ae-filter", {
        detail: {
          search: context.sourceSearch,
          from: "TransactionReference",
        },
      }),
    );
    toast.info(`Đã quay lại ${context.targetLabel} tại dòng vừa đồng bộ.`);
  }, [onTabChange, transactionReferenceReturn]);
  const [reconcileFilterStatus, setReconcileFilterStatus] = useState<
    "ALL" | "MATCHED" | "VARIANCE" | "MISSING_INFO" | "DUPLICATE" | ""
  >("");
  const [historyHasExceptions, setHistoryHasExceptions] = useState(false);
  const [viewingHistorySource, setViewingHistorySource] = useState(false);
  const handleHistoryReportStateChange = useCallback((hasExceptions: boolean, viewingSource: boolean) => {
    setHistoryHasExceptions(hasExceptions);
    setViewingHistorySource(viewingSource);
  }, []);
  const [reconcileSelectedBU, setReconcileSelectedBU] = useState<string>("ALL");
  const [reconcileSearchQuery, setReconcileSearchQuery] = useState<string>("");
  const [showNorthOnly, setShowNorthOnly] = useState(false);
  const [adjustmentFilter, setAdjustmentFilter] = useState<
    "ALL" | "HOLD" | "ADD" | "CANCEL"
  >("ALL");

  const [reconcileCurrentPage, setReconcileCurrentPage] = useState<number>(1);
  const [reconcileRowsPerPage, setReconcileRowsPerPage] = useState<number | "all">(20);
  const [selectedBUGroup, setSelectedBUGroup] = useState<string>("ALL");

  const showLeftCard =
    propShowLeftCard !== undefined ? propShowLeftCard : internalShowLeftCard;
  const setShowLeftCard =
    propSetShowLeftCard !== undefined
      ? propSetShowLeftCard
      : setInternalShowLeftCard;

  const [submittingBatchId, setSubmittingBatchId] = useState<string | null>(
    null,
  );
  const [approvedBatches, setApprovedBatches] = useState<
    Record<string, { receiptId: string; timestamp: string }>
  >({});

  const transactionBatches = useMemo(() => {
    return [
      {
        id: `BATCH-${globalMonth.replace(".", "")}-NORTH`,
        name: `Đợt Chi Lương BANK NORTH`,
        bankType: "BANK_NORTH" as BankType,
        sheet1Total: calculationSummary.sheet1Total,
        holdTotal: calculationSummary.holdTotal,
        grandTotal: calculationSummary.bankNorthTotal,
        status: approvedBatches[`BATCH-${globalMonth.replace(".", "")}-NORTH`]
          ? "APPROVED"
          : "READY_TO_EXPORT",
        description: "Toàn bộ Bank North (AHN + AHP + ATH + ATN + APT)",
        buses: ["AHN", "AHP", "ATH", "ATN", "APT"],
        txCount: bankExportData.length,
      },
      {
        id: `BATCH-${globalMonth.replace(".", "")}-OTHER`,
        name: `Đợt Chi Lương KHÁC`,
        bankType: "OTHER" as BankType,
        sheet1Total: 0,
        holdTotal: 0,
        grandTotal:
          calculationSummary.aeTotal - calculationSummary.bankNorthTotal,
        status: approvedBatches[`BATCH-${globalMonth.replace(".", "")}-OTHER`]
          ? "APPROVED"
          : "DRAFT",
        description: "Các bộ phận khác & Dự phòng",
        buses: ["OTHER"],
        txCount: 0,
      },
    ];
  }, [globalMonth, calculationSummary, bankExportData, approvedBatches]);

  const handleSubmitBatchToGateway = async (batch: any) => {
    setSubmittingBatchId(batch.id);
    try {
      const response = await fetch("/api/bulk-payments/submit", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          bankType: batch.bankType,
          batchId: batch.id,
          amount: batch.grandTotal,
          transactionsCount: batch.txCount,
          timestamp: new Date().toISOString(),
        }),
      });
      if (!response.ok) {
        throw new Error("Lỗi phản hồi từ server");
      }
      const result = await response.json();
      if (result.status === "success") {
        setApprovedBatches((prev) => ({
          ...prev,
          [batch.id]: {
            receiptId: result.data.receiptId,
            timestamp: result.data.timestamp,
          },
        }));
        toast.success(`Phê duyệt đợt chi lương thành công!`, {
          description: `Đợt ${batch.name} đã được phê duyệt thành công qua cổng thanh toán. Receipt: ${result.data.receiptId}`,
        });
      } else {
        throw new Error(result.error || "Giao dịch không thành công");
      }
    } catch (error: any) {
      console.error(error);
      toast.error("Thất bại", {
        description: `Không thể phê duyệt đợt thanh toán: ${error.message || error}`,
      });
    } finally {
      setSubmittingBatchId(null);
    }
  };

  const handleCopyReport = () => {
    const text = generateAllSummaryText();
    navigator.clipboard.writeText(text);
    toast.success("Đã copy báo cáo vào clipboard");
  };

  const handleCopyReconciliationSummary = () => {
    const month = appData.globalMonth || globalMonth || "01.2026";
    const rowCount = (appData.BankExport?.data || []).length;
    const sheet1Totals = dynamicReportStats?.sheet1Totals || {};
    const finalTotals = dynamicReportStats?.finalTotals || {};
    const holdAddItems = dynamicReportStats?.holdAddItems || [];
    const rows: Array<[string, string]> = [
      ["BÁO CÁO CHI TIẾT THEO BU", month],
      ["Số dòng dữ liệu:", `${rowCount} dòng`],
    ];

    if (activeBalanceSection === "I") {
      rows.push(["I. GROSS PAY", ""]);
      Object.entries(sheet1Totals)
        .filter(([, amount]) => Number(amount) !== 0)
        .forEach(([business, amount]) => {
          rows.push([
            `${business}:`,
            `${formatMoneyVND(Number(amount)).replace(" ₫", "")}`,
          ]);
        });
      const total = Object.values(sheet1Totals).reduce(
        (sum, amount) => sum + Number(amount || 0),
        0,
      );
      rows.push([
        "TOTAL GROSS PAY:",
        `${formatMoneyVND(total).replace(" ₫", "")}`,
      ]);
    } else if (activeBalanceSection === "II") {
      rows.push(["II. DEDUCTIONS", ""]);
      const grouped = new Map<string, number>();
      holdAddItems.forEach((item) => {
        const label = `${item.biz} ${item.type}:`;
        grouped.set(label, (grouped.get(label) || 0) + Number(item.amount || 0));
      });
      grouped.forEach((amount, label) => {
        rows.push([label, `${formatMoneyVND(amount).replace(" ₫", "")}`]);
      });
      const total = holdAddItems.reduce(
        (sum, item) => sum + item.contribution,
        0,
      );
      rows.push([
        "TOTAL DEDUCTIONS:",
        `${formatMoneyVND(total).replace(" ₫", "")}`,
      ]);
    } else if (activeBalanceSection === "III") {
      rows.push(["III. NET PAY", ""]);
      Object.entries(finalTotals)
        .filter(([, amount]) => Number(amount) !== 0)
        .forEach(([business, amount]) => {
          rows.push([
            `${business}:`,
            `${formatMoneyVND(Number(amount)).replace(" ₫", "")}`,
          ]);
        });
      const total = Object.values(finalTotals).reduce(
        (sum, amount) => sum + Number(amount || 0),
        0,
      );
      rows.push([
        "TOTAL NET PAY:",
        `${formatMoneyVND(total).replace(" ₫", "")}`,
      ]);
    } else {
      const netPay = Object.values(finalTotals).reduce(
        (sum, amount) => sum + Number(amount || 0),
        0,
      );
      const totalAcc = calculationSummary.calculatedTotal || netPay;
      const totalAe = calculationSummary.aeTotal || netPay;
      rows.push(
        ["IV. RECONCILIATION", ""],
        ["TỔNG AE:", `${formatMoneyVND(totalAe).replace(" ₫", "")}`],
        ["TỔNG ACC:", `${formatMoneyVND(totalAcc).replace(" ₫", "")}`],
        ["LỆCH (DIFF):", `${formatMoneyVND(totalAcc - totalAe).replace(" ₫", "")}`],
      );
    }

    navigator.clipboard.writeText(
      rows.map(([label, value]) => `${label}\t${value}`).join("\n"),
    );
    toast.success("Đã sao chép riêng phần báo cáo đang xem (2 cột Excel)");
  };

  const totalPayoutSum = useMemo(() => {
    const bankExportRows = appData.BankExport?.data || [];
    if (bankExportRows.length > 0) {
      const sum = bankExportRows.reduce((acc: number, r: any) => {
        return (
          acc +
          (parseMoneyToNumber(
            r["Payment Amount"] ??
              r["Amount"] ??
              r["TOTAL PAYMENT"] ??
              r["Số tiền"] ??
              r["Thành tiền"] ??
              0,
          ) || 0)
        );
      }, 0);
      if (sum > 0) return sum;
    }

    if (calculationSummary.calculatedTotal > 0)
      return calculationSummary.calculatedTotal;
    if (calculationSummary.aeTotal > 0) return calculationSummary.aeTotal;

    if (dynamicReportStats?.finalTotals) {
      const sum = Object.values(dynamicReportStats.finalTotals).reduce(
        (a, b) => a + (b || 0),
        0,
      );
      if (sum > 0) return sum;
    }
    return 0;
  }, [appData.BankExport?.data, calculationSummary, dynamicReportStats]);

  const handleAuditCellUpdate = useCallback(
    (item: any, field: string, value: any) => {
      updateAppData((prev) => {
        const newBankData = [
          ...(prev.BankExport?.data || prev.Bank_North_AE?.data || []),
        ];
        const bIndex = newBankData.findIndex(
          (r) => r === item.rawRow || r.id === item.rawRow.id,
        );
        if (bIndex !== -1) {
          newBankData[bIndex] = { ...newBankData[bIndex], [field]: value };
        }

        const cleanId = item.docId?.toLowerCase();
        const newSheet1 = [...prev.Sheet1_AE.data];
        const s1Index = newSheet1.findIndex((r) => {
          const rId = String(r["ID Number"] || r["Mã AE"] || "")
            .trim()
            .toLowerCase();
          return rId && cleanId && rId === cleanId;
        });
        if (s1Index !== -1) {
          if (field === "accountNo" || field === "benefitsAccountNo")
            newSheet1[s1Index]["Bank Account Number"] = value;
          if (field === "docId" || field === "grossPlusBenefitsId")
            newSheet1[s1Index]["ID Number"] = value;
        }

        const newHold = [...prev.Hold_AE.data];
        const hIndex = newHold.findIndex((r) => {
          const rId = String(r["ID Number"] || r["Mã AE"] || "")
            .trim()
            .toLowerCase();
          return rId && cleanId && rId === cleanId;
        });
        if (hIndex !== -1) {
          if (field === "accountNo" || field === "benefitsAccountNo")
            newHold[hIndex]["Bank Account Number"] = value;
          if (field === "docId" || field === "grossPlusBenefitsId")
            newHold[hIndex]["ID Number"] = value;
        }

        return {
          ...prev,
          BankExport: { ...prev.BankExport, data: newBankData },
          Sheet1_AE: { ...prev.Sheet1_AE, data: newSheet1 },
          Hold_AE: { ...prev.Hold_AE, data: newHold },
          TransactionActivity: markTransactionSaved(prev),
        };
      }, true, true);
      toast.success("Đã cập nhật và đồng bộ dữ liệu sang bảng gốc thành công!");
    },
    [updateAppData],
  );

  const handleAutoFillMissingAccount = useCallback(
    (item: any) => {
      if (!item.referenceTransactionKey) {
        toast.error("Không xác định được giao dịch tham chiếu duy nhất.");
        return;
      }

      updateAppData((prev) => {
        const transactionRows =
          prev.BankExport?.data?.length > 0
            ? prev.BankExport.data
            : prev.Bank_North_AE?.data || [];
        const result = applyTransactionReferenceSync({
          grossRows: prev.Sheet1_AE?.data || [],
          deductionRows: prev.Hold_AE?.data || [],
          transactionRows,
          // Same rule as the bulk lightning/menu action: Batch Payment is source.
          rawTimesheetRows: [],
          reportMonth: prev.globalMonth,
          transactionKeys: [item.referenceTransactionKey],
        });

        if (result.correctedCells === 0) {
          toast.info("Các trường tham chiếu đã khớp Transaction, không cần sửa.");
          return prev;
        }

        toast.success(
          `Đã đồng bộ ${result.correctedCells} ô trên ${result.correctedRows} dòng theo Transaction.`,
        );

        return {
          ...prev,
          // Process Sync publishes confirmed Batch Payment identity downstream only.
          Sheet1_AE: { ...prev.Sheet1_AE, data: result.grossRows },
          Hold_AE: { ...prev.Hold_AE, data: result.deductionRows },
        };
      }, true, true);
    },
    [updateAppData],
  );

  const handleApplyTransactionHistoryVersions = useCallback(
    (versions: TransactionVersion[]) => {
      updateAppData((prev) => applyLocalTransactionVersions(prev, versions), true, true);
    },
    [updateAppData],
  );

  const displayBankExportData = useMemo(() => {
    const rows = bankExportData || [];
    const sheet1Rows = appData.Sheet1_AE?.data || [];
    const holdRows = appData.Hold_AE?.data || [];

    return rows.map((row: any) => {
      const canonical = withCanonicalTransactionDocumentId(row);
      if (!canonical.BU && !canonical.Business) {
        let bu = String(row["_fileBank"] || row["Business"] || row["BU"] || "").trim();
        if (!bu) {
          const docId = String(canonical["Document ID"] || canonical["ID Number"] || "").trim();
          const accNo = String(canonical["Beneficiary Account No."] || canonical["Bank Account Number"] || "").trim();
          if (docId || accNo) {
            const matchedS1 = sheet1Rows.find((s: any) => {
              const sDoc = String(s["ID Number"] || s["CCCD"] || s["Document ID"] || "").trim();
              const sAcc = String(s["Bank Account Number"] || s["Beneficiary Account No."] || "").trim();
              return (docId && sDoc === docId) || (accNo && sAcc === accNo);
            });
            if (matchedS1) {
              bu = matchedS1.Business || matchedS1.BU || "";
            } else {
              const matchedHold = holdRows.find((h: any) => {
                const hDoc = String(h["ID Number"] || h["CCCD"] || h["Document ID"] || "").trim();
                const hAcc = String(h["Bank Account Number"] || h["Beneficiary Account No."] || "").trim();
                return (docId && hDoc === docId) || (accNo && hAcc === accNo);
              });
              if (matchedHold) {
                bu = matchedHold.Business || matchedHold.BU || "";
              }
            }
          }
        }
        if (!bu) {
          const det = String(row["Payment details"] || row["Beneficiary Name"] || "").toUpperCase();
          if (det.includes("AHP") || det.includes("HAI PHONG") || det.includes("HAIPHONG")) bu = "AHP";
          else if (det.includes("ATH") || det.includes("THANH HOA")) bu = "ATH";
          else if (det.includes("ATN") || det.includes("THAI NGUYEN")) bu = "ATN";
          else if (det.includes("APT") || det.includes("PHU THO")) bu = "APT";
          else bu = "AHN";
        }
        if (bu === "AHN_HP") bu = "AHP";
        canonical.BU = bu;
      }
      return canonical;
    });
  }, [bankExportData, appData.Sheet1_AE?.data, appData.Hold_AE?.data]);

  const reconciliationNeedsSync = useMemo(() => {
    const transactionRows = (bankExportData || []).length > 0
      ? bankExportData
      : appData.BankExport?.data?.length
        ? appData.BankExport.data
        : appData.Bank_North_AE?.data || [];
    if (!transactionRows.length) return false;
    return hasPendingReconciliationReferenceSync({
      grossRows: appData.Sheet1_AE?.data || [],
      deductionRows: appData.Hold_AE?.data || [],
      transactionRows,
      reportMonth: appData.globalMonth,
    });
  }, [bankExportData, appData.BankExport?.data, appData.Bank_North_AE?.data, appData.Sheet1_AE?.data, appData.Hold_AE?.data, appData.globalMonth]);

  const analysAnalytics = useMemo(() => {
    if (rightPanelTab !== "visuals" || displayBankExportData.length === 0) {
      return null;
    }

    return buildBulkPaymentAnalytics({
      sheet1Rows: appData.Sheet1_AE?.data || [],
      holdRows: appData.Hold_AE?.data || [],
      bankRows:
        appData.BankExport?.data?.length > 0
          ? appData.BankExport.data
          : appData.Bank_North_AE?.data || [],
      globalMonth: appData.globalMonth || "03.2026",
    });
  }, [
    appData.BankExport?.data,
    appData.Bank_North_AE?.data,
    appData.Hold_AE?.data,
    appData.globalMonth,
    appData.Sheet1_AE?.data,
    displayBankExportData.length,
    rightPanelTab,
  ]);

  const analysBusinessUnits = analysAnalytics?.businessUnits || [];

  const effectiveAnalysBusiness =
    analysSelectedBusiness === ALL_ANALYS_BUSINESS_UNITS ||
    analysSelectedBusiness === "EXCLUDE_AHP" ||
    analysBusinessUnits.includes(analysSelectedBusiness)
      ? analysSelectedBusiness
      : ALL_ANALYS_BUSINESS_UNITS;

  const bankExportTotal = useMemo(() => {
    return (appData.BankExport?.data || []).reduce((sum: number, r: any) => {
      return (
        sum +
        (parseMoneyToNumber(
          r["Payment Amount"] ??
            r["Amount"] ??
            r["TOTAL PAYMENT"] ??
            r["Số tiền"] ??
            r["Thành tiền"] ??
            0,
        ) || 0)
      );
    }, 0);
  }, [appData.BankExport?.data]);

  const hasDuplicateIds = useMemo(() => {
    if (!displayBankExportData || displayBankExportData.length === 0)
      return false;
    const seen = new Set<string>();
    for (const row of displayBankExportData) {
      const idVal =
        row["Document ID"] ||
        row["ID Number"] ||
        row["Document ID / CCCD"] ||
        row["ID"] ||
        row["CCCD"];
      if (idVal && String(idVal).trim() !== "") {
        const clean = String(idVal).trim().toLowerCase();
        if (seen.has(clean)) return true;
        seen.add(clean);
      }
    }
    return false;
  }, [displayBankExportData]);

  // Unified Reconciliation Audit computation across database tables
  const reconciliationAudit = useMemo(() => {
    const bankExportRows =
      appData.BankExport?.data?.length > 0
        ? appData.BankExport.data
        : appData.Bank_North_AE?.data || [];
    const sheet1Rows = appData.Sheet1_AE?.data || [];
    const holdRows = appData.Hold_AE?.data || [];
    const transactionReferencePlan = buildTransactionReferenceSyncPlan({
      grossRows: sheet1Rows,
      deductionRows: holdRows,
      transactionRows: bankExportRows,
      // Reconciliation Process Sync uses confirmed Batch Payment as the
      // authoritative identity source. RAWDATA must never repair Transaction here.
      rawTimesheetRows: [],
      reportMonth: appData.globalMonth,
    });
    const accountById = buildBankAccountIndex([
      { source: "Gross Pay", rows: sheet1Rows },
      { source: "Transaction", rows: appData.BankExport?.data || [] },
      { source: "Transaction", rows: appData.Bank_North_AE?.data || [] },
      { source: "HOLD AE", rows: holdRows },
    ]);

    const activeSheet1RowsList: any[] = [];
    const activeHoldRowsList: any[] = [];
    const matchedSheet1Rows = new Set<any>();
    const matchedHoldRows = new Set<any>();

    sheet1Rows.forEach((r: any) => {
      const rowMonthStr = String(
        r["Tháng báo cáo"] || r["_fileMonth"] || r["Tháng"] || r["Month"] || "",
      ).trim();
      const extracted = monMatchComp(rowMonthStr);
      if (extracted && extracted !== targetMonthLabelComp) return;
      if (!extracted && rowMonthStr && !isMonthInStrComp(rowMonthStr)) return;

      const id = String(
        r["ID Number"] ||
          r["Mã AE"] ||
          r["Mã ae"] ||
          r["Document ID"] ||
          r["CCCD"] ||
          "",
      )
        .trim()
        .toLowerCase();

      if (!id) return; // BỎ QUA NẾU TRỐNG ID NUMBER

      activeSheet1RowsList.push(r);
    });

    const holdNetByBU = new Map<string, number>();

    holdRows.forEach((r: any) => {
      const rowMonthStr = String(
        r["Tháng báo cáo"] || r["_fileMonth"] || r["Tháng"] || r["Month"] || "",
      ).trim();
      const extracted = monMatchComp(rowMonthStr);
      if (extracted && extracted !== targetMonthLabelComp) return;
      if (!extracted && rowMonthStr && !isMonthInStrComp(rowMonthStr)) return;

      const command = String(r["Lệnh"] || "")
        .trim()
        .toUpperCase();
      if (command === "-") return;

      const sheetSource = String(r["Sheet Source"] || "").toLowerCase();
      if (sheetSource.includes("sheet 1 ae") || sheetSource.includes("sheet 1"))
        return;
      if (r._dimmed) return;

      activeHoldRowsList.push(r);

      const signedAmount = getDeductionContribution(r, appData.globalMonth || "03.2026");

      let bu = r["BU"] || r["Business"] || "";
      if (bu) bu = String(bu).trim().toUpperCase();
      if (bu === "AHN_HP") bu = "AHP";

      if (!bu || bu === "Other") {
        const textToMatch = [
          r["Sheet Source"],
          r["CENTER NOTE"],
          r["Mã ae"],
          r["Note"],
          r["Full name"],
        ]
          .map((v) => String(v || "").toUpperCase())
          .join(" ");
        if (textToMatch.includes("HN") || textToMatch.includes("AHN"))
          bu = "AHN";
        else if (
          textToMatch.includes("AHP") ||
          textToMatch.includes("HAIPHONG")
        )
          bu = "AHP";
        else if (
          textToMatch.includes("ATH") ||
          textToMatch.includes("THANH HOA")
        )
          bu = "ATH";
        else if (
          textToMatch.includes("ATN") ||
          textToMatch.includes("THAI NGUYEN")
        )
          bu = "ATN";
        else if (textToMatch.includes("APT") || textToMatch.includes("PHU THO"))
          bu = "APT";
        else bu = "AHN";
      }

      if (bu) holdNetByBU.set(bu, (holdNetByBU.get(bu) || 0) + signedAmount);
    });

    const docIdCounts = new Map<string, number>();
    const bankAccountCounts = new Map<string, number>();
    transactionReferencePlan.effectiveTransactionRows.forEach((r: any) => {
      const docId = String(
        r["Document ID"] ||
          r["ID Number"] ||
          r["Document ID / CCCD"] ||
          r["CCCD"] ||
          "",
      )
        .trim()
        .toLowerCase();
      if (docId) {
        docIdCounts.set(docId, (docIdCounts.get(docId) || 0) + 1);
      }
      const account = String(
        r["Beneficiary Account No."] ||
          r["Bank Account Number"] ||
          r["Số tài khoản"] ||
          "",
      )
        .replace(/\s+/g, "")
        .trim()
        .toLowerCase();
      if (account) {
        bankAccountCounts.set(
          account,
          (bankAccountCounts.get(account) || 0) + 1,
        );
      }
    });

    const transactionAuditList: Array<{
      id: string;
      serialNo: string;
      name: string;
      docId: string;
      accountNo: string;
      bankName: string;
      bu: string;
      actualAmount: number;
      expectedAmount: number;
      sheet1Amount: number;
      holdAmount: number;
      variance: number;
      status:
        "MATCHED" | "VARIANCE" | "MISSING_INFO" | "DUPLICATE" | "NOT_IN_SHEET1";
      issues: string[];
      rawRow: any;
      benefitsAccountNo: string;
      grossPlusBenefitsId: string;
      targetTabForAccLink: "Sheet1_AE" | "Hold_AE";
      referenceTransactionKey: string;
      referenceCorrections: TransactionReferenceCorrection[];
      referenceTransactionCorrections: TransactionRawTimesheetCorrection[];
      referenceSyncReason: string;
    }> = [];

    let totalActualSum = 0;
    let totalExpectedSum = 0;
    let matchedCount = 0;
    let varianceCount = 0;
    let missingInfoCount = 0;
    let duplicateCount = 0;
    let notInSheet1Count = 0;

    bankExportRows.forEach((row: any, index: number) => {
      const serialNo = String(row["Payment Serial Number"] || index + 1);
      const name = String(
        row["Beneficiary Name"] || row["Full name"] || row["Họ tên"] || "N/A",
      ).trim();
      const rawDocId = String(
        row["Document ID"] ||
          row["ID Number"] ||
          row["Document ID / CCCD"] ||
          row["CCCD"] ||
          "",
      ).trim();
      const accountNo = String(
        row["Beneficiary Account No."] ||
          row["Bank Account Number"] ||
          row["Số tài khoản"] ||
          "",
      ).trim();
      const bankName = String(
        row["Beneficiary Bank Swift Code / IFSC Code"] ||
          row["Beneficiary Bank"] ||
          row["Ngân hàng"] ||
          "",
      ).trim();
      let bu = String(
        row["_fileBank"] || row["Business"] || row["BU"] || "Other",
      ).trim();
      if (bu === "AHN_HP") bu = "AHP";

      const actualAmount =
        parseMoneyToNumber(
          row["Payment Amount"] ??
            row["Amount"] ??
            row["TOTAL PAYMENT"] ??
            row["Số tiền"] ??
            0,
        ) || 0;

      totalActualSum += actualAmount;

      const referenceMatch = transactionReferencePlan.byTransactionIndex.get(index);
      const effectiveTransactionRow = referenceMatch?.transactionRow || row;
      const effectiveDocId = String(
        effectiveTransactionRow["Document ID"] ||
          effectiveTransactionRow["ID Number"] ||
          effectiveTransactionRow["Document ID / CCCD"] ||
          effectiveTransactionRow["CCCD"] ||
          "",
      )
        .trim()
        .toLowerCase();
      const effectiveAccount = String(
        effectiveTransactionRow["Beneficiary Account No."] ||
          effectiveTransactionRow["Bank Account Number"] ||
          effectiveTransactionRow["Số tài khoản"] ||
          "",
      )
        .replace(/\s+/g, "")
        .trim()
        .toLowerCase();
      const matchedSheet1RowsList = (referenceMatch?.grossRowIndexes || [])
        .map((rowIndex) => sheet1Rows[rowIndex])
        .filter(Boolean);
      const matchedHoldRowsList = (referenceMatch?.deductionRowIndexes || [])
        .map((rowIndex) => holdRows[rowIndex])
        .filter(Boolean);

      // Extract true ID Number from matched Sheet1 / Hold row if rawDocId is missing or equal to bank account
      let displayDocId = formatIdNumber(rawDocId);
      const primaryMatchedRow =
        matchedSheet1RowsList[0] || matchedHoldRowsList[0];
      if (!displayDocId && primaryMatchedRow) {
        const realIdFromSheet = formatIdNumber(
          primaryMatchedRow["ID Number"] ||
            primaryMatchedRow["Mã AE"] ||
            primaryMatchedRow["Mã ae"] ||
            primaryMatchedRow["CCCD"] ||
            primaryMatchedRow["Document ID"] ||
            "",
        );
        if (realIdFromSheet) {
          displayDocId = realIdFromSheet;
        }
      }

      if (displayDocId === accountNo || !displayDocId) {
        displayDocId = "";
      }

      let sheet1Amount = 0;
      const issues: string[] = [];

      if ((referenceMatch?.transactionCorrections.length || 0) > 0) {
        const fields = Array.from(
          new Set(
            referenceMatch!.transactionCorrections.map((correction) =>
              correction.field === "idNumber"
                ? "ID Number"
                : correction.field === "fullName"
                  ? "Full Name"
                  : "Bank Account Number",
            ),
          ),
        );
        issues.push(
          `Cần sửa ${fields.join(", ")} của Transaction theo RAWDATA_TIMESHEET`,
        );
      }

      if ((referenceMatch?.corrections.length || 0) > 0) {
        const fields = Array.from(
          new Set(
            referenceMatch!.corrections.map((correction) =>
              correction.field === "idNumber"
                ? "ID Number"
                : correction.field === "fullName"
                  ? "Full Name"
                  : "Bank Account Number",
            ),
          ),
        );
        issues.push(`Cần đồng bộ ${fields.join(", ")} theo Transaction`);
      }

      if (matchedSheet1RowsList.length > 0) {
        matchedSheet1RowsList.forEach((r) => {
          matchedSheet1Rows.add(r);
          const amt =
            parseMoneyToNumber(
              r["TOTAL PAYMENT"] ??
                r["Grand Total"] ??
                r["GRAND TOTAL"] ??
                r["Payment Amount"] ??
                0,
            ) || 0;
          sheet1Amount += amt;
        });
      } else {
        notInSheet1Count++;
        issues.push("Không tìm thấy trong Sheet1 AE");
      }

      let holdAmount = 0;
      matchedHoldRowsList.forEach((r) => {
        matchedHoldRows.add(r);
        holdAmount += getDeductionContribution(r, appData.globalMonth || "03.2026");
      });

      // Target = Sheet1 + Hold AE
      const expectedAmount = sheet1Amount + holdAmount;

      totalExpectedSum += expectedAmount;
      const variance = actualAmount - expectedAmount;

      let status:
        | "MATCHED"
        | "VARIANCE"
        | "MISSING_INFO"
        | "DUPLICATE"
        | "NOT_IN_SHEET1" = "MATCHED";

      if (!accountNo || accountNo.length < 3) {
        status = "MISSING_INFO";
        issues.push("Thiếu/Sai số tài khoản");
        missingInfoCount++;
      } else if (
        effectiveDocId &&
        (docIdCounts.get(effectiveDocId) || 0) > 1
      ) {
        status = "DUPLICATE";
        issues.push(
          `Trùng Document ID (${docIdCounts.get(effectiveDocId)} lần)`,
        );
        duplicateCount++;
      } else if (
        effectiveAccount &&
        (bankAccountCounts.get(effectiveAccount) || 0) > 1
      ) {
        status = "DUPLICATE";
        issues.push(
          `Trùng Bank Account (${bankAccountCounts.get(effectiveAccount)} lần)`,
        );
        duplicateCount++;
      } else if (Math.abs(variance) >= 1) {
        status = "VARIANCE";
        if (sheet1Amount > 0 && holdAmount !== 0) {
          issues.push(
            `Lệch số tiền (Sheet1: ${formatMoneyVND(sheet1Amount)}, Hold: ${holdAmount >= 0 ? "+" : ""}${formatMoneyVND(holdAmount)})`,
          );
        } else {
          issues.push(
            `Lệch số tiền (${variance > 0 ? "+" : ""}${formatMoneyVND(variance)})`,
          );
        }
        varianceCount++;
      } else if (
        (referenceMatch?.corrections.length || 0) > 0 ||
        (referenceMatch?.transactionCorrections.length || 0) > 0
      ) {
        status = "MISSING_INFO";
        missingInfoCount++;
      } else {
        status = "MATCHED";
        matchedCount++;
      }

      const accountLookupId = normalizePayrollId(
        primaryMatchedRow?.["ID Number"] ||
          primaryMatchedRow?.["Mã AE"] ||
          displayDocId,
      );
      const benefitsAccountNo =
        getBankAccount(primaryMatchedRow) ||
        accountById.get(accountLookupId)?.accountNumber ||
        getBankAccount(row);
      const grossPlusBenefitsId = formatIdNumber(
        displayDocId ||
          primaryMatchedRow?.["ID Number"] ||
          primaryMatchedRow?.["Mã AE"] ||
          "",
      );

      const s1Acc = String(
        matchedSheet1RowsList[0]?.["Bank Account Number"] ||
          matchedSheet1RowsList[0]?.["Beneficiary Account No."] ||
          "",
      )
        .replace(/\s+/g, "")
        .trim()
        .toLowerCase();
      const hAcc = String(
        matchedHoldRowsList[0]?.["Bank Account Number"] ||
          matchedHoldRowsList[0]?.["Beneficiary Account No."] ||
          "",
      )
        .replace(/\s+/g, "")
        .trim()
        .toLowerCase();
      const bAcc = accountNo.replace(/\s+/g, "").trim().toLowerCase();

      const s1HasMatch = matchedSheet1RowsList.length > 0;
      const holdHasMatch = matchedHoldRowsList.length > 0;

      const s1AccIsRight = s1HasMatch && !!bAcc && s1Acc === bAcc;
      const holdAccIsRight = holdHasMatch && !!bAcc && hAcc === bAcc;

      let targetTabForAccLink: "Sheet1_AE" | "Hold_AE" = "Sheet1_AE";

      const firstReferenceCorrection = referenceMatch?.corrections[0];
      if (firstReferenceCorrection) {
        targetTabForAccLink = firstReferenceCorrection.table;
      }

      if (!firstReferenceCorrection && s1AccIsRight) {
        // Sheet 1 AE có ID Number và Bank Acc No giống với bảng Transaction rồi -> Bỏ qua
        if (holdAccIsRight) {
          // Nếu cả 2 bảng đều đúng thì mặc định đến bảng Sheet 1 AE
          targetTabForAccLink = "Sheet1_AE";
        } else {
          // Đi đến bảng Hold AE để tìm xem có khác giá trị về Bank Acc No hay chưa có -> Đến Hold AE lọc ID Number
          targetTabForAccLink = "Hold_AE";
        }
      } else if (!firstReferenceCorrection) {
        // Sheet 1 AE chưa đúng hoặc chưa có
        if (holdAccIsRight) {
          // Hold AE đúng nhưng Sheet 1 AE chưa đúng -> đến Sheet 1 AE
          targetTabForAccLink = "Sheet1_AE";
        } else {
          // Cả 2 đều chưa đúng hoặc chưa có
          if (!s1HasMatch && holdHasMatch) {
            targetTabForAccLink = "Hold_AE";
          } else {
            targetTabForAccLink = "Sheet1_AE";
          }
        }
      }

      transactionAuditList.push({
        id: `tx-${index}`,
        serialNo,
        name,
        docId: displayDocId,
        accountNo,
        bankName,
        bu,
        actualAmount,
        expectedAmount,
        sheet1Amount,
        holdAmount,
        variance,
        status,
        issues,
        rawRow: row,
        benefitsAccountNo,
        grossPlusBenefitsId,
        targetTabForAccLink,
        referenceTransactionKey: referenceMatch?.transactionKey || "",
        referenceCorrections: referenceMatch?.corrections || [],
        referenceTransactionCorrections:
          referenceMatch?.transactionCorrections || [],
        referenceSyncReason: referenceMatch?.reason || "unmatched",
      });
    });

    const buList = ["AHN", "AHP", "ATH", "ATN", "APT", "OTHER"];
    const buMatrix: Record<
      string,
      {
        bu: string;
        expectedTotal: number;
        sheet1Total: number;
        holdTotal: number;
        actualTotal: number;
        variance: number;
        txCount: number;
        matchedTxCount: number;
        status: "MATCHED" | "VARIANCE";
      }
    > = {};

    buList.forEach((bu) => {
      const sheet1ForBu = dynamicReportStats?.sheet1Totals?.[bu] || 0;
      const holdForBu = holdNetByBU.get(bu) || 0;
      const expected = sheet1ForBu + holdForBu;

      const buTxs = transactionAuditList.filter((t) => t.bu === bu);
      const actual = buTxs.reduce((sum, t) => sum + t.actualAmount, 0);
      const varAmount = actual - expected;
      const matchedCountInBu = buTxs.filter(
        (t) => t.status === "MATCHED",
      ).length;

      buMatrix[bu] = {
        bu,
        expectedTotal: expected,
        sheet1Total: sheet1ForBu,
        holdTotal: holdForBu,
        actualTotal: actual,
        variance: varAmount,
        txCount: buTxs.length,
        matchedTxCount: matchedCountInBu,
        status: Math.abs(varAmount) < 1 ? "MATCHED" : "VARIANCE",
      };
    });

    const targetNetTotal =
      calculationSummary.calculatedTotal ?? totalExpectedSum;
    const matchedExpectedSum = totalExpectedSum;
    const unexportedAmount = targetNetTotal - matchedExpectedSum;
    const netVariance = totalActualSum - targetNetTotal;

    const getHoldRowExpectedAmount = (r: any) =>
      getDeductionContribution(r, appData.globalMonth || "03.2026");

    const unmatchedSheet1Rows = activeSheet1RowsList.filter(
      (r) => !matchedSheet1Rows.has(r),
    );
    const unmatchedHoldRows = activeHoldRowsList.filter(
      (r) => !matchedHoldRows.has(r),
    );

    let finalVarianceCount = varianceCount;
    let finalMatchedCount = matchedCount;

    // We will group unmatched rows by docId to see if Sheet 1 and Hold AE cancel each other out
    const unmatchedMap = new Map<string, any>();

    unmatchedSheet1Rows.forEach((r, idx) => {
      const expectedAmount =
        parseMoneyToNumber(
          r["TOTAL PAYMENT"] ??
            r["Grand Total"] ??
            r["GRAND TOTAL"] ??
            r["Payment Amount"] ??
            0,
        ) || 0;

      if (Math.abs(expectedAmount) < 1) return;

      const name = String(
        r["Full name"] || r["Beneficiary Name"] || "N/A",
      ).trim();
      const rawDocId = String(
        r["ID Number"] ||
          r["Mã AE"] ||
          r["Mã ae"] ||
          r["CCCD"] ||
          r["Document ID"] ||
          "",
      ).trim();
      const displayDocId = formatIdNumber(rawDocId);
      const accountNo =
        getBankAccount(r) ||
        accountById.get(normalizePayrollId(displayDocId))?.accountNumber ||
        "";
      const bankName = String(
        r["Beneficiary Bank Swift Code / IFSC Code"] ||
          r["Beneficiary Bank"] ||
          r["Ngân hàng"] ||
          "",
      ).trim();
      let bu = String(r["BU"] || r["Business"] || "Other").trim();
      if (bu === "AHN_HP") bu = "AHP";

      const key = (displayDocId || name).toLowerCase();

      if (!unmatchedMap.has(key)) {
        unmatchedMap.set(key, {
          id: `unmatched-combined-${idx}`,
          serialNo: "DISC",
          name,
          docId: displayDocId,
          accountNo,
          bankName,
          bu,
          actualAmount: 0,
          expectedAmount: 0,
          sheet1Amount: 0,
          holdAmount: 0,
          variance: 0,
          status: "VARIANCE",
          issues: [],
          rawRow: r,
          benefitsAccountNo: accountNo,
          grossPlusBenefitsId: displayDocId,
          targetTabForAccLink: "Sheet1_AE",
        });
      }

      const item = unmatchedMap.get(key);
      item.sheet1Amount += expectedAmount;
      item.expectedAmount += expectedAmount;
      item.variance -= expectedAmount;
    });

    unmatchedHoldRows.forEach((r, idx) => {
      const expectedAmount = getHoldRowExpectedAmount(r);

      if (Math.abs(expectedAmount) < 1) return;

      const name = String(
        r["Full name"] || r["Beneficiary Name"] || "N/A",
      ).trim();
      const rawDocId = String(
        r["ID Number"] ||
          r["Mã AE"] ||
          r["Mã ae"] ||
          r["CCCD"] ||
          r["Document ID"] ||
          "",
      ).trim();
      const displayDocId = formatIdNumber(rawDocId);
      const accountNo =
        getBankAccount(r) ||
        accountById.get(normalizePayrollId(displayDocId))?.accountNumber ||
        "";
      const bankName = String(
        r["Beneficiary Bank Swift Code / IFSC Code"] ||
          r["Beneficiary Bank"] ||
          r["Ngân hàng"] ||
          "",
      ).trim();
      let bu = String(r["BU"] || r["Business"] || "Other").trim();
      if (bu === "AHN_HP") bu = "AHP";

      const key = (displayDocId || name).toLowerCase();

      if (!unmatchedMap.has(key)) {
        unmatchedMap.set(key, {
          id: `unmatched-combined-hold-${idx}`,
          serialNo: "DISC",
          name,
          docId: displayDocId,
          accountNo,
          bankName,
          bu,
          actualAmount: 0,
          expectedAmount: 0,
          sheet1Amount: 0,
          holdAmount: 0,
          variance: 0,
          status: "VARIANCE",
          issues: [],
          rawRow: r,
          benefitsAccountNo: accountNo,
          grossPlusBenefitsId: displayDocId,
          targetTabForAccLink: "Hold_AE",
        });
      }

      const item = unmatchedMap.get(key);
      item.holdAmount += expectedAmount;
      item.expectedAmount += expectedAmount;
      item.variance -= expectedAmount;
    });

    // Now push them to transactionAuditList and update counts
    for (const item of unmatchedMap.values()) {
      if (Math.abs(item.variance) < 1) {
        item.status = "MATCHED";
        item.issues.push(
          "Sheet 1 và Hold bù trừ hết (Target = 0), trùng khớp với Bank AE (0)",
        );
        finalMatchedCount++;
      } else {
        item.status = "VARIANCE";
        item.issues.push(
          "Bảng nguồn có dữ liệu nhưng Bank Export không có, không bù trừ hết (Target ≠ 0)",
        );
        finalVarianceCount++;
      }
      transactionAuditList.push(item);
    }

    const isBankRowsFullyMatched =
      varianceCount === 0 && missingInfoCount === 0 && duplicateCount === 0;

    return {
      transactionAuditList,
      totalActualSum,
      totalExpectedSum: targetNetTotal,
      matchedExpectedSum,
      unexportedAmount,
      netVariance,
      matchedCount: finalMatchedCount,
      varianceCount: finalVarianceCount,
      missingInfoCount,
      duplicateCount,
      notInSheet1Count,
      buMatrix,
      isBankRowsFullyMatched,
      isFullyMatched: Math.abs(netVariance) < 1 && isBankRowsFullyMatched,
    };
  }, [
    appData.BankExport?.data,
    appData.Bank_North_AE?.data,
    appData.Sheet1_AE?.data,
    appData.Hold_AE?.data,
    appData.Q_Staff,
    appData.globalMonth,
    dynamicReportStats,
    appData.Timesheet_Roster,
    calculationSummary,
    targetMonthLabelComp,
    monMatchComp,
    isMonthInStrComp,
  ]);

  const transactionPeriodKey = `Tháng ${currentMonthNumComp}/${currentYearNumComp}`;
  const syncedTransactionVersion = appData.TrialBalanceTransactionVersions?.[transactionPeriodKey] || 0;

  useEffect(() => {
    const saveVersion = appData.TransactionActivity?.saveVersion || 0;
    if (
      saveVersion <= syncedTransactionVersion ||
      Math.abs(reconciliationAudit.netVariance) >= 1
    ) {
      return;
    }

    updateAppData((prev) => {
      const latestSaveVersion = prev.TransactionActivity?.saveVersion || 0;
      const currentSyncedVersion =
        prev.TrialBalanceTransactionVersions?.[transactionPeriodKey] || 0;
      if (
        latestSaveVersion <= currentSyncedVersion
      ) {
        return prev;
      }
      return {
        ...prev,
        TrialBalanceTransactionVersion: latestSaveVersion,
        TrialBalanceTransactionVersions: {
          ...(prev.TrialBalanceTransactionVersions || {}),
          [transactionPeriodKey]: latestSaveVersion,
        },
        TrialBalanceRefreshedAt:
          prev.TransactionActivity?.lastSavedAt || new Date().toISOString(),
      };
    });
  }, [
    appData.TransactionActivity?.saveVersion,
    syncedTransactionVersion,
    transactionPeriodKey,
    reconciliationAudit.netVariance,
    updateAppData,
  ]);

  const activeIssueCategoriesCount = 
    (reconciliationAudit.varianceCount > 0 ? 1 : 0) + 
    (reconciliationAudit.missingInfoCount > 0 ? 1 : 0) + 
    (reconciliationAudit.duplicateCount > 0 ? 1 : 0);

  const shouldShowFilterDiv = activeIssueCategoriesCount > 1;

  const effectiveReconcileFilterStatus = historyHasExceptions && reconcileFilterStatus === "MATCHED" ? "ALL" : reconcileFilterStatus !== ""
    ? reconcileFilterStatus 
    : (reconciliationAudit.varianceCount > 0 ? "VARIANCE" 
       : reconciliationAudit.duplicateCount > 0 ? "DUPLICATE" 
       : reconciliationAudit.missingInfoCount > 0 ? "MISSING_INFO" 
       : "ALL");

  const filteredTransactionAudits = useMemo(() => {
    return reconciliationAudit.transactionAuditList.filter((item) => {
      if (historyHasExceptions && item.status === "MATCHED") return false;
      if (reconcileSelectedBU === "EXCLUDE_AHP") {
        if (isAhpBuValue(item.bu)) {
          return false;
        }
      } else if (reconcileSelectedBU !== "ALL" && item.bu !== reconcileSelectedBU) {
        return false;
      }
      if (effectiveReconcileFilterStatus === "ALL") {
        if (item.status === "MATCHED") {
          return false;
        }
      } else if (item.status !== effectiveReconcileFilterStatus) {
        return false;
      }
      if (reconcileSearchQuery.trim()) {
        const q = reconcileSearchQuery.trim().toLowerCase();
        const qClean = q.replace(/^0+/, "");
        const matchName = item.name.toLowerCase().includes(q);
        const matchDocId =
          item.docId.toLowerCase().includes(q) ||
          formatIdNumber(item.docId).toLowerCase().includes(q) ||
          (qClean && item.docId.toLowerCase().includes(qClean)) ||
          (item.grossPlusBenefitsId &&
            item.grossPlusBenefitsId.toLowerCase().includes(q));
        const matchAcc =
          item.accountNo.toLowerCase().includes(q) ||
          (item.benefitsAccountNo &&
            item.benefitsAccountNo.toLowerCase().includes(q));
        const matchBu = item.bu.toLowerCase().includes(q);
        if (!matchName && !matchDocId && !matchAcc && !matchBu) {
          return false;
        }
      }
      return true;
    });
  }, [
    reconciliationAudit.transactionAuditList,
    historyHasExceptions,
    reconcileSelectedBU,
    effectiveReconcileFilterStatus,
    reconcileSearchQuery,
  ]);

  const totalItems = filteredTransactionAudits.length;
  const itemsPerPage = reconcileRowsPerPage === "all" ? totalItems : reconcileRowsPerPage;
  const totalPages = itemsPerPage > 0 ? Math.ceil(totalItems / itemsPerPage) : 1;
  const safePage = Math.min(reconcileCurrentPage, totalPages) || 1;

  const paginatedTransactionAudits = useMemo(() => {
    if (reconcileRowsPerPage === "all") return filteredTransactionAudits;
    const start = (safePage - 1) * itemsPerPage;
    return filteredTransactionAudits.slice(start, start + itemsPerPage);
  }, [filteredTransactionAudits, safePage, itemsPerPage, reconcileRowsPerPage]);

  const handleSyncTransactionFieldsToTables = useCallback(() => {
    if (hasPendingTransactionEdits) {
      toast.warning("Bấm Lưu sửa trong Batch Payment trước khi đồng bộ STK và ID.");
      return;
    }

    const transactionRows =
      appData.BankExport?.data?.length > 0
        ? appData.BankExport.data
        : appData.Bank_North_AE?.data || [];

    if (!transactionRows.length) {
      toast.info("Chưa có dữ liệu Batch Payment để đồng bộ.");
      return;
    }

    // Canonical shared action: the yellow Reconciliation lightning and every
    // "Đồng bộ từ Reconcile" entry point execute this exact same pipeline.
    const result = applyBulkReconciliationReferenceSync({
      grossRows: appData.Sheet1_AE?.data || [],
      deductionRows: appData.Hold_AE?.data || [],
      transactionRows,
      reportMonth: appData.globalMonth,
    });

    if (result.correctedCells === 0) {
      toast.info("Không còn dòng Reconciliation nào cần Process Sync.");
      return;
    }

    updateAppData((prev) => ({
      ...prev,
      // Batch Payment remains authoritative and untouched.
      Sheet1_AE: { ...prev.Sheet1_AE, data: result.grossRows },
      Hold_AE: { ...prev.Hold_AE, data: result.deductionRows },
    }), true, true);

    toast.success(
      `Đã đồng bộ ${result.correctedCells} ô trên ${result.correctedRows} dòng bằng cùng logic Process Sync. Các dòng đã khớp sẽ tự biến mất khỏi danh sách cần xử lý.`,
    );

    // Saving/checking the monthly Batch Payment snapshot remains the job of
    // "Lưu tháng" / "Check STK & ID".
  }, [
    hasPendingTransactionEdits,
    appData.BankExport?.data,
    appData.Bank_North_AE?.data,
    appData.Sheet1_AE?.data,
    appData.Hold_AE?.data,
    appData.globalMonth,
    updateAppData,
  ]);

  const reconcileTotals = useMemo(() => {
    return filteredTransactionAudits.reduce(
      (acc, item) => {
        const ae = Number(item.actualAmount) || 0;
        const accAmt = Number(item.sheet1Amount + item.holdAmount) || 0;
        const diff = Number(item.variance) || 0;
        return {
          ae: acc.ae + ae,
          accAmt: acc.accAmt + accAmt,
          diff: acc.diff + diff,
        };
      },
      { ae: 0, accAmt: 0, diff: 0 },
    );
  }, [filteredTransactionAudits]);

  const reconciliationExportRows = useMemo(() => filteredTransactionAudits.map(t => ({
      "No.": t.id.startsWith("unmatched-") ? "DISC" : t.serialNo,
      "ID NUMBER": t.docId,
      "FULL NAME": t.name,
      "Bank Acc No. from AE": t.accountNo || "",
      "Bank Acc No. from ACC": t.benefitsAccountNo || t.accountNo || "",
      "TOTAL BANK AE": t.actualAmount,
      "TOTAL BANK ACC": t.sheet1Amount + t.holdAmount,
      "Diff": t.variance,
      "Process Sync": ((t.referenceCorrections?.length || 0) + (t.referenceTransactionCorrections?.length || 0)) > 0 ? "Sync" : "",
      "Problems": t.issues.join("; "),
    })), [filteredTransactionAudits]);
  useEffect(() => registerTableExport("master-reconcile", () => ({
    schema: { columns: ["No.", "ID NUMBER", "FULL NAME", "Bank Acc No. from AE", "Bank Acc No. from ACC", "TOTAL BANK AE", "TOTAL BANK ACC", "Diff", "Process Sync", "Problems"].map(key => ({key, label: key, type: /^(TOTAL|Diff)/.test(key) ? "currency" : "text"})), hiddenColumns: [] },
    rows: reconciliationExportRows,
  })), [reconciliationExportRows]);

  const handleExportReconciliationExcel = () => {
    const wb = XLSX.utils.book_new();

    const ws2 = XLSX.utils.json_to_sheet(reconciliationExportRows);
    XLSX.utils.book_append_sheet(wb, ws2, "Chi_Tiet_Doi_Soat_Giao_Dich");

    XLSX.writeFile(
      wb,
      `Bao_Cao_Doi_Soat_Thanh_Toan_${globalMonth.replace(".", "_")}.xlsx`,
    );
    toast.success("Đã xuất file Excel Đối soát Thanh toán thành công!");
  };

  const columns = useMemo(() => {
    const baseHeaders =
      appData.BankExport?.headers && appData.BankExport.headers.length > 0
        ? [...appData.BankExport.headers]
        : [
            "Payment Serial Number",
            "Transaction Type Code",
            "Payment Type",
            "Customer Reference No",
            "Beneficiary Account No.",
            "Beneficiary Name",
            "Document ID",
            "Place of Issue",
            "ID Issuance Date",
            "Beneficiary Bank Swift Code / IFSC Code",
            "Transaction Currency",
            "Payment Amount",
            "Charge Type",
            "Payment details",
          ];

    if (displayBankExportData && displayBankExportData.length > 0) {
      const allKeys = Object.keys(displayBankExportData[0]);
      allKeys.forEach((key) => {
        const kUp = key.toUpperCase();
        if (
          !key.startsWith("_") &&
          kUp !== "ID" &&
          kUp !== "_ID" &&
          kUp !== "UUID" &&
          kUp !== "ROWID" &&
          kUp !== "RECORDID" &&
          !baseHeaders.some((h) => String(h).toUpperCase() === kUp)
        ) {
          baseHeaders.push(key);
        }
      });
    }

    let cleanBaseHeaders = canonicalTransactionHeaders(baseHeaders).filter(
      (h) => {
        const u = String(h).trim().toUpperCase();
        return (
          u !== "ID" &&
          u !== "_ID" &&
          u !== "UUID" &&
          u !== "ROWID" &&
          u !== "RECORDID" &&
          !u.startsWith("_") &&
          u !== "THÁNG BÁO CÁO" &&
          u !== "THÁNG BÁO CÁO (SHEET 1)" &&
          !u.includes("THÁNG BÁO CÁO")
        );
      },
    );

    const isNoCol = (h: string) => {
      const u = String(h).trim().toUpperCase();
      return (
        u === "NO." ||
        u === "NO" ||
        u === "STT" ||
        u === "PAYMENT SERIAL NUMBER"
      );
    };

    const firstNoIdx = cleanBaseHeaders.findIndex(isNoCol);
    if (firstNoIdx !== -1) {
      const actualNo = cleanBaseHeaders[firstNoIdx];
      cleanBaseHeaders = cleanBaseHeaders.filter(
        (h, idx) => idx === firstNoIdx || !isNoCol(h),
      );
      cleanBaseHeaders = [
        actualNo,
        ...cleanBaseHeaders.filter((h) => h !== actualNo),
      ];
    }

    return cleanBaseHeaders.map((header) => {
      const h = String(header).toUpperCase();
      let type: "text" | "number" | "currency" = "text";
      if (
        h.includes("AMOUNT") ||
        h.includes("PAYMENT AMOUNT") ||
        h.includes("TOTAL")
      ) {
        if (
          !h.includes("ACCOUNT") &&
          !h.includes("NO") &&
          !h.includes("NUMBER") &&
          !h.includes("ID") &&
          !h.includes("CODE")
        ) {
          type = "currency";
        }
      }

      return {
        key: header,
        label: header,
        type,
        align: type === "currency" ? ("right" as const) : ("left" as const),
      };
    });
  }, [appData.BankExport?.headers, displayBankExportData]);

  const useMemoHiddenBuColumns = useMemo(
    () => columns.filter((column) => column.key.toUpperCase() === "BU").map((column) => column.key),
    [columns],
  );

  useEffect(() => registerTableExport("master-transaction", () => ({
    schema: {
      columns: columns
        .filter((c) => !c.key.startsWith("_") && !/tháng\s*báo\s*cáo/i.test(c.key))
        .map((c) => ({ key: c.key, label: c.label, type: c.type })),
      hiddenColumns: [],
    },
    rows: displayBankExportData,
  })), [columns, displayBankExportData]);

  return (
    <motion.div
      variants={containerVariants}
      initial="hidden"
      animate="visible"
      className="bulk-payment-layout flex-1 w-full min-h-0 flex flex-row gap-3 bg-transparent overflow-hidden p-0 relative"
      style={{
        borderWidth: "0px",
        gap: showLeftCard ? "12px" : "0px",
        paddingBottom: "0px",
        paddingTop: "0px",
        paddingLeft: "0px",
        paddingRight: "0px",
      }}
    >
      {/* Left Panel - Actions & Info (Unified Scrollable Card) */}
      {showLeftCard && (
        <div
          className="bulk-payment-side-panel timesheet-control-panel w-[275px] sm:w-[290px] shrink-0 flex flex-col h-full select-none bg-card border border-border rounded-none p-3.5 overflow-hidden z-20 shadow-2xs"
        >
          {/* Header */}
          <div className="flex items-center justify-between pb-2 border-b border-border/70 shrink-0 box-border bg-card">
            <div className="flex flex-col gap-0.5 min-w-0">
              <span className="tabular-nums text-[9px] font-black uppercase tracking-wider text-primary block truncate">
                [01] STATEMENT
              </span>
              <h2 className="text-[11px] font-bold text-foreground uppercase tracking-tight font-sans truncate">
                Batch Payment Hub
              </h2>
            </div>


          </div>

          <div className="flex items-center gap-1 py-2 border-b border-border/70 shrink-0">
            {[
              { id: "summary", label: "Overview", icon: Layers },
              {
                id: "adjustments",
                label: `Adj (${dynamicReportStats.holdAddItems.length})`,
                icon: Wrench,
              },
              { id: "reconcile", label: `Balance`, icon: Scale },
            ].map((t) => (
              <button
                key={t.id}
                onClick={() => setActiveLeftTab(t.id as any)}
                className={`flex h-7.5 flex-1 items-center justify-center gap-1 px-1 rounded-lg text-[9px] font-bold uppercase tracking-wider transition-all cursor-pointer active:scale-[0.98] ${
                  activeLeftTab === t.id
                    ? "bg-primary text-primary-foreground border border-primary shadow-xs"
                    : "bg-muted/40 hover:bg-muted text-muted-foreground hover:text-foreground border border-border"
                }`}
              >
                <t.icon className="h-3 w-3 shrink-0" />
                <span className="truncate">{t.label}</span>
              </button>
            ))}
          </div>

          {/* Scrollable contents */}
          <div className="flex-1 overflow-y-auto custom-scrollbar pt-2 pb-4 flex flex-col gap-2.5 min-h-0">
            <AnimatePresence mode="wait">
              {activeLeftTab === "summary" && (
                <motion.div
                  key="tab-summary"
                  initial={{ opacity: 0, y: 5 }}
                  animate={{ opacity: 1, y: 0 }}
                  exit={{ opacity: 0, y: -5 }}
                  transition={{ duration: 0.15 }}
                  className="flex flex-col gap-2.5 pb-2"
                >
                  {/* Total Overview - Matched with Timesheet Hero Card */}
                  <div className="bg-primary/5 dark:bg-primary/10 border border-primary/20 rounded-xl p-2.5 flex flex-col gap-1 relative overflow-hidden shadow-2xs">
                    <div className="flex items-center justify-between">
                      <div className="flex items-center gap-1.5">
                        <Layers className="w-3.5 h-3.5 text-primary" />
                        <span className="text-[8.5px] font-bold uppercase tracking-wider text-muted-foreground">Tổng chi lương đợt này</span>
                      </div>
                      <span className="text-[8px] font-bold text-primary bg-primary/10 px-1.5 py-0.5 rounded-md uppercase">
                        {selectedBUGroup === "ALL" ? "ALL BU" : selectedBUGroup}
                      </span>
                    </div>
                    <div className="flex items-baseline justify-end gap-1 pt-0.5">
                      <span className="text-xl font-black tabular-nums tracking-tight text-foreground leading-none">
                        {formatMoneyVND(totalPayoutSum).replace(" ₫", "")}
                      </span>
                      <span className="text-[10px] font-bold text-muted-foreground uppercase tracking-wider">VND</span>
                    </div>
                  </div>

                  {/* BU breakdown metrics - REDESIGNED FOR SINGLE HEADER ROW CARD */}
                  <div className="bg-card border border-border/80 rounded-xl p-2.5 flex flex-col gap-2 shadow-2xs">
                    <div className="flex items-center justify-between border-b pb-2 gap-2">
                      <span className="text-[10px] font-bold text-foreground uppercase tracking-[0.1em] font-sans flex items-center gap-1.5 shrink-0">
                        <div className="w-1.5 h-1.5 rounded-full bg-primary" />
                        PHÁT SINH THEO BU
                      </span>

<div className="flex items-center gap-1.5 shrink-0">
                        <DropdownMenu>
                          <DropdownMenuTrigger asChild>
                            <button
                              type="button"
                              className="bu-group-menu-trigger flex items-center justify-between gap-1.5 rounded-full border border-primary/20 bg-primary/[0.035] px-2.5 py-0.5 text-left text-[10px] font-bold uppercase tracking-[0.04em] text-foreground shadow-xs transition-colors hover:border-primary/40 hover:bg-primary/[0.07] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-primary/20"
                              aria-label="Chọn nhóm BU"
                            >
                              <span className="truncate">
                                {selectedBUGroup === "ALL"
                                  ? "ALL BU"
                                  : `${selectedBUGroup.toUpperCase()}`}
                              </span>
                              <ChevronDown className="h-3 w-3 shrink-0 text-primary" />
                            </button>
                          </DropdownMenuTrigger>
                          <DropdownMenuContent
                            align="end"
                            sideOffset={6}
                            className="bu-group-menu-content w-[240px] rounded-xl border border-primary/20 bg-popover p-1.5 text-popover-foreground shadow-xl"
                          >
                            <DropdownMenuLabel className="px-2.5 py-1.5 text-[9px] font-black uppercase tracking-[0.18em] text-muted-foreground">
                              Chọn phạm vi báo cáo
                            </DropdownMenuLabel>
                            <DropdownMenuSeparator className="bg-primary/10" />
                            {["ALL", "AHN", "AHP", "ATH", "ATN", "APT", "Other"].map((biz) => {
                              const isAll = biz === "ALL";
                              const targetBUs = ["AHN", "AHP", "ATH", "ATN", "APT", "Other"];
                              const sheet1Val = isAll
                                ? targetBUs.reduce((sum, business) => sum + (dynamicReportStats.sheet1Totals[business] || 0), 0)
                                : (dynamicReportStats.sheet1Totals[biz] || 0);
                              const hasAdjustments = isAll
                                ? (dynamicReportStats.holdAddItems || []).length > 0
                                : (dynamicReportStats.holdAddItems || []).some((item) => item.biz === biz);
                              const hasData = sheet1Val !== 0 || hasAdjustments;
                              const selected = selectedBUGroup === biz;
                              return (
                                <DropdownMenuItem
                                  key={biz}
                                  onSelect={() => setSelectedBUGroup(biz)}
                                  className={`mb-0.5 flex cursor-pointer items-center justify-between rounded-lg px-2.5 py-2 text-[11px] font-bold uppercase tracking-wide outline-none last:mb-0 ${
                                    selected
                                      ? "bg-primary text-primary-foreground focus:bg-primary focus:text-primary-foreground"
                                      : "text-foreground focus:bg-primary/10 focus:text-primary"
                                  }`}
                                >
                                  <span>{isAll ? "ALL BU" : `${biz.toUpperCase()} GROUP`}</span>
                                  {selected ? (
                                    <Check className="h-3.5 w-3.5 shrink-0" />
                                  ) : !hasData ? (
                                    <span 
                                      className="text-[8px] font-bold tracking-wider text-muted-foreground inline-flex items-center"
                                      style={{ lineHeight: "18px", height: "18.9982px" }}
                                    >
                                      TRỐNG
                                    </span>
                                  ) : null}
                                </DropdownMenuItem>
                              );
                            })}
                          </DropdownMenuContent>
                        </DropdownMenu>

                      </div>
                    </div>

                    {/* Display the active group content */}
                      {(() => {
                        const biz = selectedBUGroup;
                        const isAll = biz === "ALL";
                        const targetBUs = ["AHN", "AHP", "ATH", "ATN", "APT", "Other"];

                        const sheet1Val = isAll
                          ? targetBUs.reduce((sum, b) => sum + (dynamicReportStats.sheet1Totals[b] || 0), 0)
                          : (dynamicReportStats.sheet1Totals[biz] || 0);

                        const holdAddItems = isAll
                          ? (dynamicReportStats.holdAddItems || [])
                          : (dynamicReportStats.holdAddItems || []).filter((i) => i.biz === biz);

                        const holdOnly = holdAddItems
                          .filter((i) => i.type === "HOLD")
                          .reduce((sum, i) => sum + i.amount, 0);

                        const addOnly = holdAddItems
                          .filter((i) => i.type === "ADD")
                          .reduce((sum, i) => sum + i.amount, 0);

                        const bonusOnly = holdAddItems
                          .filter((i) => i.type === "BONUS")
                          .reduce((sum, i) => sum + i.amount, 0);

                        const cancelOnly = holdAddItems
                          .filter((i) => i.type === "CANCEL")
                          .reduce((sum, i) => sum + i.amount, 0);

                        // CANCEL is displayed in the breakdown but excluded from DEDUCTIONS and NET PAY calculation
                        const deductionsSum = holdAddItems.reduce((sum, item) => sum + item.contribution, 0);

                        const finalTotal = isAll
                          ? targetBUs.reduce((sum, b) => {
                              const s1 = dynamicReportStats.sheet1Totals[b] || 0;
                              const items = (dynamicReportStats.holdAddItems || []).filter((i) => i.biz === b);
                              const h = items.filter((i) => i.type === "HOLD").reduce((acc, i) => acc + i.amount, 0);
                              const a = items.filter((i) => i.type === "ADD").reduce((acc, i) => acc + i.amount, 0);
                              return sum + (dynamicReportStats.finalTotals[b] ?? (s1 + h + a));
                            }, 0)
                          : (dynamicReportStats.finalTotals[biz] ?? (sheet1Val + deductionsSum));

                        return (
                          <div className="bu-amount-card bu-payroll-ribbon-card relative overflow-hidden rounded-2xl border p-3.5 shadow-sm transition-all">
                            <div className="bu-payroll-ribbon-orb" aria-hidden="true" />

                            <div className="bu-payroll-ribbon-heading">
                              <div className="bu-payroll-ribbon-title">
                                <span className="bu-payroll-bow-medallion">
                                  <PayrollBowIcon className="h-4 w-4" />
                                </span>
                                <span>TÓM TẮT THANH TOÁN</span>
                              </div>
                              {isAll ? <button type="button" onClick={handleCopyReport} className="inline-flex items-center justify-center p-1 text-muted-foreground hover:text-primary active:scale-[0.98]" title="Sao chép toàn bộ thông tin" aria-label="Sao chép toàn bộ thông tin"><Copy className="h-3.5 w-3.5" /></button> : <span className="bu-payroll-ribbon-chip">{biz.toUpperCase()}</span>}
                            </div>

                            <div className="bu-payroll-primary-group">
                              <div className="bu-payroll-summary-row">
                                <span className="bu-payroll-row-label">
                                  <PayrollBowIcon className="bu-payroll-row-bow" />
                                  GROSS PAY
                                </span>
                                <span className="bu-summary-value text-foreground">
                                  {formatMoneyVND(sheet1Val).replace(" ₫", "")}
                                </span>
                              </div>

                              <div className="bu-payroll-summary-row">
                                <span className="bu-payroll-row-label">
                                  <PayrollBowIcon className="bu-payroll-row-bow" />
                                  DEDUCTIONS
                                </span>
                                <span className={`bu-summary-value ${deductionsSum >= 0 ? "text-emerald-700" : "text-rose-700"}`}>
                                  {deductionsSum >= 0 ? "+" : ""}
                                  {formatMoneyVND(deductionsSum).replace(" ₫", "")}
                                </span>
                              </div>
                            </div>

                            <div className="bu-payroll-detail-panel">
                              <div className="bu-payroll-detail-row">
                                <span className="bu-summary-detail-label">
                                  <span className="bu-payroll-detail-dot bu-payroll-detail-dot--hold" />
                                  HOLD
                                </span>
                                <span className={`bu-summary-detail-value ${holdOnly !== 0 ? "text-rose-700" : "text-muted-foreground"}`}>
                                  {holdOnly !== 0
                                    ? `-${formatMoneyVND(Math.abs(holdOnly)).replace(" ₫", "")}`
                                    : "0"}
                                </span>
                              </div>

                              <div className="bu-payroll-detail-row">
                                <span className="bu-summary-detail-label">
                                  <span className="bu-payroll-detail-dot bu-payroll-detail-dot--add" />
                                  ADD
                                </span>
                                <span className={`bu-summary-detail-value ${addOnly !== 0 ? "text-emerald-700" : "text-muted-foreground"}`}>
                                  {addOnly !== 0
                                    ? `+${formatMoneyVND(Math.abs(addOnly)).replace(" ₫", "")}`
                                    : "0"}
                                </span>
                              </div>

                              <div className="bu-payroll-detail-row">
                                <span className="bu-summary-detail-label">
                                  <span className="bu-payroll-detail-dot bu-payroll-detail-dot--cancel" />
                                  CANCEL
                                </span>
                                <span className={`bu-summary-detail-value ${cancelOnly !== 0 ? "text-amber-700" : "text-muted-foreground"}`}>
                                  {cancelOnly !== 0
                                    ? `-${formatMoneyVND(Math.abs(cancelOnly)).replace(" ₫", "")}`
                                    : "0"}
                                </span>
                              </div>

                              {bonusOnly !== 0 && (
                                <div className="bu-payroll-detail-row">
                                  <span className="bu-summary-detail-label">
                                    <span className="bu-payroll-detail-dot bu-payroll-detail-dot--bonus" />
                                    BONUS
                                  </span>
                                  <span className="bu-summary-detail-value text-primary">
                                    +{formatMoneyVND(Math.abs(bonusOnly)).replace(" ₫", "")}
                                  </span>
                                </div>
                              )}
                            </div>

                            <div className="bu-payroll-net-row">
                              <span className="bu-payroll-net-label">
                                <PayrollBowIcon className="h-4 w-4" />
                                NET PAY
                              </span>
                              <span className="bu-payroll-net-value">
                                {formatMoneyVND(finalTotal).replace(" ₫", "")}
                              </span>
                            </div>
                          </div>
                        );
                      })()}
                    </div>

                    {/* General Info */}
                  <div className="grid grid-cols-2 gap-1.5">
                    <div className="bg-card border border-border/80 p-2 rounded-xl flex flex-col gap-0.5 shadow-2xs">
                      <span className="text-[8.5px] font-bold text-muted-foreground uppercase tracking-wider font-sans">
                        Tháng báo cáo
                      </span>
                      <span className="text-xs font-black text-foreground tabular-nums leading-tight">
                        {appData.globalMonth || "03.2026"}
                      </span>
                    </div>
                    <div className="bg-card border border-border/80 p-2 rounded-xl flex flex-col gap-0.5 shadow-2xs">
                      <span className="text-[8.5px] font-bold text-muted-foreground uppercase tracking-wider font-sans">
                        Số dòng dữ liệu
                      </span>
                      <span className="text-xs font-black text-foreground tabular-nums leading-tight">
                        {(appData.BankExport?.data || []).length}
                      </span>
                    </div>
                  </div>
                </motion.div>
              )}

              {activeLeftTab === "adjustments" && (
                <motion.div
                  key="tab-adjustments"
                  initial={{ opacity: 0, y: 5 }}
                  animate={{ opacity: 1, y: 0 }}
                  exit={{ opacity: 0, y: -5 }}
                  transition={{ duration: 0.15 }}
                  className="flex flex-col gap-4 pb-6"
                >
                  {(() => {
                    const sidebarAdjustmentsFiltered = (
                      dynamicReportStats?.holdAddItems || []
                    ).filter((item) => item.type === adjustmentFilter);
                    return (
                      <>
                        {/* Categorized filter selection */}
                        <div className="flex flex-wrap gap-1 bg-slate-50 p-1 rounded-xl border border-slate-100">
                          {(["ALL", "HOLD", "ADD", "CANCEL"] as const).map(
                            (f) => {
                              const count =
                                f === "ALL"
                                  ? dynamicReportStats.holdAddItems.length
                                  : dynamicReportStats.holdAddItems.filter(
                                      (item) => item.type === f,
                                    ).length;
                              return (
                                <button
                                  key={f}
                                  onClick={() => setAdjustmentFilter(f)}
                                  className={`flex-1 py-1 text-[9.5px] font-bold tracking-wide transition-all rounded-md cursor-pointer ${
                                    adjustmentFilter === f
                                      ? "bg-white text-primary shadow-xs border border-slate-200/50"
                                      : "text-slate-500 hover:text-slate-800"
                                  }`}
                                >
                                  {f} ({count})
                                </button>
                              );
                            },
                          )}
                        </div>

                        {/* Adjustment Item list */}
                        <div className="space-y-2.5">
                          {adjustmentFilter === "ALL" ? (
                            (() => {
                              const buMap = aggregateAdjustmentTotals(
                                dynamicReportStats.holdAddItems || [],
                              );

                              const activeBUs = Object.entries(buMap).filter(
                                ([_, data]) => data.totalCount > 0,
                              );

                              if (activeBUs.length === 0) {
                                return (
                                  <div className="text-[10px] text-slate-400 italic py-4 text-center font-sans">
                                    Không tìm thấy khoản điều chỉnh nào
                                  </div>
                                );
                              }

                              return activeBUs.map(([buName, buData]) => (
                                <div
                                  key={buName}
                                  className="bg-white border border-slate-200/60 rounded-2xl p-4 shadow-xs flex flex-col gap-3 hover:border-slate-300 transition-colors"
                                >
                                  <div className="flex justify-between items-center pb-2 border-b border-slate-100">
                                    <span className="font-bold text-slate-800 text-sm tracking-wide font-sans">
                                      {buName}
                                    </span>
                                    <span className="text-[10px] font-bold text-slate-400 bg-slate-50 px-2.5 py-1 rounded-full border border-slate-100/80">
                                      {buData.totalCount} khoản phát sinh
                                    </span>
                                  </div>
                                  <div className={`grid ${buData.BONUS > 0 ? "grid-cols-2" : "grid-cols-3"} gap-2 text-[11px]`}>
                                    {/* HOLD */}
                                    <div className="flex flex-col items-start justify-center gap-0.5 p-2 bg-rose-50/30 rounded-lg border border-rose-100/30">
                                      <span className="text-rose-500 font-bold uppercase tracking-wider text-[9px]">
                                        HOLD:
                                      </span>
                                      <span className="text-rose-600 font-extrabold tabular-nums text-[10px]">
                                        {buData.HOLD > 0
                                          ? `-${formatMoneyVND(buData.HOLD).replace(" ₫", "")}`
                                          : "0"}
                                      </span>
                                    </div>
                                    {/* ADD */}
                                    <div className="flex flex-col items-start justify-center gap-0.5 p-2 bg-emerald-50/30 rounded-lg border border-emerald-100/30">
                                      <span className="text-emerald-500 font-bold uppercase tracking-wider text-[9px]">
                                        ADD:
                                      </span>
                                      <span className="text-emerald-600 font-extrabold tabular-nums text-[10px]">
                                        {buData.ADD > 0
                                          ? `+${formatMoneyVND(buData.ADD).replace(" ₫", "")}`
                                          : "0"}
                                      </span>
                                    </div>
                                    {/* CANCEL */}
                                    <div className="flex flex-col items-start justify-center gap-0.5 p-2 bg-slate-50 rounded-lg border border-slate-200/30">
                                      <span className="text-slate-400 font-bold uppercase tracking-wider text-[9px]">
                                        CANCEL:
                                      </span>
                                      <span className="text-slate-500 font-extrabold tabular-nums text-[10px]">
                                        {buData.CANCEL > 0
                                          ? `-${formatMoneyVND(buData.CANCEL).replace(" ₫", "")}`
                                          : "0"}
                                      </span>
                                    </div>
                                    {buData.BONUS > 0 && (
                                      <div className="flex flex-col items-start justify-center gap-0.5 p-2 bg-emerald-50/30 rounded-lg border border-emerald-100/30">
                                        <span className="text-emerald-500 font-bold uppercase tracking-wider text-[9px]">BONUS:</span>
                                        <span className="text-emerald-600 font-extrabold tabular-nums text-[10px]">
                                          +{formatMoneyVND(buData.BONUS).replace(" ₫", "")}
                                        </span>
                                      </div>
                                    )}
                                  </div>
                                </div>
                              ));
                            })()
                          ) : sidebarAdjustmentsFiltered.length > 0 ? (
                            sidebarAdjustmentsFiltered.map((item, idx) => {
                              const isAdd = item.type === "ADD";
                              const isCancelItem = item.type === "CANCEL";

                              let badgeClass =
                                "bg-rose-50 text-rose-600 border-rose-100 text-[9px]";
                              let badgeLabel = "HOLD";
                              if (isAdd) {
                                badgeClass =
                                  "bg-emerald-50 text-emerald-600 border-emerald-100 text-[9px]";
                                badgeLabel = "ADD";
                              } else if (isCancelItem) {
                                badgeClass =
                                  "bg-slate-100 text-slate-500 border-slate-200 text-[9px]";
                                badgeLabel = "CANCEL";
                              }

                              const moneyColor =
                                isAdd
                                  ? "text-emerald-600"
                                  : "text-rose-600";
                              const moneyPrefix =
                                isAdd ? "+" : "";

                              return (
                                <div
                                  key={idx}
                                  className="flex flex-col p-2.5 bg-slate-50/50 rounded-xl border border-slate-100/50 hover:bg-slate-50 transition-colors gap-1.5"
                                >
                                  <div className="flex items-center justify-between">
                                    <div className="flex items-center gap-2">
                                      <span className="font-bold text-slate-800 font-sans text-xs">
                                        {item.biz}
                                      </span>
                                      <span
                                        className={`font-bold px-1.5 py-0.5 rounded-full border leading-none shrink-0 font-sans ${badgeClass}`}
                                      >
                                        {badgeLabel}
                                      </span>
                                    </div>
                                    <span
                                      className={`font-extrabold text-xs shrink-0 font-sans ${moneyColor}`}
                                    >
                                      {moneyPrefix}
                                      {formatMoneyVND(item.amount).replace(
                                        " ₫",
                                        "",
                                      )}
                                    </span>
                                  </div>
                                  <div className="flex items-center justify-between text-[10px] text-slate-400 font-medium">
                                    <span
                                      className="truncate max-w-[200px]"
                                      title={item.reason}
                                     >
                                       {item.reason}
                                     </span>
                                     <span className="shrink-0">
                                       {item.month}
                                     </span>
                                   </div>
                                 </div>
                               );
                             })
                           ) : (
                             <div className="text-[10px] text-slate-400 italic py-4 text-center font-sans">
                               Không tìm thấy khoản điều chỉnh nào
                             </div>
                           )}
                         </div>
                       </>
                     );
                   })()}
                 </motion.div>
               )}

               {activeLeftTab === "reconcile" && (
                <motion.div
                  key="tab-reconcile"
                  initial={{ opacity: 0, y: 5 }}
                  animate={{ opacity: 1, y: 0 }}
                  exit={{ opacity: 0, y: -5 }}
                  transition={{ duration: 0.15 }}
                  className="flex flex-col gap-3 pb-8"
                >
                  {/* SECTIONS CONSOLIDATED IN SINGLE CLEAN CARD */}
                  {(() => {
                    const sheet1TotalsMap =
                      dynamicReportStats?.sheet1Totals || {};
                    const grossPayTotal = Object.values(
                      sheet1TotalsMap,
                    ).reduce((a, b) => a + b, 0);
                    const buEntries1 = Object.entries(sheet1TotalsMap).filter(
                      ([_, amt]) => amt !== 0,
                    );

                    const holdAddItems =
                      dynamicReportStats?.holdAddItems || [];
                    const deductionsTotal = holdAddItems.reduce(
                      (sum, item) => sum + item.contribution,
                      0,
                    );

                    const formatMonthTag = (mStr: string) => {
                      if (!mStr) return "";
                      const match = String(mStr).match(
                        new RegExp("(\\d{1,2})[/._\\s-]+(\\d{2,4})"),
                      );
                      if (match) {
                        const m = match[1].padStart(2, "0");
                        const y =
                          match[2].length === 4
                            ? match[2].slice(2)
                            : match[2];
                        return `T${m}.${y}`;
                      }
                      const clean = String(mStr)
                        .replace(/^Tháng\s*/i, "T")
                        .replace(/\s+/g, "")
                        .trim();
                      return clean ? `${clean}` : "";
                    };

                    const describeAdjustment = (rawKey: string) => {
                      const cleanKey = String(rawKey || "")
                        .replace(/[[\]]/g, "")
                        .trim()
                        .toUpperCase();
                      const match = cleanKey.match(
                        /^([A-Z]+)(?:_T?(\d{1,2})[./-](\d{2,4}))?$/,
                      );
                      const type = match?.[1] || cleanKey;
                      const month = match?.[2]?.padStart(2, "0");
                      const rawYear = match?.[3];
                      const year = rawYear
                        ? rawYear.length === 2
                          ? `20${rawYear}`
                          : rawYear
                        : "";
                      const labels: Record<string, string> = {
                        HOLD: "Khoản giữ lại",
                        ADD: "Cộng thêm",
                        CANCEL: "Chỉ hiển thị",
                        BONUS: "Thưởng bổ sung",
                      };

                      return {
                        type,
                        label: labels[type] || cleanKey.replaceAll("_", " "),
                        period: month && year ? `Tháng ${month}/${year}` : "",
                      };
                    };

                    const buGroups: Record<string, { total: number; itemsMap: Record<string, number> }> = {};
                    holdAddItems.forEach((item) => {
                      if (!buGroups[item.biz])
                        buGroups[item.biz] = { total: 0, itemsMap: {} };
                      buGroups[item.biz].total += item.contribution;

                      const mTag = formatMonthTag(item.month);
                      const key = `${item.type}${mTag ? `_${mTag}` : ""}`;
                      buGroups[item.biz].itemsMap[key] =
                        (buGroups[item.biz].itemsMap[key] || 0) + item.amount;
                    });
                    const groupEntries = Object.entries(buGroups);

                    const finalTotalsMap =
                      dynamicReportStats?.finalTotals || {};
                    const netPayTotal = Object.values(finalTotalsMap).reduce(
                      (a, b) => a + b,
                      0,
                    );
                    const buEntries3 = Object.entries(finalTotalsMap).filter(
                      ([_, amt]) => amt !== 0,
                    );

                    const bankExportTotal = (
                      appData.BankExport?.data || []
                    ).reduce(
                      (sum, r) =>
                        sum +
                        (parseMoneyToNumber(
                          r["Payment Amount"] ??
                            r["Amount"] ??
                            r["TOTAL PAYMENT"] ??
                            r["Số tiền"] ??
                            r["Thành tiền"] ??
                            0,
                        ) || 0),
                      0,
                    );

                    const totalBulkPayment =
                      bankExportTotal > 0
                        ? bankExportTotal
                        : calculationSummary.aeTotal || netPayTotal;
                    const totalAcc =
                      calculationSummary.calculatedTotal || netPayTotal;
                    const bonusTotal = dynamicReportStats?.bonusTotal || 0;
                    const sameMonthHold =
                      dynamicReportStats?.sameMonthHoldTotal || 0;
                    const diffMonthAdd =
                      dynamicReportStats?.diffMonthAddTotal || 0;
                    const totalBankAe =
                      calculationSummary.calculatedTotal -
                      calculationSummary.diff;
                    const diff = totalAcc - totalBulkPayment;

                    return (
                      <div className="payroll-theme-card bu-report-card border rounded-2xl p-4 flex flex-col gap-3.5 shadow-xs transition-colors">
                        <div className="bu-report-header flex items-center justify-between border-b border-slate-200/60 dark:border-slate-800 pb-2.5">
                          <span className="text-xs font-black text-slate-800 dark:text-slate-100 uppercase tracking-wider font-sans flex items-center gap-2">
                            <Scale className="w-4 h-4 text-primary shrink-0" />
                            BÁO CÁO CHI TIẾT THEO BU ({appData.globalMonth || "01.2026"})
                          </span>
                          <button
                            onClick={handleCopyReconciliationSummary}
                            className="p-1.5 text-slate-400 hover:text-emerald-600 hover:bg-emerald-50 dark:hover:bg-emerald-950/40 rounded-lg transition-colors cursor-pointer shrink-0"
                            title="Sao chép riêng phần báo cáo đang xem (2 cột Excel)"
                          >
                            <Copy className="w-4 h-4" />
                          </button>
                        </div>

                        <div className="bu-report-records text-[11px] text-slate-500 font-sans -mt-1">
                          <Layers className="w-3.5 h-3.5" />
                          <span>Số dòng dữ liệu</span>
                          <span className="bu-report-records-value font-bold tabular-nums text-slate-800 dark:text-slate-200">
                            {(appData.BankExport?.data || []).length}
                            <small>dòng</small>
                          </span>
                        </div>

                        {/* Styled Section Dropdown Header */}
                        <div className="relative w-full">
                          <select
                            value={activeBalanceSection}
                            onChange={(e) =>
                              setActiveBalanceSection(e.target.value)
                            }
                            className={`bu-report-select appearance-none outline-none border rounded-xl px-3.5 py-2.5 font-extrabold font-sans text-xs uppercase cursor-pointer w-full tracking-wider transition-all shadow-2xs pr-8 ${
                              activeBalanceSection === "I"
                                ? "text-indigo-950 bg-indigo-50/80 border-indigo-200 focus:ring-2 focus:ring-indigo-300"
                                : activeBalanceSection === "II"
                                  ? "text-rose-950 bg-rose-50/80 border-rose-200 focus:ring-2 focus:ring-rose-300"
                                  : activeBalanceSection === "III"
                                    ? "text-emerald-950 bg-emerald-50/80 border-emerald-200 focus:ring-2 focus:ring-emerald-300"
                                    : "text-sky-950 bg-sky-50/80 border-sky-200 focus:ring-2 focus:ring-sky-300"
                            }`}
                          >
                            <option value="I" className="text-indigo-900 font-extrabold text-xs bg-white py-1">I. GROSS PAY</option>
                            <option value="II" className="text-rose-900 font-extrabold text-xs bg-white py-1">II. DEDUCTIONS</option>
                            <option value="III" className="text-emerald-900 font-extrabold text-xs bg-white py-1">III. NET PAY</option>
                            <option value="IV" className="text-sky-900 font-extrabold text-xs bg-white py-1">IV. RECONCILIATION</option>
                          </select>
                          <ChevronDown
                            className={`w-4 h-4 absolute right-3 top-1/2 -translate-y-1/2 pointer-events-none ${
                              activeBalanceSection === "I"
                                ? "text-indigo-600"
                                : activeBalanceSection === "II"
                                  ? "text-rose-600"
                                  : activeBalanceSection === "III"
                                    ? "text-emerald-600"
                                    : "text-sky-600"
                            }`}
                          />
                        </div>

                        <div className="flex flex-col gap-2 pt-1">
                          {activeBalanceSection === "I" && (
                            <>
                              {buEntries1.length > 0 ? (
                                <div className="bu-report-value-list">
                                  {buEntries1.map(([biz, amt]) => (
                                    <div
                                      key={biz}
                                      className="bu-report-value-row bu-report-value-row--gross"
                                    >
                                      <span className="bu-report-value-label text-indigo-900/90">
                                        <span className="bu-report-row-icon bu-report-row-icon--gross">
                                          <Coins />
                                        </span>
                                        {biz}:
                                      </span>
                                      <span className="font-bold text-indigo-950 tabular-nums tracking-tight text-sm">
                                        {formatMoneyVND(amt).replace(" ₫", "")}
                                      </span>
                                    </div>
                                  ))}
                                </div>
                              ) : (
                                <span className="text-xs text-indigo-400 italic py-1">
                                  - Không có dữ liệu Sheet 1
                                </span>
                              )}
                              <div className="flex justify-between items-center mt-2 pt-3 pb-1 border-t border-indigo-200/80 font-black text-indigo-950 text-xs tracking-wider uppercase">
                                <span>TOTAL GROSS PAY:</span>
                                <span className="font-black tabular-nums tracking-tight text-base text-indigo-950">
                                  {formatMoneyVND(grossPayTotal).replace(
                                    " ₫",
                                    "",
                                  )}
                                </span>
                              </div>
                            </>
                          )}
                          {activeBalanceSection === "II" && (
                            <div className="flex flex-col gap-3">
                              {groupEntries.length > 0 ? (
                                groupEntries.map(([biz, grp]) => (
                                  <div
                                    key={biz}
                                    className="bu-report-adjustment-group"
                                  >
                                    <div className="bu-report-adjustment-summary">
                                      <span className="bu-report-value-label text-rose-950">
                                        <span className="bu-report-row-icon bu-report-row-icon--deduction">
                                          <TrendingDown />
                                        </span>
                                        <span>
                                          <small>Đơn vị</small>
                                          {biz}
                                        </span>
                                      </span>
                                      <span
                                        className={`font-bold tabular-nums tracking-tight text-sm ${grp.total >= 0 ? "text-emerald-600" : "text-rose-700"}`}
                                      >
                                        {grp.total >= 0 ? "+" : ""}
                                        {formatMoneyVND(grp.total).replace(
                                          " ₫",
                                          "",
                                        )}
                                      </span>
                                    </div>
                                    <div className="bu-report-adjustment-list">
                                      {Object.entries(grp.itemsMap).map(
                                        ([key, amount]) => {
                                          const detail =
                                            describeAdjustment(key);
                                          const DetailIcon =
                                            detail.type === "ADD"
                                              ? Plus
                                              : detail.type === "BONUS"
                                                ? Sparkles
                                                : detail.type === "CANCEL"
                                                  ? RefreshCw
                                                  : AlertCircle;

                                          return (
                                            <div
                                              key={key}
                                              className={`bu-report-adjustment-row bu-report-adjustment-row--${detail.type.toLowerCase()}`}
                                            >
                                              <span className="bu-report-adjustment-copy">
                                                <span className="bu-report-adjustment-icon">
                                                  <DetailIcon />
                                                </span>
                                                <span className="bu-report-adjustment-text">
                                                  <strong>
                                                    {detail.label}
                                                  </strong>
                                                  {detail.period && (
                                                    <small>
                                                      <Calendar />
                                                      {detail.period}
                                                    </small>
                                                  )}
                                                </span>
                                              </span>
                                              <span
                                                className={`bu-report-adjustment-amount ${amount >= 0 ? "text-emerald-600" : "text-rose-600"}`}
                                              >
                                                {amount >= 0 ? "+" : ""}
                                                {formatMoneyVND(amount).replace(
                                                  " ₫",
                                                  "",
                                                )}
                                              </span>
                                            </div>
                                          );
                                        },
                                      )}
                                    </div>
                                  </div>
                                ))
                              ) : (
                                <span className="text-xs text-rose-400 italic py-1">
                                  Không có khoản điều chỉnh
                                </span>
                              )}
                              <div className="flex justify-between items-center mt-1 pt-3 pb-1 border-t border-rose-200/80 font-black text-rose-950 text-xs tracking-wider uppercase">
                                <span>TOTAL DEDUCTIONS:</span>
                                <span
                                  className={`font-black tabular-nums tracking-tight text-base ${deductionsTotal >= 0 ? "text-emerald-700" : "text-rose-700"}`}
                                >
                                  {deductionsTotal >= 0 ? "+" : ""}
                                  {formatMoneyVND(deductionsTotal).replace(
                                    " ₫",
                                    "",
                                  )}
                                </span>
                              </div>
                            </div>
                          )}
                          {activeBalanceSection === "III" && (
                            <>
                              {buEntries3.length > 0 ? (
                                <div className="bu-report-value-list">
                                  {buEntries3.map(([biz, amt]) => (
                                    <div
                                      key={biz}
                                      className="bu-report-value-row bu-report-value-row--net"
                                    >
                                      <span className="bu-report-value-label text-emerald-900/90">
                                        <span className="bu-report-row-icon bu-report-row-icon--net">
                                          <CreditCard />
                                        </span>
                                        {biz}:
                                      </span>
                                      <span className="font-bold text-emerald-950 tabular-nums tracking-tight text-sm">
                                        {formatMoneyVND(amt).replace(" ₫", "")}
                                      </span>
                                    </div>
                                  ))}
                                </div>
                              ) : (
                                <span className="text-xs text-emerald-400 italic py-1">
                                  Không có dữ liệu
                                </span>
                              )}
                              <div className="flex justify-between items-center mt-2 pt-3 pb-1 border-t border-emerald-200/80 font-black text-emerald-950 text-xs tracking-wider uppercase">
                                <span>TOTAL NET PAY:</span>
                                <span className="font-black text-emerald-700 tabular-nums tracking-tight text-base">
                                  {formatMoneyVND(netPayTotal).replace(
                                    " ₫",
                                    "",
                                  )}
                                </span>
                              </div>
                            </>
                          )}

                          {activeBalanceSection === "IV" && (
                            <div className="flex flex-col gap-2.5 pt-1 text-xs">
                              <div className="flex justify-between items-center border-b border-sky-100/80 pb-2">
                                <span className="text-sky-950 font-bold font-sans flex items-center gap-2 tracking-wide text-xs">
                                  <span className="bu-report-inline-icon bu-report-inline-icon--reconciliation">
                                    <CreditCard />
                                  </span>
                                  TỔNG AE:
                                </span>
                                <span className="font-bold text-sky-800 tabular-nums tracking-tight text-sm">
                                  {formatMoneyVND(totalBulkPayment).replace(
                                    " ₫",
                                    "",
                                  )}
                                </span>
                              </div>

                              <div className="flex justify-between items-center">
                                <span className="text-sky-950 font-bold font-sans flex items-center gap-2 tracking-wide text-xs">
                                  <span className="bu-report-inline-icon bu-report-inline-icon--reconciliation">
                                    <Scale />
                                  </span>
                                  TỔNG ACC:
                                </span>
                                <span className="font-bold text-sky-800 tabular-nums tracking-tight text-sm">
                                  {formatMoneyVND(totalAcc).replace(" ₫", "")}
                                </span>
                              </div>

                              <div className="flex flex-col gap-1.5 pl-3.5 border-l-2 border-sky-200/80 my-1 py-1">
                                <div className="flex justify-between items-center text-xs">
                                  <span className="text-sky-900/80 font-semibold font-sans flex items-center gap-1.5 text-[11px]">
                                    <span className="bu-report-inline-icon bu-report-inline-icon--hold">
                                      <AlertCircle />
                                    </span>
                                    HOLD:
                                  </span>
                                  <span className="font-semibold text-rose-600 tabular-nums tracking-tight text-xs">
                                    -
                                    {formatMoneyVND(sameMonthHold).replace(
                                      " ₫",
                                      "",
                                    )}
                                  </span>
                                </div>
                                <div className="flex justify-between items-center text-xs">
                                  <span className="text-sky-900/80 font-semibold font-sans flex items-center gap-1.5 text-[11px]">
                                    <span className="bu-report-inline-icon bu-report-inline-icon--add">
                                      <Plus />
                                    </span>
                                    ADD:
                                  </span>
                                  <span className="font-semibold text-emerald-600 tabular-nums tracking-tight text-xs">
                                    +
                                    {formatMoneyVND(diffMonthAdd).replace(
                                      " ₫",
                                      "",
                                    )}
                                  </span>
                                </div>
                              </div>

                              <div className="flex justify-between items-center border-t border-sky-100/80 pt-2">
                                <span className="text-sky-950 font-bold font-sans flex items-center gap-2 tracking-wide text-xs">
                                  <span className="bu-report-inline-icon bu-report-inline-icon--reconciliation">
                                    <Coins />
                                  </span>
                                  TỔNG BANK AE:
                                </span>
                                <span className="font-bold text-sky-800 tabular-nums tracking-tight text-sm">
                                  {formatMoneyVND(totalBankAe).replace(
                                    " ₫",
                                    "",
                                  )}
                                </span>
                              </div>

                              <div className="flex justify-between items-center border-t-2 border-sky-200 mt-1 pt-3 pb-1">
                                <span className="font-extrabold text-sky-950 font-sans text-xs tracking-wider uppercase">
                                  LỆCH (DIFF):
                                </span>
                                <span
                                  className={`font-black tabular-nums tracking-tight text-base px-2 py-0.5 rounded-md ${diff !== 0 ? "text-rose-600 bg-rose-50" : "text-emerald-600 bg-emerald-50"}`}
                                >
                                  {formatMoneyVND(diff).replace(" ₫", "")}
                                </span>
                              </div>
                            </div>
                          )}
                        </div>
                      </div>
                    );
                  })()}
                </motion.div>
              )}
            </AnimatePresence>

          </div>

        </div>
      )}

      {/* Right Panel - Data View */}
      <div
        className="bulk-payment-data-panel unified-table-frame table-container flex-1 flex flex-col min-h-0 bg-card border border-border shadow-xs relative overflow-hidden pb-0 h-full p-0"
        style={{
          borderRadius: "var(--table-radius, 12px)",
          borderWidth: "1px",
          borderColor: "var(--table-frame-border, var(--border))",
          padding: 0,
        }}
      >
        {/* Table header matching Timesheet unified-table-frame-header */}
        {rightPanelTab !== "visuals" && (
        <div
          className="unified-table-frame-header px-3 py-2 flex flex-row items-center justify-between w-full gap-3 shrink-0 select-none box-border border-b border-border bg-[var(--table-header-bg,#FAF3E8)]"
          style={{
            backgroundColor: "var(--table-header-bg, #FAF3E8)",
            minHeight: "56px",
          }}
        >
          <div className="flex min-w-0 flex-1 items-center gap-0.5">
            <div className="app-table-title-lockup min-w-0">
              <div className="app-table-title-line">
                <button
                  type="button"
                  onClick={() => setShowLeftCard(!showLeftCard)}
                  className="table-initial-toggle shrink-0 cursor-pointer transition-all active:scale-[0.98]"
                  title={showLeftCard ? "Ẩn bảng điều khiển" : "Hiện bảng điều khiển"}
                  aria-label={showLeftCard ? "Ẩn bảng điều khiển" : "Hiện bảng điều khiển"}
                  aria-expanded={showLeftCard}
                >
                  <TableInitialMark
                    label={
                      rightPanelTab === "table"
                        ? "BATCH PAYMENT"
                        : rightPanelTab === "reconcile"
                          ? "RECONCILIATION"
                          : "ACCOUNTS PAYABLE AGING"
                    }
                  />
                </button>

              <DropdownMenu>
                <DropdownMenuTrigger asChild>
                  <button
                    type="button"
                    className="inline-flex w-max max-w-none shrink-0 items-baseline bg-transparent pb-0 pr-1.5 text-primary hover:bg-primary/[0.05] transition-all active:scale-95 cursor-pointer select-none border-none shadow-none outline-none rounded-r-lg"
                    title="Chuyển bảng"
                  >
                    <TableTitleRemainder
                      className="app-table-title-remainder--expanded"
                      label={
                        rightPanelTab === "table"
                          ? "BATCH PAYMENT TRANSACTION REGISTER"
                          : rightPanelTab === "reconcile"
                            ? "RECONCILIATION OF PAYROLL PAYMENTS"
                            : "ACCOUNTS PAYABLE AGING"
                      }
                    />
                  </button>
                </DropdownMenuTrigger>
              <DropdownMenuContent
                align="start"
                sideOffset={8}
                collisionPadding={8}
                className="table-switch-menu w-52 bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 shadow-xl rounded-xl p-1 z-[99999]"
              >
                <DropdownMenuLabel className="px-3 py-1.5 text-[10px] font-black uppercase tracking-wider text-slate-400">
                  CHUYỂN BẢNG
                </DropdownMenuLabel>
                <DropdownMenuItem
                  onClick={() => {
                    setRightPanelTab("table");
                    localStorage.setItem("bulk_payment_right_tab", "table");
                  }}
                  className={`flex items-center gap-2.5 px-3 py-2 text-xs font-bold rounded-lg cursor-pointer transition-colors ${
                    rightPanelTab === "table"
                      ? "bg-primary/10 text-primary font-extrabold"
                      : "text-slate-700 dark:text-slate-200 hover:bg-slate-100 dark:hover:bg-slate-800"
                  }`}
                >
                  <Table2 className="h-4 w-4 shrink-0 text-slate-600 dark:text-slate-300" />
                  <span>Batch Payment</span>
                </DropdownMenuItem>
                <DropdownMenuItem
                  onClick={() => {
                    setRightPanelTab("reconcile");
                    localStorage.setItem("bulk_payment_right_tab", "reconcile");
                  }}
                  className={`flex items-center gap-2.5 px-3 py-2 text-xs font-bold rounded-lg cursor-pointer transition-colors ${
                    rightPanelTab === "reconcile"
                      ? "bg-primary/10 text-primary font-extrabold"
                      : "text-slate-700 dark:text-slate-200 hover:bg-slate-100 dark:hover:bg-slate-800"
                  }`}
                >
                  <Scale className="h-4 w-4 shrink-0 text-sky-600 dark:text-sky-400" />
                  <span>Reconciliation</span>
                </DropdownMenuItem>
                <DropdownMenuItem
                  onClick={() => {
                    setRightPanelTab("visuals");
                    localStorage.setItem("bulk_payment_right_tab", "visuals");
                  }}
                  className="flex items-center gap-2.5 px-3 py-2 text-xs font-bold rounded-lg cursor-pointer transition-colors text-slate-700 dark:text-slate-200 hover:bg-slate-100 dark:hover:bg-slate-800"
                >
                  <BarChart2 className="h-4 w-4 shrink-0 text-amber-600 dark:text-amber-400" />
                  <span>Accounts Payable Aging</span>
                </DropdownMenuItem>
                </DropdownMenuContent>
              </DropdownMenu>
              </div>
              <p className="app-table-title-meta max-w-[520px] truncate text-[10px] font-medium leading-3.5 text-muted-foreground">
                {rightPanelTab === "table"
                  ? `${displayBankExportData.length} giao dịch • Tổng tiền: ${formatMoneyVND(bankExportTotal)}`
                  : rightPanelTab === "reconcile"
                    ? `Đối chiếu số liệu thanh toán và bảng lương`
                    : "Theo dõi biến động và vòng đời các khoản Hold"}
              </p>
            </div>

            {rightPanelTab === "table" && transactionReferenceReturn && (
              <button
                type="button"
                onClick={handleBackFromTransactionReference}
                className="inline-flex h-7 shrink-0 items-center gap-1.5 rounded-full border border-indigo-200 bg-indigo-50 px-3 text-[10px] font-bold uppercase tracking-wider text-indigo-700 shadow-3xs transition-colors hover:bg-indigo-100"
                title={`Quay lại ${transactionReferenceReturn.targetLabel}`}
              >
                <ArrowLeft className="h-3 w-3" />
                Về {transactionReferenceReturn.targetLabel}
              </button>
            )}

            {/* General Summary Stats - Compacted */}
            {displayBankExportData.length > 0 && rightPanelTab === "table" && (
              <div
                className="hidden lg:flex items-center gap-5 border-l border-slate-200 pl-5 h-8 ml-2"
                style={{ width: "auto", paddingRight: "20px" }}
              >
                <div className="flex flex-col items-end leading-tight">
                  <span className="text-[8px] font-bold text-slate-400 uppercase tracking-widest">
                    Số nhân viên
                  </span>
                  <span className="text-[11px] font-black tabular-nums text-slate-700">
                    {displayBankExportData.length}
                  </span>
                </div>
                <div className="flex flex-col items-end leading-tight">
                  <span className="text-[8px] font-bold text-slate-400 uppercase tracking-widest">
                    Tổng ACC
                  </span>
                  <span className="text-[11px] font-black tabular-nums text-slate-700">
                    {formatMoneyVND(calculationSummary.calculatedTotal).replace(" ₫", "").replace("đ", "")}
                  </span>
                </div>
                <div className="flex flex-col items-end leading-tight">
                  <span className="text-[8px] font-bold text-slate-400 uppercase tracking-widest">
                    Tổng AE
                  </span>
                  <span className="text-[11px] font-black tabular-nums text-slate-700">
                    {formatMoneyVND(bankExportTotal).replace(" ₫", "").replace("đ", "")}
                  </span>
                </div>
              </div>
            )}

            {/* Reconciliation Summary Stats in Header */}
            {displayBankExportData.length > 0 && rightPanelTab === "reconcile" && (
              <div
                className="hidden sm:flex items-center gap-4 border-l border-slate-200 pl-4 h-8 ml-2"
                style={{ width: "auto", paddingRight: "16px" }}
              >
                <div className="flex flex-col leading-tight">
                  <span className="text-[8px] font-bold text-slate-400 uppercase tracking-widest">
                    Tổng AE
                  </span>
                  <span className="text-[11px] font-black tabular-nums text-emerald-700">
                    {formatMoneyVND(reconcileTotals.ae).replace(" ₫", "")}
                  </span>
                </div>
                <div className="flex flex-col leading-tight">
                  <span className="text-[8px] font-bold text-slate-400 uppercase tracking-widest">
                    Tổng ACC
                  </span>
                  <span className="text-[11px] font-black tabular-nums text-slate-900 dark:text-slate-100">
                    {formatMoneyVND(reconcileTotals.accAmt).replace(" ₫", "")}
                  </span>
                </div>
                <div className="flex flex-col leading-tight">
                  <span className="text-[8px] font-bold text-slate-400 uppercase tracking-widest">
                    Chênh lệch
                  </span>
                  <span className={`text-[11px] font-black tabular-nums ${reconcileTotals.diff === 0 ? "text-emerald-600" : "text-rose-600"}`}>
                    {(reconcileTotals.diff > 0 ? "+" : "") + formatMoneyVND(reconcileTotals.diff).replace(" ₫", "")}
                  </span>
                </div>
              </div>
            )}

          </div>

          <div className="flex items-center gap-2 ml-auto shrink-0">
            {(rightPanelTab === "table" || rightPanelTab === "reconcile") && (
              <div className="flex items-center gap-1.5 mr-0.5">
                <button
                  type="button"
                  id="btn-toolbar-transaction-save"
                  onClick={() => window.dispatchEvent(new Event("trigger-transaction-save"))}
                  className="px-2.5 py-1 text-xs font-semibold rounded-full border border-primary/20 bg-primary/5 hover:bg-primary/10 text-foreground transition-all cursor-pointer active:scale-[0.98] shadow-2xs whitespace-nowrap"
                  title="Lưu Batch Payment đã bấm Lưu sửa lên Supabase, thay dữ liệu đúng tháng đang chọn"
                >
                  Lưu tháng
                </button>
                <button
                  type="button"
                  id="btn-toolbar-transaction-check"
                  onClick={() => window.dispatchEvent(new Event("trigger-transaction-check"))}
                  className="px-2.5 py-1 text-xs font-semibold rounded-full border border-primary/20 bg-primary/5 hover:bg-primary/10 text-foreground transition-all cursor-pointer active:scale-[0.98] shadow-2xs whitespace-nowrap"
                  title="Tải phiên bản mới nhất của tháng này và các tháng trước từ Supabase"
                >
                  Check STK & ID
                </button>
              </div>
            )}
            <div className="relative shrink-0">
            <DropdownMenu>
              <DropdownMenuTrigger asChild>
                <button
                  className="master-square-action border text-foreground transition-all cursor-pointer flex items-center justify-center active:scale-[0.98] shadow-2xs shrink-0 hover:text-primary"
                  title="Cài đặt & Thao tác"
                >
                  <Settings className="w-4 h-4" />
                </button>
              </DropdownMenuTrigger>
              <DropdownMenuContent
                align="end"
                className="w-60"
              >
                <TableRestoreButton placement="menu" onRestore={rightPanelTab === "table" ? handleRefresh : () => updateAppData(prev => ({ ...prev }), false)} />
                <DropdownMenuItem
                  onClick={() =>
                    window.dispatchEvent(new Event("open-ui-settings"))
                  }
                  className="text-slate-700"
                >
                  <Settings className="w-4 h-4 text-slate-500" />
                  <span>Cài đặt Giao diện</span>
                </DropdownMenuItem>
                <DropdownMenuItem
                  onClick={() =>
                    chooseExcelExport(() => {
                      if (rightPanelTab === "table") {
                        downloadTransactionBankExport(displayBankExportData, undefined, columns
                            .map((c) => c.key)
                            .filter((k) => !/tháng\s*báo\s*cáo/i.test(k)));
                      } else {
                        handleExportReconciliationExcel();
                      }
                    })
                  }
                  className="text-slate-700"
                >
                  <FileSpreadsheet className="h-4 w-4 shrink-0 text-emerald-700" />
                  <span>Xuất Excel</span>
                </DropdownMenuItem>
                    <DropdownMenuItem
                      onClick={() => window.dispatchEvent(new Event("open-transaction-settings"))}
                      className="flex items-center gap-2.5 px-2.5 py-2 rounded-lg cursor-pointer hover:bg-primary/10 text-slate-700 hover:text-primary font-bold text-xs"
                    >
                      <Settings className="w-4 h-4 text-primary shrink-0" />
                      <span>Cài đặt Batch Payment</span>
                    </DropdownMenuItem>
                    <DropdownMenuSeparator className="my-1 border-slate-100" />
                    <DropdownMenuLabel className="text-[10px] font-black uppercase text-slate-400 px-2 py-1">
                      Thao tác dữ liệu
                    </DropdownMenuLabel>
                    <DropdownMenuItem
                      onClick={handleSyncTransactionFieldsToTables}
                      className="flex items-center gap-2.5 px-2.5 py-2 rounded-lg cursor-pointer hover:bg-amber-50 text-slate-700 hover:text-amber-800 font-bold text-xs"
                      title="Lấy Tên, STK và ID Number từ Batch Payment để cập nhật Gross Pay và Deductions trong tháng đang chọn"
                    >
                      <Zap className="w-4 h-4 text-amber-500 shrink-0" />
                      <span>Đồng bộ Tên · STK · ID</span>
                    </DropdownMenuItem>
                    <DropdownMenuSeparator className="my-1 border-slate-100" />
                    <DropdownMenuLabel className="text-[10px] font-black uppercase text-slate-400 px-2 py-1">
                      Xuất File Excel
                    </DropdownMenuLabel>
                    <DropdownMenuItem
                      onClick={() => {
                        chooseExcelExport(() =>
                          downloadTransactionBankExport(displayBankExportData, undefined, columns
                              .map((c) => c.key)
                              .filter((k) => !/tháng\s*báo\s*cáo/i.test(k)))
                        );
                      }}
                      className="flex items-center gap-2.5 px-2.5 py-2 rounded-lg cursor-pointer hover:bg-emerald-50 text-slate-700 hover:text-emerald-800 font-bold text-xs"
                    >
                      <FileSpreadsheet className="w-4 h-4 text-emerald-600 shrink-0" />
                      <span>Xuất Bảng kê Bank Export</span>
                    </DropdownMenuItem>
                    <DropdownMenuItem
                      onClick={() => chooseExcelExport(handleExportReconciliationExcel)}
                      className="flex items-center gap-2.5 px-2.5 py-2 rounded-lg cursor-pointer hover:bg-sky-50 text-slate-700 hover:text-sky-800 font-bold text-xs"
                    >
                      <Scale className="w-4 h-4 text-sky-600 shrink-0" />
                      <span>Xuất Báo cáo Reconciliation</span>
                    </DropdownMenuItem>
              </DropdownMenuContent>
            </DropdownMenu>
              {rightPanelTab === "reconcile" && reconciliationNeedsSync && (
                <button
                  type="button"
                  onClick={handleSyncTransactionFieldsToTables}
                  className="absolute -left-2 -top-2 z-10 flex h-5 w-5 items-center justify-center rounded-full border border-white bg-amber-400 text-amber-950 shadow-sm transition-transform hover:scale-110 active:scale-95"
                  title="Cần đồng bộ Tên, STK và ID theo Batch Payment"
                  aria-label="Đồng bộ Tên, STK và ID cho Reconciliation"
                >
                  <Zap className="h-3 w-3" />
                </button>
              )}
            </div>
          </div>
        </div>
        )}

        <div hidden={rightPanelTab === "visuals"}>
          <TransactionHistoryPanel
            hasPendingEdits={hasPendingTransactionEdits}
            syncRevision={appData.TransactionActivity?.saveVersion || 0}
            rows={appData.BankExport?.data?.length ? appData.BankExport.data : appData.Bank_North_AE?.data || []}
            month={appData.globalMonth || ""}
            showReport={rightPanelTab === "reconcile"}
            onOpenReport={() => setRightPanelTab("reconcile")}
            onApplyVersions={handleApplyTransactionHistoryVersions}
            localMonths={appData.TransactionMonthCache?.months}
            onReportStateChange={handleHistoryReportStateChange}
          />
        </div>
        {/* Dynamic Display based on empty status & current selected tab */}
        {historyHasExceptions && !viewingHistorySource && reconciliationAudit.matchedCount > 0 && <p className="px-3 py-1 text-[10px] text-muted-foreground">Đã ẩn {reconciliationAudit.matchedCount} dòng đối chiếu tiền khớp để tập trung kiểm tra STK và ID.</p>}
        {viewingHistorySource ? null : displayBankExportData.length === 0 ? (
          <div className="flex-1 flex flex-col items-center justify-center text-primary/10 bg-slate-50/20 p-8 select-none">
            <div className="max-w-xl w-full flex flex-col items-center text-center">
              <h3 className="font-serif text-2xl text-slate-800 font-bold mb-2">
                Chưa có dữ liệu bảng kê {rightPanelTab === "table" ? "Batch Payment" : "Reconciliation"}
              </h3>
              <p className="text-[10px] text-slate-400 font-sans max-w-sm mb-8 leading-relaxed font-bold uppercase tracking-wider">
                Hệ thống tự động đồng bộ chi phí AE Final và các khoản điều
                chỉnh để tạo file chuyển khoản ngân hàng.
              </p>

              <button
                onClick={handleGenerateReport}
                disabled={isGenerating}
                className="soft-button bg-primary text-white shadow-md flex items-center justify-center gap-3 px-8 py-2.5 sm:py-4 min-h-[40px] max-h-[52px] h-auto rounded-2xl font-bold uppercase text-[11px] tracking-widest hover:bg-primary/95 hover:scale-[1.02] active:scale-[0.98] cursor-pointer transition-all flex-shrink"
              >
                {isGenerating ? (
                  <RefreshCw className="w-5 h-5 animate-spin" />
                ) : (
                  <Sparkles className="w-4 h-4 shrink-0" />
                )}
                <span>
                  {rightPanelTab === "table"
                    ? "TẠO BẢNG KÊ BATCH PAYMENT NGAY"
                    : "TẠO BẢNG KÊ RECONCILIATION NGAY"}
                </span>
              </button>
            </div>
          </div>
        ) : (
          <div
            className="flex-1 min-h-0 bg-white dark:bg-card relative z-10 flex flex-col rounded-none border-0 overflow-hidden"
            style={{ backgroundColor: "var(--card, #fff)" }}
          >
            <AnimatePresence mode="wait">
              {rightPanelTab === "table" && (
                <motion.div
                  key="panel-table"
                  initial={{ opacity: 0 }}
                  animate={{ opacity: 1 }}
                  exit={{ opacity: 0 }}
                  transition={{ duration: 0.1 }}
                  className="w-full h-full flex-1 min-h-0 flex flex-col p-0 relative"
                  style={{
                    height: "100%",
                    width: "100%",
                    paddingTop: "0px",
                    paddingLeft: "0px",
                    paddingRight: "0px",
                    paddingBottom: "0px",
                    marginTop: "0px",
                    marginLeft: "0px",
                    marginRight: "0px",
                    marginBottom: "0px",
                    borderWidth: "0px",
                    borderStyle: "none",
                    borderRadius: "0px",
                    overflow: "hidden",
                  }}
                >
                  <DataTable
                    columns={columns}
                    data={displayBankExportData}
                    onCellChange={handleCellChange}
                    onDeleteRow={(row) => setDeleteConfirmTarget({ row })}
                    onDeleteRows={(rows) => setDeleteConfirmTarget({ rows })}
                    isEditable={true}
                    externalSearchTerm={searchTerm}
                    onExternalSearchChange={setSearchTerm}
                    storageKey="bulk_payment"
                    alwaysHiddenColumns={useMemoHiddenBuColumns}
                    hideBuFilter
                    ignoreSavedHiddenColumns={false}
                    showFooter={true}
                    hideSearch={true}
                    headerClassName="bg-[var(--table-column-header-bg,#F4ECD8)] text-[var(--table-column-header-text-color,#1e293b)] border-[#e7dbdc] font-bold"
                    footerClassName="bg-[var(--table-column-header-bg,#F4ECD8)] text-[var(--table-column-header-text-color,#1e293b)] border-[#e7dbdc] font-black text-[12.5px] md:text-[13px]"
                    footerActionContent={
                      <button
                        type="button"
                        onClick={handleSaveTransactionEdits}
                        disabled={!hasPendingTransactionEdits}
                        aria-label="Lưu dữ liệu Batch Payment sau chỉnh sửa"
                        title={
                          hasPendingTransactionEdits
                            ? "Lưu toàn bộ thay đổi vừa sửa trong bảng Batch Payment"
                            : "Chưa có thay đổi cần lưu"
                        }
                        className={`inline-flex h-7 shrink-0 items-center gap-1.5 rounded-full border px-3 text-[10px] font-extrabold whitespace-nowrap transition-all active:scale-[0.98] ${
                          hasPendingTransactionEdits
                            ? "border-primary bg-primary text-foreground shadow-sm hover:brightness-95 cursor-pointer"
                            : "border-border bg-background/60 text-muted-foreground opacity-55 cursor-not-allowed"
                        }`}
                      >
                        <Save className="h-3.5 w-3.5 shrink-0" />
                        <span>Lưu sửa</span>
                      </button>
                    }
                  />
                </motion.div>
              )}

              {rightPanelTab === "reconcile" && (!historyHasExceptions || reconciliationAudit.transactionAuditList.some(item => item.status !== "MATCHED")) && (
                <motion.div
                  key="panel-reconcile"
                  initial={{ opacity: 0 }}
                  animate={{ opacity: 1 }}
                  exit={{ opacity: 0 }}
                  transition={{ duration: 0.1 }}
                  className="w-full h-full flex-1 min-h-0 flex flex-col p-0 relative overflow-hidden"
                >
                  {/* Filter Tabs by Reconciliation Status */}
                  {shouldShowFilterDiv && (
                  <div
                    className="shrink-0 flex flex-wrap items-center justify-between gap-3 p-3 bg-white rounded-none border border-slate-200/80 shadow-xs"
                    style={{
                      paddingBottom: "6px",
                      height: "47.0928px",
                      borderWidth: "0px",
                      marginTop: "12px",
                      borderRadius: "0px",
                    }}
                  >
                    <div className="flex flex-wrap items-center gap-1.5">
                      {[
                        {
                          id: "ALL",
                          label: `Cần xử lý (${reconciliationAudit.varianceCount + reconciliationAudit.missingInfoCount + reconciliationAudit.duplicateCount})`,
                        },
                        {
                          id: "VARIANCE",
                          label: `⚠️ Lệch số tiền (${reconciliationAudit.varianceCount})`,
                        },
                        {
                          id: "MISSING_INFO",
                          label: `🔴 Sai/thiếu thông tin (${reconciliationAudit.missingInfoCount})`,
                        },
                        {
                          id: "DUPLICATE",
                          label: `⚠️ Trùng ID/STK (${reconciliationAudit.duplicateCount})`,
                        },
                        {
                          id: "MATCHED",
                          label: `✅ Khớp (${reconciliationAudit.matchedCount})`,
                        },
                      ].filter(tab => !historyHasExceptions || tab.id !== "MATCHED").map((tab) => (
                        <button
                          key={tab.id}
                          onClick={() =>
                            setReconcileFilterStatus(tab.id as any)
                          }
                          className={`px-3 py-1.5 rounded-[20px] text-[10px] font-extrabold uppercase tracking-wider transition-all cursor-pointer ${
                            effectiveReconcileFilterStatus === tab.id
                              ? "bg-slate-900 text-white shadow-xs"
                              : "bg-slate-100 text-slate-600 hover:bg-slate-200"
                          }`}
                          style={{ borderRadius: "20px" }}
                        >
                          {tab.label}
                        </button>
                      ))}
                    </div>

                    <div className="flex items-center gap-2 shrink-0">
                      <Popover>
                        <PopoverTrigger asChild>
                          <button
                            className="w-7 h-7 rounded-full bg-amber-100 hover:bg-amber-200 text-amber-900 flex items-center justify-center font-bold text-xs transition-colors cursor-pointer shadow-2xs"
                            title="Xem nguyên tắc đối chiếu"
                          >
                            ?
                          </button>
                        </PopoverTrigger>
                        <PopoverContent className="w-80 p-3 text-xs text-slate-600 bg-white border border-slate-200 shadow-lg rounded-lg z-[100]">
                          <p className="font-bold text-slate-800 mb-1.5">Hướng dẫn đối chiếu:</p>
                          <p className="mb-1">
                            1. <strong>Lệch số tiền:</strong> Cùng ID/STK nhưng tổng số tiền thanh toán lệch nhau.
                          </p>
                          <p className="mb-1">
                            2. <strong>Thiếu thông tin:</strong> Bản ghi thiếu STK hoặc ID Number.
                          </p>
                          <p>
                            3. <strong>Đồng bộ một lần:</strong> Hệ thống tham
                            chiếu thêm RAWDATA_TIMESHEET khi ID/STK bị trùng,
                            sau đó cập nhật đồng thời Transaction và bảng đích.
                          </p>
                        </PopoverContent>
                      </Popover>
                    </div>
                  </div>
                  )}

                  {/* BU Quick Filter Bar for Reconcile */}
                  <div
                    className="bu-filter-bar flex items-center justify-between gap-2 px-3 py-1.5 border-b shrink-0 overflow-x-auto select-none z-10"
                    style={{
                      backgroundColor: "var(--table-sub-header-bg, #EDE4DB)",
                      borderColor: "var(--grid-line-color, var(--border, #E2E8F0))",
                    }}
                  >
                    <div className="flex items-center gap-1.5 flex-wrap">
                      <span className="flex items-center gap-1 text-[10px] font-black uppercase tracking-wider text-muted-foreground mr-1">
                        <Filter className="w-3 h-3 text-primary" />
                        <span>Lọc BU:</span>
                      </span>
                      <button
                        type="button"
                        onClick={() => setReconcileSelectedBU("ALL")}
                        className={`px-2.5 py-0.5 rounded-md text-[10.5px] font-bold transition-all cursor-pointer active:scale-95 whitespace-nowrap ${
                          reconcileSelectedBU === "ALL"
                            ? "bg-primary text-primary-foreground shadow-2xs font-black"
                            : "bg-background hover:bg-muted text-foreground border border-border/80"
                        }`}
                      >
                        Tất cả
                      </button>
                      <button
                        type="button"
                        onClick={() =>
                          setReconcileSelectedBU(reconcileSelectedBU === "EXCLUDE_AHP" ? "ALL" : "EXCLUDE_AHP")
                        }
                        className={`px-2.5 py-0.5 rounded-md text-[10.5px] font-bold transition-all cursor-pointer active:scale-95 flex items-center gap-1 whitespace-nowrap ${
                          reconcileSelectedBU === "EXCLUDE_AHP"
                            ? "bg-amber-600 text-white shadow-2xs ring-1 ring-amber-600 font-black"
                            : "bg-amber-50/80 hover:bg-amber-100/80 dark:bg-amber-950/40 text-amber-800 dark:text-amber-300 border border-amber-300 dark:border-amber-700/60"
                        }`}
                        title="Lọc bảng đối chiếu trừ BU AHP (Hải Phòng)"
                      >
                        <span>Trừ AHP</span>
                        {reconcileSelectedBU === "EXCLUDE_AHP" && (
                          <span className="text-[8.5px] font-black bg-white/25 px-1 rounded">Đang lọc</span>
                        )}
                      </button>
                    </div>

                    <div className="flex items-center gap-2.5 text-[10.5px] font-medium ml-auto shrink-0 whitespace-nowrap">
                      <Tooltip>
                        <TooltipTrigger asChild>
                          <button
                            type="button"
                            className="flex h-5 w-5 items-center justify-center rounded-full border border-amber-400 bg-amber-50 text-[11px] font-black text-amber-800 transition-colors hover:bg-amber-100 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-amber-500"
                            aria-label="Các bước đồng bộ, lưu tháng và Check STK & ID"
                          >
                            !
                          </button>
                        </TooltipTrigger>
                        <TooltipContent side="top" sideOffset={8} className="max-w-[280px] border border-amber-200 bg-white p-3 text-left text-[11px] leading-relaxed text-slate-800 shadow-lg">
                          <p className="font-bold">Thứ tự Check STK & ID</p>
                          <ol className="mt-1 list-inside list-decimal space-y-1">
                            <li>Đồng bộ STK, tên và ID nếu có cảnh báo.</li>
                            <li>Nếu có sửa trực tiếp bảng Batch Payment, bấm Lưu sửa.</li>
                            <li>Bấm Lưu tháng để ghi Batch Payment đã sửa lên Supabase. Nút Đồng bộ chỉ cập nhật Gross Pay và Deductions theo Batch Payment.</li>
                            <li>Bấm Check STK & ID để đọc lại phiên bản mới nhất.</li>
                          </ol>
                        </TooltipContent>
                      </Tooltip>
                      {reconcileSelectedBU !== "ALL" && (
                        <button
                          type="button"
                          onClick={() => setReconcileSelectedBU("ALL")}
                          className="text-[10px] text-rose-600 hover:text-rose-700 font-bold hover:underline cursor-pointer flex items-center gap-0.5 ml-1"
                          title="Bỏ lọc BU"
                        >
                          <X className="w-3 h-3" />
                          Bỏ lọc
                        </button>
                      )}
                    </div>
                  </div>

                  {/* Transaction Audit Table */}
                  <div
                    className="reconcile-table-region table-body-region flex-1 min-h-0 relative rounded-none bg-white overflow-auto custom-scrollbar"
                    style={{ borderRadius: "0px" }}
                  >
                    <table aria-label="Đối chiếu tiền Reconcile" className="w-full min-w-max text-left border-separate border-spacing-0 text-[11px] font-sans">
                      <thead 
                        className="sticky top-0 z-30 shadow-sm bg-[var(--table-column-header-bg,#F4ECD8)] text-[var(--table-column-header-text-color,#1e293b)]"
                      >
                        <tr>
                          <th
                            className="group relative px-2 py-2 leading-normal font-bold uppercase tracking-wider text-[9px] w-12 text-center border-r border-b border-[var(--grid-line-color,rgba(0,0,0,0.035))] align-middle whitespace-normal cursor-pointer select-none overflow-visible"
                            style={{ textAlign: "center" }}
                          >
                            <div className="inline-flex items-center justify-center gap-1">
                              <span>No.</span>
                              <button
                                type="button"
                                onClick={(e) => {
                                  e.stopPropagation();
                                }}
                                className="opacity-0 group-hover:opacity-100 transition-opacity p-0.5 hover:bg-black/10 rounded cursor-pointer shrink-0"
                                title="Tự động căn chỉnh độ rộng cột"
                                aria-label="Tự động căn chỉnh độ rộng cột"
                              >
                                <Maximize2 className="w-3 h-3" />
                              </button>
                            </div>
                          </th>
                          <th
                            className="px-2 py-2 leading-normal font-bold uppercase tracking-wider text-[9px] border-r border-b border-[var(--grid-line-color,rgba(0,0,0,0.035))] align-middle text-center whitespace-normal overflow-visible"
                            style={{ textAlign: "center" }}
                          >
                            ID NUMBER
                          </th>
                          <th
                            className="px-2 py-2 leading-normal font-bold uppercase tracking-wider text-[9px] border-r border-b border-[var(--grid-line-color,rgba(0,0,0,0.035))] align-middle text-center whitespace-normal overflow-visible"
                            style={{ textAlign: "center" }}
                          >
                            FULL NAME
                          </th>
                          <th
                            className="px-2 py-2 leading-normal font-bold uppercase tracking-wider text-[9px] border-r border-b border-[var(--grid-line-color,rgba(0,0,0,0.035))] align-middle text-center whitespace-normal overflow-visible"
                            style={{ textAlign: "center" }}
                          >
                            Bank Acc No. from AE
                          </th>
                          <th
                            className="px-2 py-2 leading-normal font-bold uppercase tracking-wider text-[9px] border-r border-b border-[var(--grid-line-color,rgba(0,0,0,0.035))] align-middle text-center whitespace-normal overflow-visible"
                          >
                             Bank Acc No. from ACC
                          </th>
                          <th
                            className="px-2 py-2 leading-normal text-center font-bold uppercase tracking-wider text-[9px] border-r border-b border-[var(--grid-line-color,rgba(0,0,0,0.035))] align-middle whitespace-normal overflow-visible"
                          >
                            TOTAL BANK AE
                          </th>
                          <th
                            className="px-2 py-2 leading-normal text-center font-bold uppercase tracking-wider text-[9px] border-r border-b border-[var(--grid-line-color,rgba(0,0,0,0.035))] align-middle whitespace-normal overflow-visible"
                          >
                            TOTAL BANK ACC
                          </th>
                          <th
                            className="px-2 py-2 leading-normal text-center font-bold uppercase tracking-wider text-[9px] border-r border-b border-[var(--grid-line-color,rgba(0,0,0,0.035))] align-middle whitespace-normal overflow-visible"
                          >
                            Diff
                          </th>
                          <th
                            className="px-2 py-2 leading-normal text-center font-bold uppercase tracking-wider text-[9px] border-r border-b border-[var(--grid-line-color,rgba(0,0,0,0.035))] align-middle whitespace-normal overflow-visible"
                          >
                            Process Sync
                          </th>
                          <th
                            className="px-2 py-2 leading-normal text-center font-bold uppercase tracking-wider text-[9px] border-b border-[var(--table-border-color,#E7E5E4)] align-middle whitespace-normal overflow-visible"
                          >
                            Problems
                          </th>
                        </tr>
                      </thead>
                      <tbody>
                        {paginatedTransactionAudits.length === 0 ? (
                          <tr>
                            <td
                              colSpan={RECONCILE_DISPLAY_COLUMN_KEYS.length}
                              className="p-8 text-center text-slate-400 italic border-b border-slate-200"
                            >
                              Không tìm thấy giao dịch nào phù hợp với điều kiện
                              lọc
                            </td>
                          </tr>
                        ) : (
                          paginatedTransactionAudits.map((item) => {
                            const isUnmatched =
                              item.id.startsWith("unmatched-");
                            const canSync =
                              !isUnmatched &&
                              (item.referenceCorrections?.length || 0) > 0;
                            const totalTargetBankAcc =
                              item.sheet1Amount + item.holdAmount;

                            return (
                              <tr
                                key={item.id}
                                className={`border-b border-slate-200 ${
                                  isUnmatched
                                    ? "bg-amber-50/50 font-semibold"
                                    : item.status === "VARIANCE"
                                      ? "bg-amber-50/30"
                                      : item.status === "MISSING_INFO"
                                        ? "bg-rose-50/30"
                                        : item.status === "DUPLICATE"
                                          ? "bg-purple-50/30"
                                          : "bg-white"
                                }`}
                              >
                                <td className="p-2.5 text-center tabular-nums font-bold text-slate-400 text-[10px] border-b border-r border-[var(--grid-line-color,rgba(0,0,0,0.035))]">
                                  {isUnmatched ? "DISC" : item.serialNo}
                                </td>
                                <td
                                  className="group/link p-2.5 font-bold text-slate-800 border-b border-r border-[var(--grid-line-color,rgba(0,0,0,0.035))] cursor-pointer"
                                  title="Click để chuyển tới bảng nguồn Gross Pay / Hold AE"
                                  onClick={() => {
                                    if (onTabChange) {
                                      const targetTab =
                                        item.targetTabForAccLink ||
                                        (item.sheet1Amount > 0
                                          ? "Sheet1_AE"
                                          : "Hold_AE");
                                      localStorage.setItem(
                                        "bulk_payment_right_tab",
                                        "reconcile",
                                      );
                                      localStorage.setItem(
                                        "master_ae_search",
                                        item.docId || "",
                                      );
                                      onTabChange(targetTab);
                                      window.dispatchEvent(
                                        new CustomEvent("master-ae-filter", {
                                          detail: {
                                            search: item.docId || "",
                                            from: "BulkPayment",
                                          },
                                        }),
                                      );
                                      toast.info(
                                        `Đã chuyển tới bảng ${targetTab === "Sheet1_AE" ? "Gross Pay" : "HOLD AE"} và lọc ID NUMBER: ${item.docId || ""}`,
                                      );
                                    }
                                  }}
                                >
                                  <div className="tabular-nums font-bold text-slate-900 text-[13px] leading-[19.5px] flex items-center gap-1">
                                    <span>
                                      {formatIdNumber(item.docId) || "N/A"}
                                    </span>
                                    <ExternalLink className="w-3 h-3 text-sky-500 opacity-20 transition-opacity duration-200 group-hover/link:opacity-75" />
                                  </div>
                                </td>
                                <td className="p-2.5 text-[10px] text-slate-700 font-medium border-b border-r border-[var(--grid-line-color,rgba(0,0,0,0.035))]">
                                  {item.name && item.name !== "N/A"
                                    ? item.name
                                    : "-"}
                                </td>
                                <td
                                  className={`group/link p-2.5 border-b border-r border-[var(--grid-line-color,rgba(0,0,0,0.035))] ${!isUnmatched ? "cursor-pointer" : ""}`}
                                  title={
                                    !isUnmatched
                                      ? "Click để xem giao dịch bên Bank Export"
                                      : undefined
                                  }
                                  onClick={() => {
                                    if (isUnmatched) return;
                                    setSearchTerm(
                                      item.accountNo || item.docId || "",
                                    );
                                    setRightPanelTab("table");
                                    localStorage.setItem(
                                      "bulk_payment_right_tab",
                                      "table",
                                    );
                                    toast.info(
                                      `Đã chuyển tới Bank Export và lọc: ${item.accountNo || item.docId}`,
                                    );
                                  }}
                                >
                                  <div className="tabular-nums font-semibold text-sky-600 flex items-center gap-1">
                                    <span>
                                      {item.accountNo || "⚠️ Chưa có STK"}
                                    </span>
                                    {!isUnmatched && (
                                      <ExternalLink className="w-3 h-3 opacity-20 transition-opacity duration-200 group-hover/link:opacity-75" />
                                    )}
                                  </div>
                                  {item.bankName && item.bankName !== "N/A" && (
                                    <div className="text-[10px] text-slate-400 font-medium truncate max-w-[150px]">
                                      {item.bankName}
                                    </div>
                                  )}
                                </td>
                                <td
                                  className="group/link p-2.5 tabular-nums border-b border-r border-[var(--grid-line-color,rgba(0,0,0,0.035))] cursor-pointer"
                                  title="Click để chuyển tới bảng nguồn Gross Pay/Hold AE"
                                  onClick={() => {
                                    if (onTabChange) {
                                      const targetTab =
                                        item.targetTabForAccLink;
                                      localStorage.setItem(
                                        "bulk_payment_right_tab",
                                        "reconcile",
                                      );
                                      const targetSearch =
                                        item.docId ||
                                        item.grossPlusBenefitsId ||
                                        "";
                                      localStorage.setItem(
                                        "master_ae_search",
                                        targetSearch,
                                      );
                                      onTabChange(targetTab);
                                      window.dispatchEvent(
                                        new CustomEvent("master-ae-filter", {
                                          detail: {
                                            search: targetSearch,
                                            from: "BulkPayment",
                                          },
                                        }),
                                      );
                                      toast.info(
                                        `Đã chuyển tới bảng ${targetTab === "Sheet1_AE" ? "Gross Pay" : "HOLD AE"} và lọc ID NUMBER: ${targetSearch}`,
                                      );
                                    }
                                  }}
                                >
                                  <div className="tabular-nums font-semibold text-sky-600 flex items-center gap-1">
                                    <span>
                                      {item.benefitsAccountNo ||
                                        item.accountNo ||
                                        "⚠️ Chưa có STK"}
                                    </span>
                                    <ExternalLink className="w-3 h-3 opacity-20 transition-opacity duration-200 group-hover/link:opacity-75" />
                                  </div>
                                </td>
                                <td className="p-2.5 text-right tabular-nums font-bold text-emerald-700 border-b border-r border-[var(--grid-line-color,rgba(0,0,0,0.035))]">
                                  {formatMoneyVND(item.actualAmount).replace(
                                    " ₫",
                                    "",
                                  )}
                                </td>
                                <td className="p-2.5 text-right tabular-nums font-black text-slate-900 border-b border-r border-[var(--grid-line-color,rgba(0,0,0,0.035))] bg-amber-50/40">
                                  {formatMoneyVND(totalTargetBankAcc).replace(
                                    " ₫",
                                    "",
                                  )}
                                </td>
                                <td
                                  className={`p-2.5 text-right tabular-nums font-black border-b border-r border-[var(--grid-line-color,rgba(0,0,0,0.035))] ${
                                    Math.abs(item.variance) < 1
                                      ? "text-emerald-600"
                                      : "text-rose-600"
                                  }`}
                                >
                                  {item.variance > 0 ? "+" : ""}
                                  {formatMoneyVND(item.variance).replace(
                                    " ₫",
                                    "",
                                  )}
                                </td>
                                <td
                                  className={`p-2.5 text-center border-b border-r border-[var(--grid-line-color,rgba(0,0,0,0.035))] ${!canSync ? "bg-slate-100/80 select-none" : "bg-white"}`}
                                >
                                  {canSync ? (
                                    <button
                                      type="button"
                                      onClick={() =>
                                        handleAutoFillMissingAccount(item)
                                      }
                                      className="px-2.5 py-1 bg-amber-500 hover:bg-amber-600 text-white font-bold rounded-lg transition-all text-[10px] cursor-pointer shadow-xs flex items-center gap-1 mx-auto active:scale-95"
                                      title="Đồng bộ ID Number, Full Name và Bank Account Number theo Transaction"
                                    >
                                      <Zap className="w-3 h-3 fill-current shrink-0" />
                                      <span>Đồng bộ</span>
                                    </button>
                                  ) : (
                                    <span className="inline-block px-2 py-0.5 bg-slate-200/70 text-slate-400 font-semibold text-[10px] rounded-md border border-slate-200/80 cursor-not-allowed">
                                      Không cần
                                    </span>
                                  )}
                                </td>
                                <td className="p-2.5 text-center border-b border-[var(--table-border-color,#E7E5E4)]">
                                  <div className="flex flex-col items-center gap-1">
                                    <span
                                      className={`px-2 py-0.5 rounded-full text-[9px] font-extrabold border leading-none ${
                                        item.status === "MATCHED"
                                          ? "bg-emerald-50 text-emerald-600 border-emerald-200"
                                          : item.status === "VARIANCE"
                                            ? "bg-amber-50 text-amber-700 border-amber-200"
                                            : item.status === "MISSING_INFO"
                                              ? "bg-rose-50 text-rose-600 border-rose-200"
                                              : item.status === "DUPLICATE"
                                                ? "bg-purple-50 text-purple-700 border-purple-200"
                                                : "bg-slate-100 text-slate-600 border-slate-200"
                                      }`}
                                    >
                                      {item.status === "MATCHED"
                                        ? "KHỚP 100%"
                                        : item.status === "VARIANCE"
                                          ? "CHÊNH LỆCH"
                                          : item.status === "MISSING_INFO"
                                            ? "SAI/THIẾU THÔNG TIN"
                                            : item.status === "DUPLICATE"
                                              ? "TRÙNG ID/STK"
                                              : "KHÔNG CÓ TRONG SHEET1"}
                                    </span>
                                    {item.issues.length > 0 && (
                                      <span
                                        className="text-[9px] text-rose-700 font-extrabold max-w-[140px] truncate"
                                        title={item.issues.join(", ")}
                                      >
                                        {item.issues.join(", ")}
                                      </span>
                                    )}
                                  </div>
                                </td>
                              </tr>
                            );
                          })
                        )}
                      </tbody>
                      <tfoot className="sticky bottom-0 z-30 bg-[var(--table-column-header-bg,#F4ECD8)]">
                        <tr className="total-row">
                          {Array.from(
                            { length: RECONCILE_DISPLAY_COLUMN_KEYS.length },
                            (_, columnIndex) => {
                              if (columnIndex > 0 && columnIndex < 5) {
                                return null;
                              }
                              if (columnIndex === 0) {
                                return (
                                  <td
                                    key="reconcile-total-header"
                                    colSpan={5}
                                    className="p-2.5 border-b border-t border-[var(--table-border-color,#E7E5E4)] border-r-0 border-l-0 text-[var(--table-column-header-text-color,inherit)] text-right font-extrabold uppercase tracking-wider text-[12px]"
                                    style={{
                                      backgroundColor: "var(--table-column-header-bg, #F4ECD8)",
                                    }}
                                  >
                                    TỔNG TIỀN:
                                  </td>
                                );
                              }
                              return (
                                <td
                                  key={`reconcile-total-${columnIndex}`}
                                  className={`p-2.5 border-b border-t border-[var(--table-border-color,#E7E5E4)] border-r-0 border-l-0 text-[var(--table-column-header-text-color,inherit)] ${
                                    columnIndex === 5
                                      ? "text-right tabular-nums font-black text-[12.5px]"
                                      : columnIndex === 6
                                        ? "text-right tabular-nums font-black text-[12.5px]"
                                        : columnIndex === 7
                                          ? `text-right tabular-nums font-black text-[13px] ${reconcileTotals.diff === 0 ? "text-emerald-300" : "text-amber-300"}`
                                          : ""
                                  }`}
                                  style={{
                                    backgroundColor:
                                      "var(--table-column-header-bg, #F4ECD8)",
                                  }}
                                >
                                  {columnIndex === 5
                                    ? formatMoneyVND(reconcileTotals.ae).replace(" ₫", "")
                                    : columnIndex === 6
                                      ? formatMoneyVND(reconcileTotals.accAmt).replace(" ₫", "")
                                      : columnIndex === 7
                                        ? (reconcileTotals.diff > 0 ? "+" : "") + formatMoneyVND(reconcileTotals.diff).replace(" ₫", "")
                                        : ""}
                                </td>
                              );
                            }
                          )}
                        </tr>
                      </tfoot>
                    </table>
                  </div>

                  {/* Pagination Footer */}
                  <div
                    className="flex items-center justify-between shrink-0 z-10 relative table-footer-pagination border-t border-slate-200"
                    style={{
                      height: "44.9802px",
                      backgroundColor:
                        "var(--table-footer-bg, var(--table-header-bg, #FAF3E8))",
                      paddingRight: "12px",
                      paddingLeft: "12px",
                      paddingTop: "3px",
                      paddingBottom: "3px"
                    }}
                  >
                    <div className="flex items-center gap-3">
                      <div className="flex items-center gap-1.5">
                        <span className="text-[12px] font-medium text-slate-500 whitespace-nowrap">
                          Hiển thị:
                        </span>
                        <Select
                          value={reconcileRowsPerPage === "all" ? "all" : String(reconcileRowsPerPage)}
                          onValueChange={(val) => {
                            setReconcileRowsPerPage(val === "all" ? "all" : Number(val));
                            setReconcileCurrentPage(1);
                          }}
                        >
                          <SelectTrigger
                            className="rounded-full px-2.5 text-[12px] font-bold font-sans normal-case text-slate-700 border-slate-200 bg-white hover:bg-slate-50 transition-colors shadow-2xs h-[22px] py-0"
                            style={{ height: "22px", width: "90px", fontSize: "12px" }}
                          >
                            <SelectValue placeholder="Chọn..." className="font-sans normal-case text-[12px]" />
                          </SelectTrigger>
                          <SelectContent className="bg-[var(--popover,#fff)] border-[#e7dbdc] z-[99999] opacity-100 font-sans shadow-xl rounded-xl">
                            <SelectItem value="10" className="text-[12px] font-medium font-sans normal-case cursor-pointer">10 dòng</SelectItem>
                            <SelectItem value="20" className="text-[12px] font-medium font-sans normal-case cursor-pointer">20 dòng</SelectItem>
                            <SelectItem value="50" className="text-[12px] font-medium font-sans normal-case cursor-pointer">50 dòng</SelectItem>
                            <SelectItem value="100" className="text-[12px] font-medium font-sans normal-case cursor-pointer">100 dòng</SelectItem>
                            <SelectItem value="all" className="text-[12px] font-medium font-sans normal-case cursor-pointer">Tất cả</SelectItem>
                          </SelectContent>
                        </Select>
                      </div>
                    </div>

                    {/* Pagination Controls */}
                    <div className="flex items-center gap-1">
                      <button
                        type="button"
                        disabled={safePage === 1}
                        onClick={() => setReconcileCurrentPage(1)}
                        className="flex items-center justify-center w-6 h-6 rounded-md border border-[#e7dbdc] bg-white text-slate-700 hover:bg-slate-50 disabled:opacity-40 disabled:hover:bg-white disabled:cursor-not-allowed transition-all shadow-2xs active:scale-95 cursor-pointer select-none"
                        title="Trang đầu"
                      >
                        <ChevronsLeft className="w-3.5 h-3.5" />
                      </button>
                      <button
                        type="button"
                        disabled={safePage === 1}
                        onClick={() => setReconcileCurrentPage((p) => Math.max(1, p - 1))}
                        className="flex items-center justify-center w-6 h-6 rounded-md border border-[#e7dbdc] bg-white text-slate-700 hover:bg-slate-50 disabled:opacity-40 disabled:hover:bg-white disabled:cursor-not-allowed transition-all shadow-2xs active:scale-95 cursor-pointer select-none"
                        title="Trang trước"
                      >
                        <ChevronLeft className="w-3.5 h-3.5" />
                      </button>
                      
                      <span className="text-[12px] font-bold text-slate-700 px-1.5 tabular-nums whitespace-nowrap text-center min-w-[70px]">
                        TRANG {safePage} / {totalPages || 1}
                      </span>

                      <button
                        type="button"
                        disabled={safePage >= totalPages || totalPages === 0}
                        onClick={() => setReconcileCurrentPage((p) => Math.min(totalPages, p + 1))}
                        className="flex items-center justify-center w-6 h-6 rounded-md border border-[#e7dbdc] bg-white text-slate-700 hover:bg-slate-50 disabled:opacity-40 disabled:hover:bg-white disabled:cursor-not-allowed transition-all shadow-2xs active:scale-95 cursor-pointer select-none"
                        title="Trang sau"
                      >
                        <ChevronRight className="w-3.5 h-3.5" />
                      </button>
                      <button
                        type="button"
                        disabled={safePage >= totalPages || totalPages === 0}
                        onClick={() => setReconcileCurrentPage(totalPages)}
                        className="flex items-center justify-center w-6 h-6 rounded-md border border-[#e7dbdc] bg-white text-slate-700 hover:bg-slate-50 disabled:opacity-40 disabled:hover:bg-white disabled:cursor-not-allowed transition-all shadow-2xs active:scale-95 cursor-pointer select-none"
                        title="Trang cuối"
                      >
                        <ChevronsRight className="w-3.5 h-3.5" />
                      </button>
                    </div>
                  </div>
                </motion.div>
              )}

              {rightPanelTab === "visuals" && (
                <motion.div
                  key="panel-visuals"
                  initial={{ opacity: 0 }}
                  animate={{ opacity: 1 }}
                  exit={{ opacity: 0 }}
                  transition={{ duration: 0.1 }}
                  className="flex-1 min-h-0 overflow-hidden"
                >
                  {analysAnalytics && (
                    <BulkPaymentAnalytics
                      analytics={analysAnalytics}
                      selectedBusiness={effectiveAnalysBusiness}
                      allBusinessUnitsValue={ALL_ANALYS_BUSINESS_UNITS}
                      searchTerm={analysSearchTerm}
                      onSearchTermChange={setAnalysSearchTerm}
                      onSelectedBusinessChange={setAnalysSelectedBusiness}
                      searchVisible={analysSearchVisible}
                      onSearchVisibleChange={setAnalysSearchVisible}
                      onResetFilters={() => {
                        setAnalysSelectedBusiness(
                          ALL_ANALYS_BUSINESS_UNITS,
                        );
                        setAnalysSearchTerm("");
                      }}
                      isBulkPaymentCardVisible={showLeftCard}
                      onToggleBulkPaymentCard={() =>
                        setShowLeftCard((visible) => !visible)
                      }
                      onViewChange={(view) => {
                        setRightPanelTab(view);
                        localStorage.setItem("bulk_payment_right_tab", view);
                      }}
                      onDrilldownToTransaction={(bu, month) => {
                        setRightPanelTab("table");
                        localStorage.setItem("bulk_payment_right_tab", "table");
                        setSearchTerm(month);
                        toast.success(`Đang lọc giao dịch: ${bu} - Tháng phát sinh ${month}`);
                      }}
                    />
                  )}
                </motion.div>
              )}
            </AnimatePresence>
          </div>
        )}
      </div>

      <Dialog open={showClearDialog} onOpenChange={setShowClearDialog}>
        <DialogContent className="sm:max-w-md border border-primary/10 shadow-2xl bg-white rounded-[2.5rem] p-10">
          <DialogHeader>
            <DialogTitle className="font-bold uppercase tracking-[0.2em] text-primary text-sm">
              Xác nhận xoá dữ liệu
            </DialogTitle>
            <DialogDescription className="font-bold text-slate-400 text-[11px] uppercase tracking-widest mt-4 leading-relaxed">
              Bạn có chắc chắn muốn xóa toàn bộ dữ liệu bảng kê? Hành động này
              không thể hoàn tác.
            </DialogDescription>
          </DialogHeader>
          <DialogFooter className="gap-4 mt-10">
            <Button
              variant="outline"
              onClick={() => setShowClearDialog(false)}
              className="border-primary/10 bg-white font-bold uppercase text-[10px] tracking-[0.2em] px-8 py-3 h-12 rounded-2xl hover:bg-primary/5 transition-all cursor-pointer"
            >
              Hủy
            </Button>
            <Button
              variant="destructive"
              onClick={() => {
                handleClearReport();
                setShowClearDialog(false);
              }}
              className="bg-rose-500 text-white font-bold uppercase text-[10px] tracking-[0.2em] px-8 py-3 h-12 rounded-2xl hover:bg-rose-600 shadow-rose-500/20 transition-all cursor-pointer border-0"
            >
              Xác nhận xoá
            </Button>
          </DialogFooter>
        </DialogContent>
      </Dialog>

      {/* Delete Confirmation Modal Dialog with Save */}
      <Dialog
        open={!!deleteConfirmTarget}
        onOpenChange={(open) => !open && setDeleteConfirmTarget(null)}
      >
        <DialogContent className="sm:max-w-md bg-white border border-slate-200 shadow-2xl rounded-2xl p-6">
          <DialogHeader className="gap-2">
            <DialogTitle className="flex items-center gap-2.5 text-rose-600 font-extrabold text-base">
              <div className="w-9 h-9 rounded-xl bg-rose-50 border border-rose-200 flex items-center justify-center shrink-0">
                <Trash2 className="w-5 h-5 text-rose-600" />
              </div>
              <span>Xác nhận xóa dòng giao dịch</span>
            </DialogTitle>
            <DialogDescription className="text-xs text-slate-600 leading-relaxed pt-1">
              Bạn có chắc chắn muốn xóa dòng giao dịch này không?
            </DialogDescription>
          </DialogHeader>

          {deleteConfirmTarget?.row && (
            <div className="bg-slate-50 rounded-xl p-3.5 border border-slate-200/80 text-xs flex flex-col gap-1.5 text-slate-800">
              <div className="flex justify-between items-center">
                <span className="text-slate-500 font-semibold">
                  Người thụ hưởng:
                </span>
                <strong className="text-slate-900 font-bold">
                  {deleteConfirmTarget.row["Beneficiary Name"] ||
                    deleteConfirmTarget.row["Full name"] ||
                    "N/A"}
                </strong>
              </div>
              <div className="flex justify-between items-center">
                <span className="text-slate-500 font-semibold">
                  STK Ngân hàng:
                </span>
                <strong className="tabular-nums text-slate-900">
                  {deleteConfirmTarget.row["Beneficiary Account No."] ||
                    deleteConfirmTarget.row["Bank Account Number"] ||
                    "N/A"}
                </strong>
              </div>
              <div className="flex justify-between items-center">
                <span className="text-slate-500 font-semibold">
                  Số tiền chuyển:
                </span>
                <strong className="tabular-nums text-emerald-700 font-extrabold">
                  {formatMoneyVND(
                    parseMoneyToNumber(
                      deleteConfirmTarget.row["Payment Amount"] ||
                        deleteConfirmTarget.row["Payment amount"] ||
                        0,
                    ),
                  )}
                </strong>
              </div>
            </div>
          )}

          {deleteConfirmTarget?.rows && deleteConfirmTarget.rows.length > 0 && (
            <div className="bg-slate-50 rounded-xl p-3.5 border border-slate-200/80 text-xs text-slate-800">
              Đang chọn xóa{" "}
              <strong className="text-rose-600 font-extrabold">
                {deleteConfirmTarget.rows.length}
              </strong>{" "}
              dòng giao dịch khỏi Bảng kê Bank Export.
            </div>
          )}

          <div className="bg-amber-50/90 border border-amber-200 p-3 rounded-xl text-[11px] text-amber-900 leading-relaxed font-medium">
            💡 Sau khi bạn bấm <strong>Bấm lưu & Cập nhật</strong>, dòng giao
            dịch sẽ bị xóa và các chỉ số tổng hợp BU, chênh lệch ròng cùng bảng
            đối soát sẽ tự động tính toán lại theo số liệu mới.
          </div>

          <DialogFooter className="gap-2 sm:gap-2 pt-2 border-t border-slate-100">
            <Button
              type="button"
              variant="outline"
              onClick={() => setDeleteConfirmTarget(null)}
              className="px-4 py-2 text-xs font-bold rounded-xl border-slate-200 text-slate-700 hover:bg-slate-100"
            >
              Hủy bỏ
            </Button>
            <Button
              type="button"
              onClick={() => {
                if (deleteConfirmTarget?.row) {
                  handleDeleteRow(deleteConfirmTarget.row);
                } else if (
                  deleteConfirmTarget?.rows &&
                  deleteConfirmTarget.rows.length > 0
                ) {
                  handleDeleteRows(deleteConfirmTarget.rows);
                }
                setDeleteConfirmTarget(null);
                toast.success(
                  "Đã xóa dòng giao dịch và lưu cập nhật dữ liệu thành công!",
                );
              }}
              className="px-4 py-2 text-xs font-extrabold bg-rose-600 hover:bg-rose-700 text-white rounded-xl shadow-xs flex items-center gap-1.5 active:scale-95 transition-all cursor-pointer"
            >
              <Save className="w-3.5 h-3.5" />
              <span>Bấm lưu & Cập nhật</span>
            </Button>
          </DialogFooter>
        </DialogContent>
      </Dialog>
    </motion.div>
  );
}
