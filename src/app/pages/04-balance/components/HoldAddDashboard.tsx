import { autoApproveTrialBalanceRow } from "../../../lib/utils/trial-balance-approval";
import { trialBalancePeriod, trialBalanceRowLabel, trialBalanceRowOrder } from "../../../lib/utils/trial-balance-presentation";
import { hideInactivePastHold, nextTrialBalanceCarry } from "../../../lib/utils/trial-balance-carry";
import { TableRestoreButton } from '../../../components/TableRestoreButton';
import { chooseExcelExport } from "../../../components/ExportScopeDialog";
/* eslint-disable @typescript-eslint/no-explicit-any */
import React, { useState, useMemo, useCallback } from "react";
import { useNavigate } from "react-router";
import { buildBulkPaymentAnalytics } from "../../../lib/utils/bulk-payment-analytics";
import {
  Download,
  Settings,
  CheckCircle2,
  ChevronDown,
  ChevronRight,
  ChevronsDownUp,
  ChevronsUpDown,
  Trash2,
} from "lucide-react";
import { Button } from "../../../components/ui/button";
import { toast } from "sonner";
import { useUiSettings } from "../../../lib/ui-settings";
import { useAppData } from "../../../lib/contexts/AppDataContext";
import { ConfirmDialog } from "../../../components/shared/ConfirmDialog";
import {
  TableInitialMark,
  TableTitleRemainder,
} from "../../../components/TableInitialMark";
import { parseMoneyToNumber, removeVietnameseTones } from "../../../lib/utils/data-utils";
import { clearBalancePageData } from "../../../lib/utils/data-clear-scopes";
import {
  resolveL07BuFromAeCode,
  getCenterInfoByAECode,
  getCenterInfoByL07,
} from "../../../lib/utils/center-utils";
import {
  DropdownMenu,
  DropdownMenuContent,
  DropdownMenuTrigger,
} from "../../../components/ui/dropdown-menu";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "../../../components/ui/select";
import { downloadHierarchicalWorkbook } from "../../../lib/utils/excel-export";
import { calculateTrialBalanceHeaderTotals } from "../../../lib/utils/trial-balance-header-totals";
import { grossPayTotalsByMonth, normalizeGrossPayBusiness } from "../../../lib/utils/gross-pay";
import { parsePayrollMonth } from "../../../lib/utils/hold-carryover";

const fmt = (n: number) => {
  const rounded = Math.round(n);
  return new Intl.NumberFormat("vi-VN", {
    minimumFractionDigits: 0,
    maximumFractionDigits: 0,
  }).format(rounded);
};

const toRoman = (num: number) => {
  const roman = [
    "",
    "I",
    "II",
    "III",
    "IV",
    "V",
    "VI",
    "VII",
    "VIII",
    "IX",
    "X",
    "XI",
    "XII",
  ];
  return roman[num] || num.toString();
};

const parseMonthFromFileName = (
  fileName: string,
  globalMonth?: string,
): string | null => {
  if (!fileName) return null;
  const match = fileName.match(/\b(0?[1-9]|1[0-2])[./-](20\d{2})\b/);
  if (match) {
    const m = parseInt(match[1], 10);
    const y = parseInt(match[2], 10);
    return `${m < 10 ? "0" + m : m}.${y}`;
  }
  const tMatch = fileName.match(/(Th\w*|T|Month\s*)(0?[1-9]|1[0-2])\b/i);
  if (tMatch) {
    const m = parseInt(tMatch[2], 10);
    const ref = globalMonth || "03.2026";
    const refParts = ref.split(".");
    const currentMonthNum = parseInt(refParts[0], 10) || 3;
    const currentYearNum = parseInt(refParts[1], 10) || 2026;
    const y = m > currentMonthNum ? currentYearNum - 1 : currentYearNum;
    return `${m < 10 ? "0" + m : m}.${y}`;
  }
  return null;
};

const getNextMonthStr = (periodStr: string): string => {
  if (!periodStr) return "";
  const regex = /(?:tháng|thang|t)?\s*(\d{1,2})[/\-.]\s*(\d{4})/i;
  const match = periodStr.match(regex);
  if (match) {
    let m = parseInt(match[1], 10);
    let y = parseInt(match[2], 10);
    m++;
    if (m > 12) {
      m = 1;
      y++;
    }
    return `Tháng ${m}/${y}`;
  }
  const parts = periodStr.split(".");
  if (parts.length === 2) {
    let m = parseInt(parts[0], 10);
    let y = parseInt(parts[1], 10);
    m++;
    if (m > 12) {
      m = 1;
      y++;
    }
    return `${m < 10 ? "0" + m : m}.${y}`;
  }
  return periodStr;
};

const isSameMonthForSumIf = (rowMonthRaw?: string, workMonthRaw?: string): boolean => {
  if (!workMonthRaw) return false;
  if (!rowMonthRaw) return true; // If no month is specified in row, assume it matches the target context
  
  const parsePartsStr = (s: string) => {
    // Clean string: trim, lower case, remove spaces, replace Vietnamese "tháng" if any
    const clean = s.trim().toLowerCase()
      .replace(/\s+/g, "")
      .replace(/^tháng/, ""); // remove prefix "tháng" if any
    
    // Now standard separators are /, -, _, .
    const dotSep = clean.replace(/[-_/]/g, ".");
    const parts = dotSep.split(".");
    
    if (parts.length >= 2) {
      // Find year (4 digits)
      const yearIdx = parts.findIndex(p => p.length === 4 && !isNaN(parseInt(p, 10)));
      if (yearIdx !== -1) {
        const year = parseInt(parts[yearIdx], 10);
        let month = 0;
        if (yearIdx === 1) {
          month = parseInt(parts[0], 10);
        } else if (yearIdx === 2) {
          month = parseInt(parts[1], 10);
        } else if (yearIdx === 0) {
          month = parseInt(parts[1], 10);
        }
        if (month >= 1 && month <= 12 && year > 0) {
          return { month, year };
        }
      } else {
        // Fallback: e.g. "01.26" -> [01, 26]
        const first = parseInt(parts[0], 10);
        const second = parseInt(parts[1], 10);
        if (!isNaN(first) && !isNaN(second)) {
          if (first > 12) {
            return { month: second, year: first < 100 ? first + 2000 : first };
          } else {
            return { month: first, year: second < 100 ? second + 2000 : second };
          }
        }
      }
    }
    
    // Try matching formats like "tháng 3/2026", etc via regex
    const match = s.match(/(tháng|thg)?\s*(\d{1,2})\s*([./-])\s*(\d{4})/i);
    if (match) {
      return { month: parseInt(match[2], 10), year: parseInt(match[3], 10) };
    }
    
    return null;
  };

  const p1 = parsePartsStr(rowMonthRaw);
  const p2 = parsePartsStr(workMonthRaw);
  
  if (p1 && p2) {
    return p1.month === p2.month && p1.year === p2.year;
  }
  
  // Last resort: simple clean direct compare
  const directClean = (str: string) => str.trim().toLowerCase()
    .replace(/\s+/g, "")
    .replace(/^tháng/, "")
    .replace(/[/_]/g, ".");
  return directClean(rowMonthRaw) === directClean(workMonthRaw);
};

const isExactReportPeriod = (
  row: Pick<BuRow, "reportMonth" | "month">,
  period: string,
) => {
  const reportMonth = String(row.reportMonth || row.month || "").trim();
  return reportMonth.length > 0 && isSameMonthForSumIf(reportMonth, period);
};

const isPastMonthHold = (row: any, currentMonthNum: number, currentYearNum: number): boolean => {
  if (!row) return false;

  const rawSource = String(row["Sheet Source"] || "").toUpperCase();
  const isBonusSummer = rawSource.includes("BONUS") && (
    rawSource.includes("SUMMER") || 
    rawSource.includes("INSTRUCTOR") || 
    rawSource.includes("INTROSTION")
  );
  if (isBonusSummer) return false;

  let phatSinh = "";
  for (const k of Object.keys(row)) {
    const kNorm = k.toLowerCase()
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .replace(/\u0111/g, "d")
      .trim();
    if (kNorm === "thang phat sinh") {
      phatSinh = String(row[k] || "").trim();
      break;
    }
  }

  if (!phatSinh) {
    phatSinh = String(
      row["Tháng phát sinh"] || 
      row["tháng phát sinh"] || 
      row["Thang phat sinh"] || 
      row["thang phat sinh"] || 
      row["Tháng Phát Sinh"] ||
      ""
    ).trim();
  }

  if (phatSinh) {
    const parts = phatSinh.split(/[./-]/);
    if (parts.length === 2) {
      const m = parseInt(parts[0], 10);
      const y = parseInt(parts[1], 10);
      let nghiepVu = String(row["Nghiệp vụ"] || "").trim().toUpperCase();
      if (!nghiepVu) {
        for (const k of Object.keys(row)) {
          const kLower = k.toLowerCase();
          if (kLower.includes("nghiệp vụ") || kLower.includes("nghiệp vụ") || kLower.includes("nghiep vu")) {
            nghiepVu = String(row[k] || "").trim().toUpperCase();
            break;
          }
        }
      }
      const isHoldOrCancel = nghiepVu.includes("HOLD") || nghiepVu.includes("CANCEL");
      if (isHoldOrCancel && !isNaN(m) && !isNaN(y)) {
        if (y < currentYearNum || (y === currentYearNum && m < currentMonthNum)) {
          return true;
        }
      }
    }
  }

  // Extract from any possible "Tình trạng thanh toán" or "Trạng thái" key
  let tttt = "";
  for (const k of Object.keys(row)) {
    const kLower = k.toLowerCase();
    if (kLower.includes("tình trạng thanh toán") || kLower.includes("tình trạng") || kLower.includes("tttt") || kLower.includes("tinh trang thanh toan")) {
      tttt = String(row[k] || "");
      break;
    }
  }
  if (!tttt) {
    tttt = String(row["Tình trạng thanh toán"] || row["Tình Trạng Thanh Toán"] || row["TÌNH TRẠNG THANH TOÁN"] || row["Tháng phát sinh"] || row["tháng phát sinh"] || row["Trạng thái"] || row["trạng thái"] || "");
  }

  const ttttLower = tttt.toLowerCase().trim();
  if (ttttLower.includes("pending")) {
    const pendingMatch = ttttLower.match(/pending\s*(?:tháng|thg|t)?\s*(\d+)(\s*([./-])\s*(\d+))?/i);
    if (pendingMatch) {
      const m = parseInt(pendingMatch[1], 10);
      let y = currentYearNum;
      
      const yrMatch = ttttLower.match(/\b(202\d)\b/);
      if (yrMatch) {
        y = parseInt(yrMatch[1], 10);
      } else if (m === 11 || m === 12) {
        y = 2025;
      }
      
      // If of a past month
      if (y < currentYearNum || (y === currentYearNum && m < currentMonthNum)) {
        return true;
      }
    }
  }
  return false;
};

interface BuRow {
  id: string;
  month: string;
  reportMonth?: string;
  bu: string;
  thu: number;
  chi: number;
  add: number;
  hold: number;
  cancel: number;
  bonus: number;
  rawAdd: number;
  rawHold: number;
  rawCancel: number;
  rawBonus: number;
  rawHoldPending: number;
  ghiChu: string;
  confirmed: boolean;
  displayMonth?: string;
  customMonthDisplay?: string;
  openHold?: number;
  rawOpenHold?: number;
  _dimmed?: boolean;
  _excludeFromTotals?: boolean;
  _isPastHoldApprove?: boolean;
  _isOpeningHold?: boolean;
  lenh?: string;
  isPaidStatus?: boolean;
}

export function HoldAddDashboard() {
  const uiSettings = useUiSettings();
  const { appData, updateAppData } = useAppData();
  const navigate = useNavigate();
  const [yearFilter, setYearFilter] = useState("all");
  const [itemsPerPage, setItemsPerPage] = useState<number | "all">("all");
  const [showClearBalancePageDialog, setShowClearBalancePageDialog] = useState(false);
  const [showDeletePeriodDialog, setShowDeletePeriodDialog] = useState(false);
  const currentPeriodVal = appData.globalMonth || "03.2026";
  const currentPeriodParts = currentPeriodVal.split(".");
  const currentPeriodMonthNum = parseInt(currentPeriodParts[0], 10) || 3;
  const currentPeriodYearNum = parseInt(currentPeriodParts[1], 10) || 2026;
  const currentPeriod = `Tháng ${currentPeriodMonthNum}/${currentPeriodYearNum}`;

  const [expanded, setExpanded] = useState<Set<string>>(
    new Set([
      "Kỳ hiện tại",
      currentPeriod,
      "Tháng 1/2026",
      "Tháng 12/2025",
      "Tháng 2/2026",
    ]),
  );

  const getSavedDataForPeriod = useCallback((map: Record<string, any> | undefined, month: string) => {
    if (!map || !month) return undefined;
    if (map[month]) return map[month];
    const matchM = month.match(/\b([1-9]|1[0-2])\b/);
    const matchY = month.match(/20\d{2}/);
    if (matchM) {
      const mNum = parseInt(matchM[0], 10);
      const yNum = matchY ? parseInt(matchY[0], 10) : 2026;
      const v1 = `Tháng ${mNum}/${yNum}`;
      const v2 = `Tháng ${String(mNum).padStart(2, "0")}/${yNum}`;
      const v3 = `${String(mNum).padStart(2, "0")}.${yNum}`;
      const v4 = `Tháng ${mNum}.${yNum}`;
      const v5 = `Tháng ${String(mNum).padStart(2, "0")}.${yNum}`;
      return map[v1] || map[v2] || map[v3] || map[v4] || map[v5];
    }
    return undefined;
  }, []);

  const isPeriodSaved = useCallback(
    (month: string) => {
      if (!month || !appData.SavedPeriods_HoldAdd) return false;
      return !!getSavedDataForPeriod(appData.SavedPeriods_HoldAdd, month);
    },
    [appData.SavedPeriods_HoldAdd, getSavedDataForPeriod],
  );

  const handleDeleteSavedPeriod = useCallback(() => {
    updateAppData((prev: any) => {
      const nextSavedAll = prev.SavedBal_PayrollTrial
        ? JSON.parse(JSON.stringify(prev.SavedBal_PayrollTrial))
        : {};
      const nextSavedPeriods = prev.SavedPeriods_HoldAdd
        ? JSON.parse(JSON.stringify(prev.SavedPeriods_HoldAdd))
        : {};
      const nextSavedRows = prev.SavedRows_HoldAdd
        ? JSON.parse(JSON.stringify(prev.SavedRows_HoldAdd))
        : {};
      const nextSavedRowsMeta = prev.SavedRows_HoldAdd_Meta
        ? JSON.parse(JSON.stringify(prev.SavedRows_HoldAdd_Meta))
        : {};

      const currentMonthVal = appData.globalMonth || "03.2026";
      const currentMonthNum = parseInt(currentMonthVal.split(".")[0], 10) || 3;
      const currentYearNum = parseInt(currentMonthVal.split(".")[1], 10) || 2026;
      const variations = [
        currentPeriod,
        `Tháng ${currentMonthNum}/${currentYearNum}`,
        `Tháng ${String(currentMonthNum).padStart(2, "0")}/${currentYearNum}`,
        `Tháng ${currentMonthNum}.${currentYearNum}`,
        `Tháng ${String(currentMonthNum).padStart(2, "0")}.${currentYearNum}`,
        `${String(currentMonthNum).padStart(2, "0")}.${currentYearNum}`,
      ];

      variations.forEach((v) => {
        delete nextSavedPeriods[v];
        delete nextSavedRows[v];
        delete nextSavedRowsMeta[v];
      });

      const nextMonthStr = getNextMonthStr(currentPeriod);
      if (nextMonthStr) {
        delete nextSavedAll[nextMonthStr];
        const nextMonthNum = (currentMonthNum % 12) + 1;
        const nextYearNum = currentMonthNum === 12 ? currentYearNum + 1 : currentYearNum;
        const nextVariations = [
          nextMonthStr,
          `Tháng ${nextMonthNum}/${nextYearNum}`,
          `Tháng ${String(nextMonthNum).padStart(2, "0")}/${nextYearNum}`,
          `Tháng ${nextMonthNum}.${nextYearNum}`,
          `Tháng ${String(nextMonthNum).padStart(2, "0")}.${nextYearNum}`,
          `${String(nextMonthNum).padStart(2, "0")}.${nextYearNum}`,
        ];
        nextVariations.forEach((nv) => {
          delete nextSavedAll[nv];
        });
      }

      return {
        ...prev,
        SavedBal_PayrollTrial: nextSavedAll,
        SavedPeriods_HoldAdd: nextSavedPeriods,
        SavedRows_HoldAdd: nextSavedRows,
        SavedRows_HoldAdd_Meta: nextSavedRowsMeta,
      };
    });
    toast.success(`Đã xóa dữ liệu đã lưu ${currentPeriod}!`);
  }, [updateAppData, currentPeriod, appData.globalMonth]);

  const handleClearBalancePage = useCallback(() => {
    updateAppData(clearBalancePageData);
    setShowClearBalancePageDialog(false);
    toast.success("Đã xóa dữ liệu trang Balance; Timesheet, Audit và Master được giữ nguyên");
  }, [updateAppData]);

  // Extracted month conversion utility that preserves the year
  const extractMonth = useCallback(
    (str: string) => {
      if (!str) return null;
      const s = str.toUpperCase().trim();

      // 1. Matches "Tháng MM/YYYY", "T MM/YYYY", "MM/YYYY"
      const yrMatch = s.match(
        /(?:THÁNG|THANG|T)?\s*(\d{1,2})(?:[./\- ]|NAM\s+|YEAR\s+)+(\d{4})/i,
      );
      if (yrMatch) {
         const m = parseInt(yrMatch[1], 10);
         const y = parseInt(yrMatch[2], 10);
         if (m >= 1 && m <= 12) return `Tháng ${m}/${y}`;
      }

      // 2. Matches [D]D/[M]M e.g. "15/1", "1/2"
      const dmMatch = s.match(/\b\d{1,2}[./-]\s*(\d{1,2})\b/);
      if (dmMatch) {
         const m = parseInt(dmMatch[1], 10);
         if (m >= 1 && m <= 12) {
           let y = currentPeriodYearNum;
           // If extracted month > current month, it's likely from the previous year
           if (m > currentPeriodMonthNum && (currentPeriodYearNum === 2025 || currentPeriodYearNum === 2026)) {
             y = currentPeriodYearNum - 1;
           }
           return `Tháng ${m}/${y}`;
         }
      }

      // 3. Matches "T MM", "Tháng MM"
      const tMatch = s.match(/T[HÁNG]*\s*(\d+)/i);
      if (tMatch) {
        const m = parseInt(tMatch[1], 10);
        if (m >= 1 && m <= 12) {
          let y = currentPeriodYearNum;
          if (m === 11 || m === 12) {
            y = currentPeriodYearNum === 2025 ? 2025 : (currentPeriodYearNum === 2026 ? 2025 : currentPeriodYearNum);
          } else if (m > currentPeriodMonthNum && (currentPeriodYearNum === 2025 || currentPeriodYearNum === 2026)) {
            y = currentPeriodYearNum - 1;
          }
          return `Tháng ${m}/${y}`;
        }
      }

      // 4. Matches "MM.YYYY" (e.g. 02.2026)
      const dotMatch = s.match(/^(\d{1,2})\.(\d{4})$/);
      if (dotMatch) {
        return `Tháng ${parseInt(dotMatch[1], 10)}/${parseInt(dotMatch[2], 10)}`;
      }

      return null;
    },
    [currentPeriodMonthNum, currentPeriodYearNum],
  );

  const getMonthNum = useCallback(
    (mStr: string) => {
      if (!mStr) return 0;
      const match = mStr.match(/(\d+)\/(\d+)/);
      if (match) {
        const m = parseInt(match[1], 10);
        const y = parseInt(match[2], 10);
        return y * 100 + m;
      }
      const matchOnlyMonth = mStr.match(/\d+/);
      if (matchOnlyMonth) {
        const m = parseInt(matchOnlyMonth[0], 10);
        let y = currentPeriodYearNum;
        if (m > currentPeriodMonthNum && (currentPeriodYearNum === 2025 || currentPeriodYearNum === 2026)) {
          y = currentPeriodYearNum - 1;
        }
        return y * 100 + m;
      }
      return 0;
    },
    [currentPeriodMonthNum, currentPeriodYearNum],
  );
  const currentPeriodNum = useMemo(
    () => getMonthNum(currentPeriod),
    [currentPeriod, getMonthNum],
  );

  // Comparison metrics check
  const selectedMonth = appData.globalMonth || "03.2026";
  const fileMonths = useMemo(() => {
    return (
      appData.Ae_Global_Inputs?.map(
        (f: any) =>
          f?.month ||
          parseMonthFromFileName(f?.name || "", appData.globalMonth) ||
          "03.2026",
      ) || []
    );
  }, [appData.Ae_Global_Inputs, appData.globalMonth]);
  const isMonthMatched = fileMonths.includes(selectedMonth);

  const bulkPaymentDiff = useMemo(() => {
    const currentMonthVal = appData.globalMonth || "03.2026";
    const currentPeriodParts = currentMonthVal.split(".");
    const currentMonthNum = parseInt(currentPeriodParts[0], 10) || 3;
    const currentYearNum = parseInt(currentPeriodParts[1], 10) || 2026;
    const targetMonthLabelComp = `Tháng ${currentMonthNum}/${currentYearNum}`;
    const monthShortStrComp = `T${currentMonthNum}`;
    const monthDashStrComp = `${currentMonthNum}/${currentYearNum}`;
  
    const monMatchComp = (s: string) => {
      if (!s) return null;
      const up = String(s).toUpperCase().trim();
      const yrMatch = up.match(
        /(?:THÁNG|THANG|T)?\s*(\d{1,2})(?:[./\- ]|NAM\s+|YEAR\s+)+(\d{4})/i,
      );
      if (yrMatch) {
        const m = parseInt(yrMatch[1], 10);
        const y = parseInt(yrMatch[2], 10);
        return `Tháng ${m}/${y}`;
      }
      const mMatch = up.match(/(?:THÁNG|THANG|T)\s*(\d+)/i);
      if (mMatch) {
        const m = parseInt(mMatch[1], 10);
        let y = currentYearNum;
        if (m === 11 || m === 12) {
          y = currentYearNum === 2025 ? 2025 : (currentYearNum === 2026 ? 2025 : currentYearNum);
        } else if (m > currentMonthNum && (currentYearNum === 2025 || currentYearNum === 2026)) {
          y = currentYearNum - 1;
        }
        return `Tháng ${m}/${y}`;
      }
      const dmMatch = up.match(/\b\d{1,2}[./-]\s*(\d{1,2})\b/);
      if (dmMatch) {
        const m = parseInt(dmMatch[1], 10);
        let y = currentYearNum;
        if (m > currentMonthNum && (currentYearNum === 2025 || currentYearNum === 2026)) {
          y = currentYearNum - 1;
        }
        return `Tháng ${m}/${y}`;
      }
      return null;
    };
  
    const isMonthInStrComp = (s: string) => {
      const up = String(s || "").toUpperCase();
      return (
        up.includes(targetMonthLabelComp.toUpperCase()) ||
        up.includes(monthShortStrComp.toUpperCase()) ||
        up.includes(monthDashStrComp)
      );
    };

    const sheet1Total = appData.Sheet1_AE?.data?.reduce(
      (sum: number, r: any) => {
        const rowMonthStr = String(r["Tháng báo cáo"] || r["_fileMonth"] || r["Tháng"] || "").trim();
        if (rowMonthStr && !isSameMonthForSumIf(rowMonthStr, currentMonthVal)) return sum;
        
        let biz = String(r["Business"] || r["BU"] || "Unknown").trim().toUpperCase();
        if (biz === "AHN_HP") biz = "AHP";
        if (biz !== "AHN" && biz !== "AHP" && biz !== "ATH" && biz !== "ATN" && biz !== "APT") return sum;
        
        return sum + parseMoneyToNumber(
          r["TOTAL PAYMENT"] || 
          r["Total Payment"] || 
          r["Payment Amount"] || 
          r["Grand Total"] || 
          r["GRAND TOTAL"] || 
          r["PAYMENT AMOUNT"] || 
          0
        );
      },
      0
    ) || 0;

    const idToSheet1: Record<string, string> = {};
    const nameToSheet1: Record<string, string> = {};
    const accToSheet1: Record<string, string> = {};

    appData.Sheet1_AE?.data?.forEach((row: any) => {
      const id = String(row["ID Number"] || row["Mã AE"] || row["Mã ae"] || "").trim();
      const name = removeVietnameseTones(row["Full name"] || row["Beneficiary Name"] || "").toUpperCase();
      const acc = String(row["Bank Account Number"] || row["Beneficiary Account No."] || "").trim();
      
      let biz = row["Business"] || row["BU"] || "Unknown";
      if (biz === "AHN_HP") biz = "AHP";
      
      if (id) idToSheet1[id] = biz;
      if (name) nameToSheet1[name] = biz;
      if (acc) accToSheet1[acc] = biz;
    });

    let holdTotal = 0;
    appData.Hold_AE?.data?.forEach((row: any) => {
      const rowMonthRaw = String(row["Tháng báo cáo"] || row["_fileMonth"] || row["Tháng"] || "").trim();
      const extracted = monMatchComp(rowMonthRaw);

      if (extracted && extracted !== targetMonthLabelComp) return;
      if (!extracted && rowMonthRaw && !isMonthInStrComp(rowMonthRaw)) return;

      let biz = row["BU"] || row["Business"] || "";
      if (biz) biz = String(biz).trim().toUpperCase();
      if (biz === "AHN_HP") biz = "AHP";

      const id = String(row["ID Number"] || row["Mã AE"] || row["Mã ae"] || "").trim();
      const name = removeVietnameseTones(row["Full name"] || row["Beneficiary Name"] || "").toUpperCase();
      const acc = String(row["Bank Account Number"] || row["Beneficiary Account No."] || "").trim();

      if (!biz || biz === "UNKNOWN") biz = idToSheet1[id];
      if ((!biz || biz === "UNKNOWN") && acc) biz = accToSheet1[acc];
      if ((!biz || biz === "UNKNOWN") && name) biz = nameToSheet1[name];

      // fallback
      if (!biz || biz === "UNKNOWN") {
        const textToMatch = [ row["Sheet Source"], row["CENTER NOTE"], row["Mã ae"], row["Note"], row["Full name"] ]
          .map(v => String(v || "").toUpperCase()).join(" ");
        if (textToMatch.includes("HN") || textToMatch.includes("AHN")) biz = "AHN";
        else if (textToMatch.includes("AHP") || textToMatch.includes("HAIPHONG")) biz = "AHP";
        else if (textToMatch.includes("ATH") || textToMatch.includes("THANH HOA")) biz = "ATH";
        else if (textToMatch.includes("ATN") || textToMatch.includes("THAI NGUYEN")) biz = "ATN";
        else if (textToMatch.includes("APT") || textToMatch.includes("PHU THO")) biz = "APT";
        else biz = "AHN";
      }

      if (biz === "AHN_HP") biz = "AHP";

      if (biz !== "AHN" && biz !== "AHP" && biz !== "ATH" && biz !== "ATN" && biz !== "APT") return;

      let val = parseMoneyToNumber(row["TOTAL PAYMENT"] || row["Grand Total"] || row["GRAND TOTAL"] || row["Payment Amount"] || 0);
      const nghiepVu = String(row["Nghiệp vụ"] || "").toLowerCase();
      const label = String(row["Sheet Source"] || "").toUpperCase() || (val >= 0 ? "ADD" : "HOLD");
      const isHold = label.includes("HOLD") || nghiepVu.includes("hold");
      const isAdd = label.includes("ADD") || (!isHold && val > 0) || nghiepVu.includes("add");

      const command = String(row["Lệnh"] || "").trim().toUpperCase();
      if (command === "-") return;
      
      const sheetSource = String(row["Sheet Source"] || "").toLowerCase();

      if (sheetSource.includes("sheet 1 ae") || sheetSource.includes("sheet 1")) return;
      if (row._dimmed) return;

      const phatSinhRaw = String(row["Tháng phát sinh"] || row["tháng phát sinh"] || row["Thang phat sinh"] || "").trim();
      const psMonth = monMatchComp(phatSinhRaw) || extracted;
      const getMonthNumLocal = (mStr: string) => {
        if (!mStr) return 0;
        const match = mStr.match(/(\d+)\/(\d+)/);
        if (match) {
          return parseInt(match[2], 10) * 100 + parseInt(match[1], 10);
        }
        return 0;
      };
      const isTargetHC = (isHold || nghiepVu === "cancel" || nghiepVu.includes("cancel")) &&
                          extracted === targetMonthLabelComp &&
                          (psMonth ? getMonthNumLocal(psMonth) <= getMonthNumLocal(targetMonthLabelComp) : true);

      if (isPastMonthHold(row, currentMonthNum, currentYearNum) && !isTargetHC) {
        val = 0;
      } else if (isAdd) {
        val = Math.abs(val);
      } else {
        val = -Math.abs(val);
      }

      if (val !== 0) {
        holdTotal += val;
      }
    });

    const calculatedTotal = sheet1Total + holdTotal;

    const aeTotal = appData.Bank_North_AE?.data?.reduce(
      (sum: number, r: any) => {
        const rowMonthRaw = String(r["Tháng báo cáo"] || r["_fileMonth"] || r["Tháng"] || "").trim();
        const extracted = monMatchComp(rowMonthRaw);
        
        const fMonthRaw = String(r["_fileMonth"] || "").trim();
        const fMonthExtracted = monMatchComp(fMonthRaw);
        
        let isMonthMatch = false;
        if (fMonthExtracted === targetMonthLabelComp) {
          isMonthMatch = true;
        } else if (extracted && extracted === targetMonthLabelComp) {
          isMonthMatch = true;
        } else if (!extracted && rowMonthRaw && isMonthInStrComp(rowMonthRaw)) {
          isMonthMatch = true;
        }

        if (!isMonthMatch) return sum;

        // Resolve business unit for each bank row to ensure we only sum AHN & AHP
        let biz = r["Business"] || r["BU"] || "";
        if (biz) biz = String(biz).trim().toUpperCase();
        if (biz === "AHN_HP") biz = "AHP";

        const id = String(r["ID Number"] || r["Mã AE"] || r["Mã ae"] || "").trim();
        const name = removeVietnameseTones(r["Full name"] || r["Beneficiary Name"] || "").toUpperCase();
        const acc = String(r["Bank Account Number"] || r["Beneficiary Account No."] || "").trim();

        if (!biz || biz === "UNKNOWN") biz = idToSheet1[id];
        if ((!biz || biz === "UNKNOWN") && acc) biz = accToSheet1[acc];
        if ((!biz || biz === "UNKNOWN") && name) biz = nameToSheet1[name];

        // fallback based on text in the bank row
        if (!biz || biz === "UNKNOWN") {
          const textToMatch = [ r["Sheet Source"], r["CENTER NOTE"], r["Mã ae"], r["Note"], r["Full name"], r["Payment details"] ]
            .map(v => String(v || "").toUpperCase()).join(" ");
          if (textToMatch.includes("HN") || textToMatch.includes("AHN")) biz = "AHN";
          else if (textToMatch.includes("AHP") || textToMatch.includes("HAIPHONG")) biz = "AHP";
          else if (textToMatch.includes("ATH") || textToMatch.includes("THANH HOA")) biz = "ATH";
          else if (textToMatch.includes("ATN") || textToMatch.includes("THAI NGUYEN")) biz = "ATN";
          else if (textToMatch.includes("APT") || textToMatch.includes("PHU THO")) biz = "APT";
          else biz = "AHN"; // default fallback to AHN
        }

        if (biz === "AHN_HP") biz = "AHP";

        if (biz !== "AHN" && biz !== "AHP" && biz !== "ATH" && biz !== "ATN" && biz !== "APT") return sum;

        return sum + (parseMoneyToNumber(r["TOTAL PAYMENT"]) || 0);
      },
      0
    ) || 0;

    return calculatedTotal - aeTotal;
  }, [
    appData.Sheet1_AE?.data,
    appData.Hold_AE?.data,
    appData.Bank_North_AE?.data,
    appData.globalMonth
  ]);

  const isBulkDiffZero = Math.abs(bulkPaymentDiff) < 1;
  const isOkToDisplayValues = isMonthMatched && isBulkDiffZero;

  // eslint-disable-next-line @typescript-eslint/no-unused-vars
  const _discrepancyCandidates = useMemo(() => {
    const list: { type: string; label: string; amount: number }[] = [];
    const targetDiff = Math.abs(bulkPaymentDiff);
    if (targetDiff < 1) return list;

    const holdRows = appData.Hold_AE?.data || [];
    const bankRows = appData.Bank_North_AE?.data || [];

    // Single rows from Hold_AE matching targetDiff
    holdRows.forEach((r: any, idx: number) => {
      const val = parseMoneyToNumber(r["TOTAL PAYMENT"] || r["TOTAL"] || r["Total"] || 0);
      const absVal = Math.abs(val);
      if (Math.abs(absVal - targetDiff) < 10) {
        list.push({
          type: "Hold/Add Row",
          label: `Bảng Hold AE (Dòng #${idx + 1}): AE "${r["Full name"] || "N/A"}" có số tiền ${val.toLocaleString()} (Bản ghi ở nguồn Hold đang ${val < 0 ? "trừ" : "cộng"} số tiền này).`,
          amount: val
        });
      }
    });

    // Single rows from Bank_North_AE matching targetDiff
    bankRows.forEach((r: any, idx: number) => {
      const val = parseMoneyToNumber(r["TOTAL PAYMENT"] || r["Payment Amount"] || 0);
      const absVal = Math.abs(val);
      if (Math.abs(absVal - targetDiff) < 10) {
        list.push({
          type: "Bank North Row",
          label: `Bảng Bank North (Dòng #${idx + 1}): AE "${r["Full name"] || r["Beneficiary Name"] || "N/A"}" có số tiền thanh toán thực tế là ${val.toLocaleString()}.`,
          amount: val
        });
      }
    });

    // Pair of rows from Hold_AE matching targetDiff (if row counts under 500 to keep it highly performant)
    if (list.length < 5 && holdRows.length < 500) {
      for (let i = 0; i < holdRows.length; i++) {
        const val1 = parseMoneyToNumber(holdRows[i]["TOTAL PAYMENT"] || 0);
        const name1 = holdRows[i]["Full name"] || "N/A";
        for (let j = i + 1; j < holdRows.length; j++) {
          const val2 = parseMoneyToNumber(holdRows[j]["TOTAL PAYMENT"] || 0);
          const name2 = holdRows[j]["Full name"] || "N/A";

          if (Math.abs(Math.abs(val1) + Math.abs(val2) - targetDiff) < 10) {
            list.push({
              type: "Hold Pair (Sum)",
              label: `Khớp tổng 2 dòng Hold: "${name1}" (${val1.toLocaleString()}) + "${name2}" (${val2.toLocaleString()}) = ${(Math.abs(val1) + Math.abs(val2)).toLocaleString()}.`,
              amount: Math.abs(val1) + Math.abs(val2)
            });
          }
          if (Math.abs(val1 + val2 - targetDiff) < 10) {
            list.push({
              type: "Hold Pair (Sum)",
              label: `Khớp tổng 2 dòng Hold: "${name1}" (${val1.toLocaleString()}) & "${name2}" (${val2.toLocaleString()}) = ${(val1 + val2).toLocaleString()}.`,
              amount: val1 + val2
            });
          }
        }
      }
    }

    return list.slice(0, 5); // display top 5 most relevant possibilities
  }, [appData.Hold_AE?.data, appData.Bank_North_AE?.data, bulkPaymentDiff]);

  const data = useMemo(() => {
    const sheet1Rows = appData.Sheet1_AE?.data || [];
    const holdRows = appData.Hold_AE?.data || [];

    let defaultMonth = currentPeriod;
    for (const r of sheet1Rows) {
      const m = extractMonth(
        String(r["_fileMonth"] || r["Tháng"] || r["Tháng báo cáo"] || ""),
      );
      if (m) {
        defaultMonth = m;
        break;
      }
    }

    const buStats: Record<
      string,
      BuRow & {
        rawThu: number;
        rawChi: number;
      }
    > = {};

    const getBuKey = (m: string, b: string) => `${m}_${b}`;

    const accToBU: Record<string, string> = {};
    const extractBU = (str: string) => {
      const u = str.toUpperCase();
      if (u.includes("HN") || u.includes("AHN")) return "AHN";
      if (u.includes("HP") || u.includes("AHP") || u.includes("HAI PHONG"))
        return "AHP";
      if (u.includes("TH") || u.includes("ATH") || u.includes("THANH HOA"))
        return "ATH";
      if (u.includes("TN") || u.includes("ATN") || u.includes("THAI NGUYEN"))
        return "ATN";
      if (u.includes("VT") || u.includes("AVT")) return "AVT";
      return "UNKNOWN";
    };

    sheet1Rows.forEach((r: any) => {
      let biz = String(r["Business"] || "").trim();
      const l07 = String(r["L07"] || r["Mã AE"] || r["Mã ae"] || "").trim();
      if (!biz && l07) {
        const resolved =
          resolveL07BuFromAeCode(l07) ||
          getCenterInfoByAECode(l07) ||
          getCenterInfoByL07(l07);
        if (resolved) {
          biz =
            ("bu" in resolved
              ? resolved.bu
              : "bus" in resolved
                ? resolved.bus
                : "") || biz;
        }
      }
      if (!biz)
        biz = extractBU(
          String(r["CENTER NOTE"] || r["Note"] || r["L07"] || ""),
        );
      if (!biz || biz === "UNKNOWN") biz = "AHN";
      if (biz === "AHN_HP") biz = "AHP";

      const acc = String(
        r["Bank Account Number"] || r["Beneficiary Account No."] || "",
      ).trim();
      if (acc) accToBU[acc] = biz;

      const rawMonthStr = String(
        r["Tháng báo cáo"] || "",
      ).trim();
      const monthStr = extractMonth(rawMonthStr);
      
      // Also check transaction month for Sheet 1 if available
      const transStr = String(r["Tháng phát sinh"] || r["tháng phát sinh"] || r["Thang phat sinh"] || "").trim();
      const psMonth = extractMonth(transStr);

      const month = monthStr || psMonth || defaultMonth;
      
      // Enforce: only allow the selected month (currentPeriod) to appear on the Trial Balance dashboard
      if (month !== currentPeriod) return;

      const key = getBuKey(month, biz);

      if (!buStats[key]) {
        buStats[key] = {
          id: key,
          month,
          reportMonth: month,
          displayMonth: month,
          bu: biz,
          rawThu: 0,
          rawChi: 0,
          thu: 0,
          chi: 0,
          add: 0,
          hold: 0,
          cancel: 0,
          rawAdd: 0,
          rawHold: 0,
          rawCancel: 0,
          bonus: 0,
          rawBonus: 0,
          rawHoldPending: 0,
          ghiChu: "",
          lenh: "",
          confirmed: false,
        };
      }
      buStats[key].rawThu += parseMoneyToNumber(
        r["TOTAL PAYMENT"] ||
          r["Grand Total"] ||
          r["GRAND TOTAL"] ||
          r["Payment Amount"] ||
          0,
      );
      buStats[key].thu = buStats[key].rawThu;
    });

    // Bank Export (bulkRows) is NO LONGER USED to define `chi` (Lương Hold của tháng).
    // Specifically: "lương hold của tháng tại trial balance = cột nghiệp vụ bảng hold ae = HOLD +
    // sheet source chứa tháng báo cáo hiện tại + tháng báo cáo = card chọn tháng chọn = giá trị tại cột +
    // cột total payment mang giá trị âm, không đưa giá trị sheet 1 ae vào cột lương hold"

    holdRows.forEach((r: any) => {
      let biz = String(r["BU"] || r["Business"] || "")
        .trim()
        .toUpperCase();
      if (biz === "AHN_HP") biz = "AHP";
      const acc = String(
        r["Bank Account Number"] || r["Beneficiary Account No."] || "",
      ).trim();
      const l07 = String(r["L07"] || r["Mã AE"] || r["Mã ae"] || "").trim();
      
      const rowMonthRaw = String(r["Tháng báo cáo"] || r["_fileMonth"] || r["Tháng"] || "").trim();
      const rowMonthLabel = extractMonth(rowMonthRaw);

      const transactionMonthStr = String(
        r["Tháng phát sinh"] ||
          r["tháng phát sinh"] ||
          r["Thang phat sinh"] ||
          r["Trạng thái"] ||
          r["Sheet Source"] ||
          r["Tình trạng thanh toán"] ||
          r["Tháng"] ||
          r["_fileMonth"] ||
          r["Tháng báo cáo"] ||
          "",
      );

      const ttttUpper = String(r["Tình trạng thanh toán"] || "").toUpperCase();
      const noteUpper = String(r["Note"] || "").toUpperCase();
      const notesUpper = String(r["Notes"] || "").toUpperCase();

      const groupMonthStr = String(
        r["Tháng báo cáo"] ||
          r["_fileMonth"] ||
          r["Tháng"] ||
          defaultMonth
      );
      const month = extractMonth(groupMonthStr) || defaultMonth; 

      // Enforce: only allow the selected month (currentPeriod) to appear on the Trial Balance dashboard
      if (month !== currentPeriod) {
        return;
      }

      const displayMonth = extractMonth(transactionMonthStr) || month;

      // Skip future arising months, but allow historical and current entries
      if (getMonthNum(displayMonth) > currentPeriodNum) {
        return;
      }
      
      const monthShortStr = `T${currentPeriodMonthNum}`;
      const monthDashStr = `${currentPeriodMonthNum}/${currentPeriodYearNum}`;
      const isMonthInStr = (s: string) => {
        if (!s) return true;
        const up = s.toUpperCase();
        return up.includes(currentPeriod.toUpperCase()) || 
               up.includes(monthShortStr.toUpperCase()) || 
               up.includes(monthDashStr);
      };

      // If no month label at all, and no mention of the current month in status fields, skip.
      if (!rowMonthLabel && rowMonthRaw && !isMonthInStr(rowMonthRaw)) {
        return;
      }

      let command = String(r["Lệnh"] || "")
        .trim()
        .toUpperCase();

      const isPaidStatus =
        ttttUpper.includes("ĐÃ THANH TOÁN") ||
        ttttUpper.includes("THANH TOÁN") ||
        ttttUpper.includes("ĐÒ TT") ||
        ttttUpper.includes("PAID") ||
        ttttUpper.includes("ĐÒ CHI") ||
        noteUpper.includes("ĐÃ THANH TOÁN") ||
        noteUpper.includes("THANH TOÁN") ||
        noteUpper.includes("ĐÒ TT") ||
        noteUpper.includes("PAID") ||
        noteUpper.includes("ĐÒ CHI") ||
        notesUpper.includes("ĐÃ THANH TOÁN") ||
        notesUpper.includes("THANH TOÁN") ||
        notesUpper.includes("ĐÒ TT") ||
        notesUpper.includes("PAID") ||
        notesUpper.includes("ĐÒ CHI") ||
        command === "OK";

      if (isPaidStatus) {
        command = "-";
      }


      if (!biz || biz === "UNKNOWN") {
        if (l07) {
          const resolved =
            resolveL07BuFromAeCode(l07) ||
            getCenterInfoByAECode(l07) ||
            getCenterInfoByL07(l07);
          if (resolved) {
            biz =
              ("bu" in resolved
                ? resolved.bu
                : "bus" in resolved
                  ? resolved.bus
                  : "") || biz;
          }
        }
      }
      if (!biz || biz === "UNKNOWN") biz = accToBU[acc];
      if (!biz || biz === "UNKNOWN")
        biz = extractBU(String(r["Note"] || r["Sheet Source"] || ""));
      if (!biz || biz === "UNKNOWN") biz = "AHN";
      if (biz === "AHN_HP") biz = "AHP";

      const nv = String(r["Nghiệp vụ"] || "").toLowerCase();
      const st = String(r["Trạng thái"] || "").toLowerCase();
      const ss = String(r["Sheet Source"] || "").toLowerCase();

      // If Trạng thái or Nghiệp vụ or Note contains bonus or thưởng, add directly to the base row's rawThu/thu
      const noteStr = String(r["Note"] || r["Notes"] || "").toLowerCase();
      const isBonusRow =
        nv.includes("bonus") ||
        st.includes("bonus") ||
        ss.includes("bonus") ||
        noteStr.includes("bonus") ||
        nv.includes("thưởng") ||
        st.includes("thưởng") ||
        noteStr.includes("thưởng");
      if (isBonusRow) {
        const tpRaw = parseMoneyToNumber(
          r["TOTAL PAYMENT"] ||
            r["Grand Total"] ||
            r["GRAND TOTAL"] ||
            r["Payment Amount"] ||
            0,
        );
        const baseKey = getBuKey(month, biz);
        if (!buStats[baseKey]) {
          buStats[baseKey] = {
            id: baseKey,
            month,
            reportMonth: month,
            displayMonth: month,
            bu: biz,
            rawThu: 0,
            rawChi: 0,
            thu: 0,
            chi: 0,
            add: 0,
            hold: 0,
            cancel: 0,
            rawAdd: 0,
            rawHold: 0,
            rawCancel: 0,
            bonus: 0,
            rawBonus: 0,
            rawHoldPending: 0,
            ghiChu: "",
            confirmed: false,
            lenh: "",
          };
        }
        buStats[baseKey].rawThu += tpRaw;
        buStats[baseKey].thu = buStats[baseKey].rawThu;
        return;
      }

      let type = "add";

      const nvLower = String(r["Nghiệp vụ"] || r["NGHIỆP VỤ"] || r["Nghiep vu"] || r["nghiepVu"] || r["Loại"] || "").trim().toLowerCase();
      // If Trạng thái or Nghiệp vụ contains bonus
      /* bonus merged */
      if (nvLower === "cancel" || nvLower.includes("cancel") || nvLower === "c" || nv.includes("cancel") || st.includes("cancel") || ss.includes("cancel") || ttttUpper.includes("CANCEL")) {
        type = "cancel";
      }
      // Khong dua gia tri sheet 1 ae vao cot luong hold
      else if (ss.includes("sheet 1 ae")) {
        // Not considered hold
        type = "other";
      }
      // If Trạng thái or Nghiệp vụ contains hold but not add
      else if (
        (nv === "hold" || nv.includes("hold") || st.includes("hold") || ss.includes("hold")) &&
        !(nv.includes("add") || st.includes("add") || ss.includes("add"))
      ) {
        type = "hold";
      }

      // Khoản add phải có tháng báo cáo trùng card chọn tháng
      if (type === "add" && month !== currentPeriod) {
        return;
      }

      // Khoản cancel phải có tháng báo cáo trùng card chọn tháng
      if (type === "cancel" && month !== currentPeriod) {
        return;
      }

      const tpRaw = parseMoneyToNumber(
        r["TOTAL PAYMENT"] ||
          r["Grand Total"] ||
          r["GRAND TOTAL"] ||
          r["Payment Amount"] ||
          0,
      );

      const isCurrentPeriodHold = type === "hold" && isSameMonthForSumIf(displayMonth, month);
      const isPastMonth = getMonthNum(displayMonth) < getMonthNum(month);
      // We don't dim past holds if they are processed in the current file month! They are legitimate transactions of the current month.
      const isDimmedHold = false;

      const groupMonth = displayMonth || month;

      // Accumulate ADD/HOLD/CANCEL directly into distinct summary rows separated by displayMonth
      let mergeType = type;
      if (type === "add" || type === "hold" || type === "cancel") {
        mergeType = "adjustment";
      }
      const key = isCurrentPeriodHold
        ? `${month}_${biz}_${displayMonth}_hold_${month}`
        : `${groupMonth}_${biz}_${displayMonth}_${mergeType}_${month}`;
      
      // Do not merge hold/add/cancel into base row.
      // If the user wants Bonus merged into the base row, we'll handle it below.
      // For now, keep keys distinct so they are processed properly.
      // const key = `${groupMonth}_${biz}_${displayMonth}_${type}_${month}`; // already set

      const formatAdjustmentMonth = (value: string) => {
        const raw = String(value || "").trim();
        const match = raw.match(/(?:Th[aá]ng\s*)?(\d{1,2})[/-]\s*(\d{4})/i);
        if (!match) return raw;
        const m = parseInt(match[1], 10);
        return `${m < 10 ? "0" + m : m}.${match[2]}`;
      };

      let customMonthDisplay = ``;
      const rawNv = String(r["Nghiệp vụ"] || "").trim();
      const rawThangPhatSinh = String(r["Tháng phát sinh"] || "").trim();

      if (rawNv) {
        const capitalizedNv = rawNv.charAt(0).toUpperCase() + rawNv.slice(1).toLowerCase();
        const finalPhatSinh = rawThangPhatSinh || displayMonth || month;
        customMonthDisplay = `${capitalizedNv} lương tháng ${finalPhatSinh}`;
      } else {
        const stDisplay = String(r["Trạng thái"] || "").trim();
        const stDisplayUpper = stDisplay.toUpperCase();

        if (stDisplayUpper.startsWith("HOLD T") || stDisplayUpper === "HOLD") {
          customMonthDisplay = `Hold lương tháng ${displayMonth}`;
        } else if (stDisplayUpper.startsWith("ADD T") || stDisplayUpper === "ADD") {
          customMonthDisplay = `Add lương tháng ${displayMonth}`;
        } else if (stDisplayUpper.startsWith("CANCEL T") || stDisplayUpper === "CANCEL") {
          customMonthDisplay = `Cancel lương tháng ${displayMonth}`;
        } else if (stDisplay) {
          if (stDisplayUpper.includes("HOLD")) {
            customMonthDisplay = `Hold lương tháng ${displayMonth}`;
          } else if (stDisplayUpper.includes("ADD")) {
            customMonthDisplay = `Add lương tháng ${displayMonth}`;
          } else if (stDisplayUpper.includes("CANCEL")) {
            customMonthDisplay = `Cancel lương tháng ${displayMonth}`;
          } else {
            customMonthDisplay = stDisplay;
          }
        } else if (type === "add") {
          customMonthDisplay = `Add lương tháng ${displayMonth}`;
        } else if (type === "hold") {
          customMonthDisplay = `Hold lương tháng ${displayMonth}`;
        } else if (type === "cancel") {
          customMonthDisplay = `Cancel lương tháng ${displayMonth}`;
        }
      }

      if (type === "hold" || type === "add" || type === "cancel" || false) {
        const label =
          type === "hold" ? "Hold" : type === "add" ? "Add" : "Cancel";
        const descriptionMonth = formatAdjustmentMonth(
          rawThangPhatSinh || displayMonth || month,
        );
        customMonthDisplay = `${label} lương tháng ${descriptionMonth}`;
      }

      if (key === getBuKey(month, biz)) {
        customMonthDisplay = ""; // Base row shouldn't have adjustment labels
      }

      if (!buStats[key]) {
        buStats[key] = {
          id: key,
          month: groupMonth,
          reportMonth: month,
          displayMonth,
          bu: biz,
          rawThu: 0,
          rawChi: 0,
          thu: 0,
          chi: 0,
          add: 0,
          hold: 0,
          cancel: 0,
          rawAdd: 0,
          rawHold: 0,
          rawCancel: 0,
          bonus: 0,
          rawBonus: 0,
          rawHoldPending: 0,
          ghiChu: "",
          confirmed: false,
          lenh: command,
          customMonthDisplay,
          rawOpenHold: 0,
          isPaidStatus,
        };
      } else {
        if (command === "OK" || (command === "-" && buStats[key].lenh !== "OK")) {
          buStats[key].lenh = command;
        }
        if (isPaidStatus) {
          buStats[key].isPaidStatus = true;
        }
        if (type === "hold" && buStats[key].customMonthDisplay?.toLowerCase().startsWith("add")) {
          buStats[key].customMonthDisplay = customMonthDisplay;
        }
      }

      if (isDimmedHold) {
        // We will calculate openHold/open balances chronologically
        buStats[key].rawOpenHold = 0;
      }

      if (type === "cancel") {
        buStats[key].rawCancel += tpRaw;
        if (!buStats[key].ghiChu) {
          buStats[key].ghiChu = "cancel";
        }
      } else if (type === "hold") {
        buStats[key].rawHold += tpRaw;
        const tttt = String(
          r["_originalTinhTrangThanhToan"] !== undefined
            ? r["_originalTinhTrangThanhToan"]
            : r["Tình trạng thanh toán"] || "",
        ).toUpperCase().trim();
        if (tttt.includes("PENDING") || tttt === "") {
           buStats[key].rawHoldPending += tpRaw;
        }
      } else {
        buStats[key].rawAdd += tpRaw;
        if (isPastMonth) {
          buStats[key].ghiChu = "Add hold tháng quá khứ";
        }
      }

      if (r["Notes"] || r["Note"]) {
        const notesText = String(r["Notes"] || r["Note"] || "");
        if (notesText) {
          buStats[key].ghiChu =
            isPastMonth && type === "add"
              ? "Add hold tháng quá khứ"
              : notesText;
        }
      }
    });

    // --- USE STANDARD COMPILED STATS ---
    const testBuStats = { ...buStats };

    Object.values(testBuStats).forEach((s) => {
      const isAddOnly = (s.rawAdd !== undefined && s.rawAdd > 0) && (!s.rawHold || s.rawHold === 0) && (!s.rawCancel || s.rawCancel === 0);
      const totalArising = Math.abs(s.rawHold || 0) + Math.abs(s.rawCancel || 0);
      const isAdjustmentType =
        String(s.id).includes("_adjustment_") ||
        String(s.id).includes("_hold") ||
        (s.rawAdd !== undefined && s.rawAdd > 0) ||
        (s.rawHold !== undefined && s.rawHold > 0) ||
        (s.rawCancel !== undefined && s.rawCancel > 0);

      if (isAdjustmentType && (totalArising > 0 || isAddOnly)) {
        const isCurrentMonthHold = s.displayMonth && s.reportMonth && isSameMonthForSumIf(s.displayMonth, s.reportMonth) && (s.rawHold || 0) > 0 && !s.rawAdd && !s.rawCancel;
        if (!isAddOnly && !isCurrentMonthHold) {
          s.openHold = Math.max(s.rawOpenHold || 0, totalArising);
          s.rawOpenHold = s.openHold;
        } else {
          s.openHold = 0;
          s.rawOpenHold = 0;
        }
        s.thu = 0;
        s.chi = isCurrentMonthHold ? Math.abs(s.rawHold || 0) : 0;
        s.add = Math.abs(s.rawAdd || 0);
        s.hold = Math.abs(s.rawHold || 0);
        s.cancel = Math.abs(s.rawCancel || 0);
      } else {
        if (s.rawCancel !== undefined) s.cancel = s.rawCancel;
        if (s.rawHold !== undefined) s.hold = s.rawHold;
        if (s.rawAdd !== undefined) s.add = s.rawAdd;
        if (s.rawBonus !== undefined) s.bonus = isNaN(s.rawBonus) ? 0 : s.rawBonus;
        if (s.rawOpenHold !== undefined) s.openHold = s.rawOpenHold;
      }
    });

    // Make sure we have currentPeriod slots for all unique BUs so they can receive carried-forward amounts
    const allBUs = new Set(Object.values(testBuStats).map((b) => b.bu));
    allBUs.forEach((biz) => {
      const currentKey = getBuKey(currentPeriod, biz);
      if (!testBuStats[currentKey]) {
        testBuStats[currentKey] = {
          id: currentKey,
          month: currentPeriod,
          reportMonth: currentPeriod,
          displayMonth: currentPeriod,
          bu: biz,
          rawThu: 0,
          rawChi: 0,
          thu: 0,
          chi: 0,
          add: 0,
          hold: 0,
          cancel: 0,
          rawAdd: 0,
          rawHold: 0,
          rawCancel: 0,
          bonus: 0,
          rawBonus: 0,
          rawHoldPending: 0,
          ghiChu: "",
          confirmed: false,
          openHold: 0,
        };
      }
    });

    const baseRows = Object.values(testBuStats);
    const adjustmentRows: typeof baseRows = [];
    const openingHoldRows: typeof baseRows = [];

    const formatOpeningHoldMonth = (value: string) => {
      const raw = String(value || "").trim();
      const match = raw.match(/(?:Th[aá]ng\s*)?(\d{1,2})[/-]\s*(\d{4})/i);
      if (!match) return raw;
      const month = match[1].padStart(2, "0");
      return `${month}.${match[2]}`;
    };

    const analytics = buildBulkPaymentAnalytics({
      sheet1Rows: appData.Sheet1_AE?.data || [],
      holdRows: appData.Hold_AE?.data || [],
      bankRows: appData.Bank_North_AE?.data || [],
      globalMonth: currentPeriodVal,
    });

    const finalOpenBal: Record<string, Record<string, number>> = {};
    
    analytics.summaryRows.forEach(row => {
      const bu = row["BU"];
      const thangHold = row["Tháng HOLD"];
      const openingBal = row["Số dư HOLD đầu kỳ"];
      
      if (openingBal > 0) {
        if (!finalOpenBal[bu]) finalOpenBal[bu] = {};
        finalOpenBal[bu][thangHold] = openingBal;
      }
    });

    // 3. Merge Opening Balance with ADD rows if they share the same BU and the same arising month (Tháng phát sinh)
    // Khoản nào cùng tháng phát sinh cùng BU thì cho cùng 1 dòng nếu có số dư đầu kỳ, phát sinh add
    Object.entries(finalOpenBal).forEach(([bu, months]) => {
      Object.entries(months).forEach(([displayMonth, amount]) => {
        if (amount <= 0) return;

        // Find if there is an adjustment row in testBuStats for this BU and matching displayMonth / tháng phát sinh
        const matchingAddRow = Object.values(testBuStats).find((r) => {
          if (r.bu !== bu) return false;
          const isAdjustmentType =
            String(r.id).includes("_adjustment_") ||
            (r.rawAdd && r.rawAdd > 0) ||
            (r.add && r.add > 0) ||
            (r.rawHold && r.rawHold > 0) ||
            (r.hold && r.hold > 0) ||
            (r.rawCancel && r.rawCancel > 0) ||
            (r.cancel && r.cancel > 0);
          if (!isAdjustmentType) return false;

          const rMonth = r.displayMonth || r.month;
          return isSameMonthForSumIf(rMonth, displayMonth);
        });

        if (matchingAddRow) {
          // Merge Opening Balance into the row: cùng 1 dòng
          const totalArising = Math.abs(matchingAddRow.rawHold || 0) + Math.abs(matchingAddRow.rawAdd || 0) + Math.abs(matchingAddRow.rawCancel || 0);
          matchingAddRow.openHold = Math.max(amount, totalArising);
          matchingAddRow.rawOpenHold = matchingAddRow.openHold;
          matchingAddRow._isOpeningHold = true;
          const descriptionMonth = formatOpeningHoldMonth(displayMonth);
          if (!matchingAddRow.customMonthDisplay) {
            matchingAddRow.customMonthDisplay = `Hold lương tháng ${descriptionMonth}`;
          }
        } else {
          // No matching ADD row: create separate Opening Hold row
          const descriptionMonth = formatOpeningHoldMonth(displayMonth);
          openingHoldRows.push({
            id: `${currentPeriod}_${bu}_opening_hold_${displayMonth}`,
            month: currentPeriod,
            reportMonth: currentPeriod,
            displayMonth,
            bu,
            rawThu: 0,
            rawChi: 0,
            thu: 0,
            chi: 0,
            add: 0,
            hold: 0,
            cancel: 0,
            rawAdd: 0,
            rawHold: 0,
            rawCancel: 0,
            bonus: 0,
            rawBonus: 0,
            rawHoldPending: 0,
            ghiChu: "Hold tháng trong quá khứ",
            confirmed: true,
            customMonthDisplay: `Hold lương tháng ${descriptionMonth}`,
            openHold: amount,
            rawOpenHold: amount,
            _isOpeningHold: true,
            lenh: "OK",
          });
        }
      });
    });

    // Preserve the opening detail allocation while approving all movements.
    baseRows.forEach((row) => {
      const totalArising = Math.abs(row.rawHold || 0) + Math.abs(row.rawAdd || 0) + Math.abs(row.rawCancel || 0);
      if (totalArising > 0 && getMonthNum(row.displayMonth || row.month) < getMonthNum(row.reportMonth || row.month)) {
        row.openHold = Math.max(row.openHold || 0, totalArising);
        row.rawOpenHold = row.openHold;
      }
    });

    let processedResult = [...baseRows, ...openingHoldRows, ...adjustmentRows].filter(r => {
      if (r._excludeFromTotals) {
        const repNum = getMonthNum(r.reportMonth || r.month || "");
        const dNum = getMonthNum(r.displayMonth || r.month || "");
        if (repNum > currentPeriodNum && dNum > currentPeriodNum) return false;
      }
      return true;
    });

    // Restore current monthly snapshots before applying automatic approval.
    const savedPeriods = appData.SavedPeriods_HoldAdd || {};
    const savedRowsMap = appData.SavedRows_HoldAdd || {};
    const savedRowsMeta = appData.SavedRows_HoldAdd_Meta || {};
    const transactionVersions = appData.TrialBalanceTransactionVersions || {};
    const finalRows: typeof processedResult = [];

    // Group the raw computed result rows by reporting month to align snapshot and live data
    const resultMap = new Map<string, typeof processedResult>();
    processedResult.forEach((r) => {
      const repM = r.reportMonth || r.month;
      if (!resultMap.has(repM)) {
        resultMap.set(repM, []);
      }
      resultMap.get(repM)!.push(r);
    });

    const allMonths = new Set([
      ...resultMap.keys(),
      ...Object.keys(savedRowsMap),
    ]);

    allMonths.forEach((m) => {
      const isSaved = !!getSavedDataForPeriod(savedPeriods, m);
      const live = resultMap.get(m) || [];
      const snapshot = getSavedDataForPeriod(savedRowsMap, m);
      const snapshotMeta = getSavedDataForPeriod(savedRowsMeta, m);
      const snapshotVersion = snapshotMeta?.transactionVersion || 0;
      const requiredTransactionVersion =
        getSavedDataForPeriod(transactionVersions, m) || 0;
      const snapshotIsCurrent =
        snapshotVersion >= requiredTransactionVersion;

      if (isSaved && Array.isArray(snapshot) && snapshotIsCurrent) {
        // A saved Trial Balance is a complete monthly snapshot. Restore every
        // standard and adjustment row exactly as saved so changing months does
        // not discard ordinary transactions. A newer Transaction save bumps
        // the period version, making this snapshot stale and selecting `live`.
        finalRows.push(...snapshot.map((row: any) => ({ ...row })));
      } else {
        finalRows.push(...live);
      }
    });

    // CANCEL is a movement of one report period, not a carried balance item.
    // Snapshots from older periods may still be needed for HOLD opening balances,
    // retain CANCEL history for carry calculation; grouped hides it from later views.
    processedResult = finalRows.map(autoApproveTrialBalanceRow);

    return processedResult.sort((a, b) => {
      const mA = a.reportMonth === currentPeriod ? 99999999 : getMonthNum(a.month);
      const mB = b.reportMonth === currentPeriod ? 99999999 : getMonthNum(b.month);
      if (mA !== mB) return mB - mA; // newest first

      if (a.bu !== b.bu) return a.bu.localeCompare(b.bu); // Group by BU

      const dMa = getMonthNum(a.displayMonth || a.month);
      const dMb = getMonthNum(b.displayMonth || b.month);
      if (dMa !== dMb) return dMa - dMb; // source month A -> B

      const typeWeight = (id: string) => {
        const sid = String(id);
        if (!sid.includes("_adjustment_") && !sid.includes("_hold") && !sid.includes("_add") && !sid.includes("_cancel") && !sid.includes("_past_")) return 0;
        if (sid.includes("_hold") || sid.includes("_adjustment_")) return 1;
        if (sid.includes("_add")) return 2;
        if (sid.includes("_cancel")) return 3;
        return 4;
      };

      const wA = typeWeight(String(a.id));
      const wB = typeWeight(String(b.id));
      if (wA !== wB) return wA - wB; // standard -> hold -> add -> cancel

      return 0;
    });
  }, [
    appData.Sheet1_AE?.data,
    appData.Hold_AE?.data,
    appData.Bank_North_AE?.data,
    appData.SavedPeriods_HoldAdd,
    appData.SavedRows_HoldAdd,
    appData.SavedRows_HoldAdd_Meta,
    appData.TrialBalanceTransactionVersions,
    currentPeriod,
    currentPeriodVal,
    currentPeriodNum,
    getMonthNum,
    extractMonth,
    getSavedDataForPeriod,
    currentPeriodMonthNum,
    currentPeriodYearNum,
  ]);

  const handleSaveBalances = () => {
    updateAppData((prev: any) => {
      const nextMonthStr = getNextMonthStr(currentPeriod);
      // Giữ lại state cũ
      const nextSavedAll = prev.SavedBal_PayrollTrial
        ? JSON.parse(JSON.stringify(prev.SavedBal_PayrollTrial))
        : {};

      const currentRows = data.filter((d) => d.reportMonth === currentPeriod);
      const buMap: Record<
        string,
        {
          thu: number;
          chi: number;
          chiForSDCK: number;
          hold: number;
          openHold: number;
          add: number;
          cancel: number;
        }
      > = {};
      currentRows.forEach((r) => {
        if (!buMap[r.bu]) {
          buMap[r.bu] = {
            thu: 0,
            chi: 0,
            chiForSDCK: 0,
            hold: 0,
            openHold: 0,
            add: 0,
            cancel: 0,
          };
        }
        const isAdjustment = String(r.id).includes("_adjustment_");
        const isHold = isAdjustment || String(r.id).includes("_hold");
        const isPastHold = isHold && getMonthNum(r.displayMonth || r.month) < getMonthNum(currentPeriod);
        const isCancel = String(r.id).includes("_cancel");
        buMap[r.bu].thu += !r._excludeFromTotals ? r.thu : 0;
        buMap[r.bu].chi += !r._excludeFromTotals ? r.chi : 0;
        let rowChiForSDCK = !r._excludeFromTotals && !isPastHold && !(isCancel && !isAdjustment) ? r.chi : 0;
        if (isHold && !isPastHold) {
          rowChiForSDCK = Math.abs(r.rawHoldPending || r.chi || r.hold || 0);
        }
        buMap[r.bu].chiForSDCK += rowChiForSDCK;
        buMap[r.bu].hold += !r._excludeFromTotals ? r.hold : 0;
        buMap[r.bu].openHold += !r._excludeFromTotals ? r.openHold || 0 : 0;
        buMap[r.bu].add += !r._excludeFromTotals ? r.add : 0;
        buMap[r.bu].cancel += !r._excludeFromTotals ? r.cancel : 0;
      });

      // Tạo object lưu trữ cho tháng N+1
      const nextMonthData: Record<string, any> =
        nextSavedAll[nextMonthStr] || {};

      Object.keys(buMap).forEach((bu) => {
        const openingBalByMonth = buBalancesByMonth[currentPeriod]?.[bu]?.openBalByMonth || {};
        const nextOpenBalByMonth = nextTrialBalanceCarry(
          openingBalByMonth, currentRows.filter(r => r.bu === bu), currentPeriod,
        );

        const totalHoldToTransfer = Object.values(nextOpenBalByMonth).reduce((s, v) => s + v, 0);

        nextMonthData[bu] = {
          openBal: totalHoldToTransfer,
          openBalByMonth: nextOpenBalByMonth,
          holdAmt: totalHoldToTransfer,
        };

        // Self-sanity check: Clear faulty legacy own-month carryover in the underlying saved states
        if (nextSavedAll[currentPeriod] && nextSavedAll[currentPeriod][bu]) {
          const currentSavedOpenBal =
            nextSavedAll[currentPeriod][bu].openBal || 0;
          if (
            Math.round(currentSavedOpenBal) === Math.round(totalHoldToTransfer)
          ) {
            delete nextSavedAll[currentPeriod][bu];
            if (Object.keys(nextSavedAll[currentPeriod]).length === 0) {
              delete nextSavedAll[currentPeriod];
            }
          }
        }
      });

      nextSavedAll[nextMonthStr] = nextMonthData;

      const nextSavedPeriods = prev.SavedPeriods_HoldAdd
        ? JSON.parse(JSON.stringify(prev.SavedPeriods_HoldAdd))
        : {};
      nextSavedPeriods[currentPeriod] = true;

      // Snapshot formulation for currentPeriod
      // Just save rows as they are currently rendered!
      const snapshotRows = currentRows.map((r) => {
        return { ...r };
      });

      const nextSavedRows = prev.SavedRows_HoldAdd
        ? JSON.parse(JSON.stringify(prev.SavedRows_HoldAdd))
        : {};
      nextSavedRows[currentPeriod] = snapshotRows;

      const nextSavedRowsMeta = prev.SavedRows_HoldAdd_Meta
        ? JSON.parse(JSON.stringify(prev.SavedRows_HoldAdd_Meta))
        : {};
      nextSavedRowsMeta[currentPeriod] = {
        transactionVersion:
          prev.TrialBalanceTransactionVersions?.[currentPeriod] || 0,
        savedAt: new Date().toISOString(),
      };

      return {
        ...prev,
        SavedBal_PayrollTrial: nextSavedAll,
        SavedPeriods_HoldAdd: nextSavedPeriods,
        SavedRows_HoldAdd: nextSavedRows,
        SavedRows_HoldAdd_Meta: nextSavedRowsMeta,
      };
    });
    toast.success("Đã lưu dữ liệu: Chuyển Lương sang SDĐK kỳ sau!");
  };

  // Group by Report Month - Show ALL rows grouped by their reportMonth
  const grouped = useMemo(() => {
    const map = new Map<string, BuRow[]>();
    for (const row of data) {
      if (String(row.id).includes("_cancel") && !isExactReportPeriod(row, currentPeriod)) continue;
      // Use reportMonth if set, otherwise use current period
      const rowReportMonth = row.reportMonth || currentPeriod;
      
      if (rowReportMonth) {
        if (!map.has(rowReportMonth)) map.set(rowReportMonth, []);
        map.get(rowReportMonth)!.push(row);
      }
    }
    return map;
  }, [data, currentPeriod]);
  const monthKeys = useMemo(() => {
    return [...grouped.keys()].sort((a, b) => getMonthNum(b) - getMonthNum(a));
  }, [grouped, getMonthNum]);
  const grossPayByMonth = useMemo(() => {
    const totals = grossPayTotalsByMonth(appData.Sheet1_AE.data);
    for (const [month, saved] of Object.entries(appData.SavedGrossPayByMonth || {})) {
      if (totals.has(month)) continue;
      const savedTotals = grossPayTotalsByMonth(saved.data, month).get(month);
      if (savedTotals) totals.set(month, savedTotals);
    }
    return totals;
  }, [appData.Sheet1_AE.data, appData.SavedGrossPayByMonth]);

  const allOpen =
    monthKeys.length > 0 && monthKeys.every((k) => expanded.has(k));
  const toggleAll = () => setExpanded(allOpen ? new Set() : new Set(monthKeys));
  const toggle = (mk: string) =>
    setExpanded((prev) => {
      const s = new Set(prev);
      if (s.has(mk)) {
        s.delete(mk);
      } else {
        s.add(mk);
      }
      return s;
    });

  const buBalancesByMonth = useMemo(() => {
    const balances: Record<
      string,
      Record<string, { openBal: number; openBalByMonth: Record<string, number>; closeBal: number }>
    > = {};

    const allMonths = [...new Set(data.map(row => row.reportMonth || row.month))]
      .sort((a, b) => getMonthNum(a) - getMonthNum(b));
    
    allMonths.forEach((mk) => {
      balances[mk] = {};
      const rows = data.filter(row => (row.reportMonth || row.month) === mk);
      const uniqueBUsInGroup = Array.from(new Set(rows.map((e) => e.bu)));
      const sortedIdx = allMonths.indexOf(mk);

      uniqueBUsInGroup.forEach((bu) => {
        const savedData = appData.SavedBal_PayrollTrial?.[mk]?.[bu];
        const savedVal = savedData?.openBal || 0;
        const savedOpenBalByMonth = savedData?.openBalByMonth;

        let openBal = 0;
        let openBalByMonth: Record<string, number> = {};

        if (savedData) {
          openBal = Math.abs(savedVal);
          if (savedOpenBalByMonth) {
            openBalByMonth = { ...savedOpenBalByMonth };
          } else {
            openBalByMonth = { "general": openBal };
          }
        } else if (sortedIdx > 0) {
          const prevMk = allMonths[sortedIdx - 1];
          const prevBalInfo = balances[prevMk]?.[bu];
          
          if (prevBalInfo) {
            openBalByMonth = { ...(prevBalInfo.openBalByMonth || {}) };
            const prevRows = data.filter(row => (row.reportMonth || row.month) === prevMk);
            const prevRowsForBu = prevRows.filter((r) => r.bu === bu);
            
            openBalByMonth = nextTrialBalanceCarry(openBalByMonth, prevRowsForBu, prevMk);

            openBal = Object.values(openBalByMonth).reduce((s, v) => s + v, 0);
          }
        }

        const rowsForBu = rows.filter((r) => r.bu === bu);
        rowsForBu.forEach((r) => {
          const isAddRow = ((r.rawAdd || 0) > 0 || (r.add || 0) > 0) && (r.rawHold || 0) === 0 && (r.rawCancel || 0) === 0 && !r._isOpeningHold;
          if (!isAddRow && r.openHold && r.openHold > 0 && getMonthNum(r.displayMonth || r.month) < getMonthNum(mk)) {
            const dMonth = r.displayMonth || r.month;
            openBalByMonth[dMonth] = Math.max(openBalByMonth[dMonth] || 0, r.openHold);
          }
        });
        openBal = Math.max(openBal, Object.values(openBalByMonth).reduce((s, v) => s + v, 0));

        const buThu = rowsForBu.reduce((s, r) => s + (r.thu || 0), 0);

        // SDCK = Lương TA của tháng - Số dư ĐK - Lương Hold của tháng
        // IMPORTANT: Khoản HOLD/CANCEL có tháng phát sinh < tháng báo cáo (quá khứ) thì không trừ vào SDCK,
        // nghĩa là chi=0 khi tính closeBal. Lương TA (thu) thì luôn luôn của kỳ này.
        // Tương tự, khoản hold của các tháng trước (holdRows) đã được xử lý ở đầu và có _isPastHoldApprove=true
        // khi nó được load lên để approved ở tháng này. Nhưng ở đây, chỉ là `buChi`.
        let adjustedChi = 0;
        rowsForBu.forEach((r) => {
           if (r._isPastHoldApprove) {
             // Do nothing - do not subtract it from closeBal
           } else {
             adjustedChi += (r.chi || 0);
           }
        });

        // SDCK = Lương TA của tháng - Lương Hold của tháng (THÁNG PHÁT SINH = THÁNG BÁO CÁO)
        const closeBal = buThu - adjustedChi;

        balances[mk][bu] = { openBal, openBalByMonth, closeBal };
      });
    });

    return balances;
  }, [data, appData.SavedBal_PayrollTrial, getMonthNum]);

  // The displayed remainder uses exactly the carry calculation used by Save.
  const remainingHoldByMonth = useMemo(() => {
    const result: Record<string, { total: number; byBu: Record<string, number> }> = {};
    for (const mk of monthKeys) {
      const byBu: Record<string, number> = {};
      for (const [bu, balance] of Object.entries(buBalancesByMonth[mk] || {})) {
        const carry = nextTrialBalanceCarry(
          balance.openBalByMonth,
          data.filter(row => row.bu === bu && (row.reportMonth || row.month) === mk),
          mk,
        );
        byBu[bu] = Object.values(carry).reduce((sum, amount) => sum + amount, 0);
      }
      result[mk] = { byBu, total: Object.values(byBu).reduce((sum, amount) => sum + amount, 0) };
    }
    return result;
  }, [monthKeys, buBalancesByMonth, data]);

  const computedMonthTotals = useMemo(() => {
    const totals: Record<
      string,
      {
        openBal: number;
        psThu: number;
        psChi: number;
        closeBal: number;
        addAmt: number;
        holdAmt: number;
        cancelAmt: number;
        bonusAmt: number;
      }
    > = {};

    monthKeys.forEach((mk) => {
      const rows = grouped.get(mk) || [];
      const uniqueBUsInGroup = Array.from(new Set(rows.map((e) => e.bu)));

      let openBal = 0;
      uniqueBUsInGroup.forEach((bu) => {
        openBal += buBalancesByMonth[mk]?.[bu]?.openBal || 0;
      });

      // Only sum rows that belong to this period month
      const rowsThisMonth = rows.filter((row) =>
        isExactReportPeriod(row, mk),
      );

      const psThu = rowsThisMonth
        .filter((e) => !e._excludeFromTotals)
        .reduce((s, e) => s + (e.thu || 0), 0);

      const psChi = rowsThisMonth
        .filter((e) => !e._excludeFromTotals)
        .reduce((s, e) => s + Math.abs(e.chi || 0), 0);

      const addAmt = rowsThisMonth
        .filter((e) => !e._excludeFromTotals)
        .reduce((s, e) => s + Math.abs(e.add || 0), 0);

      const holdAmt = rowsThisMonth
        .filter((e) => !e._excludeFromTotals)
        .reduce((s, e) => s + Math.abs(e.hold || 0), 0);

      const cancelAmt = rowsThisMonth
        .filter((row) => isExactReportPeriod(row, mk) && !row._excludeFromTotals)
        .reduce((s, e) => s + Math.abs(e.cancel || 0), 0);

      const bonusAmt = rowsThisMonth
        .filter((e) => !e._excludeFromTotals)
        .reduce((s, e) => s + Math.abs(isNaN(e.bonus) ? 0 : e.bonus), 0);

      // SDCK = Lương TA của tháng - Lương Hold của tháng (THÁNG PHÁT SINH = THÁNG BÁO CÁO)
      let adjustedPsChi = 0;
      rowsThisMonth.filter((e) => !e._excludeFromTotals).forEach(r => {
         if (r._isPastHoldApprove) {
           // Do not subtract from SDCK
         } else {
           adjustedPsChi += (r.chi || 0);
         }
      });
      const closeBal = psThu - adjustedPsChi;

      totals[mk] = {
        openBal: Math.abs(openBal),
        psThu,
        psChi,
        closeBal,
        addAmt,
        holdAmt,
        cancelAmt,
        bonusAmt,
      };
    });

    return totals;
  }, [grouped, monthKeys, buBalancesByMonth]);

  const {
    grandOpenBal,
    grandThu,
    grandChi,
    grandBal,
    filteredData,
    rowRCloseBalances,
    monthCloseBalances,
  } = useMemo(() => {
    let open = 0;
    let thu = 0;
    let chi = 0;
    let add = 0;
    let hold = 0;
    let cancel = 0;

    monthKeys.forEach((mk) => {
      const t = computedMonthTotals[mk];
      if (t) {
        open += t.openBal;
        thu += t.psThu;
        chi += t.psChi;
        add += t.addAmt;
        hold += t.holdAmt;
        cancel += t.cancelAmt;
      }
    });

    const rowRCloseBalances: Record<string, number> = {};
    const buCloseBalances: Record<string, number> = {};
    const monthCloseBalances: Record<string, number> = {};

    monthKeys.forEach((mk) => {
      let monthCloseBalSum = 0;
      const rows = grouped.get(mk) || [];
      const uniqueBUs = Array.from(new Set(rows.map((r) => r.bu)));

      uniqueBUs.forEach((bu) => {
        let buCloseBalSum = 0;
        const buRows = rows.filter((r) => r.bu === bu);

        buRows.forEach((e) => {
          const isRowDimmed = !!e._dimmed && isPeriodSaved(e.month);
          const displayedThu = isRowDimmed ? 0 : (e.thu || 0);
          const displayedChi = isRowDimmed ? 0 : Math.abs(e.chi || 0);

          let adjustedDisplayedChi = displayedChi;
          if (e._isPastHoldApprove) {
            adjustedDisplayedChi = 0; // Exclude past holds/cancels from SDCK calculation
          }

          // SDCK = Lương TA của tháng - Lương Hold của tháng (standalone per row, KHÔNG CỘNG DÒNG BÊN TRÊN NỮA)
          const rowCloseBal = displayedThu - adjustedDisplayedChi;
          rowRCloseBalances[e.id] = rowCloseBal;

          buCloseBalSum += rowCloseBal;
        });

        buCloseBalances[`${mk}_${bu}`] = buCloseBalSum;
        monthCloseBalSum += buCloseBalSum;
      });
      monthCloseBalances[mk] = monthCloseBalSum;
    });

    const filtered = data.filter((r) => !r._excludeFromTotals && (!String(r.id).includes("_cancel") || isExactReportPeriod(r, currentPeriod)));

    let globalAdjustedChi = 0;
    filtered.forEach(r => {
      if (r._isPastHoldApprove) {
         // Excluded
      } else {
         globalAdjustedChi += Math.abs(r.chi || 0);
      }
    });

    const grandBal = thu - globalAdjustedChi;

    return {
      grandOpenBal: open,
      grandThu: thu,
      grandChi: chi,
      grandAdd: add,
      grandHold: hold,
      grandCancel: cancel,
      grandBal,
      filteredData: filtered,
      rowRCloseBalances,
      buCloseBalances,
      monthCloseBalances,
    };
  }, [monthKeys, computedMonthTotals, data, grouped, currentPeriod, isPeriodSaved]);

  const normalizeMonthLabel = useCallback(
    (value?: string) => {
      if (!value) return "";
      const extracted = extractMonth(value);
      if (extracted) return extracted;
      return String(value).trim();
    },
    [extractMonth],
  );

  const isCurrentPeriodRow = useCallback(
    (row: any) => {
      const reportMonth = row?.reportMonth || row?.month;
      return (
        normalizeMonthLabel(String(reportMonth || "")) === currentPeriod
      );
    },
    [currentPeriod, normalizeMonthLabel],
  );

  const currentPeriodRows = useMemo(() => {
    return data.filter(isCurrentPeriodRow);
  }, [data, isCurrentPeriodRow]);

  const countBusinesses = useCallback((rows: BuRow[]) => {
    return new Set(
      rows
        .filter((row) => !row._excludeFromTotals && row.bu)
        .map((row) => row.bu),
    ).size;
  }, []);

  const trialBalanceHeaderTotals = useMemo(
    () => calculateTrialBalanceHeaderTotals(currentPeriodRows, currentPeriod),
    [currentPeriodRows, currentPeriod],
  );
  const chiPhiLuongTaPillValue = trialBalanceHeaderTotals.payrollCost;
  const holdPillValue = trialBalanceHeaderTotals.hold;
  const grandAddPillValue = trialBalanceHeaderTotals.add;
  const cancelPillValue = trialBalanceHeaderTotals.cancel;

  const handleExportExcel = useCallback(() => {
    const exportRows = monthKeys.flatMap(mk => {
      const rows = grouped.get(mk) || [];
      return [...new Set(rows.map(row => row.bu))].flatMap((bu, buIndex) => {
        let detailIndex = 0;
        const buRowsForExport = rows.filter(row => row.bu === bu && !hideInactivePastHold(row, mk))
          .sort((a, b) => trialBalanceRowOrder(a) - trialBalanceRowOrder(b));
        const hasDetailOpenHold = buRowsForExport.some(r => !!r.customMonthDisplay && (r.openHold || 0) > 0);
        const hasDetailHoldOrOpen = buRowsForExport.some(r => !!r.customMonthDisplay);

        return buRowsForExport.map(row => {
          const isDetail = !!row.customMonthDisplay;
          const buBalInfo = buBalancesByMonth[mk]?.[bu];
          const rowOpenBal = isDetail ? (row.openHold || 0) : (hasDetailOpenHold ? 0 : (buBalInfo?.openBal || 0));

          const origin = trialBalancePeriod(row.displayMonth || row.month);
          const report = trialBalancePeriod(row.reportMonth || row.month || mk);
          const isCancel =
            /cancel/i.test(row.customMonthDisplay || "") ||
            String(row.id).includes("_cancel") ||
            /cancel/i.test(row.ghiChu || "") ||
            !!(row.rawCancel || row.cancel);
          const isCurrentHold =
            !isCancel &&
            origin === report &&
            (!!(row.rawHold || row.hold) || /_hold\b/i.test(row.id) || /\bhold\b/i.test(row.customMonthDisplay || ""));
          const rowRemainingHold = isCancel
            ? 0
            : isDetail
              ? (isCurrentHold
                  ? (Math.abs(row.chi || 0) || row.rawHold || row.hold || 0)
                  : Math.max(0, (row.openHold || 0) - (row.thu || 0) - (row.rawCancel || row.cancel || 0)))
              : (hasDetailHoldOrOpen ? 0 : (remainingHoldByMonth[mk]?.byBu[bu] || 0));

          return {
            "#": isDetail ? `${buIndex + 1}.${++detailIndex}` : String(buIndex + 1),
            "Ngày / Tháng": trialBalanceRowLabel(row),
            "Business": bu,
            "Số dư Hold ĐK": rowOpenBal,
            "Lương TA của tháng": row.thu || 0,
            "Lương Hold của tháng": Math.abs(row.chi || 0),
            "Tổng PS tại kỳ": isDetail ? (row.thu || 0) : (rowRCloseBalances[row.id] ?? 0),
            "Số dư Hold còn lại": rowRemainingHold,
            "Note": row.ghiChu || "",
          };
        });
      });
    });

    void downloadHierarchicalWorkbook({
      title: `BALANCE · ${currentPeriodVal}`,
      fileName: `Trial_Balance_${new Date().toISOString().split("T")[0]}.xlsx`,
      pages: [
        {
          title: "Trial Balance",
          children: [
            {
              title: "Payroll Trial Balance",
              sheetName: "Trial Balance",
              table: {
                rows: exportRows,
                cards: [
                  { label: "Period", value: currentPeriodVal },
                  {
                    label: "Payroll Cost",
                    value: chiPhiLuongTaPillValue,
                  },
                  { label: "HOLD", value: holdPillValue },
                  { label: "ADD", value: grandAddPillValue },
                  { label: "CANCEL", value: cancelPillValue },
                  { label: "Business Units", value: countBusinesses(currentPeriodRows) },
                ],
              },
            },
          ],
        },
      ],
    })
      .then(() =>
        toast.success("Đã xuất Trial Balance kèm toàn bộ dữ liệu card."),
      )
      .catch((error) => {
        console.error("Trial Balance workbook export failed:", error);
        toast.error("Không thể xuất workbook Trial Balance.");
      });
  }, [
    monthKeys,
    grouped,
    buBalancesByMonth,
    remainingHoldByMonth,
    cancelPillValue,
    chiPhiLuongTaPillValue,
    countBusinesses,
    currentPeriodRows,
    currentPeriodVal,
    grandAddPillValue,
    holdPillValue,
    rowRCloseBalances,
  ]);

  return (
    <div 
      className="trial-balance-frame unified-table-frame table-container h-full flex-1 flex flex-col min-h-0 overflow-hidden bg-card w-full shadow-xs" 
      style={{ 
        borderRadius: "0px",
        border: "1px solid #dfd0d6",
        padding: "0px",
        gap: "0px" 
      }}
    >
      {/* Toolbar */}
      <div
        className="trial-balance-header unified-table-frame-header flex-shrink-0 py-0 flex items-center justify-between bg-[var(--table-header-bg,#FAF3E8)] border-b border-[#dfd0d6]"
        style={{ minHeight: "56px", paddingTop: "0px", paddingBottom: "0px", paddingLeft: "0px", paddingRight: "0px", backgroundColor: "var(--table-header-bg, #FAF3E8)", borderBottom: "1px solid #dfd0d6", boxSizing: "border-box" }}
      >
        <div className="trial-balance-header-content w-full flex items-center justify-between flex-wrap py-2 h-full" style={{ paddingLeft: "12px", paddingRight: "12px", gap: "24px", boxSizing: "border-box" }}>
          {/* Summary Pills on Top Header Bar (Replacing Payroll Hub title & icon) */}
          <div id="trial-balance-summary" className="flex items-baseline gap-4 flex-wrap" style={{ minHeight: "38px", paddingLeft: "0px", paddingRight: "0px" }}>
            <span
              id="trial-balance-title"
              className="trial-balance-title font-bold text-foreground font-sans mr-0 inline-flex items-baseline gap-1"
              style={{ height: "30.3px", marginBottom: "0px" }}
            >
              <TableInitialMark
                label="TRIAL BALANCE"
                className="shrink-0 text-primary self-center"
                style={{ fontSize: "20px", height: "30px", lineHeight: "30px" }}
                glyphStyle={{ fontSize: "20px", height: "30px", width: "23.8281px" }}
              />
              <TableTitleRemainder
                label="TRIAL BALANCE"
                className="app-table-title-remainder--expanded"
                style={{ fontSize: "18px", lineHeight: "21.3px", marginTop: "0px", marginBottom: "0px", paddingTop: "0px", paddingBottom: "0px" }}
              />
            </span>
            <span className="text-[14px] bg-transparent border-0 rounded-none px-1 py-0.5 text-foreground inline-flex items-baseline gap-1.5 shadow-none" style={{ lineHeight: "19px" }}>
              <span className="badge-chiphi text-[#475569] dark:text-slate-400 font-bold uppercase tracking-wider text-[11px] font-sans self-center" style={{ borderWidth: "0px", borderRadius: "0px", color: "#475569" }}>
                CHI PHÍ LƯƠNG TA
              </span>
              <span
                className="font-sans font-bold text-[#1d4ed8] dark:text-blue-400 text-[14px]"
              >
                {fmt(chiPhiLuongTaPillValue)}
              </span>
            </span>
            <span className="text-[14px] bg-transparent border-0 rounded-none px-1 py-0.5 text-foreground inline-flex items-baseline gap-1.5 shadow-none">
              <span className="text-muted-foreground font-bold uppercase tracking-wider text-[11px] font-sans">HOLD</span>
              <span className="font-sans font-bold text-rose-600 dark:text-rose-400 text-[14px]">
                {fmt(holdPillValue)}
              </span>
            </span>
            <span className="text-[14px] bg-transparent border-0 rounded-none px-1 py-0.5 text-foreground inline-flex items-baseline gap-1.5 shadow-none">
              <span className="text-muted-foreground font-bold uppercase tracking-wider text-[11px] font-sans">ADD</span>
              <span
                className="font-sans font-bold text-[#68182e] dark:text-rose-300 text-[14px]"
              >
                {fmt(grandAddPillValue)}
              </span>
            </span>
            <span className="text-[14px] bg-transparent border-0 rounded-none px-1 py-0.5 text-foreground inline-flex items-baseline gap-1.5 shadow-none">
              <span className="text-muted-foreground font-bold uppercase tracking-wider text-[11px] font-sans">CANCEL</span>
              <span
                className="font-sans font-bold text-[#e65100] dark:text-amber-400 text-[14px]"
              >
                {fmt(cancelPillValue)}
              </span>
            </span>

            {!isOkToDisplayValues && (
              <div className="group relative mt-0 ml-1">
                <button
                  onClick={() => {
                    localStorage.setItem("master_ae_active_tab", "Hold_AE");
                    navigate("/master-ae");
                  }}
                  className="flex items-center justify-center rounded-full bg-rose-50 border border-rose-200 hover:bg-rose-100 transition-colors shadow-xs cursor-pointer text-[12px]"
                  style={{ width: "20px", height: "20px" }}
                >
                  ⚠️
                </button>
                <div className="absolute left-full top-1/2 -translate-y-1/2 ml-2 w-72 p-3 bg-white border border-rose-200 rounded-lg shadow-lg opacity-0 group-hover:opacity-100 pointer-events-none transition-opacity z-50">
                  <p className="text-xs text-rose-800" style={{ lineHeight: "1.4" }}>
                    {!isMonthMatched
                      ? `Tháng được chọn trên sidebar (${selectedMonth}) chưa trùng khớp với bất kỳ dòng THÁNG nào của file AE tải lên (${fileMonths.join(", ") || "Chưa tải file"}).`
                      : `Sự chênh lệch tài chính giữa (Sheet 1 + Hoạt động Hold AE) & Bank North AE khác 0 (Chênh lệch hiện tại: ${fmt(bulkPaymentDiff)}).`}
                  </p>
                </div>
              </div>
            )}
          </div>
          {/* Controls */}
          <div className="flex gap-2 items-center flex-wrap">
            <DropdownMenu>
              <DropdownMenuTrigger asChild>
                <Button
                  variant="outline"
                  size="sm"
                  className="p-0 bg-white dark:bg-card border border-[#e7dbdc] text-foreground hover:bg-muted/50 cursor-pointer shadow-xs rounded-[18px]"
                  style={{ width: "32px", height: "32px" }}
                >
                  <Settings className="w-4 h-4" />
                </Button>
              </DropdownMenuTrigger>
              <DropdownMenuContent
                align="end"
                className="w-[290px] bg-white dark:bg-card border border-[#e7dbdc] dark:border-border shadow-2xl p-2.5 flex flex-col gap-2 rounded-xl overflow-hidden z-[99999]"
              >
                <div className="h-8 flex items-center justify-center px-4 bg-primary/10 text-primary rounded-lg font-bold tracking-wide text-xs">
                  Kỳ: {currentPeriod}
                </div>
                
                {/* Save and Delete action buttons inside Settings dropdown */}
                <TableRestoreButton placement="menu" onRestore={() => updateAppData(prev => ({ ...prev }), false)} />
                <Button
                  size="sm"
                  onClick={handleSaveBalances}
                  disabled={isPeriodSaved(currentPeriod)}
                  className="h-8 text-xs normal-case tracking-normal w-full justify-start gap-2.5 rounded-lg font-bold shadow-xs border disabled:opacity-60 disabled:cursor-not-allowed cursor-pointer px-3 overflow-hidden"
                  style={{
                    backgroundColor: isPeriodSaved(currentPeriod) ? "#9ca3af" : "#b183ad",
                    color: "#ffffff",
                    borderColor: "#e8eae9",
                  }}
                >
                  <CheckCircle2 className="w-3.5 h-3.5 shrink-0" />
                  <span className="truncate flex-1 text-left">{isPeriodSaved(currentPeriod) ? "Đã lưu dữ liệu" : "Lưu dữ liệu"}</span>
                </Button>
                <Button
                  size="sm"
                  onClick={() => setShowDeletePeriodDialog(true)}
                  disabled={!isPeriodSaved(currentPeriod)}
                  className="h-8 text-xs normal-case tracking-normal w-full justify-start gap-2.5 rounded-lg font-bold shadow-xs bg-rose-600 hover:bg-rose-700 text-white border border-rose-500 cursor-pointer disabled:cursor-not-allowed disabled:opacity-45 px-3 overflow-hidden"
                >
                  <Trash2 className="w-3.5 h-3.5 shrink-0" />
                  <span className="truncate flex-1 text-left">Xóa dữ liệu kỳ hiện tại</span>
                </Button>

                <Button
                  size="sm"
                  onClick={() => setShowClearBalancePageDialog(true)}
                  className="h-8 text-xs normal-case tracking-normal w-full justify-start gap-2.5 rounded-lg font-bold shadow-xs bg-rose-50 hover:bg-rose-100 text-rose-700 border border-rose-200 cursor-pointer px-3 overflow-hidden"
                >
                  <Trash2 className="w-3.5 h-3.5 shrink-0" />
                  <span className="truncate flex-1 text-left">Xóa dữ liệu trang Balance</span>
                </Button>

                <Select value={yearFilter} onValueChange={setYearFilter}>
                  <SelectTrigger className="h-8 text-xs normal-case tracking-normal w-full bg-background border-[#e7dbdc] text-foreground rounded-lg px-3 font-medium">
                    <SelectValue placeholder="Năm" />
                  </SelectTrigger>
                  <SelectContent className="bg-white dark:bg-card border border-[#e7dbdc] z-[99999]">
                    <SelectItem value="all">Tất cả</SelectItem>
                    <SelectItem value="2026">2026</SelectItem>
                  </SelectContent>
                </Select>
                <Button
                  onClick={() => window.dispatchEvent(new Event("open-ui-settings"))}
                  variant="outline"
                  size="sm"
                  className="h-8 text-xs normal-case tracking-normal w-full justify-start gap-2.5 bg-background border-[#e7dbdc] text-foreground hover:bg-muted cursor-pointer px-3 font-semibold overflow-hidden"
                >
                  <Settings className="w-3.5 h-3.5 text-slate-500 shrink-0" />
                  <span className="truncate flex-1 text-left">Cài đặt giao diện</span>
                </Button>
                <Button
                  variant="outline"
                  size="sm"
                  onClick={() => chooseExcelExport(handleExportExcel, handleExportExcel)}
                  className="h-8 text-xs normal-case tracking-normal w-full justify-start gap-2.5 bg-background border-[#e7dbdc] text-foreground hover:bg-muted cursor-pointer px-3 font-semibold overflow-hidden"
                >
                  <Download className="w-3.5 h-3.5 shrink-0" />
                  <span className="truncate flex-1 text-left">Xuất Excel Trial Balance</span>
                </Button>
              </DropdownMenuContent>
            </DropdownMenu>
          </div>
        </div>
      </div>
      <div
        id="trial-balance-table-body"
        className="table-body-region flex-1 min-h-0 overflow-y-auto custom-scrollbar border-0 shadow-none"
        style={{ borderRadius: "0px", height: "calc(100vh - 250px)" }}
      >
          <table
            className="trial-balance-table w-full border-separate border-spacing-0 table-auto bg-white dark:bg-card"
            style={{
              fontFamily: uiSettings.tableFont || "var(--font-main)",
              fontSize: uiSettings.fontSize || "13px",
            }}
          >
            <thead className="sticky top-0 z-20 shadow-sm border-b border-[#dfd0d6] dark:border-slate-700 bg-[var(--table-column-header-bg,#F4F2EE)] text-[var(--table-column-header-text-color,#1e293b)]">
              <tr>
                <th
                  className="border-r border-b border-[#dfd0d6] dark:border-slate-800 px-3 py-2.5 text-center font-sans font-bold text-[12px] uppercase tracking-wider bg-[var(--table-column-header-bg,#F4F2EE)] text-[var(--table-column-header-text-color,#1e293b)] whitespace-nowrap"
                  rowSpan={2}
                >
                  <div className="flex items-center justify-center gap-2 w-full">
                    <span>#</span>
                    <button
                      onClick={toggleAll}
                      title={allOpen ? "Thu gọn tất cả" : "Mở tất cả"}
                      className="p-1 rounded hover:bg-black/10 dark:hover:bg-white/10 opacity-80 hover:opacity-100 transition-opacity flex items-center justify-center cursor-pointer"
                    >
                      {allOpen ? (
                        <ChevronsDownUp className="w-4 h-4" />
                      ) : (
                        <ChevronsUpDown className="w-4 h-4" />
                      )}
                    </button>
                  </div>
                </th>
                <th
                  className="border-r border-b border-[#dfd0d6] dark:border-slate-800 px-3 py-2.5 text-center font-sans font-bold text-[12px] uppercase tracking-wider bg-[var(--table-column-header-bg,#F4F2EE)] text-[var(--table-column-header-text-color,#1e293b)] whitespace-nowrap"
                  rowSpan={2}
                >
                  Ngày / Tháng
                </th>
                <th
                  className="border-r border-b border-[#dfd0d6] dark:border-slate-800 px-4 py-2.5 text-center font-sans font-bold text-[12px] uppercase tracking-wider bg-[var(--table-column-header-bg,#F4F2EE)] text-[var(--table-column-header-text-color,#1e293b)] min-w-[200px]"
                  rowSpan={2}
                >
                  Business
                </th>
                <th
                  className="border-r border-b border-[#dfd0d6] dark:border-slate-800 px-3 py-2.5 text-center font-sans font-bold text-[12px] uppercase tracking-wider bg-[var(--table-column-header-bg,#F4F2EE)] text-[var(--table-column-header-text-color,#1e293b)] whitespace-nowrap min-w-[100px]"
                  rowSpan={2}
                >
                  Số dư Hold ĐK
                </th>
                <th
                  className="border-r border-b border-[#dfd0d6] dark:border-slate-800 px-3 py-2.5 text-center font-sans font-bold text-[12px] uppercase tracking-wider bg-[var(--table-column-header-bg,#F4F2EE)] text-[var(--table-column-header-text-color,#1e293b)] whitespace-nowrap"
                  colSpan={2}
                >
                  Phát sinh trong kỳ
                </th>
                <th
                  className="border-r border-b border-[#dfd0d6] dark:border-slate-800 px-3 py-2.5 text-center font-sans font-bold text-[12px] uppercase tracking-wider bg-[var(--table-column-header-bg,#F4F2EE)] text-[var(--table-column-header-text-color,#1e293b)] whitespace-nowrap min-w-[100px]"
                  colSpan={2}
                >
                  Số dư CK
                </th>
                <th
                  className="note-column-header border-b border-[#dfd0d6] px-3 py-2.5 text-center font-sans font-bold text-[12px] uppercase tracking-wider bg-[var(--table-column-header-bg,#F4F2EE)] text-[var(--table-column-header-text-color,#1e293b)] whitespace-nowrap min-w-[200px]"
                  rowSpan={2}
                  style={{ borderLeft: 0, borderRight: "1px solid #dfd0d6" }}
                >
                  NOTE
                </th>
              </tr>
              <tr>
                <th 
                  className="border-r border-b border-[#dfd0d6] dark:border-slate-800 px-3 py-1.5 text-center font-sans font-bold text-[11px] uppercase tracking-wider bg-[var(--table-column-header-bg,#F4F2EE)] text-[var(--table-column-header-text-color,#1e293b)] whitespace-nowrap min-w-[100px]"
                >
                  Lương TA của tháng
                </th>
                <th 
                  className="border-r border-b border-[#dfd0d6] dark:border-slate-800 px-3 py-1.5 text-center font-sans font-bold text-[11px] uppercase tracking-wider bg-[var(--table-column-header-bg,#F4F2EE)] text-[var(--table-column-header-text-color,#1e293b)] whitespace-nowrap min-w-[100px]"
                >
                  Lương Hold của tháng
                </th>
                <th 
                  className="border-r border-b border-[#dfd0d6] dark:border-slate-800 px-3 py-1.5 text-center font-sans font-bold text-[11px] uppercase tracking-wider bg-[var(--table-column-header-bg,#F4F2EE)] text-[var(--table-column-header-text-color,#1e293b)] whitespace-nowrap min-w-[100px]"
                >
                  Tổng PS tại kỳ
                </th>
                <th 
                  className="border-r border-b border-[#dfd0d6] dark:border-slate-800 px-3 py-1.5 text-center font-sans font-bold text-[11px] uppercase tracking-wider bg-[var(--table-column-header-bg,#F4F2EE)] text-[var(--table-column-header-text-color,#1e293b)] whitespace-nowrap min-w-[100px]"
                >
                  Số dư Hold còn lại
                </th>
              </tr>
            </thead>

            <tbody>
              {monthKeys.length === 0 && (
                <tr>
                  <td
                    colSpan={9}
                    className="p-10 text-center text-muted-foreground italic font-medium bg-white dark:bg-card border-r border-b border-[#e7dbdc] dark:border-slate-800"
                  >
                    Không có dữ liệu phù hợp thỏa mãn điều kiện tìm kiếm.
                  </td>
                </tr>
              )}

              {(() => {
                return monthKeys.map((mk, mi) => {
                  const rows = grouped.get(mk) || [];

                    const {
                      openBal,
                      psThu,
                      psChi,
                    } = computedMonthTotals[mk] || {
                      openBal: 0,
                      psThu: 0,
                      psChi: 0,
                      closeBal: 0,
                      addAmt: 0,
                      holdAmt: 0,
                      cancelAmt: 0,
                    };
                  const isOpen = expanded.has(mk);

                  return [
                    // Month row
                    <tr
                      key={`m-${mk}`}
                      className="trial-month-total-row cursor-pointer group transition-colors"
                      onClick={() => toggle(mk)}
                    >
                      <td className="border-r border-b border-[#dfd0d6] dark:border-slate-800 p-3 text-center text-slate-800 dark:text-slate-200 font-bold !bg-[#FAF9F6]/80 dark:!bg-slate-800/60 whitespace-nowrap text-[13px]">
                        {toRoman(mi + 1)}
                      </td>
                      <td className="border-r border-b border-[#dfd0d6] dark:border-slate-800 p-3 text-left text-slate-800 dark:text-slate-200 font-bold !bg-[#FAF9F6]/80 dark:!bg-slate-800/60 whitespace-nowrap text-[13px] pl-3">
                        <span className="flex items-center gap-2">
                          {isOpen ? (
                            <ChevronDown className="w-4 h-4 text-slate-800 dark:text-slate-200 stroke-[2.5px] shrink-0" />
                          ) : (
                            <ChevronRight className="w-4 h-4 text-slate-800 dark:text-slate-200 stroke-[2.5px] shrink-0" />
                          )}
                          {mk === currentPeriod ? `Tháng ${currentPeriodVal}` : mk}
                        </span>
                      </td>
                      <td className="border-r border-b border-[#dfd0d6] dark:border-slate-800 p-3 text-center text-slate-800 dark:text-slate-200 font-bold !bg-[#FAF9F6]/80 dark:!bg-slate-800/60 whitespace-nowrap text-[13px]">
                        {countBusinesses(rows)} BU
                      </td>
                      <td className="border-r border-b border-[#dfd0d6] dark:border-slate-800 p-3 text-right text-slate-800 dark:text-slate-200 font-bold tabular-nums text-xs !bg-[#FAF9F6]/80 dark:!bg-slate-800/60 whitespace-nowrap">
                        {openBal !== 0 ? fmt(openBal) : "0"}
                      </td>
                      <td className="border-r border-b border-[#dfd0d6] dark:border-slate-800 p-3 text-right text-slate-800 dark:text-slate-200 font-bold tabular-nums text-xs !bg-[#FAF9F6]/80 dark:!bg-slate-800/60 whitespace-nowrap">
                        {psThu !== 0 ? fmt(psThu) : "0"}
                      </td>
                      <td className="border-r border-b border-[#dfd0d6] dark:border-slate-800 p-3 text-right text-slate-800 dark:text-slate-200 font-bold tabular-nums text-xs !bg-[#FAF9F6]/80 dark:!bg-slate-800/60 whitespace-nowrap">
                        {psChi !== 0 ? fmt(psChi) : "0"}
                      </td>
                      <td className="trial-month-close-balance border-r border-b border-[#dfd0d6] dark:border-slate-800 p-3 text-right text-slate-800 dark:text-slate-200 font-bold tabular-nums text-xs !bg-[#FAF9F6]/80 dark:!bg-slate-800/60 whitespace-nowrap">
                        {fmt(monthCloseBalances[mk] ?? 0)}
                      </td>
                      <td className="border-r border-b border-[#dfd0d6] dark:border-slate-800 p-3 text-right text-slate-800 dark:text-slate-200 font-bold tabular-nums text-xs !bg-[#FAF9F6]/80 dark:!bg-slate-800/60 whitespace-nowrap">
                        {fmt(remainingHoldByMonth[mk]?.total || 0)}
                      </td>

                      <td className="note-column-cell border-b border-[#dfd0d6] min-w-[200px] !bg-[#FAF9F6]/80 dark:!bg-slate-800/60" style={{ borderLeft: 0, borderRight: "1px solid #dfd0d6" }}></td>
                    </tr>,

                    // Detail rows
                    ...(isOpen
                      ? (() => {
                          const uniqueBUs = Array.from(new Set(rows.map(r => r.bu)));
                          let businessIndex = 0;
                          return uniqueBUs.flatMap((bu) => {
                            const buRows = rows.filter(r => r.bu === bu && !hideInactivePastHold(r, mk) && (!String(r.id).includes("_cancel") || isExactReportPeriod(r, currentPeriod)))
                              .sort((a, b) => trialBalanceRowOrder(a) - trialBalanceRowOrder(b));
                            const mainNumber = ++businessIndex;
                            let detailIndex = 0;
                            
                            const sumOpenBal = buBalancesByMonth[mk]?.[bu]?.openBal || 0;
                            let sumThu = 0;
                            let sumChi = 0;
                            
                            const hasDetailOpenHold = buRows.some(r => !!r.customMonthDisplay && (r.openHold || 0) > 0);
                            const hasDetailHoldOrOpen = buRows.some(r => !!r.customMonthDisplay);

                            const renderedBuRows = buRows.map((e) => {
                              const isDetail = !!e.customMonthDisplay;
                              const rowNumber = isDetail ? `${mainNumber}.${++detailIndex}` : String(mainNumber);
                              const buBalInfo = buBalancesByMonth[mk]?.[e.bu];
                              const rowOpenBal = isDetail ? (e.openHold || 0) : (hasDetailOpenHold ? 0 : (buBalInfo?.openBal || 0));
                              
                              const origin = trialBalancePeriod(e.displayMonth || e.month);
                              const report = trialBalancePeriod(e.reportMonth || e.month || mk);
                              const isCancel =
                                /cancel/i.test(e.customMonthDisplay || "") ||
                                String(e.id).includes("_cancel") ||
                                /cancel/i.test(e.ghiChu || "") ||
                                !!(e.rawCancel || e.cancel);
                              const isCurrentHold =
                                !isCancel &&
                                origin === report &&
                                (!!(e.rawHold || e.hold) || /_hold\b/i.test(e.id) || /\bhold\b/i.test(e.customMonthDisplay || ""));
                              const displayedThu = e.thu;
                              const displayedChi = Math.abs(e.chi);
                              const rowRemainingHold = isCancel
                                ? 0
                                : isDetail
                                  ? (isCurrentHold
                                      ? (displayedChi || e.rawHold || e.hold || 0)
                                      : Math.max(0, (e.openHold || 0) - displayedThu - (e.rawCancel || e.cancel || 0)))
                                  : (hasDetailHoldOrOpen ? 0 : (remainingHoldByMonth[mk]?.byBu[bu] || 0));
                              const rClose = rowRCloseBalances[e.id] ?? 0;
                              const isRowDimmed =
                                !!e._dimmed && isPeriodSaved(e.month);
                              const displayedRCloseStr = rClose !== 0 ? fmt(rClose) : "0";
                              
                              sumThu += isRowDimmed ? 0 : displayedThu;
                              sumChi += isRowDimmed ? 0 : displayedChi;

                              return (
                                <tr
                                  key={e.id}
                                  data-trial-detail={isDetail ? "true" : undefined}
                                  className={`group ${
                                    isRowDimmed
                                      ? "opacity-35 select-none bg-slate-100/50 dark:bg-slate-800/10 italic text-muted-foreground/60 line-through"
                                      : !isDetail
                                        ? "bg-[#F8F9FA] dark:bg-slate-800/50 hover:bg-[#F1F3F5] dark:hover:bg-slate-800/70 font-medium"
                                        : "bg-white dark:bg-card hover:bg-slate-50/70 dark:hover:bg-slate-800/30"
                                  } transition-colors`}
                                >
                                  <td
                                    className={`border-r border-b border-[#dfd0d6] dark:border-slate-800 ${
                                      !isDetail ? "font-bold text-slate-800 dark:text-slate-100 text-[12px]" : "font-normal text-slate-400 dark:text-slate-500 text-[11px]"
                                    } p-2 text-center whitespace-nowrap`}
                                  >
                                    {rowNumber}
                                  </td>
                                  <td
                                    className={`border-r border-b border-[#dfd0d6] dark:border-slate-800 ${
                                      !isDetail
                                        ? "px-3 py-2 text-left whitespace-nowrap min-w-[120px] text-[12.5px] font-bold text-slate-900 dark:text-slate-50 tracking-tight"
                                        : "px-3 py-2 text-left whitespace-nowrap min-w-[120px] text-[12px] text-slate-600 dark:text-slate-300 font-normal pl-5"
                                    }`}
                                    title={trialBalanceRowLabel(e)}
                                  >
                                    {isDetail ? (
                                      <span className="inline-flex items-center gap-1.5">
                                        <span className="text-slate-400 dark:text-slate-500 text-[10px] font-mono select-none">↳</span>
                                        <span>{trialBalanceRowLabel(e)}</span>
                                      </span>
                                    ) : (
                                      trialBalanceRowLabel(e)
                                    )}
                                  </td>
                                  <td
                                    className="border-r border-b border-[#dfd0d6] dark:border-slate-800 px-3 py-2 text-center whitespace-nowrap text-[12px]"
                                    title={e.bu}
                                  >
                                    {!isDetail ? (
                                      <span className="font-bold text-slate-800 dark:text-slate-200 bg-slate-200/60 dark:bg-slate-700/60 px-1.5 py-0.5 rounded text-[11px] tracking-wide">
                                        {e.bu}
                                      </span>
                                    ) : (
                                      <span className="text-slate-400 dark:text-slate-500 text-[11px] font-normal">
                                        {e.bu}
                                      </span>
                                    )}
                                  </td>
                                  <td
                                    className={`border-r border-b border-[#dfd0d6] dark:border-slate-800 ${
                                      !isDetail ? "font-semibold text-slate-900 dark:text-slate-100" : "font-normal text-slate-600 dark:text-slate-400"
                                    } p-2 text-right tabular-nums text-xs whitespace-nowrap min-w-[75px]`}
                                  >
                                    {rowOpenBal !== 0 ? (
                                      <span
                                        className={`trial-opening-amount ${
                                          !isDetail ? "text-slate-900 dark:text-slate-100 font-semibold text-[12px]" : "text-slate-600 dark:text-slate-400 font-normal text-[12px]"
                                        }`}
                                      >
                                        {fmt(rowOpenBal)}
                                      </span>
                                    ) : (
                                      "0"
                                    )}
                                  </td>
                                  <td
                                    className="border-r border-b border-[#dfd0d6] dark:border-slate-800 p-2 text-right tabular-nums text-xs whitespace-nowrap min-w-[80px]"
                                  >
                                    {!isRowDimmed && displayedThu !== 0 ? (
                                      <span
                                        className={`${
                                          !isDetail ? "text-slate-900 dark:text-slate-100 font-semibold text-[12px]" : "text-slate-600 dark:text-slate-400 font-normal text-[12px]"
                                        }`}
                                      >
                                        {fmt(displayedThu)}
                                      </span>
                                    ) : (
                                      "0"
                                    )}
                                  </td>
                                  <td
                                    className="border-r border-b border-[#dfd0d6] dark:border-slate-800 p-2 text-right tabular-nums text-xs whitespace-nowrap min-w-[80px]"
                                  >
                                    {!isRowDimmed && displayedChi !== 0 ? (
                                      <span
                                        className={`${
                                          !isDetail ? "text-slate-900 dark:text-slate-100 font-semibold text-[12px]" : "text-slate-600 dark:text-slate-400 font-normal text-[12px]"
                                        }`}
                                      >
                                        {fmt(displayedChi)}
                                      </span>
                                    ) : (
                                      "0"
                                    )}
                                  </td>
                                  <td
                                    className="border-r border-b border-[#dfd0d6] dark:border-slate-800 p-2 text-right tabular-nums text-xs whitespace-nowrap min-w-[80px]"
                                  >
                                    <span
                                      className={`${
                                        !isDetail ? "text-rose-600 dark:text-rose-400 font-bold text-[12px]" : "text-rose-600/80 dark:text-rose-400/80 font-medium text-[12px]"
                                      }`}
                                    >
                                      {isDetail ? fmt(displayedThu) : displayedRCloseStr}
                                    </span>
                                  </td>
                                  <td
                                    className={`border-r border-b border-[#dfd0d6] dark:border-slate-800 ${
                                      !isDetail
                                        ? "font-semibold text-slate-900 dark:text-slate-100"
                                        : "font-normal text-slate-600 dark:text-slate-400"
                                    } p-2 text-right tabular-nums text-xs whitespace-nowrap min-w-[80px]`}
                                  >
                                    {fmt(rowRemainingHold)}
                                  </td>

                                  <td
                                    className={`note-column-cell border-b border-[#dfd0d6] ${
                                      !isDetail ? "text-slate-700 dark:text-slate-300 text-[12px]" : "text-muted-foreground/80 text-[11px] italic"
                                    } p-2 text-left min-w-[200px]`}
                                    style={{ borderLeft: 0, borderRight: "1px solid #dfd0d6" }}
                                    title={e.ghiChu}
                                  >
                                    {e.ghiChu || ""}
                                  </td>
                                </tr>
                              );
                            });
                            
                            // Sum the displayed closing values, which already normalize HOLD signs.
                            const finalCloseBal = buRows.reduce(
                              (sum, row) => sum + (rowRCloseBalances[row.id] ?? 0), 0,
                            );

                            const subtotalRow = (
                              <tr key={`subtotal-${mk}-${bu}`} className="!bg-[#F2EADB] dark:!bg-slate-800/90 font-bold border-y-2 border-[#dfd0d6] dark:border-slate-700">
                                <td colSpan={3} className="border-r border-b border-[#dfd0d6] dark:border-slate-700 p-2 text-center text-slate-800 dark:text-slate-200 font-sans uppercase tracking-wider text-[11px] !bg-[#F2EADB] dark:!bg-slate-800/90">
                                  TỔNG BU - {bu}
                                </td>
                                <td className="border-r border-b border-[#dfd0d6] dark:border-slate-700 p-2 text-right text-slate-800 dark:text-slate-100 tabular-nums text-xs whitespace-nowrap !bg-[#F2EADB] dark:!bg-slate-800/90">
                                  {sumOpenBal !== 0 ? fmt(sumOpenBal) : "0"}
                                </td>
                                <td className="border-r border-b border-[#dfd0d6] dark:border-slate-700 p-2 text-right text-slate-800 dark:text-slate-100 tabular-nums text-xs whitespace-nowrap !bg-[#F2EADB] dark:!bg-slate-800/90">
                                  {sumThu !== 0 ? fmt(sumThu) : "0"}
                                </td>
                                <td className="border-r border-b border-[#dfd0d6] dark:border-slate-700 p-2 text-right text-slate-800 dark:text-slate-100 tabular-nums text-xs whitespace-nowrap !bg-[#F2EADB] dark:!bg-slate-800/90">
                                  {sumChi !== 0 ? fmt(sumChi) : "0"}
                                </td>
                                <td className="border-r border-b border-[#dfd0d6] dark:border-slate-700 p-2 text-right tabular-nums text-xs whitespace-nowrap text-rose-600 dark:text-rose-400 font-bold !bg-[#F2EADB] dark:!bg-slate-800/90">
                                  {finalCloseBal !== 0 ? fmt(finalCloseBal) : "0"}
                                </td>
                                <td className="border-r border-b border-[#dfd0d6] dark:border-slate-700 p-2 text-right text-slate-800 dark:text-slate-200 tabular-nums text-xs whitespace-nowrap !bg-[#F2EADB] dark:!bg-slate-800/90">
                                  {fmt(remainingHoldByMonth[mk]?.byBu[bu] || 0)}
                                </td>

                                <td className="note-column-cell border-b border-[#dfd0d6] dark:border-slate-700 p-2 !bg-[#F2EADB] dark:!bg-slate-800/90" style={{ borderLeft: 0, borderRight: "1px solid #dfd0d6" }}></td>
                              </tr>
                            );
                            
                            const firstDetail = buRows.findIndex(row => Boolean(row.customMonthDisplay));
                            const insertAt = firstDetail < 0 ? renderedBuRows.length : firstDetail;
                            const grossTotals = grossPayByMonth.get(parsePayrollMonth(mk)?.dot || "");
                            const grossPayRow = (
                              <tr key={`gross-pay-${mk}-${bu}`} className="bg-primary/5 font-bold text-foreground">
                                <td className="border-b border-border" />
                                <th scope="row" className="border-b border-border px-3 py-2 text-left text-xs whitespace-nowrap">
                                  Tổng phát sinh CP Intern
                                </th>
                                <td className="border-b border-border p-2 text-center text-xs">{bu}</td>
                                <td colSpan={3} className="border-b border-border" />
                                <td className="border-b border-border p-2 text-right tabular-nums text-xs whitespace-nowrap">
                                  {grossTotals ? fmt(grossTotals.get(normalizeGrossPayBusiness(bu)) || 0) : "—"}
                                </td>
                                <td colSpan={2} className="border-b border-border" />
                              </tr>
                            );
                            return [...renderedBuRows.slice(0, insertAt), grossPayRow, ...renderedBuRows.slice(insertAt), subtotalRow];
                          });
                        })()
                      : []),
                  ];
                });
              })()}
            </tbody>

            <tfoot className="sticky bottom-0 z-20 shadow-[0_-2px_10px_rgba(0,0,0,0.08)] bg-[var(--table-column-header-bg,#E8DEC8)] text-[var(--table-column-header-text-color,inherit)]">
              <tr className="border-t-2 border-b-2 border-[#dfd0d6] dark:border-slate-700 bg-[var(--table-column-header-bg,#E8DEC8)] text-[var(--table-column-header-text-color,inherit)]">
                <td
                  colSpan={3}
                  className="border-r border-b border-[#dfd0d6] dark:border-slate-700 px-3 py-3 text-center font-sans font-bold text-[12px] uppercase tracking-wider text-[var(--table-column-header-text-color,inherit)] whitespace-nowrap bg-[var(--table-column-header-bg,#E8DEC8)] sticky bottom-0 z-20"
                >
                  TỔNG CỘNG THÁNG {currentPeriodVal} —{" "}
                  <span className="opacity-75 font-bold ml-1 tracking-normal tabular-nums text-xs text-[var(--table-column-header-text-color,inherit)]">
                    {countBusinesses(filteredData)} BU
                  </span>
                </td>
                <td 
                  className="border-r border-b border-[#dfd0d6] dark:border-slate-700 p-3 text-right tabular-nums text-xs font-bold text-[var(--table-column-header-text-color,inherit)] whitespace-nowrap bg-[var(--table-column-header-bg,#E8DEC8)] sticky bottom-0 z-20"
                >
                  {grandOpenBal !== 0 ? fmt(grandOpenBal) : "0"}
                </td>
                <td 
                  className="border-r border-b border-[#dfd0d6] dark:border-slate-700 p-3 text-right tabular-nums text-xs font-bold text-emerald-400 whitespace-nowrap bg-[var(--table-column-header-bg,#E8DEC8)] sticky bottom-0 z-20"
                >
                  {grandThu !== 0 ? fmt(grandThu) : "0"}
                </td>
                <td 
                  className="border-r border-b border-[#dfd0d6] dark:border-slate-700 p-3 text-right tabular-nums text-xs font-bold text-rose-400 whitespace-nowrap bg-[var(--table-column-header-bg,#E8DEC8)] sticky bottom-0 z-20"
                >
                  {grandChi !== 0 ? fmt(grandChi) : "0"}
                </td>
                <td
                  className={`border-r border-b border-[#dfd0d6] dark:border-slate-700 p-3 text-right tabular-nums text-xs font-bold whitespace-nowrap bg-[var(--table-column-header-bg,#E8DEC8)] sticky bottom-0 z-20 ${grandBal >= 0 ? "text-[var(--table-column-header-text-color,inherit)]" : "text-rose-400 font-extrabold"}`}
                >
                  {fmt(grandBal)}
                </td>
                <td 
                  className="border-r border-b border-[#dfd0d6] dark:border-slate-700 p-3 text-right tabular-nums text-xs font-bold text-[var(--table-column-header-text-color,inherit)] whitespace-nowrap bg-[var(--table-column-header-bg,#E8DEC8)] sticky bottom-0 z-20"
                >
                  {fmt(monthKeys.reduce((sum, mk) => sum + (remainingHoldByMonth[mk]?.total || 0), 0))}
                </td>
                <td 
                  className="note-column-cell border-b border-[#dfd0d6] min-w-[200px] bg-[var(--table-column-header-bg,#E8DEC8)] sticky bottom-0 z-20"
                  style={{ borderLeft: 0, borderRight: "1px solid #dfd0d6" }}
                />
              </tr>
            </tfoot>
          </table>
      </div>

      {/* Caption & Footer Pagination Controls */}
      <div
        className="table-footer-pagination flex-shrink-0 py-2 px-4 flex items-center justify-between gap-3 text-xs w-full bg-[var(--table-footer-bg,var(--table-header-bg,#FAF3E8))]"
        style={{ height: "50.9848px", minHeight: "50.9848px", backgroundColor: "var(--table-footer-bg, var(--table-header-bg, #FAF3E8))" }}
      >
          <div className="flex items-center gap-2">
            <span className="text-[12px] font-semibold text-slate-700 dark:text-slate-300 font-sans">
              Hiển thị:
            </span>
            <Select
              value={String(itemsPerPage)}
              onValueChange={(val) => {
                setItemsPerPage(val === "all" ? "all" : Number(val));
              }}
            >
              <SelectTrigger 
                className="rounded-full px-3 text-[12px] font-bold font-sans text-slate-700 dark:text-slate-200 border-[#e7dbdc] bg-white dark:bg-slate-800 hover:bg-slate-50 transition-colors shadow-2xs cursor-pointer"
                style={{ height: "26px", width: "100px", fontSize: "12px" }}
              >
                <SelectValue placeholder="Tất cả" className="text-[12px]" />
              </SelectTrigger>
              <SelectContent className="bg-white dark:bg-slate-800 border-[#e7dbdc] z-[99999] opacity-100 font-sans">
                <SelectItem value="10" className="text-[12px] font-medium font-sans">10 dòng</SelectItem>
                <SelectItem value="20" className="text-[12px] font-medium font-sans">20 dòng</SelectItem>
                <SelectItem value="50" className="text-[12px] font-medium font-sans">50 dòng</SelectItem>
                <SelectItem value="100" className="text-[12px] font-medium font-sans">100 dòng</SelectItem>
                <SelectItem value="all" className="text-[12px] font-medium font-sans">Tất cả</SelectItem>
              </SelectContent>
            </Select>
          </div>
          <div className="text-[12px] font-semibold text-slate-600 dark:text-slate-400 font-sans">
            Tổng số: {monthKeys.length} kỳ ({currentPeriodRows.filter(row => !hideInactivePastHold(row, currentPeriod)).length} dòng dữ liệu)
          </div>
      </div>
      <ConfirmDialog
        isOpen={showClearBalancePageDialog}
        onClose={() => setShowClearBalancePageDialog(false)}
        onConfirm={handleClearBalancePage}
        title="Xóa dữ liệu trang Balance?"
        description="Toàn bộ kỳ đã lưu, số dư chuyển kỳ và trạng thái xác nhận của Balance sẽ bị xóa. Dữ liệu Timesheet, Audit và Master được giữ nguyên."
        confirmText="XÓA TRANG BALANCE"
        variant="destructive"
      />
      <ConfirmDialog
        isOpen={showDeletePeriodDialog}
        onClose={() => setShowDeletePeriodDialog(false)}
        onConfirm={() => {
          handleDeleteSavedPeriod();
          setShowDeletePeriodDialog(false);
        }}
        title={`Xóa dữ liệu đã lưu ${currentPeriod}?`}
        description={`Bạn có chắc chắn muốn xóa dữ liệu đã lưu của ${currentPeriod}? Số dư chuyển tiếp kỳ sau liên quan cũng sẽ được tính toán lại. Dữ liệu các kỳ khác được giữ nguyên.`}
        confirmText="XÓA DỮ LIỆU KỲ NÀY"
        variant="destructive"
      />
    </div>
  );
}
