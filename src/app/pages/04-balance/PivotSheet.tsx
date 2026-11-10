import { TableRestoreButton } from '../../components/TableRestoreButton';
import { parsePayrollMonth } from "../../lib/utils/hold-carryover";
import { findMasterRosterSheetName } from "../../lib/utils/master-sheet-utils";
import { resolveMasterSpecialCenter } from "../../lib/utils/master-special-centers";
import { registerTableExport } from "../../lib/utils/table-excel";
import { chooseExcelExport } from "../../components/ExportScopeDialog";
import { ConfirmDialog } from "../../components/shared/ConfirmDialog";
import { isAhpBuValue } from "../../components/DataTable";
/* eslint-disable @typescript-eslint/no-explicit-any, react-hooks/set-state-in-effect, react-hooks/purity, @typescript-eslint/no-unused-vars */
import React, { useState, useEffect, useCallback, useRef, useMemo } from "react";
import * as XLSX from "xlsx";
import {
  Download,
  RefreshCw,
  FileSpreadsheet,
  Eye,
  Upload,
  SlidersHorizontal,
  ChevronLeft,
  ChevronRight,
  ChevronsLeft,
  ChevronsRight,
  ChevronUp,
  ChevronDown,
  ArrowUpDown,
  Trash2,
  Columns,
  Rows,
  Plus,
  X,
  Settings,
  Maximize2,
  Filter,
} from "lucide-react";
import { useAppData } from "../../lib/contexts/AppDataContext";
import {
  applyPivotMktTypeCache,
  buildPivotFromAppData,
  formatPivotTypeHeader,
  getPivotDataMonths,
  getPivotSourceLabels,
  markPivotZhnSource,
  normalizePivotMonth,
  omitPivotMonths,
  readPivotMktTypeCache,
  replacePivotCacheMonth,
  PIVOT_CACHE_VERSION,
  updatePivotMktTypeCache,
  writePivotMktTypeCache,
  type PivotGroupedData,
} from "../../lib/utils/pivot-utils";
import { resolveMktRosterCenter } from "../../lib/utils/center-utils";
import { parseDurationToHours } from "../../lib/schemas/excel-schema";
import { parseMoneyToNumber as parseExcelMoney } from "../../lib/utils/data-utils";
import { toast } from "sonner";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "../../components/ui/select";
import {
  TableInitialMark,
  TableTitleRemainder,
} from "../../components/TableInitialMark";

// ==========================================
// LOCAL HELPER UTILITIES
// ==========================================

function parseMoneyToNumber(val: any): number {
  return parseExcelMoney(val);
}

function formatNumber(val: any): string {
  const n = parseMoneyToNumber(val);
  const rounded = Math.round(n);
  return rounded.toLocaleString('vi-VN', {
    minimumFractionDigits: 0,
    maximumFractionDigits: 0,
  });
}

function formatMoneyVND(val: any): string {
  const n = parseMoneyToNumber(val);
  const rounded = Math.round(n);
  return rounded.toLocaleString('vi-VN', {
    minimumFractionDigits: 0,
    maximumFractionDigits: 0,
  });
}

function removeVietnameseTones(str: string): string {
  if (!str) return '';
  return str.normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/đ/g, "d").replace(/Đ/g, "D");
}

function formatIdNumber(id: any): string {
  return String(id || '').trim();
}

function prepareDataForExport(data: any[]): any[] {
  return data;
}

function parseAnyDate(dateStr: string): Date | null {
  if (!dateStr) return null;
  const d = new Date(dateStr);
  return isNaN(d.getTime()) ? null : d;
}

function getVal(row: any, key: string): any {
  return row ? row[key] : null;
}

function parseTimeStrToHours(timeStr: string): number {
  if (!timeStr) return 0;
  const [h, m] = String(timeStr).split(':').map(Number);
  return (h || 0) + (m || 0) / 60;
}

async function getExcelFileBuffer(
  file: File,
): Promise<{ buffer: ArrayBuffer; name: string }> {
  if (!file) {
    throw new Error("Không tìm thấy thông tin file để đọc.");
  }
  return {
    buffer: await file.arrayBuffer(),
    name: file.name,
  };
}

function formatTime12Hour(timeStr: string): string {
  return String(timeStr);
}

const COMMON_FIELD_ALIASES: Record<string, string[]> = {};
function scoreMatch(a: string, b: string): number { return a === b ? 1 : 0; }
function normalizeId(id: any): string { return String(id || ''); }
function toVietnamDateString(date: Date): string { return String(date); }
function generateUUID(): string { return Math.random().toString(36).substring(2, 9); }
async function fetchGoogleSheetAsFile(url: string, name: string): Promise<File> { return new File([], name); }
function isMoneyColumn(col: string): boolean { return Boolean(col && col.toLowerCase().includes('money')); }
async function fetchWithBackoff(fn: any): Promise<any> { return await fn(); }

function mergePivotTypeColumns(
  currentColumns: string[],
  incomingColumns: string[],
): string[] {
  const result: string[] = [];
  const seen = new Set<string>();

  [...(currentColumns || []), ...(incomingColumns || [])].forEach((column) => {
    const normalized = String(column || "").trim().toUpperCase();
    if (!normalized || seen.has(normalized)) return;
    seen.add(normalized);
    result.push(normalized);
  });

  const priorityOrder = [
    "MKT LOCAL",
    "LXO",
    "EC",
    "PT-DEMO",
    "LDEC01",
    "LDEM01",
    "LPAR01",
    "LRET01",
    "MOTH01",
    "RENEWAL PROJECTS",
    "DISCOVERY CAMP",
    "SUMMER OUTING",
    "SUMMER INSTRUCTORS",
    "EXTRA SUMMER INSTRUCTORS",
    "EXTRA INSTRUCTORS",
    "OTHER",
  ];

  return result.sort((a, b) => {
    if (a === "UNSPECIFIED") return 1;
    if (b === "UNSPECIFIED") return -1;
    const idxA = priorityOrder.indexOf(a);
    const idxB = priorityOrder.indexOf(b);
    if (idxA !== -1 && idxB !== -1) return idxA - idxB;
    if (idxA !== -1) return -1;
    if (idxB !== -1) return 1;
    return a.localeCompare(b);
  });
}

// ==========================================
// MAPPING DEFINITIONS & LOGIC FROM SPEC
// ==========================================

const rawCenterToMktMap: Record<string, string> = {
  "Ly Thai To": "BN0001.LTT", "Tu Son": "BN0002.TSN", "Pho Hue": "HN0001.PHY",
  "Thai Ha": "HN0002.THA", "Hoang Quoc Viet": "HN0003.HQV", "Lieu Giai": "HN0004.LGI",
  "Nguyen Van Linh": "HN0005.NVL", "Van Quan": "HN0007.VQN", "The Garden": "HN0010.MDH",
  "Nguyen Huu Tho": "HN0012.NHT", "Tan Mai": "HN0014.TMI", "Van Phu": "HN0015.VPU",
  "Phan Dinh Phung": "HN0016.PDP", "Ham Nghi": "HN0017.HNI", "Vu Tong Phan": "HN0018.VTP",
  "Nguyen Tuan": "HN0019.NTN", "Ngoai Giao Doan": "HN0021.NGD", "Mo Lao": "HN0022.NVO",
  "Linh Dam": "HN0023.LDM", "Times City": "HN0024.TCY", "Le Trong Tan": "HN0025.LTT",
  "Viet Hung": "HN0026.VHG", "Ocean Park": "HN0027.OPK", "Pham Van Dong": "HN0028.PVD",
  "Vu Pham Ham": "HN0029.VPH", "An Khanh": "HN0030.AKH", "An Hung": "HN0031.AHG",
  "Lac Long Quan": "HN0032.LLQ", "Dong Anh": "HN0033.DAH", "Hong Tien": "HN0034.HTN",
  "Ecopark": "HY0001.ECP", "Hai Phong": "Hai Phong", "Quang Ninh": "QN0001.HLG",
  "Vinh": "VIN001.CTG", "Vinh Phuc": "VP0001.PCT", "Thanh Hoa": "TH0001.TPU",
  "Thai Nguyen": "TN0001.LNQ", "Phu Tho": "PT0001.HVG", "NTW": "NTW"
};

const aeCodeToL07Map: Record<string, string> = {
  "Ngo Si Lien": "BN0001.LTT",
  "BN01.LTT": "BN0001.LTT",
  "Tu Son": "BN0002.TSN",
  "BN02.TUS": "BN0002.TSN",
  "Pho Hue Junior": "HN0001.PHY",
  "Pho Hue": "HN0001.PHY",
  "Thai Ha": "HN0002.THA",
  "Thai Ha (center Láng Hạ)": "HN0002.THA",
  "Thai Ha (center Lang Ha)": "HN0002.THA",
  "Hoang Quoc Viet": "HN0003.HQV",
  "Lieu Giai": "HN0004.LGI",
  "Nguyen Van Linh": "HN0005.NVL",
  "Van Quan": "HN0007.VQN",
  "My Dinh": "HN0010.MDH",
  "The Garden": "HN0010.MDH",
  "Hoang Mai": "HN0012.NHT",
  "Nguyen Huu Tho": "HN0012.NHT",
  "Tan Mai": "HN0014.TMI",
  "Van Phu": "HN0015.VPU",
  "Phan Dinh Phung": "HN0016.PDP",
  "Ham Nghi": "HN0017.HNI",
  "Vu Tong Phan": "HN0018.VTP",
  "Nguyen Tuan": "HN0019.NTN",
  "Ngoai Giao Doan": "HN0021.NGD",
  "Nguyen Van Loc": "HN0022.NVO",
  "Mo Lao": "HN0022.NVO",
  "Linh Dam": "HN0023.LDM",
  "TIMES CITY": "HN0024.TCY",
  "Le Trong Tan": "HN0025.LTT",
  "Viet Hung": "HN0026.VHG",
  "Ocepark": "HN0027.OPK",
  "Ocean Park": "HN0027.OPK",
  "Pham Van Dong": "HN0028.PVD",
  "Vu Pham Ham": "HN0029.VPH",
  "An Khanh": "HN0030.AKH",
  "An Hung": "HN0031.AHG",
  "Xuan Dieu (đổi thành Lạc Long Quân)": "HN0032.LLQ",
  "Xuan Dieu": "HN0032.LLQ",
  "Lac Long Quan": "HN0032.LLQ",
  "HN33.DAH": "HN0033.DAH",
  "Dong Anh": "HN0033.DAH",
  "HN34.HTN": "HN0034.HTN",
  "Hong Tien": "HN0034.HTN",
  "Ecopark": "HY0001.ECP",
  "Hai Phong": "Hai Phong",
  "Hai Phong 1": "HP0001.LHP",
  "HP1.LHP": "HP0001.LHP",
  "Hai Phong 2": "HP0002.HBT",
  "Hai Phong 3": "HP0003.VIN",
  "Ha Long": "QN0001.HLG",
  "Quang Ninh": "QN0001.HLG",
  "QN01.HL": "QN0001.HLG",
  "Vinh": "VIN001.CTG",
  "VIN01.CT": "VIN001.CTG",
  "Vinh Phuc": "VP0001.PCT",
  "TH01.TPU": "TH0001.TPU",
  "Thanh Hoa": "TH0001.TPU",
  "TN01.LNQ": "TN0001.LNQ",
  "Thai Nguyen": "TN0001.LNQ",
  "PT01.HVG": "PT0001.HVG",
  "Phu Tho": "PT0001.HVG",
  "Apollo Advance -South": "AA",
  "ASP - HN": "HN0200.ASP",
  "MKT LOCAL NORTH": "MKT LOCAL NORTH",
  "Cambridge": "CAMBRIDGE",
  "MKT HP": "MKT LOCAL NORTH_HP",
  "MKT TN01.LNQ": "MKT LOCAL NORTH_TN",
  "MKT PT01.HVG": "MKT LOCAL NORTH_PT",
  "MKT TH01.TPU": "MKT LOCAL NORTH_TH",
  "NTW": "NTW",
  "Contest": "CONTEST",
  "Job Fair": "JOB FAIR"
};

function extractBankName(fileName: string, bankLabel?: string) {
  if (bankLabel) {
    const upperBank = bankLabel.toUpperCase();
    if (upperBank.includes("MKT")) return "MKT LOCAL";
  }
  const name = fileName.toUpperCase().replace(/\.[^/.]+$/, "");
  const tokens = name.replace(/[^A-Z0-9]/g, ' ').split(/\s+/);
  
  if (tokens.includes('MKT')) return 'MKT LOCAL';
  if (tokens.includes('TH')) return 'TH';
  if (tokens.includes('HP')) return 'HP';
  if (tokens.includes('TN')) return 'TN';
  if (tokens.includes('PT')) return 'PT';
  if (tokens.includes('NORTH')) return 'NORTH';
  
  return 'NORTH';
}

function processNorthLogic(rawCenter: string) {
  const specialCenter = resolveMasterSpecialCenter(rawCenter);
  if (specialCenter) return { chargeToCenterMkt: "", l07: specialCenter.l07, bu: specialCenter.business };
  const cleaned = rawCenter ? String(rawCenter).trim() : "";
  let l07 = cleaned;

  for (const [key, value] of Object.entries(aeCodeToL07Map)) {
    if (key.toUpperCase() === cleaned.toUpperCase()) {
      l07 = value;
      break;
    }
  }

  if (l07 === cleaned) {
    const upperClean = cleaned.toUpperCase();
    if (upperClean.includes("THAI HA") || upperClean.includes("THÁI HÀ")) l07 = "HN0002.THA";
    else if (upperClean.includes("XUAN DIEU") || upperClean.includes("XUÂN DIỆU") || upperClean.includes("LAC LONG QUAN") || upperClean.includes("LẠC LONG QUÂN")) l07 = "HN0032.LLQ";
    else if (upperClean.includes("OCEAN PARK") || upperClean.includes("OCEPARK")) l07 = "HN0027.OPK";
  }

  let bu = "OTHER";
  if (l07 === "AA" || l07 === "ZHN0000.GY" || l07 === "HN0200.ASP" || l07.startsWith("HN") || l07.startsWith("BN") || l07.startsWith("HY") || l07.startsWith("QN") || l07.startsWith("VIN") || l07.startsWith("VP") || l07 === "MKT LOCAL NORTH") {
    bu = "AHN";
  } else if (l07.startsWith("HP") || l07.toUpperCase() === "HAI PHONG" || l07 === "MKT LOCAL NORTH_HP") {
    bu = "AHP";
  } else if (l07.startsWith("TN") || l07 === "MKT LOCAL NORTH_TN") {
    bu = "ATN";
  } else if (l07.startsWith("TH") || l07 === "MKT LOCAL NORTH_TH") {
    bu = "ATH";
  } else if (l07.startsWith("PT") || l07 === "MKT LOCAL NORTH_PT") {
    bu = "APT";
  }

  return { chargeToCenterMkt: "", l07, bu };
}

function processTimesheetMktLogic(inputData: { chargetocenterCode: string }) {
  const resolved = resolveMktRosterCenter(inputData.chargetocenterCode);
  return {
    chargeToCenterMkt: resolved.chargeToCenterMkt,
    l07: resolved.l07,
    bu: resolved.business || "OTHER",
  };
}

function parseMonthFromFileName(fileName: string, globalMonth?: string): string | null {
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
    let y = currentYearNum;
    if (m === 11 || m === 12) {
      y = 2025;
    } else if (m > currentMonthNum) {
      y = currentYearNum - 1;
    }
    return `${m < 10 ? "0" + m : m}.${y}`;
  }
  return null;
}

// ==========================================
// MAIN PIVOT SHEET REACT COMPONENT
// ==========================================

export function PivotSheet() {
  const { appData, updateAppData, isLoading: isAppDataLoading } = useAppData();
  const [isProcessing, setIsProcessing] = useState(false);
  const [selectedMonthFilter, setSelectedMonthFilter] = useState<string>(() => {
    try {
      const saved = localStorage.getItem("pivot_master_selected_month_filter");
      if (saved) return normalizePivotMonth(saved);
    } catch {
      // ignore
    }
    return "ALL";
  });

  const [currentPage, setCurrentPage] = useState<number>(1);
  const [buFilterMode, setBuFilterMode] = useState<string>("ALL");

  const handleSetBuFilterMode = useCallback((mode: string) => {
    setBuFilterMode(mode);
    setCurrentPage(1);
  }, []);

  const [groupedData, setGroupedData] = useState<Record<string, Record<string, Record<string, Record<string, number>>>>>(() => {
    try {
      const cached = localStorage.getItem("pivot_master_processed_data");
      if (cached) {
        const parsed = JSON.parse(cached);
        return parsed.cacheVersion === PIVOT_CACHE_VERSION
          ? parsed.groupedData || {}
          : {};
      }
    } catch (e) {
      console.warn("Error reading pivot cache", e);
    }
    return {};
  });

  const [typeColumns, setTypeColumns] = useState<string[]>(() => {
    try {
      const cached = localStorage.getItem("pivot_master_processed_data");
      if (cached) {
        const parsed = JSON.parse(cached);
        return parsed.cacheVersion === PIVOT_CACHE_VERSION
          ? parsed.typeColumns || []
          : [];
      }
    } catch {
      // ignore cache error
    }
    return [];
  });

  const [diagnosticLogs, setDiagnosticLogs] = useState<any[]>(() => {
    try {
      const cached = localStorage.getItem("pivot_master_processed_data");
      if (cached) {
        const parsed = JSON.parse(cached);
        return parsed.cacheVersion === PIVOT_CACHE_VERSION
          ? parsed.diagnosticLogs || []
          : [];
      }
    } catch {
      // ignore cache error
    }
    return [];
  });

  const [_sourceInfo, _setSourceInfo] = useState<string>(() => {
    try {
      const cached = localStorage.getItem("pivot_master_processed_data");
      if (cached) {
        const parsed = JSON.parse(cached);
        return parsed.cacheVersion === PIVOT_CACHE_VERSION
          ? parsed.sourceInfo || ""
          : "";
      }
    } catch {
      // ignore cache error
    }
    return "";
  });
  const [isSettingsOpen, setIsSettingsOpen] = useState(false);
  const [rowsPerPage, setRowsPerPage] = useState<number>(() => {
    try {
      const savedValue = localStorage.getItem("pivot_master_rows_per_page");
      if (savedValue === "all") return Infinity;
      const parsedValue = Number(savedValue);
      return Number.isFinite(parsedValue) && parsedValue > 0
        ? parsedValue
        : 50;
    } catch {
      return 50;
    }
  });
  const [hiddenColumns, setHiddenColumns] = useState<Record<string, boolean>>({ month: true });
  const [sortField, setSortField] = useState<string | null>(null);
  const [sortDirection, setSortDirection] = useState<"asc" | "desc">("asc");

  const safeTypeColumns = useMemo(() => (Array.isArray(typeColumns) ? typeColumns : []), [typeColumns]);
  const safeGroupedData = useMemo(() => omitPivotMonths(
    groupedData && typeof groupedData === "object" ? groupedData : {},
    appData.PivotMasterClearedMonths || [],
  ), [groupedData, appData.PivotMasterClearedMonths]);

  const [columnWidths, setColumnWidths] = useState<Record<string, number>>(() => {
    try {
      const cached = localStorage.getItem("pivot_master_column_widths");
      if (cached) return JSON.parse(cached);
    } catch {
      // ignore
    }
    return {
      no: 50,
      month: 90,
      business: 90,
      charge: 220,
      grandTotal: 140,
    };
  });

  const resizingColRef = useRef<{ colKey: string; startX: number; startWidth: number } | null>(null);

  const handleResizeStart = (e: React.MouseEvent, colKey: string, defaultW = 120) => {
    e.preventDefault();
    e.stopPropagation();
    const startX = e.clientX;
    const startWidth = columnWidths[colKey] || defaultW;
    resizingColRef.current = { colKey, startX, startWidth };
    let latestWidth = startWidth;
    let resizeFrame: number | null = null;

    const handleMouseMove = (moveEvent: MouseEvent) => {
      if (!resizingColRef.current) return;
      const targetColKey = resizingColRef.current.colKey;
      const deltaX = moveEvent.clientX - resizingColRef.current.startX;
      latestWidth = Math.max(45, resizingColRef.current.startWidth + deltaX);
      if (resizeFrame !== null) return;
      resizeFrame = window.requestAnimationFrame(() => {
        resizeFrame = null;
        setColumnWidths((prev) =>
          targetColKey ? { ...prev, [targetColKey]: latestWidth } : prev,
        );
      });
    };

    const handleMouseUp = () => {
      const targetColKey = resizingColRef.current?.colKey;
      if (resizeFrame !== null) {
        window.cancelAnimationFrame(resizeFrame);
        resizeFrame = null;
      }
      if (targetColKey) {
        setColumnWidths((prev) => {
          const next = { ...prev, [targetColKey]: latestWidth };
          try {
            localStorage.setItem(
              "pivot_master_column_widths",
              JSON.stringify(next),
            );
          } catch {
            // Ignore local storage quota/privacy errors.
          }
          return next;
        });
      }
      resizingColRef.current = null;
      window.removeEventListener("mousemove", handleMouseMove);
      window.removeEventListener("mouseup", handleMouseUp);
    };

    window.addEventListener("mousemove", handleMouseMove);
    window.addEventListener("mouseup", handleMouseUp);
  };

  const [editingCell, setEditingCell] = useState<{
    bu: string;
    l07: string;
    month: string;
    field: string;
  } | null>(null);

  const [editValue, setEditValue] = useState<string>("");

  const handleStartEdit = (bu: string, l07: string, month: string, field: string, currentValue: any) => {
    setEditingCell({ bu, l07, month, field });
    setEditValue(currentValue === undefined || currentValue === null ? "" : String(currentValue));
  };

  const saveToCache = (newGroupedData: any, newTypeColumns = safeTypeColumns) => {
    try {
      localStorage.setItem("pivot_master_processed_data", JSON.stringify({
        cacheVersion: PIVOT_CACHE_VERSION,
        groupedData: newGroupedData,
        typeColumns: newTypeColumns,
        diagnosticLogs,
        sourceInfo: _sourceInfo,
        filter: selectedMonthFilter,
        reportingMonth: appData.globalMonth || "03.2026",
        updatedAt: Date.now()
      }));
    } catch (e) {
      console.warn("Failed saving pivot data to cache", e);
    }
  };

  const handleSaveEdit = () => {
    if (!editingCell) return;
    const { bu, l07, month, field } = editingCell;

    setGroupedData(prev => {
      const nextData = JSON.parse(JSON.stringify(prev));
      if (!nextData[bu] || !nextData[bu][l07] || !nextData[bu][l07][month]) return prev;

      if (field === "bu") {
        const newBu = editValue.trim().toUpperCase() || "UNKNOWN";
        if (newBu !== bu) {
          if (!nextData[newBu]) nextData[newBu] = {};
          if (!nextData[newBu][l07]) nextData[newBu][l07] = {};
          nextData[newBu][l07][month] = nextData[bu][l07][month];
          delete nextData[bu][l07][month];
          if (Object.keys(nextData[bu][l07]).length === 0) delete nextData[bu][l07];
          if (Object.keys(nextData[bu]).length === 0) delete nextData[bu];
        }
      } else if (field === "l07") {
        const newL07 = editValue.trim() || "UNKNOWN";
        if (newL07 !== l07) {
          if (!nextData[bu][newL07]) nextData[bu][newL07] = {};
          nextData[bu][newL07][month] = nextData[bu][l07][month];
          delete nextData[bu][l07][month];
          if (Object.keys(nextData[bu][l07]).length === 0) delete nextData[bu][l07];
        }
      } else if (field === "month") {
        const newMonth = editValue.trim() || "UNKNOWN";
        if (newMonth !== month) {
          nextData[bu][l07][newMonth] = nextData[bu][l07][month];
          delete nextData[bu][l07][month];
        }
      } else {
        const rawNum = editValue.replace(/,/g, "").trim();
        const numVal = parseFloat(rawNum);
        const finalVal = isNaN(numVal) ? 0 : numVal;
        nextData[bu][l07][month][field] = finalVal;
      }

      saveToCache(nextData);
      return nextData;
    });

    setEditingCell(null);
    setEditValue("");
    toast.success("Đã cập nhật dữ liệu ô Pivot Master");
  };

  const handleCancelEdit = () => {
    setEditingCell(null);
    setEditValue("");
  };

  const handleAddRow = () => {
    const defaultBU = "AHN";
    const defaultL07 = `CENTER_${Date.now().toString().slice(-4)}`;
    const defaultMonth = appData.globalMonth || "03.2026";
    setGroupedData(prev => {
      const nextData = JSON.parse(JSON.stringify(prev));
      if (!nextData[defaultBU]) nextData[defaultBU] = {};
      if (!nextData[defaultBU][defaultL07]) nextData[defaultBU][defaultL07] = {};
      nextData[defaultBU][defaultL07][defaultMonth] = {};
      safeTypeColumns.forEach(t => {
        nextData[defaultBU][defaultL07][defaultMonth][t] = 0;
      });
      saveToCache(nextData);
      return nextData;
    });
    toast.success(`Đã thêm dòng mới với Center/L07: ${defaultL07}`);
  };

  const handleAddColumn = () => {
    const colName = window.prompt("Nhập tên cột mới:");
    if (!colName) return;
    const trimmed = colName.trim();
    if (!trimmed) return;
    
    if (safeTypeColumns.includes(trimmed)) {
      toast.error(`Cột "${trimmed}" đã tồn tại!`);
      return;
    }
    
    setTypeColumns(prev => {
      const next = [...prev, trimmed];
      try {
        const cached = localStorage.getItem("pivot_master_processed_data");
        if (cached) {
          const parsed = JSON.parse(cached);
          parsed.cacheVersion = PIVOT_CACHE_VERSION;
          parsed.typeColumns = next;
          localStorage.setItem("pivot_master_processed_data", JSON.stringify(parsed));
        }
      } catch {
        // ignore
      }
      return next;
    });
    
    toast.success(`Đã thêm cột mới: ${trimmed}`);
  };

  const [deleteRowTarget, setDeleteRowTarget] = useState<{ bu: string; l07: string; month: string } | null>(null);

  const handleDeleteRow = (bu: string, l07: string, month: string) => {
    setDeleteRowTarget({ bu, l07, month });
  };

  const confirmDeleteRow = () => {
    if (!deleteRowTarget) return;
    const { bu, l07, month } = deleteRowTarget;
    setGroupedData(prev => {
      const nextData = JSON.parse(JSON.stringify(prev));
      if (nextData[bu] && nextData[bu][l07] && nextData[bu][l07][month]) {
        delete nextData[bu][l07][month];
        if (Object.keys(nextData[bu][l07]).length === 0) delete nextData[bu][l07];
        if (Object.keys(nextData[bu]).length === 0) delete nextData[bu];
      }
      saveToCache(nextData);
      return nextData;
    });
    toast.success(`Đã xóa dòng ${bu} - ${l07}`);
    setDeleteRowTarget(null);
  };
  
  const [deleteMonthTarget, setDeleteMonthTarget] = useState<string | null>(null);

  const restorePivotMonths = useCallback((months: readonly string[]) => {
    const reopened = new Set(months.map(normalizePivotMonth));
    updateAppData(previous => {
      const cleared = previous.PivotMasterClearedMonths || [];
      const remaining = cleared.filter(month => !reopened.has(normalizePivotMonth(month)));
      return remaining.length === cleared.length ? previous : { ...previous, PivotMasterClearedMonths: remaining };
    }, false, true);
  }, [updateAppData]);

  const confirmDeleteMonth = () => {
    if (!deleteMonthTarget || isProcessing || isAppDataLoading) return;
    const month = deleteMonthTarget;
    updateAppData(previous => ({
      ...previous,
      PivotMasterClearedMonths: [...new Set([...(previous.PivotMasterClearedMonths || []).map(normalizePivotMonth), month])],
    }), false, true);
    const remaining = omitPivotMonths(groupedData, [month]);
    writePivotMktTypeCache(updatePivotMktTypeCache(
      readPivotMktTypeCache(groupedData, safeTypeColumns), {}, [], [month],
    ));
    saveToCache(remaining);
    setGroupedData(remaining);
    setEditingCell(null);
    setCurrentPage(1);
    setDeleteMonthTarget(null);
    toast.success(`Đã xóa dữ liệu Pivot Master tháng ${month}.`);
  };

  const settingsMenuRef = useRef<HTMLDivElement>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    function handleClickOutside(event: MouseEvent) {
      if (settingsMenuRef.current && !settingsMenuRef.current.contains(event.target as Node)) {
        setIsSettingsOpen(false);
      }
    }
    document.addEventListener("mousedown", handleClickOutside);
    return () => document.removeEventListener("mousedown", handleClickOutside);
  }, []);

  useEffect(() => {
    setCurrentPage(1);
  }, [rowsPerPage, groupedData, typeColumns]);

  useEffect(() => {
    try {
      localStorage.setItem(
        "pivot_master_rows_per_page",
        rowsPerPage === Infinity ? "all" : String(rowsPerPage),
      );
    } catch {
      // ignore pagination preference write error
    }
  }, [rowsPerPage]);

  useEffect(() => {
    const handlePivotDataUpdated = (event: Event) => {
      const detail = (event as CustomEvent).detail || {};
      if (!detail.groupedData || typeof detail.groupedData !== "object") return;

      const incomingTypeColumns = Array.isArray(detail.typeColumns)
        ? detail.typeColumns
        : [];
      const mktTypeCache = readPivotMktTypeCache(
        detail.groupedData,
        incomingTypeColumns,
      );
      const restoredGroupedData = applyPivotMktTypeCache(
        detail.groupedData,
        mktTypeCache,
      );
      const restoredTypeColumns = mergePivotTypeColumns(
        incomingTypeColumns,
        mktTypeCache.typeColumns,
      );

      setGroupedData(restoredGroupedData);
      setTypeColumns(restoredTypeColumns);
      setDiagnosticLogs(
        Array.isArray(detail.diagnosticLogs) ? detail.diagnosticLogs : [],
      );
      _setSourceInfo(String(detail.sourceInfo || ""));
    };

    window.addEventListener("pivot-data-updated", handlePivotDataUpdated);
    return () =>
      window.removeEventListener("pivot-data-updated", handlePivotDataUpdated);
  }, []);

  const processFileBuffers = useCallback(async (fileList: { name: string; bank?: string; buffer: ArrayBuffer; month: string }[]) => {
    const newGroupedData: Record<string, Record<string, Record<string, Record<string, number>>>> = {};
    const mktGroupedData: PivotGroupedData = {};
    const uniqueTypes = new Set<string>();
    const mktTypeColumns = new Set<string>();
    const mktMonths = new Set<string>();
    let processedMktFiles = 0;
    const newLogs: any[] = [];

    // Within one month the last valid MKT workbook replaces earlier versions.
    for (const item of [...fileList].reverse()) {
      try {
        const displayBankName = extractBankName(item.name, item.bank);
        let processType = (displayBankName === 'MKT LOCAL') ? "MKT LOCAL NORTH" : "NORTH";
        
        const workbook = XLSX.read(item.buffer, { type: "array" });
        let targetSheetName = "";

        const rosterSheet = findMasterRosterSheetName(
          workbook.SheetNames.filter(name => workbook.Sheets[name]?.["!ref"]),
        );

        if (processType === "MKT LOCAL NORTH" || rosterSheet) {
          processType = "MKT LOCAL NORTH";
          targetSheetName = rosterSheet || workbook.SheetNames[0];
        } else {
          targetSheetName = workbook.SheetNames.find(n => {
            const normName = (n || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/đ/g, "d").replace(/Đ/g, "D").toUpperCase().trim().replace(/\s+/g, " ");
            return normName === 'SHEET 1' || normName.includes('SHEET 1') || normName.includes('INTERN') || normName.includes('REPORT');
          }) || workbook.SheetNames.find(n => {
            const normName = (n || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/đ/g, "d").replace(/Đ/g, "D").toUpperCase().trim().replace(/\s+/g, " ");
            return normName.includes('DATA') || normName.includes('DU LIEU') || n.toUpperCase().includes('DỮ LIỆU');
          }) || workbook.SheetNames[0];
        }

        const worksheet = workbook.Sheets[targetSheetName];
        if (!worksheet) continue;

        const jsonData = XLSX.utils.sheet_to_json<any[]>(worksheet, { header: 1 });
        if (jsonData.length === 0 || (processType === "NORTH" && jsonData.length === 1)) continue;

        let headerRowIdx = 0;
        for (let r = 0; r < Math.min(15, jsonData.length); r++) {
          const row = jsonData[r];
          if (!row || row.length === 0) continue;
          const rowStr = row.join(' ').toUpperCase();
          if (rowStr.includes('CENTER') || rowStr.includes('CHARGE') || rowStr.includes('TYPE') || rowStr.includes('MÃ TT')) {
            headerRowIdx = r;
            break;
          }
        }

        const headers = jsonData[headerRowIdx];
        if (!headers) continue;

        if (processType === "NORTH") {
          const centerColIdx = headers.findIndex((h: any) => {
            if (!h) return false;
            const val = String(h).trim().toUpperCase();
            return val === 'CENTER' || val === 'CENTERS' || val === 'CENTER CODE' || val === 'MÃ TT' || val.includes('TRUNG TÂM');
          });

          const bankAccColIdx = headers.findIndex((h: any) => {
            if (!h) return false;
            const val = String(h).trim().toUpperCase();
            return val === 'BANK ACCOUNT NUMBER' || val.includes('BANK ACCOUNT');
          });

          if (centerColIdx === -1 || bankAccColIdx === -1) continue;

          const chargeCols: { index: number; label: string }[] = [];
          const seenLabels = new Set<string>();
          headers.forEach((h: any, idx: number) => {
            if (h) {
              const strH = String(h).trim();
              const uH = strH.toUpperCase();
              if (
                uH.includes("CENTER") ||
                uH.includes("TRUNG TÂM") ||
                uH.includes("NOTE") ||
                uH.includes("STATUS") ||
                uH.includes("ACCOUNT") ||
                uH.includes("NAME") ||
                uH.includes("TOTAL") ||
                uH.includes("CODE") ||
                uH.includes("THÁNG")
              ) return;
              if (uH.includes("CHARGE") || uH === "LXO" || uH === "EC" || uH === "PT-DEMO" || uH.includes("EXTRA") || uH.includes("INSTRUCTOR") || uH.includes("SUMMER") || uH.includes("BONUS")) {
                const label = formatPivotTypeHeader(strH);
                if (label && label !== "EXCLUDE" && !seenLabels.has(label)) {
                  seenLabels.add(label);
                  chargeCols.push({ index: idx, label });
                  uniqueTypes.add(label);
                }
              }
            }
          });

          for (let i = headerRowIdx + 1; i < jsonData.length; i++) {
            const row = jsonData[i];
            if (!row || row.length === 0) continue;

            const bankAccVal = row[bankAccColIdx];
            if (bankAccVal === undefined || bankAccVal === null || String(bankAccVal).trim() === "") {
              continue;
            }

            const rawCenter = row[centerColIdx];
            if (!rawCenter) continue;

            if (String(rawCenter).toUpperCase().includes("MKT")) {
              continue;
            }

            const mapped = processNorthLogic(String(rawCenter));
            const { bu, l07 } = mapped;

            if (!newGroupedData[bu]) newGroupedData[bu] = {};
            if (!newGroupedData[bu][l07]) newGroupedData[bu][l07] = {};
            if (!newGroupedData[bu][l07][item.month]) newGroupedData[bu][l07][item.month] = {};
            markPivotZhnSource(newGroupedData[bu][l07][item.month], l07, rawCenter);

            chargeCols.forEach(col => {
              const rawVal = row[col.index];
              const val = parseExcelMoney(rawVal);

              if (val === 0 && rawVal) {
                newLogs.push({ Source: "NORTH", File: item.name, RawCenter: rawCenter, Column: col.label, RawValue: rawVal });
              }

              if (!newGroupedData[bu][l07][item.month][col.label]) {
                newGroupedData[bu][l07][item.month][col.label] = 0;
              }
              newGroupedData[bu][l07][item.month][col.label] += val;
            });
          }
        } else if (processType === "MKT LOCAL NORTH") {
          if (mktMonths.has(item.month)) continue;
          const centerColIdx = headers.findIndex((h: any) => {
            if (!h) return false;
            const val = String(h).trim().toUpperCase();
            return (
              val === 'CHARGE TO CENTER' ||
              val === 'CHARGE TO CENTER MKT' ||
              val === 'CHARGETOCENTER' ||
              val === 'CHARGETOCENTERCODE' ||
              val.includes('CHARGE TO CENTER') ||
              val === 'CENTER' ||
              val === 'CENTERS' ||
              val === 'MÃ TT' ||
              val === 'TRUNG TÂM' ||
              val === 'CƠ SỞ'
            );
          });
          const allocationColIdx = headers.findIndex((header: unknown) =>
            /^CHARGE\s*TO\s*CENTER/i.test(String(header || "").trim()),
          );
          const classColIdx = headers.findIndex((header: unknown) => String(header || "").trim().toUpperCase() === "CLASS");
          const typeColIdx = headers.findIndex((h: any) => {
            if (!h) return false;
            const val = String(h).trim().toUpperCase();
            return (
              val === 'TYPE' ||
              val === 'CODE' ||
              val === 'TASK TYPE' ||
              val === 'TYPE CODE' ||
              val === 'LOẠI' ||
              val === 'PHÂN LOẠI' ||
              val === 'MÃ' ||
              val.includes('TYPE') ||
              val.includes('CODE') ||
              val.includes('LOẠI')
            );
          });
          const durationColIdx = headers.findIndex((h: any) => {
            if (!h) return false;
            const val = String(h).trim().toUpperCase();
            return (
              val === 'DURATION' ||
              val === 'HOURS' ||
              val === 'HOUR' ||
              val === 'SỐ GIỜ' ||
              val === 'GIỜ' ||
              val === 'TOTAL HOURS' ||
              val.includes('DURATION') ||
              val.includes('HOURS') ||
              val.includes('GIỜ')
            );
          });

          const totalColIdx = headers.findIndex((h: unknown) => {
            const label = String(h || "").trim().toUpperCase();
            return label.includes("TOTAL") || label.includes("THỰC NHẬN") ||
              label.includes("LƯƠNG") || label.includes("SỐ TIỀN");
          });
          if (centerColIdx === -1 || (durationColIdx === -1 && totalColIdx === -1)) continue;
          processedMktFiles += 1;
          mktMonths.add(item.month);

          for (let i = headerRowIdx + 1; i < jsonData.length; i++) {
            const row = jsonData[i];
            if (!row || row.length === 0) continue;

            const sourceCenter = row[centerColIdx];
            const resolvedSource = resolveMktRosterCenter(String(sourceCenter || "")).l07;
            const rawCenter = row[allocationColIdx] ||
              (resolvedSource === "MKT LOCAL NORTH" && row[classColIdx]) ||
              sourceCenter;
            if (!rawCenter) continue;

            const mapped = processTimesheetMktLogic({ chargetocenterCode: String(rawCenter) });
            const { bu, l07 } = mapped;

            // Bỏ qua MKT LOCAL NORTH và các trạm con của nó vì lương đã được tách vào LDEM/LRET/LDEC...
            if (!l07 || l07.toUpperCase().includes("MKT LOCAL NORTH")) continue;

            let durationHours = 0;
            if (durationColIdx !== -1 && row[durationColIdx] !== undefined) {
              durationHours = parseDurationToHours(row[durationColIdx]);
            }
            let calculatedSalary = durationHours * 20000;

            if (calculatedSalary === 0 && totalColIdx !== -1) {
              calculatedSalary = parseExcelMoney(row[totalColIdx]);
            }

            if (calculatedSalary === 0) continue;

            const rawType = (typeColIdx !== -1 && row[typeColIdx] !== undefined)
              ? String(row[typeColIdx]).trim()
              : "MKT LOCAL";

            const typeVal = formatPivotTypeHeader(rawType || "UNSPECIFIED");
            if (typeVal === "EXCLUDE") continue;
            uniqueTypes.add(typeVal);
            mktTypeColumns.add(typeVal);

            if (!newGroupedData[bu]) newGroupedData[bu] = {};
            if (!newGroupedData[bu][l07]) newGroupedData[bu][l07] = {};
            if (!newGroupedData[bu][l07][item.month]) newGroupedData[bu][l07][item.month] = {};

            if (!newGroupedData[bu][l07][item.month][typeVal]) {
              newGroupedData[bu][l07][item.month][typeVal] = 0;
            }
            newGroupedData[bu][l07][item.month][typeVal] += calculatedSalary;

            if (!mktGroupedData[bu]) mktGroupedData[bu] = {};
            if (!mktGroupedData[bu][l07]) mktGroupedData[bu][l07] = {};
            if (!mktGroupedData[bu][l07][item.month]) {
              mktGroupedData[bu][l07][item.month] = {};
            }
            if (!mktGroupedData[bu][l07][item.month][typeVal]) {
              mktGroupedData[bu][l07][item.month][typeVal] = 0;
            }
            mktGroupedData[bu][l07][item.month][typeVal] += calculatedSalary;
          }
        }
      } catch (err) {
        console.error("Error processing file buffer:", item.name, err);
      }
    }

    let unspecifiedTotal = 0;
    for (const bu in newGroupedData) {
      for (const l07 in newGroupedData[bu]) {
        for (const month in newGroupedData[bu][l07]) {
          unspecifiedTotal += newGroupedData[bu][l07][month]["UNSPECIFIED"] || 0;
        }
      }
    }
    if (unspecifiedTotal === 0) {
      uniqueTypes.delete("UNSPECIFIED");
    }

    return {
      groupedData: newGroupedData,
      typeColumns: Array.from(uniqueTypes).sort(),
      mktGroupedData,
      mktTypeColumns: Array.from(mktTypeColumns).sort(),
      mktMonths: Array.from(mktMonths).sort(),
      processedMktFiles,
      logs: newLogs,
    };
  }, []);

  const loadMasterData = useCallback(async (showToastMsg = false, restoreOriginal = false) => {
    const cachedStr = restoreOriginal ? null : localStorage.getItem("pivot_master_processed_data");
    let cachedGroupedData: PivotGroupedData = {};
    let cachedTypeColumns: string[] = [];
    let resetMonths: string[] = [];
    if (cachedStr) {
      try {
        const parsed = JSON.parse(cachedStr);
        if (parsed.cacheVersion === PIVOT_CACHE_VERSION) {
          cachedGroupedData = parsed.groupedData || {};
          resetMonths = Array.isArray(parsed.resetMonths) ? parsed.resetMonths : [];
          cachedTypeColumns = Array.isArray(parsed.typeColumns)
            ? parsed.typeColumns
            : [];
        }
        if (
          !showToastMsg && resetMonths.length === 0 &&
          parsed.cacheVersion === PIVOT_CACHE_VERSION &&
          parsed.reportingMonth === (appData.globalMonth || "03.2026") &&
          parsed.groupedData &&
          Object.keys(parsed.groupedData).length > 0
        ) {
          const cleanCachedGrouped: PivotGroupedData = {};
          Object.entries(parsed.groupedData || {}).forEach(([bu, l07Rows]: [string, any]) => {
            Object.entries(l07Rows || {}).forEach(([l07, monthRows]: [string, any]) => {
              const uL07 = String(l07).trim().toUpperCase();
              if (
                uL07.includes("MKT LOCAL NORTH") ||
                uL07.startsWith("MKT LOCAL") ||
                uL07.includes("MKT_LOCAL")
              ) {
                return;
              }
              if (!cleanCachedGrouped[bu]) cleanCachedGrouped[bu] = {};
              cleanCachedGrouped[bu][l07] = monthRows;
            });
          });
          setGroupedData(cleanCachedGrouped);
          setTypeColumns(parsed.typeColumns || []);
          setDiagnosticLogs(parsed.diagnosticLogs || []);
          _setSourceInfo(parsed.sourceInfo || "Đã tải từ cache");
          return;
        }
      } catch {
        // ignore cache parse error
      }
    }

    setIsProcessing(true);

    try {
      const processedSheet1 = (appData.Sheet1_AE?.data || []).filter((r: any) => String(r["ID Number"] || r["id_number"] || "").trim() !== "");
      const processedRoster = (appData.Master_Roster || []).filter((row: any) => {
        const sourceFile = String(row?._sourceFile || "").trim().toUpperCase();
        const rowId = String(row?._rowId || "").trim().toLowerCase();
        return sourceFile !== "MOCK_ROSTER.XLSX" && !rowId.startsWith("mock-row-");
      });

      // Ưu tiên dữ liệu đã được bảng Master xử lý để đồng bộ nhanh và không
      // parse lại hàng loạt file Excel trên main thread.
      if (processedSheet1.length > 0 || processedRoster.length > 0) {
        const baseResult = buildPivotFromAppData(
          processedSheet1,
          [],
          [],
          appData.globalMonth || "03.2026",
        );
        const rosterResult = buildPivotFromAppData(
          [],
          [],
          processedRoster,
          appData.globalMonth || "03.2026",
        );
        let mktTypeCache = readPivotMktTypeCache(
          cachedGroupedData,
          cachedTypeColumns,
        );

        // Chỉ thay snapshot TYPE khi Master thực sự có dữ liệu MKT Local.
        // Nếu đợt tải Master thiếu file MKT, cache cũ được giữ nguyên.
        if (processedRoster.length > 0) {
          mktTypeCache = updatePivotMktTypeCache(
            mktTypeCache,
            rosterResult.groupedData || {},
            rosterResult.typeColumns || [],
            getPivotDataMonths(rosterResult.groupedData || {}).filter(month =>
              resetMonths.length ? resetMonths.includes(month) : restoreOriginal || !mktTypeCache.manualMonths?.includes(month),
            ),
          );
          writePivotMktTypeCache(mktTypeCache);
        }

        const rebuiltGroupedData = applyPivotMktTypeCache(
          baseResult.groupedData || {},
          mktTypeCache,
        );
        const restoredGroupedData = resetMonths.length
          ? resetMonths.reduce((current, month) => replacePivotCacheMonth(current, rebuiltGroupedData, month), cachedGroupedData)
          : rebuiltGroupedData;
        const restoredTypeColumns = mergePivotTypeColumns(
          resetMonths.length ? [...cachedTypeColumns, ...(baseResult.typeColumns || [])] : baseResult.typeColumns || [],
          mktTypeCache.typeColumns,
        );

        setGroupedData(restoredGroupedData);
        setTypeColumns(restoredTypeColumns);
        setDiagnosticLogs([]);
        const infoStr = `Đồng bộ từ Master (${processedSheet1.length} dòng Gross Pay, ${processedRoster.length} dòng MKT Local; đã giữ TYPE MKT đã lưu)`;
        _setSourceInfo(infoStr);

        try {
          localStorage.setItem("pivot_master_processed_data", JSON.stringify({
            cacheVersion: PIVOT_CACHE_VERSION,
            groupedData: restoredGroupedData,
            typeColumns: restoredTypeColumns,
            diagnosticLogs: [],
            sourceInfo: infoStr,
            filter: selectedMonthFilter,
            reportingMonth: appData.globalMonth || "03.2026",
            updatedAt: Date.now(),
          }));
        } catch {
          // ignore cache write error
        }

        if (showToastMsg) {
          if (getPivotDataMonths(restoredGroupedData).includes(selectedMonthFilter)) restorePivotMonths([selectedMonthFilter]);
          toast.success(
            `Đã đồng bộ Pivot Master từ dữ liệu đã xử lý trong bảng Master`,
          );
        }
        return;
      }

      const masterRows = appData.Ae_Global_Inputs || [];
      const fileBuffers: { name: string; bank?: string; buffer: ArrayBuffer; month: string }[] = [];

      for (const row of masterRows) {
        const rowMonth =
          row.month ||
          row.fileMonth ||
          parseMonthFromFileName(row.fileName || row.name || "", appData.globalMonth) ||
          appData.globalMonth ||
          "03.2026";

        if (row.fileObj && row.fileObj instanceof File) {
          try {
            const buffer = await row.fileObj.arrayBuffer();
            fileBuffers.push({ name: row.fileName || row.fileObj.name, bank: row.bank, buffer, month: normalizePivotMonth(rowMonth) });
          } catch (fileErr) {
            console.warn(`Không thể đọc file buffer cho file ${row.fileName || row.name}:`, fileErr);
          }
        } else if (row.buffer && row.buffer instanceof ArrayBuffer) {
          fileBuffers.push({ name: row.fileName || row.name, bank: row.bank, buffer: row.buffer, month: normalizePivotMonth(rowMonth) });
        }
      }

      let resGrouped: Record<string, Record<string, Record<string, Record<string, number>>>> = {};
      let resTypes: string[] = [];
      let restoredLogs: typeof diagnosticLogs = [];
      let infoStr = "";

      if (fileBuffers.length > 0) {
        const res = await processFileBuffers(fileBuffers);
        resGrouped = res?.groupedData || {};
        resTypes = res?.typeColumns || [];
        restoredLogs = res?.logs || [];
        infoStr = `Đồng bộ từ ${fileBuffers.length} file Master`;

        if ((res?.processedMktFiles || 0) > 0) {
          const currentMktTypeCache = readPivotMktTypeCache(
            cachedGroupedData,
            cachedTypeColumns,
          );
          const nextMktTypeCache = updatePivotMktTypeCache(
            currentMktTypeCache,
            res?.mktGroupedData || {},
            res?.mktTypeColumns || [],
            resetMonths.length ? (res?.mktMonths || []).filter(month => resetMonths.includes(month)) : res?.mktMonths || [],
          );
          writePivotMktTypeCache(nextMktTypeCache);
        }
      }

      setDiagnosticLogs(restoredLogs);

      if (Object.keys(resGrouped).length === 0 && appData.Sheet1_AE?.data && appData.Sheet1_AE.data.length > 0) {
        const filteredSheet1 = appData.Sheet1_AE.data || [];
        const filteredRoster = appData.Master_Roster || [];

        const sheet1Res = buildPivotFromAppData(
          filteredSheet1,
          [],
          filteredRoster,
          appData.globalMonth || "03.2026",
        );
        const sheet1Grouped = sheet1Res?.groupedData || {};
        const sheet1Types = sheet1Res?.typeColumns || [];

        // Merge Sheet1_AE pivot data with fileBuffers pivot data if both present
        for (const bu in sheet1Grouped) {
          if (!resGrouped[bu]) resGrouped[bu] = {};
          for (const l07 in sheet1Grouped[bu]) {
            if (!resGrouped[bu][l07]) resGrouped[bu][l07] = {};
            for (const month in sheet1Grouped[bu][l07]) {
              if (!resGrouped[bu][l07][month]) resGrouped[bu][l07][month] = {};
              for (const col in sheet1Grouped[bu][l07][month]) {
                const val = sheet1Grouped[bu][l07][month][col] || 0;
                resGrouped[bu][l07][month][col] = (resGrouped[bu][l07][month][col] || 0) + val;
              }
            }
          }
        }

        const typeSet = new Set([...resTypes, ...sheet1Types]);
        resTypes = Array.from(typeSet);
        infoStr = infoStr ? `${infoStr} & Sheet1 AE` : `Đồng bộ từ dữ liệu Sheet1 AE (${filteredSheet1.length} dòng)`;
      }

      if (Object.keys(resGrouped).length > 0) {
        const mktTypeCache = readPivotMktTypeCache(
          cachedGroupedData,
          cachedTypeColumns,
        );
        const rebuiltGroupedData = applyPivotMktTypeCache(
          resGrouped,
          mktTypeCache,
        );
        const restoredGroupedData = resetMonths.length
          ? resetMonths.reduce((current, month) => replacePivotCacheMonth(current, rebuiltGroupedData, month), cachedGroupedData)
          : rebuiltGroupedData;
        const restoredTypeColumns = mergePivotTypeColumns(
          resetMonths.length ? [...cachedTypeColumns, ...resTypes] : resTypes,
          mktTypeCache.typeColumns,
        );

        setGroupedData(restoredGroupedData);
        setTypeColumns(restoredTypeColumns);
        _setSourceInfo(infoStr);

        try {
          localStorage.setItem("pivot_master_processed_data", JSON.stringify({
            cacheVersion: PIVOT_CACHE_VERSION,
            groupedData: restoredGroupedData,
            typeColumns: restoredTypeColumns,
            diagnosticLogs: restoredLogs,
            sourceInfo: infoStr,
            filter: selectedMonthFilter,
            reportingMonth: appData.globalMonth || "03.2026",
            updatedAt: Date.now()
          }));
        } catch {
          // ignore
        }

        if (showToastMsg) {
          if (getPivotDataMonths(restoredGroupedData).includes(selectedMonthFilter)) restorePivotMonths([selectedMonthFilter]);
          toast.success("Đã đồng bộ lại dữ liệu Pivot Master từ bảng Setting Master");
        }
      } else {
        if (showToastMsg) {
          toast.info(
            "Chưa có dữ liệu Master đã xử lý. Đang mở bảng Setting Master.",
          );
          window.dispatchEvent(new Event("master-ae-request-upload"));
        }
      }
    } catch (err) {
      console.error("Error loading master data:", err);
      if (showToastMsg) {
        toast.error("Lỗi khi xử lý dữ liệu từ bảng Master");
      }
    } finally {
      setIsProcessing(false);
    }
  }, [
    appData.Ae_Global_Inputs,
    appData.globalMonth,
    appData.Sheet1_AE?.data,
    appData.Master_Roster,
    processFileBuffers,
    restorePivotMonths,
    selectedMonthFilter,
  ]);

  useEffect(() => {
    if (!isAppDataLoading) void loadMasterData(false);
  }, [isAppDataLoading, loadMasterData]);

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (!files || files.length === 0) return;

    setIsProcessing(true);
    try {
      const fileBuffers: { name: string; buffer: ArrayBuffer; month: string }[] = [];
      for (let i = 0; i < files.length; i++) {
        const file = files[i];
        try {
          const buffer = await file.arrayBuffer();
          const selectedMonth = selectedMonthFilter === "ALL" ? appData.globalMonth : selectedMonthFilter;
          const fileMonth = normalizePivotMonth(parsePayrollMonth(file.name)?.dot || selectedMonth || "06.2026");
          fileBuffers.push({ name: file.name, buffer, month: fileMonth });
        } catch (fErr) {
          console.warn(`Lỗi khi đọc file ${file.name}:`, fErr);
        }
      }

      const res = await processFileBuffers(fileBuffers);
      if ((res?.processedMktFiles || 0) === 0) {
        toast.error(
          "Không tìm thấy sheet Roster hoặc Q_Roster hợp lệ trong file MKT Local North.",
        );
        return;
      }

      const currentMktTypeCache = readPivotMktTypeCache(
        safeGroupedData,
        safeTypeColumns,
      );
      const nextMktTypeCache = updatePivotMktTypeCache(
        currentMktTypeCache,
        res?.mktGroupedData || {},
        res?.mktTypeColumns || [],
        res.mktMonths,
        true,
      );
      writePivotMktTypeCache(nextMktTypeCache);

      // Áp dụng snapshot TYPE bằng phép thay thế có chọn lọc. Các cột không
      // thuộc MKT Local North không bị xóa, cộng thêm hay thay đổi giá trị.
      const mergedGroupedData = applyPivotMktTypeCache(
        groupedData,
        nextMktTypeCache,
        res.mktMonths,
      );
      const mergedTypeColumns = mergePivotTypeColumns(
        safeTypeColumns,
        nextMktTypeCache.typeColumns,
      );
      const mergedLogs = [...diagnosticLogs, ...(res?.logs || [])];

      setGroupedData(mergedGroupedData);
      setTypeColumns(mergedTypeColumns);
      setDiagnosticLogs(mergedLogs);
      const infoStr = `Đã cập nhật ${res.processedMktFiles} file MKT Local (${nextMktTypeCache.typeColumns.length} cột TYPE); giữ nguyên các cột khác`;
      _setSourceInfo(infoStr);

      try {
        localStorage.setItem("pivot_master_processed_data", JSON.stringify({
          cacheVersion: PIVOT_CACHE_VERSION,
          groupedData: mergedGroupedData,
          typeColumns: mergedTypeColumns,
          diagnosticLogs: mergedLogs,
          sourceInfo: infoStr,
          filter: selectedMonthFilter,
          reportingMonth: appData.globalMonth || "03.2026",
          updatedAt: Date.now()
        }));
      } catch {
        // ignore cache write error
      }
      restorePivotMonths(res.mktMonths);
      toast.success(
        `Đã cập nhật ${res.processedMktFiles} file MKT Local; ${nextMktTypeCache.typeColumns.length} cột TYPE đã được ghi nhớ, các cột khác giữ nguyên`,
      );
    } catch (err) {
      console.error("Error processing manual upload:", err);
      toast.error("Lỗi khi xử lý các file vừa chọn");
    } finally {
      setIsProcessing(false);
      if (fileInputRef.current) {
        fileInputRef.current.value = "";
      }
    }
  };

  const handleDownloadDiagnosticCSV = () => {
    setIsSettingsOpen(false);
    if (diagnosticLogs.length === 0) {
      toast.info("Không có dòng log lỗi nào cần kiểm tra.");
      return;
    }

    const headers = ["Source", "File", "RawCenter", "Column", "RawValue"];
    const csvRows = [headers.join(",")];

    diagnosticLogs.forEach(log => {
      const row = [
        `"${log.Source || ""}"`,
        `"${log.File || ""}"`,
        `"${String(log.RawCenter || "").replace(/"/g, '""')}"`,
        `"${log.Column || ""}"`,
        `"${String(log.RawValue || "").replace(/"/g, '""')}"`
      ];
      csvRows.push(row.join(","));
    });

    const blob = new Blob([csvRows.join("\n")], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.setAttribute("download", `pivot-diagnostic-logs-${Date.now()}.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  const handleExportExcel = () => {
    setIsSettingsOpen(false);
    if (Object.keys(safeGroupedData).length === 0) {
      toast.error("Không có dữ liệu để xuất Excel!");
      return;
    }

    const visibleTypes = safeTypeColumns
      .map((type, idx) => ({ type, idx }))
      ;

    const headers: string[] = [];
    headers.push("No.");
    headers.push("Business");
    headers.push("L07");
    headers.push("Tháng");
    visibleTypes.forEach(({ type }) => headers.push(type));
    headers.push("TỔNG CỘNG");

    const wsData: any[][] = [];
    wsData.push(headers);

    let rowId = 1;
    const excelGrandTotals = new Array(visibleTypes.length).fill(0);
    let superGrandTotal = 0;
    const sortedBUs = Object.keys(safeGroupedData)
      .filter((bu) => {
        if (buFilterMode === "ALL") return true;
        if (buFilterMode === "EXCLUDE_AHP") return !isAhpBuValue(bu);
        return bu === buFilterMode;
      })
      .sort();

    sortedBUs.forEach(bu => {
      const buTotals = new Array(visibleTypes.length).fill(0);
      let buGrandTotal = 0;
      const l07s = Object.keys(safeGroupedData[bu] || {}).sort();

      l07s.forEach(l07 => {
        const uL07 = l07.toUpperCase().trim();
        if (
          uL07.includes("MKT LOCAL NORTH") ||
          uL07.startsWith("MKT LOCAL") ||
          uL07.includes("MKT_LOCAL") ||
          uL07 === "MKT"
        ) {
          return;
        }
        const months = Object.keys(safeGroupedData[bu][l07] || {}).sort();
        months.forEach(month => {
          if (selectedMonthFilter !== "ALL" && normalizePivotMonth(month) !== selectedMonthFilter) return;
          
          let rowTotal = 0;
          const rowVals = visibleTypes.map(({ type }, vIdx) => {
            const val = safeGroupedData[bu][l07][month][type] || 0;
            buTotals[vIdx] += val;
            excelGrandTotals[vIdx] += val;
            rowTotal += val;
            return val;
          });
          buGrandTotal += rowTotal;
          superGrandTotal += rowTotal;
          const sourceLabels = getPivotSourceLabels(safeGroupedData[bu][l07][month]);
          const displayL07 = sourceLabels.length > 0
            ? `${l07} — ${sourceLabels.join(" / ")}`
            : l07;

          const rowData: any[] = [];
          rowData.push(rowId++);
          rowData.push(bu);
          rowData.push(displayL07);
          rowData.push(month);
          rowData.push(...rowVals);
          rowData.push(rowTotal);

          wsData.push(rowData);
        });
      });

      const buRowData: any[] = [];
      buRowData.push("");
      buRowData.push(bu);
      buRowData.push(`${bu} Total`);
      buRowData.push("");
      buRowData.push(...buTotals);
      buRowData.push(buGrandTotal);

      wsData.push(buRowData);
    });

    const totalRowData: any[] = [];
    totalRowData.push("");
    totalRowData.push("");
    totalRowData.push("");
    totalRowData.push("TỔNG CỘNG");
    totalRowData.push(...excelGrandTotals);
    totalRowData.push(superGrandTotal);

    wsData.push(totalRowData);

    const ws = XLSX.utils.aoa_to_sheet(wsData);
    ws["!cols"] = [
      { hidden: hiddenColumns.no }, { hidden: hiddenColumns.business },
      { hidden: hiddenColumns.charge }, { hidden: hiddenColumns.month },
      ...safeTypeColumns.map((type, idx) => ({ hidden: isTypeColHidden(type, idx) })),
      { hidden: hiddenColumns.grandTotal },
    ];
    Object.values(ws).forEach(cell => { if (cell?.t === "n") cell.z = "General"; });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Pivot_Data");
    XLSX.writeFile(wb, "Pivot_Salary_Report.xlsx");
    toast.success("Đã xuất báo cáo Excel thành công!");
  };

  const availableMonths = useMemo(() => {
    const months = new Set(getPivotDataMonths(safeGroupedData));
    if (selectedMonthFilter !== "ALL") months.add(selectedMonthFilter);
    return [...months].sort();
  }, [safeGroupedData, selectedMonthFilter]);
  const canDeleteSelectedMonth = !isAppDataLoading && !isProcessing &&
    Boolean(parsePayrollMonth(selectedMonthFilter)) && getPivotDataMonths(safeGroupedData).includes(selectedMonthFilter);

  const { allFlatRows, totalCenters, totalSalarySum, grandTotals, superGrandTotal } = useMemo(() => {
    let totalCenters = 0;
    let totalSalarySum = 0;
    const grandTotals = new Array(safeTypeColumns.length).fill(0);
    let superGrandTotal = 0;

    const allFlatRows: Array<{
      globalRowId: number;
      month: string;
      bu: string;
      l07: string;
      values: number[];
      rowTotal: number;
      sourceLabels: string[];
    }> = [];

    let rIdx = 1;
    const currentSortedBUs = Object.keys(safeGroupedData)
      .filter((bu) => {
        if (buFilterMode === "ALL") return true;
        if (buFilterMode === "EXCLUDE_AHP") return !isAhpBuValue(bu);
        return String(bu).trim().toUpperCase() === buFilterMode.trim().toUpperCase();
      })
      .sort();
    currentSortedBUs.forEach(bu => {
      const l07s = Object.keys(safeGroupedData[bu] || {}).sort();
      l07s.forEach(l07 => {
        const uL07 = l07.toUpperCase().trim();
        if (
          uL07.includes("MKT LOCAL NORTH") ||
          uL07.startsWith("MKT LOCAL") ||
          uL07.includes("MKT_LOCAL") ||
          uL07 === "MKT"
        ) {
          return;
        }
        const months = Object.keys(safeGroupedData[bu][l07] || {}).sort();
        months.forEach(month => {
          if (selectedMonthFilter !== "ALL" && normalizePivotMonth(month) !== selectedMonthFilter) return;

          let rowTotal = 0;
          const values = safeTypeColumns.map((type, idx) => {
            const val = safeGroupedData[bu][l07][month][type] || 0;
            grandTotals[idx] += val;
            rowTotal += val;
            return val;
          });

          if (rowTotal === 0 && bu === "OTHER" && (l07 === "UNKNOWN" || !l07)) {
            return;
          }

          totalCenters++;
          superGrandTotal += rowTotal;
          totalSalarySum += rowTotal;

          allFlatRows.push({
            globalRowId: rIdx++,
            month,
            bu,
            l07,
            values,
            rowTotal,
            sourceLabels: getPivotSourceLabels(safeGroupedData[bu][l07][month]),
          });
        });
      });
    });
    return { allFlatRows, totalCenters, totalSalarySum, grandTotals, superGrandTotal };
  }, [safeGroupedData, safeTypeColumns, selectedMonthFilter, buFilterMode]);

  const isTypeColHidden = useCallback((type: string, idx: number) => {
    if (hiddenColumns[`type_${type}`] !== undefined) {
      return hiddenColumns[`type_${type}`];
    }
    return (grandTotals[idx] || 0) === 0;
  }, [grandTotals, hiddenColumns]);

  const sortedFlatRows = useMemo(() => {
    if (!sortField) return allFlatRows;
    return [...allFlatRows].sort((a, b) => {
      let valA: any = 0;
      let valB: any = 0;
      if (sortField === "no") {
        valA = a.globalRowId;
        valB = b.globalRowId;
      } else if (sortField === "month") {
        valA = a.month;
        valB = b.month;
      } else if (sortField === "bu") {
        valA = a.bu;
        valB = b.bu;
      } else if (sortField === "l07") {
        valA = a.l07;
        valB = b.l07;
      } else if (sortField === "rowTotal") {
        valA = a.rowTotal;
        valB = b.rowTotal;
      } else if (sortField.startsWith("type_")) {
        const typeName = sortField.replace("type_", "");
        const typeIdx = safeTypeColumns.indexOf(typeName);
        if (typeIdx !== -1) {
          valA = a.values[typeIdx] || 0;
          valB = b.values[typeIdx] || 0;
        }
      }
      if (typeof valA === "string") {
        const cmp = valA.localeCompare(valB, 'vi');
        return sortDirection === "asc" ? cmp : -cmp;
      }
      return sortDirection === "asc" ? valA - valB : valB - valA;
    });
  }, [allFlatRows, safeTypeColumns, sortField, sortDirection]);

  useEffect(() => {
    const cols = [
      {key: "No.", label: "No.", hidden: !!hiddenColumns.no},
      {key: "Business", label: "Business", hidden: !!hiddenColumns.business},
      {key: "L07", label: "L07", hidden: !!hiddenColumns.charge},
      {key: "Month", label: "Tháng", hidden: !!hiddenColumns.month},
      ...safeTypeColumns.map((type, idx) => ({key: type, label: type, type: "currency", hidden: isTypeColHidden(type, idx)})),
      {key: "Grand Total", label: "TỔNG CỘNG", type: "currency", hidden: !!hiddenColumns.grandTotal},
    ];
    return registerTableExport("master-pivot", () => ({
      schema: {columns: cols, hiddenColumns: cols.filter(c => c.hidden).map(c => c.key)},
      rows: sortedFlatRows.map(row => ({"No.": row.globalRowId, Business: row.bu,
        L07: row.sourceLabels.length ? row.l07 + " — " + row.sourceLabels.join(" / ") : row.l07,
        Month: row.month, ...Object.fromEntries(safeTypeColumns.map((type, idx) => [type, row.values[idx]])), "Grand Total": row.rowTotal})),
    }));
  });

  const autoFitAllColumns = useCallback(() => {
    const canvas = document.createElement("canvas");
    const context = canvas.getContext("2d");
    const rowsToMeasure = sortedFlatRows.slice(0, 500);

    const measureCol = (headerText: string, values: (string | number | undefined | null)[], minWidth = 60, maxWidth = 1000) => {
      let maxW = minWidth;
      if (context) {
        context.font = "bold 11px Inter, sans-serif";
        maxW = Math.max(maxW, context.measureText(headerText).width + 72);

        context.font = "12px 'Plus Jakarta Sans', 'Inter', sans-serif";
        for (const val of values) {
          if (val === undefined || val === null || val === "") continue;
          const text = typeof val === "number" ? formatNumber(val) : String(val);
          const w = context.measureText(text).width + 40;
          if (w > maxW) maxW = w;
        }
      } else {
        const longest = values.reduce<number>(
          (length, value) => Math.max(length, String(value ?? "").length),
          headerText.length
        );
        maxW = Math.max(minWidth, longest * 9 + 72);
      }
      return Math.min(maxWidth, Math.max(minWidth, Math.ceil(maxW)));
    };

    const nextWidths: Record<string, number> = {
      no: measureCol("No.", rowsToMeasure.map((_, i) => i + 1), 60, 100),
      business: measureCol("Business", rowsToMeasure.map(r => r.bu), 90, 350),
      charge: measureCol(
        "L07",
        rowsToMeasure.map(r => [r.l07, ...r.sourceLabels].join(" · ")),
        160,
        600,
      ),
      month: measureCol("Tháng", rowsToMeasure.map(r => r.month), 80, 180),
      grandTotal: measureCol("TỔNG CỘNG", rowsToMeasure.map(r => r.rowTotal), 130, 350),
    };

    safeTypeColumns.forEach((type, typeIndex) => {
      if (isTypeColHidden(type, typeIndex)) return;
      nextWidths[`type_${type}`] = measureCol(
        type,
        rowsToMeasure.map(r => r.values[typeIndex] || 0),
        95,
        600
      );
    });

    setColumnWidths(nextWidths);
    try {
      localStorage.setItem("pivot_master_column_widths", JSON.stringify(nextWidths));
    } catch {
      // Ignore local storage quota/privacy errors.
    }
    toast.success("Đã tự động căn chỉnh độ rộng cột Pivot Master theo dữ liệu!");
  }, [safeTypeColumns, sortedFlatRows, isTypeColHidden]);

  const visibleTableWidth = useMemo(() => {
    let total = 0;
    if (!hiddenColumns.no) total += columnWidths.no || 50;
    if (!hiddenColumns.business) total += columnWidths.business || 90;
    if (!hiddenColumns.charge) total += columnWidths.charge || 220;
    if (!hiddenColumns.month) total += columnWidths.month || 90;
    safeTypeColumns.forEach((type, idx) => {
      if (!isTypeColHidden(type, idx)) {
        total += columnWidths[`type_${type}`] || 120;
      }
    });
    if (!hiddenColumns.grandTotal) total += columnWidths.grandTotal || 140;
    return Math.max(total, 640);
  }, [columnWidths, hiddenColumns, safeTypeColumns, isTypeColHidden]);

  const totalRowsCount = sortedFlatRows.length;
  const totalPages =
    rowsPerPage === Infinity
      ? 1
      : Math.max(1, Math.ceil(totalRowsCount / rowsPerPage));
  const validCurrentPage = Math.min(currentPage, totalPages);
  const startIndex =
    totalRowsCount === 0 || rowsPerPage === Infinity
      ? 0
      : (validCurrentPage - 1) * rowsPerPage;
  const endIndex =
    rowsPerPage === Infinity
      ? totalRowsCount
      : Math.min(startIndex + rowsPerPage, totalRowsCount);
  const paginatedRows =
    rowsPerPage === Infinity
      ? sortedFlatRows
      : sortedFlatRows.slice(startIndex, endIndex);
  const buSubtotals = new Map<
    string,
    { values: number[]; rowTotal: number }
  >();
  allFlatRows.forEach((row) => {
    const subtotal = buSubtotals.get(row.bu) || {
      values: new Array(safeTypeColumns.length).fill(0),
      rowTotal: 0,
    };
    row.values.forEach((value, index) => {
      subtotal.values[index] += value;
    });
    subtotal.rowTotal += row.rowTotal;
    buSubtotals.set(row.bu, subtotal);
  });
  const lastRowIdByBu = new Map<string, number>();
  sortedFlatRows.forEach((row) => {
    lastRowIdByBu.set(row.bu, row.globalRowId);
  });
  const pivotLabelColumnSpan = ["no", "business", "charge", "month"].filter(
    (key) => !hiddenColumns[key],
  ).length;

  const renderRows = () => {
    if (paginatedRows.length === 0) {
      return (
        <tr>
          <td colSpan={6 + safeTypeColumns.length} className="bg-card py-12 text-center text-sm text-muted-foreground">
            <span>Chưa có dữ liệu. Vui lòng tải file ở bảng <span className="font-semibold text-foreground/70">Setting Master</span> và nhấn <span className="font-semibold text-foreground/70">Xử lý dữ liệu</span>.</span>
          </td>
        </tr>
      );
    }

    return paginatedRows.map((item) => {
      const isEditingThisRow = editingCell?.bu === item.bu && editingCell?.l07 === item.l07 && editingCell?.month === item.month;
      const showBuSubtotal = lastRowIdByBu.get(item.bu) === item.globalRowId;
      const buSubtotal = buSubtotals.get(item.bu);

      return (
        <React.Fragment key={`${item.bu}-${item.l07}-${item.month}`}>
          <tr
            className="pivot-master-data-row border-b border-border bg-[var(--table-data-bg,var(--card))] transition-colors"
          >
          {!hiddenColumns.no && (
            <td 
              style={{ width: columnWidths["no"] || 50, minWidth: columnWidths["no"] || 50, maxWidth: columnWidths["no"] || 50 }}
              className="group/no relative align-middle border-r border-b border-[var(--grid-line-color,rgba(0,0,0,0.035))] px-2 py-2 text-center text-xs tabular-nums text-muted-foreground"
            >
              <span>{item.globalRowId}</span>
              <button
                onClick={() => handleDeleteRow(item.bu, item.l07, item.month)}
                className="opacity-0 group-hover/no:opacity-100 absolute right-1 top-1/2 -translate-y-1/2 text-rose-500 hover:text-rose-700 transition-opacity p-0.5 cursor-pointer"
                title="Xóa dòng"
              >
                <Trash2 className="w-3 h-3" />
              </button>
            </td>
          )}
          {!hiddenColumns.business && (
            <td 
              style={{ width: columnWidths["business"] || 90, minWidth: columnWidths["business"] || 90, maxWidth: columnWidths["business"] || 90 }}
              onDoubleClick={() => handleStartEdit(item.bu, item.l07, item.month, "bu", item.bu)}
              className="cursor-pointer align-middle whitespace-normal break-words border-r border-b border-[var(--grid-line-color,rgba(0,0,0,0.035))] px-2.5 py-2 text-center text-xs font-semibold text-card-foreground transition-colors"
              title="Nhấp đúp để sửa Business"
            >
              {isEditingThisRow && editingCell.field === "bu" ? (
                <input
                  autoFocus
                  value={editValue}
                  onChange={(e) => setEditValue(e.target.value)}
                  onKeyDown={(e) => {
                    if (e.key === "Enter") handleSaveEdit();
                    if (e.key === "Escape") handleCancelEdit();
                  }}
                  onBlur={handleSaveEdit}
                  className="w-full rounded-md border border-primary/35 bg-card px-1.5 py-0.5 text-center text-xs font-semibold text-foreground outline-none ring-2 ring-primary/15"
                />
              ) : (
                <span>{item.bu}</span>
              )}
            </td>
          )}
          {!hiddenColumns.charge && (
            <td 
              style={{ width: columnWidths["charge"] || 220, minWidth: columnWidths["charge"] || 220, maxWidth: columnWidths["charge"] || 220 }}
              onDoubleClick={() => handleStartEdit(item.bu, item.l07, item.month, "l07", item.l07)}
              className="cursor-pointer align-middle whitespace-normal break-words border-r border-b border-[var(--grid-line-color,rgba(0,0,0,0.035))] px-2.5 py-2 text-left text-xs font-medium text-card-foreground transition-colors"
              title={item.sourceLabels.length > 0
                ? `L07 ${item.l07} — Khoản: ${item.sourceLabels.join(" / ")}. Nhấp đúp để sửa L07.`
                : `Nhấp đúp để sửa L07 (${item.l07})`}
            >
              {isEditingThisRow && editingCell.field === "l07" ? (
                <input
                  autoFocus
                  value={editValue}
                  onChange={(e) => setEditValue(e.target.value)}
                  onKeyDown={(e) => {
                    if (e.key === "Enter") handleSaveEdit();
                    if (e.key === "Escape") handleCancelEdit();
                  }}
                  onBlur={handleSaveEdit}
                  className="w-full rounded-md border border-primary/35 bg-card px-1.5 py-0.5 text-xs font-medium text-foreground outline-none ring-2 ring-primary/15"
                />
              ) : (
                <div className="flex min-w-0 flex-wrap items-center gap-1.5">
                  <span className="break-all">{item.l07}</span>
                  {item.sourceLabels.length > 0 && (
                    <span className="whitespace-normal break-words rounded border border-primary/15 bg-primary/8 px-1.5 py-0.5 text-[9px] font-semibold uppercase tracking-wide text-primary/75">
                      Khoản: {item.sourceLabels.join(" / ")}
                    </span>
                  )}
                </div>
              )}
            </td>
          )}
          {!hiddenColumns.month && (
            <td 
              style={{ width: columnWidths["month"] || 90, minWidth: columnWidths["month"] || 90, maxWidth: columnWidths["month"] || 90 }}
              onDoubleClick={() => handleStartEdit(item.bu, item.l07, item.month, "month", item.month)}
              className="cursor-pointer align-middle whitespace-normal break-words border-r border-b border-[var(--grid-line-color,rgba(0,0,0,0.035))] px-2.5 py-2 text-center text-xs tabular-nums text-card-foreground transition-colors"
              title="Nhấp đúp để sửa Tháng"
            >
              {isEditingThisRow && editingCell.field === "month" ? (
                <input
                  autoFocus
                  value={editValue}
                  onChange={(e) => setEditValue(e.target.value)}
                  onKeyDown={(e) => {
                    if (e.key === "Enter") handleSaveEdit();
                    if (e.key === "Escape") handleCancelEdit();
                  }}
                  onBlur={handleSaveEdit}
                  className="w-full rounded-md border border-primary/35 bg-card px-1.5 py-0.5 text-center text-xs tabular-nums text-foreground outline-none ring-2 ring-primary/15"
                />
              ) : (
                <span>{item.month}</span>
              )}
            </td>
          )}
          {safeTypeColumns.map((type, tIdx) => {
            if (isTypeColHidden(type, tIdx)) return null;
            const val = item.values[tIdx];
            const colKey = `type_${type}`;
            const w = columnWidths[colKey] || 120;

            return (
              <td 
                key={colKey}
                style={{ width: w, minWidth: w, maxWidth: w }}
                onDoubleClick={() => handleStartEdit(item.bu, item.l07, item.month, type, val)}
                className={`cursor-pointer align-middle whitespace-normal break-words border-r border-b border-[var(--grid-line-color,rgba(0,0,0,0.035))] px-2.5 py-2 text-right text-xs font-normal tabular-nums transition-colors ${val > 0 ? "text-card-foreground" : "text-muted-foreground"}`}
                title={`Nhấp đúp để sửa ${type}`}
              >
                {isEditingThisRow && editingCell.field === type ? (
                  <input
                    autoFocus
                    value={editValue}
                    onChange={(e) => setEditValue(e.target.value)}
                    onKeyDown={(e) => {
                      if (e.key === "Enter") handleSaveEdit();
                      if (e.key === "Escape") handleCancelEdit();
                    }}
                    onBlur={handleSaveEdit}
                    className="w-full rounded-md border border-primary/35 bg-card px-1.5 py-0.5 text-right text-xs tabular-nums text-foreground outline-none ring-2 ring-primary/15"
                  />
                ) : (
                  <span>{val ? formatNumber(val) : "0"}</span>
                )}
              </td>
            );
          })}
          {!hiddenColumns.grandTotal && (
            <td 
              style={{ width: columnWidths["grandTotal"] || 140, minWidth: columnWidths["grandTotal"] || 140, maxWidth: columnWidths["grandTotal"] || 140 }}
              className="align-middle whitespace-normal break-words border-b border-border px-3 py-2 text-right text-xs font-semibold tabular-nums text-primary"
            >
              {item.rowTotal ? formatNumber(item.rowTotal) : "0"}
            </td>
          )}
          </tr>
          {showBuSubtotal && buSubtotal && (
            <tr className="pivot-master-subtotal-row total-row border-b-2 border-border font-bold text-primary shadow-xs">
              {pivotLabelColumnSpan > 0 && (
                <td
                  colSpan={pivotLabelColumnSpan}
                  className="border-r border-[var(--grid-line-color,rgba(0,0,0,0.035))] px-2.5 py-2 text-left text-xs font-bold text-primary"
                >
                  TỔNG CỘNG {item.bu}
                </td>
              )}
              {buSubtotal.values.map((value, index) => {
                const type = safeTypeColumns[index];
                if (isTypeColHidden(type, index)) return null;
                const colKey = `type_${type}`;
                const width = columnWidths[colKey] || 120;
                return (
                  <td
                    key={`${item.bu}-${type}`}
                    style={{ width, minWidth: width, maxWidth: width }}
                    className="border-r border-[var(--grid-line-color,rgba(0,0,0,0.035))] px-2.5 py-2 text-right text-xs font-bold tabular-nums text-primary"
                  >
                    {value ? formatNumber(value) : "0"}
                  </td>
                );
              })}
              {!hiddenColumns.grandTotal && (
                <td
                  style={{
                    width: columnWidths["grandTotal"] || 140,
                    minWidth: columnWidths["grandTotal"] || 140,
                    maxWidth: columnWidths["grandTotal"] || 140,
                  }}
                  className="px-3 py-2 text-right text-xs font-bold tabular-nums text-primary"
                >
                  {buSubtotal.rowTotal ? formatNumber(buSubtotal.rowTotal) : "0"}
                </td>
              )}
            </tr>
          )}
        </React.Fragment>
      );
    });
  };

  const toggleSort = (field: string) => {
    if (sortField === field) {
      if (sortDirection === "asc") {
        setSortDirection("desc");
      } else {
        setSortField(null);
        setSortDirection("asc");
        toast.success("Đã xóa sắp xếp cột");
      }
    } else {
      setSortField(field);
      setSortDirection("asc");
    }
  };

  const renderSortIndicator = (field: string) => {
    if (sortField !== field) return null;
    return (
      <div className="inline-flex items-center gap-0.5 ml-1 shrink-0">
        <span className="text-current font-bold">
          {sortDirection === "asc" ? (
            <ChevronUp className="w-3 h-3 stroke-[2.5]" />
          ) : (
            <ChevronDown className="w-3 h-3 stroke-[2.5]" />
          )}
        </span>
        <button
          type="button"
          onClick={(e) => {
            e.stopPropagation();
            e.preventDefault();
            setSortField(null);
            setSortDirection("asc");
            toast.success("Đã xóa sắp xếp cột");
          }}
          className="p-0.5 rounded hover:bg-white/20 text-current/80 hover:text-white transition-colors cursor-pointer"
          title="Xóa sắp xếp cột này"
          aria-label="Xóa sắp xếp cột này"
        >
          <X className="w-2.5 h-2.5 stroke-[2.5]" />
        </button>
      </div>
    );
  };

  return (
    <div 
      className="pivot-master-frame unified-table-frame table-container relative flex h-full w-full flex-col gap-0 overflow-hidden border border-border bg-card p-0 text-card-foreground shadow-xs"
      style={{
        borderRadius: "var(--table-radius, 12px)",
        borderWidth: "1px",
        borderColor: "var(--table-frame-border, var(--border))",
        padding: 0,
      }}
    >
      {/* HEADER SECTION */}
      <div 
        className="unified-table-frame-header flex min-h-[56px] shrink-0 items-center justify-between gap-3 border-b border-border bg-[var(--table-header-bg,#FAF3E8)] px-3 py-2"
        style={{ backgroundColor: "var(--table-header-bg, #FAF3E8)", minHeight: "56px" }}
      >
        <div className="flex w-full min-w-0 items-center">
          <div className="app-table-title-lockup min-w-0">
            <div className="app-table-title-line">
              <TableInitialMark
                label="PIVOT MASTER COST ALLOCATION BY BU, L07 & TASK TYPE"
                className="shrink-0 text-primary"
              />
              <h3 className="truncate text-[15px] font-bold leading-[22px] tracking-tight text-foreground">
                <TableTitleRemainder
                  label="PIVOT MASTER COST ALLOCATION BY BU, L07 & TASK TYPE"
                  className="app-table-title-remainder--expanded text-[15px]"
                />
              </h3>
            </div>
            <p className="app-table-title-meta truncate text-[10px] font-medium leading-[14px] text-muted-foreground">
              Tổng hợp chi phí theo BU, L07 và loại · {totalCenters} trung tâm
            </p>
          </div>

          {/* Compact right-side controls. Keep their combined width bounded so
              the month selector cannot be pushed outside the table card. */}
          <div className="ml-auto grid shrink-0 grid-cols-[minmax(108px,auto)_108px_28px] items-end gap-1.5">
            {/* TỔNG TIỀN */}
            <div className="flex min-w-[108px] flex-col items-stretch border-l border-border/60 pl-2">
              <span className="whitespace-nowrap text-right pr-[5px] text-[9px] font-bold uppercase tracking-tighter text-foreground/60">
                TỔNG TIỀN
              </span>
              <div className="mt-0.5 flex h-7 min-w-[100px] items-center justify-center rounded-md border border-border/70 bg-card px-3 shadow-2xs">
                <span className="whitespace-nowrap text-[13px] px-1 font-black tracking-tight text-primary tabular-nums">
                  {formatNumber(superGrandTotal)}
                </span>
              </div>
            </div>

            {/* Month Filter Selector */}
            <div className="flex w-[108px] min-w-0 flex-col items-stretch">
              <span className="whitespace-nowrap text-right pl-0.5 pr-1.5 text-[9px] font-bold uppercase tracking-tighter text-foreground/60">
                THÁNG
              </span>
              <div className="relative mt-0.5 flex h-7 min-w-0 items-center rounded-full border border-border bg-card px-2 shadow-2xs">
                <select
                  value={selectedMonthFilter}
                  onChange={(e) => {
                    const val = e.target.value;
                    setSelectedMonthFilter(val);
                    try {
                      localStorage.setItem("pivot_master_selected_month_filter", val);
                    } catch {
                      // ignore
                    }
                  }}
                  className="w-full min-w-0 cursor-pointer appearance-none bg-transparent py-0 pl-0 pr-[6px] leading-[18.5px] text-center text-[10px] font-bold text-foreground focus:outline-none"
                  aria-label="Chọn tháng Pivot Master"
                >
                  <option value="ALL">Tất cả</option>
                  {availableMonths.map((m) => (
                    <option key={m} value={m}>
                      {m}
                    </option>
                  ))}
                </select>
                <ChevronDown className="pointer-events-none absolute right-1.5 top-1/2 h-3 w-3 -translate-y-1/2 text-muted-foreground" />
              </div>
            </div>

            {/* Settings button & dropdown */}
            <div className="relative" ref={settingsMenuRef}>
              <button
                type="button"
                onClick={() => setIsSettingsOpen(!isSettingsOpen)}
                className="flex h-7 w-7 cursor-pointer items-center justify-center rounded-full border border-border bg-card text-card-foreground shadow-2xs transition-all hover:border-primary/40 hover:bg-primary/[0.05] hover:text-primary active:scale-95"
                title="Cài đặt & Tiện ích Pivot Master"
                aria-label="Cài đặt Pivot Master"
              >
                <Settings className="h-3.5 w-3.5 text-primary" />
              </button>

              {isSettingsOpen && (
                <div className="absolute top-full right-0 z-[9999] mt-2 flex w-72 flex-col gap-3 rounded-xl border border-border bg-card p-3.5 text-card-foreground shadow-2xl">
                  <div className="flex items-center justify-between border-b border-slate-100 pb-2">
                    <span className="text-xs font-bold text-slate-800 flex items-center gap-1.5">
                      <Settings className="w-4 h-4 text-slate-600" /> Cài đặt & Thao tác
                    </span>
                    <button onClick={() => setIsSettingsOpen(false)} className="text-slate-400 hover:text-slate-600 p-0.5 cursor-pointer">
                      <X className="w-4 h-4" />
                    </button>
                  </div>

                  {/* ACTION BUTTONS SECTION */}
                  <TableRestoreButton placement="plain" onRestore={() => { setIsSettingsOpen(false); setEditingCell(null); void loadMasterData(true, true); }} />
                  <div className="flex flex-col gap-1.5">
                    <span className="text-[11px] font-semibold text-slate-500 uppercase tracking-wider">Thao tác dữ liệu</span>
                    <div className="grid grid-cols-2 gap-1.5">
                      <button
                        onClick={() => {
                          setIsSettingsOpen(false);
                          loadMasterData(true);
                        }}
                        disabled={isProcessing}
                        className="flex items-center justify-center gap-1.5 px-2.5 py-1.5 bg-slate-50 border border-slate-200 hover:bg-slate-100 text-slate-700 text-xs font-medium rounded-lg transition-all disabled:opacity-50 cursor-pointer"
                        title="Tải lại dữ liệu từ bảng Cài đặt"
                      >
                        <RefreshCw className={`w-3.5 h-3.5 ${isProcessing ? "animate-spin text-amber-600" : ""}`} />
                        <span>{isProcessing ? "Đang xử lý..." : "Đồng bộ"}</span>
                      </button>

                      <label className="flex items-center justify-center gap-1.5 px-2.5 py-1.5 bg-slate-50 border border-slate-200 hover:bg-slate-100 text-slate-700 text-xs font-medium rounded-lg cursor-pointer transition-all">
                        <Upload className="w-3.5 h-3.5 text-blue-600" />
                        <span>Tải file</span>
                        <input
                          ref={fileInputRef}
                          type="file"
                          disabled={isProcessing}
                          multiple
                          accept=".xlsx,.xls,.csv"
                          onChange={(e) => {
                            setIsSettingsOpen(false);
                            handleFileUpload(e);
                          }}
                          className="hidden"
                        />
                      </label>

                      <button
                        onClick={() => {
                          setIsSettingsOpen(false);
                          handleAddRow();
                        }}
                        className="flex items-center justify-center gap-1.5 px-2.5 py-1.5 bg-emerald-50 border border-emerald-200 hover:bg-emerald-100 text-emerald-700 text-xs font-medium rounded-lg transition-all cursor-pointer"
                      >
                        <Plus className="w-3.5 h-3.5" />
                        <span>Thêm dòng</span>
                      </button>

                      <button
                        onClick={() => {
                          setIsSettingsOpen(false);
                          handleAddColumn();
                        }}
                        className="flex items-center justify-center gap-1.5 px-2.5 py-1.5 bg-indigo-50 border border-indigo-200 hover:bg-indigo-100 text-indigo-700 text-xs font-medium rounded-lg transition-all cursor-pointer"
                      >
                        <Plus className="w-3.5 h-3.5" />
                        <span>Thêm cột</span>
                      </button>
                    </div>



                    <button
                      type="button"
                      disabled={!canDeleteSelectedMonth}
                      title={selectedMonthFilter === "ALL" ? "Chọn một tháng để xóa Pivot Master" : `Xóa Pivot Master tháng ${selectedMonthFilter}`}
                      onClick={() => {
                        setIsSettingsOpen(false);
                        setDeleteMonthTarget(selectedMonthFilter);
                      }}
                      className="flex w-full items-center gap-1.5 rounded-lg border border-destructive/20 px-2.5 py-1.5 text-xs font-semibold text-destructive transition-colors hover:bg-destructive/10 disabled:cursor-not-allowed disabled:opacity-40"
                    >
                      <Trash2 className="h-3.5 w-3.5" />
                      Xóa dữ liệu Pivot Master theo tháng
                    </button>

                    <button
                      onClick={() => chooseExcelExport(handleExportExcel)}
                      className="w-full flex items-center justify-center gap-1.5 px-3 py-1.5 bg-amber-600 hover:bg-amber-700 text-white text-xs font-medium rounded-lg shadow-2xs transition-all cursor-pointer"
                    >
                      <Download className="w-3.5 h-3.5" />
                      <span>Xuất Excel</span>
                    </button>
                  </div>

                  <div className="border-t border-slate-100 pt-2 flex flex-col gap-1.5">
                    <span className="text-[11px] font-semibold text-slate-500 uppercase tracking-wider">Giao diện bảng</span>
                    <button
                      type="button"
                      onClick={() => {
                        setIsSettingsOpen(false);
                        window.dispatchEvent(new Event("open-ui-settings"));
                      }}
                      className="flex w-full cursor-pointer items-center justify-center gap-1.5 rounded-lg border border-primary/20 bg-primary/[0.05] px-2.5 py-2 text-xs font-semibold text-primary transition-colors hover:bg-primary/10"
                    >
                      <SlidersHorizontal className="h-3.5 w-3.5" />
                      <span>Chỉnh sửa giao diện bảng</span>
                    </button>
                  </div>

                  <div className="border-t border-slate-100 pt-2 flex flex-col gap-1.5">
                    <span className="text-[11px] font-semibold text-slate-500 uppercase tracking-wider">Ẩn/Hiện cột</span>
                    <div className="max-h-40 overflow-y-auto flex flex-col gap-1 pr-1">
                      <label className="flex items-center gap-2 hover:bg-slate-50 p-1 rounded cursor-pointer">
                        <input
                          type="checkbox"
                          checked={!hiddenColumns.no}
                          onChange={(e) => setHiddenColumns(prev => ({ ...prev, no: !e.target.checked }))}
                          className="rounded border-slate-300 text-amber-600 focus:ring-amber-500 w-3.5 h-3.5"
                        />
                        <span className="text-xs font-medium text-slate-700">No.</span>
                      </label>
                      <label className="flex items-center gap-2 hover:bg-slate-50 p-1 rounded cursor-pointer">
                        <input
                          type="checkbox"
                          checked={!hiddenColumns.business}
                          onChange={(e) => setHiddenColumns(prev => ({ ...prev, business: !e.target.checked }))}
                          className="rounded border-slate-300 text-amber-600 focus:ring-amber-500 w-3.5 h-3.5"
                        />
                        <span className="text-xs font-medium text-slate-700">Business</span>
                      </label>
                      <label className="flex items-center gap-2 hover:bg-slate-50 p-1 rounded cursor-pointer">
                        <input
                          type="checkbox"
                          checked={!hiddenColumns.charge}
                          onChange={(e) => setHiddenColumns(prev => ({ ...prev, charge: !e.target.checked }))}
                          className="rounded border-slate-300 text-amber-600 focus:ring-amber-500 w-3.5 h-3.5"
                        />
                        <span className="text-xs font-medium text-slate-700">L07</span>
                      </label>
                      <label className="flex items-center gap-2 hover:bg-slate-50 p-1 rounded cursor-pointer">
                        <input
                          type="checkbox"
                          checked={!hiddenColumns.month}
                          onChange={(e) => setHiddenColumns(prev => ({ ...prev, month: !e.target.checked }))}
                          className="rounded border-slate-300 text-amber-600 focus:ring-amber-500 w-3.5 h-3.5"
                        />
                        <span className="text-xs font-medium text-slate-700">Tháng</span>
                      </label>
                      {safeTypeColumns.map((type, idx) => (
                        <label key={type} className="flex items-center gap-2 hover:bg-slate-50 p-1 rounded cursor-pointer">
                          <input
                            type="checkbox"
                            checked={!isTypeColHidden(type, idx)}
                            onChange={(e) => setHiddenColumns(prev => ({ ...prev, [`type_${type}`]: !e.target.checked }))}
                            className="rounded border-slate-300 text-amber-600 focus:ring-amber-500 w-3.5 h-3.5"
                          />
                          <span className="text-xs font-medium text-slate-700">{type}</span>
                        </label>
                      ))}
                      <label className="flex items-center gap-2 hover:bg-slate-50 p-1 rounded cursor-pointer">
                        <input
                          type="checkbox"
                          checked={!hiddenColumns.grandTotal}
                          onChange={(e) => setHiddenColumns(prev => ({ ...prev, grandTotal: !e.target.checked }))}
                          className="rounded border-slate-300 text-amber-600 focus:ring-amber-500 w-3.5 h-3.5"
                        />
                        <span className="text-xs font-medium text-slate-700">TỔNG CỘNG</span>
                      </label>
                    </div>
                  </div>

                  {diagnosticLogs.length > 0 && (
                    <div className="border-t border-slate-100 pt-2">
                      <button
                        onClick={handleDownloadDiagnosticCSV}
                        className="w-full flex items-center justify-center gap-1.5 px-2.5 py-1.5 bg-rose-50 hover:bg-rose-100 text-rose-700 text-xs font-medium rounded-lg border border-rose-200 transition-colors cursor-pointer"
                      >
                        <Download className="w-3.5 h-3.5" />
                        <span>Tải CSV Log lỗi ({diagnosticLogs.length})</span>
                      </button>
                    </div>
                  )}
                </div>
              )}
            </div>
          </div>
        </div>
      </div>

      {/* BU Filter Bar */}
      <div
        className="bu-filter-bar flex items-center justify-between gap-2 px-3 border-b border-[var(--border)] flex-none overflow-x-auto select-none z-10 h-9 min-h-9 max-h-9"
        style={{
          backgroundColor: "var(--table-sub-header-bg, #EDE4DB)",
          height: "36px",
          minHeight: "36px",
          maxHeight: "36px",
          flex: "0 0 36px"
        }}
      >
        <div className="flex items-center gap-1.5 flex-wrap">
          <span className="flex items-center gap-1 text-[10px] font-black uppercase tracking-wider text-muted-foreground mr-1">
            <Filter className="w-3 h-3 text-primary" />
            <span>Lọc BU:</span>
          </span>
          {/* Nút Tất cả */}
          <button
            type="button"
            onClick={() => handleSetBuFilterMode("ALL")}
            className={`px-2.5 py-0.5 rounded-md text-[10.5px] font-bold transition-all cursor-pointer active:scale-95 whitespace-nowrap ${
              buFilterMode === "ALL"
                ? "bg-primary text-primary-foreground shadow-2xs font-black"
                : "bg-background hover:bg-muted text-foreground border border-border/80"
            }`}
          >
            Tất cả
          </button>
          {/* Nút Trừ AHP */}
          <button
            type="button"
            onClick={() =>
              handleSetBuFilterMode(buFilterMode === "EXCLUDE_AHP" ? "ALL" : "EXCLUDE_AHP")
            }
            className={`px-2.5 py-0.5 rounded-md text-[10.5px] font-bold transition-all cursor-pointer active:scale-95 flex items-center gap-1 whitespace-nowrap ${
              buFilterMode === "EXCLUDE_AHP"
                ? "bg-amber-600 text-white shadow-2xs ring-1 ring-amber-600 font-black"
                : "bg-amber-50/80 hover:bg-amber-100/80 dark:bg-amber-950/40 text-amber-800 dark:text-amber-300 border border-amber-300 dark:border-amber-700/60"
            }`}
            title="Lọc toàn bộ bảng trừ BU AHP (Hải Phòng)"
          >
            <span>Trừ AHP</span>
            {buFilterMode === "EXCLUDE_AHP" && (
              <span className="text-[8.5px] font-black bg-white/25 px-1 rounded">Đang lọc</span>
            )}
          </button>
        </div>
        
        {/* Counter & quick reset */}
        <div className="flex items-center gap-2 text-[10.5px] font-medium text-muted-foreground ml-auto shrink-0 whitespace-nowrap">
          <span className="tabular-nums">
            {buFilterMode !== "ALL" ? (
              <span>
                Đang lọc BU: <strong className="text-primary font-bold">{sortedFlatRows.length}</strong> dòng
              </span>
            ) : (
              <span>{sortedFlatRows.length} dòng</span>
            )}
          </span>
          {buFilterMode !== "ALL" && (
            <button
              type="button"
              onClick={() => handleSetBuFilterMode("ALL")}
              className="text-primary hover:text-primary/80 font-bold underline ml-2 transition-colors cursor-pointer"
            >
              Bỏ lọc
            </button>
          )}
        </div>
      </div>

      {/* MAIN DATA TABLE */}
      <div 
        className="table-body-region flex min-h-0 flex-1 flex-col overflow-hidden bg-[var(--table-data-bg,var(--card,#fff))]"
      >
        <div className="relative flex-1 overflow-auto custom-scrollbar">
          <table
            className="pivot-master-table border-separate border-spacing-0 select-none table-fixed bg-[var(--table-data-bg,var(--card,#fff))] text-left text-xs"
            style={{
              fontFamily: "var(--font-table, var(--font-main))",
              fontSize: "13px",
              width: `${visibleTableWidth}px`,
              minWidth: `${visibleTableWidth}px`,
            }}
          >
            <colgroup>
              {!hiddenColumns.no && <col style={{ width: `${columnWidths.no || 60}px` }} />}
              {!hiddenColumns.business && <col style={{ width: `${columnWidths.business || 90}px` }} />}
              {!hiddenColumns.charge && <col style={{ width: `${columnWidths.charge || 220}px` }} />}
              {!hiddenColumns.month && <col style={{ width: `${columnWidths.month || 90}px` }} />}
              {safeTypeColumns.map((type, idx) =>
                isTypeColHidden(type, idx) ? null : (
                  <col key={type} style={{ width: `${columnWidths[`type_${type}`] || 120}px` }} />
                ),
              )}
              {!hiddenColumns.grandTotal && <col style={{ width: `${columnWidths.grandTotal || 140}px` }} />}
            </colgroup>
            <thead className="sticky top-0 z-20 border-b border-border bg-[var(--table-column-header-bg,#F4ECD8)] font-bold text-[var(--table-column-header-text-color,var(--table-column-header-text,#1e293b))] shadow-2xs">
              <tr>
                {!hiddenColumns.no && (
                  <th 
                    onClick={() => toggleSort("no")}
                    className="group relative cursor-pointer align-middle border-r border-[var(--grid-line-color,rgba(0,0,0,0.035))] bg-[var(--table-column-header-bg,#F4ECD8)] px-2 py-2.5 text-center text-[10px] font-semibold tracking-wider text-[var(--table-column-header-text-color,var(--table-column-header-text,#1e293b))] transition-colors hover:bg-white/10"
                    style={{ width: columnWidths["no"] || 60, minWidth: columnWidths["no"] || 60, maxWidth: columnWidths["no"] || 60, textTransform: "none" }}
                    title="Nhấp để sắp xếp (Tăng dần → Giảm dần → Hủy sắp xếp)"
                  >
                    <div className="inline-flex items-center justify-center gap-1">
                      <span>No.</span>
                      {renderSortIndicator("no")}
                      <button
                        type="button"
                        onClick={(event) => {
                          event.stopPropagation();
                          autoFitAllColumns();
                        }}
                        onMouseDown={(event) => event.stopPropagation()}
                        className="inline-flex h-5 w-5 items-center justify-center rounded-full text-current opacity-70 group-hover:opacity-100 hover:bg-white/20 transition-all cursor-pointer"
                        title="Căn độ rộng tất cả cột theo dữ liệu"
                        aria-label="Căn độ rộng tất cả cột theo dữ liệu"
                      >
                        <Maximize2 className="w-3 h-3 text-current" />
                      </button>
                    </div>
                    <div
                      onMouseDown={(e) => handleResizeStart(e, "no", 60)}
                      className="absolute right-0 top-0 bottom-0 w-1.5 cursor-col-resize hover:bg-amber-400 opacity-0 hover:opacity-100 transition-opacity"
                    />
                  </th>
                )}

                {!hiddenColumns.business && (
                  <th 
                    style={{ width: columnWidths["business"] || 90, minWidth: columnWidths["business"] || 90, maxWidth: columnWidths["business"] || 90 }}
                    onClick={() => toggleSort("bu")}
                    className="relative cursor-pointer border-r border-[var(--grid-line-color,rgba(0,0,0,0.035))] bg-[var(--table-column-header-bg,#F4ECD8)] px-2.5 py-2.5 text-center text-[10px] font-bold uppercase tracking-wider text-[var(--table-column-header-text-color,var(--table-column-header-text,#1e293b))] transition-colors hover:bg-white/10"
                    title="Nhấp để sắp xếp (Tăng dần → Giảm dần → Hủy sắp xếp)"
                  >
                    <div className="inline-flex items-center justify-center gap-1">
                      <span>Business</span>
                      {renderSortIndicator("bu")}
                    </div>
                    <div
                      onMouseDown={(e) => handleResizeStart(e, "business", 90)}
                      className="absolute right-0 top-0 bottom-0 w-1.5 cursor-col-resize hover:bg-amber-400 opacity-0 hover:opacity-100 transition-opacity"
                    />
                  </th>
                )}

                {!hiddenColumns.charge && (
                  <th 
                    style={{ width: columnWidths["charge"] || 220, minWidth: columnWidths["charge"] || 220, maxWidth: columnWidths["charge"] || 220 }}
                    onClick={() => toggleSort("l07")}
                    className="relative cursor-pointer border-r border-[var(--grid-line-color,rgba(0,0,0,0.035))] bg-[var(--table-column-header-bg,#F4ECD8)] px-2.5 py-2.5 text-left text-[10px] font-bold uppercase tracking-wider text-[var(--table-column-header-text-color,var(--table-column-header-text,#1e293b))] transition-colors hover:bg-white/10"
                    title="Nhấp để sắp xếp (Tăng dần → Giảm dần → Hủy sắp xếp)"
                  >
                    <div className="inline-flex items-center gap-1">
                      <span>L07</span>
                      {renderSortIndicator("l07")}
                    </div>
                    <div
                      onMouseDown={(e) => handleResizeStart(e, "charge", 220)}
                      className="absolute right-0 top-0 bottom-0 w-1.5 cursor-col-resize hover:bg-amber-400 opacity-0 hover:opacity-100 transition-opacity"
                    />
                  </th>
                )}

                {!hiddenColumns.month && (
                  <th 
                    style={{ width: columnWidths["month"] || 90, minWidth: columnWidths["month"] || 90, maxWidth: columnWidths["month"] || 90 }}
                    onClick={() => toggleSort("month")}
                    className="relative cursor-pointer border-r border-[var(--grid-line-color,rgba(0,0,0,0.035))] bg-[var(--table-column-header-bg,#F4ECD8)] px-2.5 py-2.5 text-center text-[10px] font-bold uppercase tracking-wider text-[var(--table-column-header-text-color,var(--table-column-header-text,#1e293b))] transition-colors hover:bg-white/10"
                    title="Nhấp để sắp xếp (Tăng dần → Giảm dần → Hủy sắp xếp)"
                  >
                    <div className="inline-flex items-center justify-center gap-1">
                      <span>Tháng</span>
                      {renderSortIndicator("month")}
                    </div>
                    <div
                      onMouseDown={(e) => handleResizeStart(e, "month", 90)}
                      className="absolute right-0 top-0 bottom-0 w-1.5 cursor-col-resize hover:bg-amber-400 opacity-0 hover:opacity-100 transition-opacity"
                    />
                  </th>
                )}

                {safeTypeColumns.map((type, idx) => {
                  if (isTypeColHidden(type, idx)) return null;
                  const colKey = `type_${type}`;
                  const w = columnWidths[colKey] || 120;
                  return (
                    <th 
                      key={type}
                      style={{ width: w, minWidth: w, maxWidth: w }}
                      onClick={() => toggleSort(colKey)}
                      className="relative cursor-pointer border-r border-[var(--grid-line-color,rgba(0,0,0,0.035))] bg-[var(--table-column-header-bg,#F4ECD8)] px-2.5 py-2.5 text-right text-[10px] font-bold uppercase tracking-wider text-[var(--table-column-header-text-color,var(--table-column-header-text,#1e293b))] transition-colors hover:bg-white/10"
                      title="Nhấp để sắp xếp (Tăng dần → Giảm dần → Hủy sắp xếp)"
                    >
                      <div className="flex w-full min-w-0 items-center justify-end gap-1">
                        <span className="min-w-0 whitespace-normal break-words text-right leading-tight" title={type}>{type}</span>
                        {renderSortIndicator(colKey)}
                      </div>
                      <div
                        onMouseDown={(e) => handleResizeStart(e, colKey, 120)}
                        className="absolute right-0 top-0 bottom-0 w-1.5 cursor-col-resize hover:bg-amber-400 opacity-0 hover:opacity-100 transition-opacity"
                      />
                    </th>
                  );
                })}

                {!hiddenColumns.grandTotal && (
                  <th 
                    style={{ width: columnWidths["grandTotal"] || 140, minWidth: columnWidths["grandTotal"] || 140, maxWidth: columnWidths["grandTotal"] || 140 }}
                    onClick={() => toggleSort("rowTotal")}
                    className="relative cursor-pointer bg-[var(--table-column-header-bg,#F4ECD8)] px-3 py-2.5 text-right text-[10px] font-extrabold uppercase tracking-wider text-[var(--table-column-header-text-color,var(--table-column-header-text,#1e293b))] transition-colors hover:bg-white/15"
                    title="Nhấp để sắp xếp (Tăng dần → Giảm dần → Hủy sắp xếp)"
                  >
                    <div className="inline-flex items-center justify-end gap-1 w-full">
                      <span>TỔNG CỘNG</span>
                      {renderSortIndicator("rowTotal")}
                    </div>
                    <div
                      onMouseDown={(e) => handleResizeStart(e, "grandTotal", 140)}
                      className="absolute right-0 top-0 bottom-0 w-1.5 cursor-col-resize hover:bg-amber-400 opacity-0 hover:opacity-100 transition-opacity"
                    />
                  </th>
                )}
              </tr>
            </thead>
            <tbody>
              {renderRows()}
            </tbody>

            {/* GRAND TOTAL FOOTER ROW */}
            {paginatedRows.length > 0 && (
              <tfoot className="sticky bottom-0 z-10 border-t border-border bg-[var(--table-column-header-bg,#F4ECD8)] font-black text-[var(--table-column-header-text-color,var(--table-column-header-text,inherit))] shadow-sm">
                <tr className="total-row">
                  {pivotLabelColumnSpan > 0 && (
                    <td
                      colSpan={pivotLabelColumnSpan}
                      className="border-r-0 border-l-0 bg-[var(--table-column-header-bg,#F4ECD8)] px-2.5 py-2.5 text-left font-black text-[var(--table-column-header-text-color,var(--table-column-header-text,inherit))]"
                    >
                      TỔNG CỘNG TẤT CẢ
                    </td>
                  )}
                  {grandTotals.map((v, idx) => {
                    const type = safeTypeColumns[idx];
                    if (isTypeColHidden(type, idx)) return null;
                    const colKey = `type_${type}`;
                    const w = columnWidths[colKey] || 120;
                    return (
                      <td 
                        key={idx} 
                        style={{ width: w, minWidth: w, maxWidth: w }}
                        className="border-r-0 border-l-0 bg-[var(--table-column-header-bg,#F4ECD8)] px-2.5 py-2.5 text-right text-xs font-black tabular-nums text-[var(--table-column-header-text-color,var(--table-column-header-text,inherit))]"
                      >
                        {v ? formatNumber(v) : "0"}
                      </td>
                    );
                  })}
                  {!hiddenColumns.grandTotal && (
                    <td 
                      style={{ width: columnWidths["grandTotal"] || 140, minWidth: columnWidths["grandTotal"] || 140, maxWidth: columnWidths["grandTotal"] || 140 }}
                      className="border-r-0 border-l-0 bg-[var(--table-column-header-bg,#F4ECD8)] px-3 py-2.5 text-right text-xs font-black tabular-nums text-[var(--table-column-header-text-color,var(--table-column-header-text,inherit))]"
                    >
                      {superGrandTotal ? formatNumber(superGrandTotal) : "0"}
                    </td>
                  )}
                </tr>
              </tfoot>
            )}
          </table>
        </div>

        {/* PAGINATION FOOTER — đồng bộ với các bảng Master */}
        <div
          className="table-footer-pagination unified-table-frame-footer flex h-[52px] shrink-0 items-center justify-between gap-3 border-t border-border bg-[var(--table-footer-bg,var(--table-header-bg,#FAF3E8))] px-3 py-1.5 text-muted-foreground"
        >
          <div className="flex flex-wrap items-center gap-3">
            <div className="flex items-center gap-1.5">
              <span className="whitespace-nowrap text-[12px] font-medium text-muted-foreground">
                Hiển thị:
              </span>
              <Select
                value={rowsPerPage === Infinity ? "all" : String(rowsPerPage)}
                onValueChange={(value) => {
                  setRowsPerPage(
                    value === "all" ? Infinity : Number(value),
                  );
                  setCurrentPage(1);
                }}
              >
                <SelectTrigger
                  className="h-[22px] w-[100px] rounded-full border-border bg-card px-2.5 py-0 text-[12px] font-bold normal-case text-foreground shadow-2xs transition-colors hover:bg-muted/60"
                  style={{ width: "100px", height: "22px", fontSize: "12px" }}
                >
                  <SelectValue placeholder="Chọn..." className="text-[12px]" />
                </SelectTrigger>
                <SelectContent className="z-[99999] border-border bg-popover font-sans opacity-100">
                  <SelectItem value="10" className="font-sans text-[12px] font-medium normal-case">10 dòng</SelectItem>
                  <SelectItem value="20" className="font-sans text-[12px] font-medium normal-case">20 dòng</SelectItem>
                  <SelectItem value="50" className="font-sans text-[12px] font-medium normal-case">50 dòng</SelectItem>
                  <SelectItem value="100" className="font-sans text-[12px] font-medium normal-case">100 dòng</SelectItem>
                  <SelectItem value="all" className="font-sans text-[12px] font-medium normal-case">Tất cả</SelectItem>
                </SelectContent>
              </Select>
            </div>

            <span className="whitespace-nowrap border-l border-border pl-3 text-[12px] font-medium text-muted-foreground">
              {totalRowsCount === 0
                ? "0 dòng"
                : `Hiển thị ${startIndex + 1} - ${endIndex} / ${totalRowsCount} dòng`}
            </span>

          </div>

          <div className="flex h-6 items-center gap-1 border-l border-border pl-4">
            <button
              type="button"
              onClick={() => setCurrentPage(1)}
              disabled={validCurrentPage <= 1}
              className="flex h-7 w-7 cursor-pointer select-none items-center justify-center rounded-full border border-border bg-card text-foreground/70 shadow-3xs transition-all hover:bg-muted/60 hover:text-foreground active:scale-95 disabled:cursor-not-allowed disabled:opacity-30"
              title="Trang đầu"
            >
              <ChevronsLeft className="h-3.5 w-3.5" />
            </button>
            <button
              type="button"
              onClick={() => setCurrentPage((page) => Math.max(1, page - 1))}
              disabled={validCurrentPage <= 1}
              className="flex h-7 w-7 cursor-pointer select-none items-center justify-center rounded-full border border-border bg-card text-foreground/70 shadow-3xs transition-all hover:bg-muted/60 hover:text-foreground active:scale-95 disabled:cursor-not-allowed disabled:opacity-30"
              title="Trang trước"
            >
              <ChevronLeft className="h-3.5 w-3.5" />
            </button>

            <span 
              className="min-w-[90px] whitespace-nowrap px-3 text-center font-display uppercase tracking-widest text-foreground/70"
              style={{ fontWeight: "normal", fontSize: "12px", lineHeight: "16px" }}
            >
              TRANG {validCurrentPage} / {totalPages}
            </span>

            <button
              type="button"
              onClick={() =>
                setCurrentPage((page) => Math.min(totalPages, page + 1))
              }
              disabled={validCurrentPage >= totalPages}
              className="flex h-7 w-7 cursor-pointer select-none items-center justify-center rounded-full border border-border bg-card text-foreground/70 shadow-3xs transition-all hover:bg-muted/60 hover:text-foreground active:scale-95 disabled:cursor-not-allowed disabled:opacity-30"
              title="Trang sau"
            >
              <ChevronRight className="h-3.5 w-3.5" />
            </button>
            <button
              type="button"
              onClick={() => setCurrentPage(totalPages)}
              disabled={validCurrentPage >= totalPages}
              className="flex h-7 w-7 cursor-pointer select-none items-center justify-center rounded-full border border-border bg-card text-foreground/70 shadow-3xs transition-all hover:bg-muted/60 hover:text-foreground active:scale-95 disabled:cursor-not-allowed disabled:opacity-30"
              title="Trang cuối"
            >
              <ChevronsRight className="h-3.5 w-3.5" />
            </button>
          </div>
        </div>
      </div>
      <ConfirmDialog
        isOpen={!!deleteMonthTarget}
        onClose={() => setDeleteMonthTarget(null)}
        onConfirm={confirmDeleteMonth}
        title={`Xóa dữ liệu Pivot Master tháng ${deleteMonthTarget || ""}?`}
        description="Thao tác áp dụng cho toàn bộ BU và Center trong Pivot Master của tháng này."
        confirmText="XÓA PIVOT THÁNG NÀY"
        variant="destructive"
      />
      <ConfirmDialog
        isOpen={!!deleteRowTarget}
        onClose={() => setDeleteRowTarget(null)}
        onConfirm={confirmDeleteRow}
        title="Xác nhận xóa dòng Pivot"
        description={`Bạn có chắc chắn muốn xóa dòng ${deleteRowTarget?.bu} - ${deleteRowTarget?.l07} (Tháng ${deleteRowTarget?.month}) khỏi bảng Pivot? Thao tác này không thể hoàn tác.`}
        confirmText="XÓA DÒNG"
        variant="destructive"
      />
    </div>
  );
}
