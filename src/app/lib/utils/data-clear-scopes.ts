import { cloneDefaultAllowedTaRules } from "./allowed-ta-rules";
import { INITIAL_APP_DATA } from "../../constants/initial-data";
import type { AppData, TableData } from "../../types";

export type AuditTableKey = "main" | "detail" | "rules";

const emptyTable = (table: TableData): TableData => ({
  headers: [...(table?.headers || [])],
  data: [],
});

export function clearMasterTableData(
  previous: AppData,
  table: "Sheet1_AE" | "Bank_North_AE" | "Hold_AE" | "BankExport",
): AppData {
  return {
    ...previous,
    [table]: emptyTable(previous[table]),
    ...(table === "Hold_AE" ? { HoldCarrySnapshots: {} } : {}),
    ...(table === "BankExport" ? { ReconciliationByMonth: {}, TransactionMonthCache: undefined, TransactionHistoryRestoreDisabled: true } : {}),
    updatedAt: new Date().toISOString(),
  };
}

export function clearMasterPageData(previous: AppData): AppData {
  return {
    ...previous,
    Ae_Global_Inputs: [],
    Final_AE: emptyTable(previous.Final_AE),
    Sheet1_AE: emptyTable(previous.Sheet1_AE),
    Hold_AE: emptyTable(previous.Hold_AE),
    Bank_North_AE: emptyTable(previous.Bank_North_AE),
    SoSanh_AE: emptyTable(previous.SoSanh_AE),
    BankExport: emptyTable(previous.BankExport),
    TransactionMonthCache: undefined,
    TransactionHistoryRestoreDisabled: true,
    ReconciliationByMonth: {},
    HoldCarrySnapshots: {},
    CustomReport: emptyTable(previous.CustomReport),
    Master_Roster: [],
    Master_RosterFileName: "",
    updatedAt: new Date().toISOString(),
  };
}

export function clearAuditTableData(
  previous: AppData,
  table: AuditTableKey,
): AppData {
  if (table === "rules") {
    return {
      ...previous,
      Q_AllowedTARules: [],
      updatedAt: new Date().toISOString(),
    };
  }

  return {
    ...previous,
    AuditReport: emptyTable(previous.AuditReport),
    AuditClearedTables: {
      ...(previous.AuditClearedTables || {}),
      [table]: true,
    },
    updatedAt: new Date().toISOString(),
  };
}

export function clearAuditPageData(previous: AppData): AppData {
  return {
    ...previous,
    AuditReport: emptyTable(previous.AuditReport),
    Q_TeacherHours: [],
    Q_TeacherHoursFileName: "",
    Q_CheckTAs: [],
    Q_CheckTAsFileName: "",
    Q_BonusData: [],
    Q_BonusSheetName: "",
    Q_AllowedTARules: [],
    AuditClearedTables: { main: true, detail: true, rules: true },
    updatedAt: new Date().toISOString(),
  };
}

export function clearBalancePageData(previous: AppData): AppData {
  return {
    ...previous,
    SavedBal_PayrollTrial: {},
    SavedPeriods_HoldAdd: {},
    SavedRows_HoldAdd: {},
    SavedRows_HoldAdd_Meta: {},
    ConfirmedIds_HoldAdd: [],
    TransactionActivity: undefined,
    TrialBalanceTransactionVersion: 0,
    TrialBalanceTransactionVersions: {},
    TrialBalanceRefreshedAt: undefined,
    updatedAt: new Date().toISOString(),
  };
}

export function createClearedWebData(previous: AppData): AppData {
  const cleanInitial = JSON.parse(JSON.stringify(INITIAL_APP_DATA)) as AppData;

  return {
    ...cleanInitial,
    TableOriginals: {},
    TransactionHistoryRestoreDisabled: true,
    globalMonth: previous.globalMonth,
    Timesheet_InputList: [],
    Timesheet_Roster: [],
    Timesheet_RosterFileName: "",
    Timesheet_RosterEditHistory: [],
    Master_Roster: [],
    Master_RosterFileName: "",
    Q_Staff: [],
    Q_Salary_Scale: [],
    Q_Cache: [],
    Q_CheckTAs: [],
    Q_CheckTAsFileName: "",
    Q_TeacherHours: [],
    Q_TeacherHoursFileName: "",
    Q_BonusData: [],
    Q_BonusSheetName: "",
    Q_AllowedTARules: cloneDefaultAllowedTaRules(),
    Timesheets: [],
    TA_Employee_Summary: { headers: [], data: [] },
    TA_Center_Summary: { headers: [], data: [] },
    SavedBal_PayrollTrial: {},
    SavedPeriods_HoldAdd: {},
    SavedRows_HoldAdd: {},
    SavedRows_HoldAdd_Meta: {},
    ConfirmedIds_HoldAdd: [],
    TransactionActivity: undefined,
    TrialBalanceTransactionVersion: 0,
    TrialBalanceTransactionVersions: {},
    TrialBalanceRefreshedAt: undefined,
    ReconciliationByMonth: {},
    HoldCarrySnapshots: {},
    AuditClearedTables: {},
    Timesheet_SkipSupabaseRestore: true,
    Timesheet_LocalClearedAt: new Date().toISOString(),
    lastSupabaseSyncAt: undefined,
    updatedAt: new Date().toISOString(),
  };
}
