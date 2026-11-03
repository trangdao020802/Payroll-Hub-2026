import { masterTableExportColumns } from "./master-table-columns";
import type { AppData } from "../../types";
import {
  accountOf,
  buildBulkPaymentAnalytics,
  classifyHoldOperation,
  comparePeriods,
  dimensionsOf,
  employeeIdOf,
  extractOccurrencePeriod,
  extractReportPeriod,
  formatPeriod,
  fullNameOf,
  moneyOf,
  parseMonthPeriod,
  periodFromParts,
} from "./bulk-payment-analytics";
import { parseMoneyToNumber, removeVietnameseTones } from "./data-utils";
import { hasRequiredDeductionsFields } from "./deductions-row-validation";
import {
  BANK_TRANSACTION_EXPORT_HEADERS,
  prepareTransactionBankExportRows,
  type WorkbookExportDefinition,
} from "./excel-export";
import { buildPivotFromAppData, getPivotSourceLabels, omitPivotMonths } from "./pivot-utils";
import {
  canonicalTransactionHeaders,
  withCanonicalTransactionDocumentId,
} from "./transaction-history";

type DataRow = Record<string, unknown>;

const readFirst = (row: DataRow, keys: string[]): unknown => {
  for (const key of keys) {
    const value = row?.[key];
    if (value !== undefined && value !== null && String(value).trim() !== "") {
      return value;
    }
  }
  return "";
};

const normalizeMonth = (value: unknown): string => {
  const raw = String(value ?? "").trim();
  const match = raw.match(/(?:THÁNG|THANG|T)?\s*(\d{1,2})[./_\- ](\d{4})/i);
  if (!match) return raw.toUpperCase();
  return `${String(Number(match[1])).padStart(2, "0")}.${match[2]}`;
};

const rowMonth = (row: DataRow): string =>
  normalizeMonth(
    readFirst(row, [
      "Tháng báo cáo",
      "THÁNG BÁO CÁO",
      "_fileMonth",
      "Tháng",
      "Month",
    ]),
  );

const matchesReportingMonth = (row: DataRow, period: string): boolean => {
  const value = rowMonth(row);
  return !value || value === normalizeMonth(period);
};

const rowId = (row: DataRow): string =>
  String(
    readFirst(row, [
      "ID Number",
      "Document ID",
      "Document ID / CCCD",
      "Mã nhân viên",
      "Mã NV",
      "Mã AE",
      "Mã ae",
      "CCCD",
    ]),
  )
    .trim()
    .toUpperCase();

const rowName = (row: DataRow): string =>
  removeVietnameseTones(
    String(
      readFirst(row, [
        "Full name",
        "Full Name",
        "Beneficiary Name",
        "Họ và tên",
        "Họ tên",
      ]),
    ),
  )
    .replace(/\s+/g, " ")
    .trim()
    .toUpperCase();

const rowAccount = (row: DataRow): string =>
  String(
    readFirst(row, [
      "Bank Account Number",
      "Beneficiary Account No.",
      "Account Number",
      "Số tài khoản",
      "STK",
    ]),
  )
    .replace(/\s+/g, "")
    .trim()
    .toUpperCase();

const rowBusiness = (row: DataRow): string =>
  String(readFirst(row, ["BU", "Business", "business"]) || "OTHER")
    .trim()
    .toUpperCase()
    .replace(/^AHN_HP$/, "AHP");

const rowL07 = (row: DataRow): string =>
  String(readFirst(row, ["L07", "l07", "Center", "Center Code"]))
    .trim()
    .toUpperCase();

const rowAmount = (row: DataRow): number =>
  parseMoneyToNumber(
    readFirst(row, [
      "Payment Amount",
      "TOTAL PAYMENT",
      "Total Payment",
      "Grand Total",
      "GRAND TOTAL",
      "Số tiền",
      "Thành tiền",
    ]),
  );

const sumRows = (rows: DataRow[]): number =>
  rows.reduce((sum, row) => sum + rowAmount(row), 0);

const sumField = (rows: DataRow[], field: string): number =>
  rows.reduce((sum, row) => sum + Number(row?.[field] || 0), 0);

const buildLifecycleExportRows = (
  holdRows: DataRow[],
  reportingMonth: string,
) => {
  const currentPeriod =
    parseMonthPeriod(reportingMonth) ||
    periodFromParts(new Date().getMonth() + 1, new Date().getFullYear())!;

  const details = holdRows.flatMap((row, index) => {
    const operation = classifyHoldOperation(row);
    if (!operation) return [];

    const reportPeriod = extractReportPeriod(row, currentPeriod);
    const occurrencePeriod = extractOccurrencePeriod(row, reportPeriod);
    if (
      comparePeriods(reportPeriod, currentPeriod) > 0 ||
      comparePeriods(occurrencePeriod, currentPeriod) > 0
    ) {
      return [];
    }

    const dimensions = dimensionsOf(row);
    return [
      {
        "No.": index + 1,
        BU: dimensions.business,
        "Tháng phát sinh": formatPeriod(occurrencePeriod),
        "Kỳ báo cáo": formatPeriod(reportPeriod),
        "ID Number": employeeIdOf(row),
        "Full Name": fullNameOf(row),
        L07: dimensions.l07,
        "Bank Account Number": accountOf(row),
        "Nghiệp vụ": operation,
        Amount: Math.abs(moneyOf(row)),
        "Sheet Source": readFirst(row, ["Sheet Source", "SHEET SOURCE"]),
        Note: readFirst(row, ["Note", "Ghi chú", "GHI CHÚ"]),
      },
    ];
  });

  const employees = new Map<
    string,
    {
      BU: string;
      occurrenceMonth: string;
      idNumber: string;
      fullName: string;
      l07: string;
      hold: number;
      paid: number;
      cancel: number;
    }
  >();
  details.forEach((row) => {
    const key = [
      row.BU,
      row["Tháng phát sinh"],
      row["ID Number"] || row["Full Name"],
    ].join("||");
    const current = employees.get(key) || {
      BU: String(row.BU || ""),
      occurrenceMonth: String(row["Tháng phát sinh"] || ""),
      idNumber: String(row["ID Number"] || ""),
      fullName: String(row["Full Name"] || ""),
      l07: String(row.L07 || ""),
      hold: 0,
      paid: 0,
      cancel: 0,
    };
    const amount = Number(row.Amount || 0);
    if (row["Nghiệp vụ"] === "HOLD") current.hold += amount;
    if (row["Nghiệp vụ"] === "ADD") current.paid += amount;
    if (row["Nghiệp vụ"] === "CANCEL") current.cancel += amount;
    employees.set(key, current);
  });

  const employeeRows = Array.from(employees.values()).map((row, index) => ({
    "No.": index + 1,
    BU: row.BU,
    "Tháng phát sinh": row.occurrenceMonth,
    "ID Number": row.idNumber,
    "Full Name": row.fullName,
    L07: row.l07,
    "Tổng HOLD gốc": row.hold,
    "Tổng đã thanh toán": row.paid,
    "Tổng CANCEL": row.cancel,
    "Số dư HOLD còn lại": Math.max(0, row.hold - row.paid - row.cancel),
  }));

  return { details, employeeRows };
};

interface ReconciliationAggregate {
  identity: string;
  idNumber: string;
  fullName: string;
  bankAccount: string;
  business: string;
  l07: string;
  grossPay: number;
  deductions: number;
  actual: number;
  transactionCount: number;
}

const buildReconciliationRows = (
  grossRows: DataRow[],
  deductionRows: DataRow[],
  transactionRows: DataRow[],
) => {
  const sourceByIdentity = new Map<string, ReconciliationAggregate>();
  const accountToIdentity = new Map<string, string>();
  const nameToIdentity = new Map<string, string>();

  const sourceIdentity = (row: DataRow, fallback: string): string => {
    const id = rowId(row);
    if (id) return `ID:${id}`;
    const account = rowAccount(row);
    if (account) return `ACCOUNT:${account}`;
    const name = rowName(row);
    return name ? `NAME:${name}` : fallback;
  };

  const addSourceRows = (rows: DataRow[], field: "grossPay" | "deductions") => {
    rows.forEach((row, index) => {
      const identity = sourceIdentity(row, `${field}:${index}`);
      const current = sourceByIdentity.get(identity) || {
        identity,
        idNumber: rowId(row),
        fullName: String(
          readFirst(row, ["Full name", "Full Name", "Beneficiary Name"]),
        ).trim(),
        bankAccount: rowAccount(row),
        business: rowBusiness(row),
        l07: rowL07(row),
        grossPay: 0,
        deductions: 0,
        actual: 0,
        transactionCount: 0,
      };
      current[field] += rowAmount(row);
      current.idNumber ||= rowId(row);
      current.fullName ||= String(
        readFirst(row, ["Full name", "Full Name", "Beneficiary Name"]),
      ).trim();
      current.bankAccount ||= rowAccount(row);
      current.business ||= rowBusiness(row);
      current.l07 ||= rowL07(row);
      sourceByIdentity.set(identity, current);
      if (current.bankAccount) accountToIdentity.set(current.bankAccount, identity);
      const normalizedName = rowName(row);
      if (normalizedName) nameToIdentity.set(normalizedName, identity);
    });
  };

  addSourceRows(grossRows, "grossPay");
  addSourceRows(deductionRows, "deductions");

  transactionRows.forEach((row, index) => {
    const id = rowId(row);
    const account = rowAccount(row);
    const normalizedName = rowName(row);
    const identity =
      (id && sourceByIdentity.has(`ID:${id}`) ? `ID:${id}` : "") ||
      accountToIdentity.get(account) ||
      nameToIdentity.get(normalizedName) ||
      (id ? `ID:${id}` : account ? `ACCOUNT:${account}` : normalizedName ? `NAME:${normalizedName}` : `BANK:${index}`);
    const current = sourceByIdentity.get(identity) || {
      identity,
      idNumber: id,
      fullName: String(readFirst(row, ["Beneficiary Name", "Full name"])).trim(),
      bankAccount: account,
      business: rowBusiness(row),
      l07: rowL07(row),
      grossPay: 0,
      deductions: 0,
      actual: 0,
      transactionCount: 0,
    };
    current.actual += rowAmount(row);
    current.transactionCount += 1;
    current.idNumber ||= id;
    current.fullName ||= String(readFirst(row, ["Beneficiary Name", "Full name"])).trim();
    current.bankAccount ||= account;
    sourceByIdentity.set(identity, current);
  });

  const details = Array.from(sourceByIdentity.values())
    .map((item, index) => {
      const expected = item.grossPay + item.deductions;
      const variance = item.actual - expected;
      return {
        "No.": index + 1,
        BU: item.business || "OTHER",
        L07: item.l07,
        "ID Number": item.idNumber,
        "Full Name": item.fullName,
        "Bank Account Number": item.bankAccount,
        "Gross Pay": item.grossPay,
        Deductions: item.deductions,
        "Expected Payment": expected,
        "Actual Transaction": item.actual,
        Variance: variance,
        "Transaction Count": item.transactionCount,
        Status: Math.abs(variance) < 1 ? "MATCHED" : "VARIANCE",
      };
    })
    .sort((left, right) =>
      String(left.BU).localeCompare(String(right.BU)) ||
      String(left["ID Number"]).localeCompare(String(right["ID Number"])),
    );

  const byBusiness = new Map<
    string,
    {
      grossPay: number;
      deductions: number;
      expected: number;
      actual: number;
      variance: number;
      transactions: number;
    }
  >();
  details.forEach((row) => {
    const business = String(row.BU || "OTHER");
    const current = byBusiness.get(business) || {
      grossPay: 0,
      deductions: 0,
      expected: 0,
      actual: 0,
      variance: 0,
      transactions: 0,
    };
    current.grossPay += row["Gross Pay"];
    current.deductions += row.Deductions;
    current.expected += row["Expected Payment"];
    current.actual += row["Actual Transaction"];
    current.variance += row.Variance;
    current.transactions += row["Transaction Count"];
    byBusiness.set(business, current);
  });

  const byBusinessRows = Array.from(byBusiness.entries())
    .sort(([left], [right]) => left.localeCompare(right))
    .map(([business, totals], index) => ({
      "No.": index + 1,
      BU: business,
      "Gross Pay": totals.grossPay,
      Deductions: totals.deductions,
      "Expected Payment": totals.expected,
      "Actual Transaction": totals.actual,
      Variance: totals.variance,
      "Transaction Count": totals.transactions,
      Status: Math.abs(totals.variance) < 1 ? "MATCHED" : "VARIANCE",
    }));

  return { details, byBusinessRows };
};

const buildPivotRows = (
  grossRows: DataRow[],
  rosterRows: DataRow[],
  reportingMonth: string,
  clearedMonths: readonly string[],
) => {
  const result = buildPivotFromAppData(
    grossRows,
    [],
    rosterRows,
    reportingMonth,
  );
  const groupedData = omitPivotMonths(result.groupedData, clearedMonths);
  const typeColumns = result.typeColumns || [];
  const rows: DataRow[] = [];
  const grandTotals = Object.fromEntries(typeColumns.map((type) => [type, 0]));
  let grandTotal = 0;
  let rowNumber = 1;

  Object.keys(groupedData)
    .sort()
    .forEach((business) => {
      const businessTotals = Object.fromEntries(
        typeColumns.map((type) => [type, 0]),
      );
      let businessTotal = 0;
      Object.keys(groupedData[business] || {})
        .sort()
        .forEach((l07) => {
          Object.keys(groupedData[business][l07] || {})
            .sort()
            .forEach((month) => {
              const bucket = groupedData[business][l07][month];
              const sourceLabels = getPivotSourceLabels(bucket);
              const row: DataRow = {
                "No.": rowNumber,
                Business: business,
                L07:
                  sourceLabels.length > 0
                    ? `${l07} — ${sourceLabels.join(" / ")}`
                    : l07,
                Month: month,
              };
              let total = 0;
              typeColumns.forEach((type) => {
                const amount = Number(bucket[type] || 0);
                row[type] = amount;
                businessTotals[type] += amount;
                grandTotals[type] += amount;
                total += amount;
              });
              row["Grand Total"] = total;
              rows.push(row);
              rowNumber += 1;
              businessTotal += total;
              grandTotal += total;
            });
        });
      rows.push({
        "No.": "",
        Business: business,
        L07: `${business} Total`,
        Month: "",
        ...businessTotals,
        "Grand Total": businessTotal,
      });
    });

  rows.push({
    "No.": "",
    Business: "",
    L07: "GRAND TOTAL",
    Month: "",
    ...grandTotals,
    "Grand Total": grandTotal,
  });
  return { rows, grandTotal };
};

export function createMasterExportDefinition(
  appData: AppData,
): WorkbookExportDefinition {
  const reportingMonth = normalizeMonth(appData.globalMonth || "03.2026");
  const grossRows = (appData.Sheet1_AE?.data || []).filter(
    (row) => rowId(row) && matchesReportingMonth(row, reportingMonth),
  );
  const allValidDeductions = (appData.Hold_AE?.data || []).filter(
    hasRequiredDeductionsFields,
  );
  const deductionRows = allValidDeductions.filter((row) => {
    const operation = String(row?.["Nghiệp vụ"] || "").toUpperCase();
    const source = String(row?.["Sheet Source"] || "").toUpperCase();
    return (
      matchesReportingMonth(row, reportingMonth) &&
      !operation.includes("BONUS") &&
      !source.includes("BONUS")
    );
  });
  const rawTransactions =
    appData.BankExport?.data?.length > 0
      ? appData.BankExport.data
      : appData.Bank_North_AE?.data || [];
  const transactionRows = rawTransactions.map(
    withCanonicalTransactionDocumentId,
  );
  const transactionExportRows = prepareTransactionBankExportRows(transactionRows);
  const reconciliation = buildReconciliationRows(
    grossRows,
    deductionRows,
    rawTransactions,
  );
  const expectedTotal = sumRows(grossRows) + sumRows(deductionRows);
  const actualTotal = sumRows(rawTransactions);
  const analytics = buildBulkPaymentAnalytics({
    sheet1Rows: grossRows,
    holdRows: allValidDeductions,
    bankRows: rawTransactions,
    globalMonth: reportingMonth,
  });
  const lifecycle = buildLifecycleExportRows(
    allValidDeductions,
    reportingMonth,
  );
  const pivot = buildPivotRows(
    grossRows,
    (appData.Master_Roster || []).filter((row: DataRow) => {
      const source = String(row?._sourceFile || "").toUpperCase();
      const id = String(row?._rowId || "").toLowerCase();
      return source !== "MOCK_ROSTER.XLSX" && !id.startsWith("mock-row-");
    }),
    reportingMonth,
    appData.PivotMasterClearedMonths || [],
  );

  const dateStamp = new Date().toISOString().slice(0, 10);
  return {
    title: `MASTER · ${reportingMonth}`,
    fileName: `Payroll_Hub_Master_${reportingMonth.replace(".", "_")}_${dateStamp}.xlsx`,
    pages: [
      {
        title: "Gross Pay",
        children: [
          {
            title: "Gross Pay Details",
            sheetName: "Gross Pay",
            table: {
              storageKey: "master-ae-Sheet1_AE",
              columns: masterTableExportColumns(appData.Sheet1_AE, "Sheet1_AE"),
              rows: grossRows,
              cards: [
                { label: "Reporting Month", value: reportingMonth },
                { label: "Employees", value: grossRows.length },
                { label: "Total Gross Pay", value: sumRows(grossRows) },
              ],
            },
          },
        ],
      },
      {
        title: "Deductions",
        children: [
          {
            title: "Deductions & Benefits",
            sheetName: "Deductions",
            table: {
              storageKey: "master_ae_Hold_AE",
              columns: masterTableExportColumns(appData.Hold_AE, "Hold_AE"),
              rows: deductionRows,
              cards: [
                { label: "Reporting Month", value: reportingMonth },
                { label: "Rows", value: deductionRows.length },
                { label: "Total Deductions", value: sumRows(deductionRows) },
              ],
            },
          },
        ],
      },
      {
        title: "Bulk Payment",
        children: [
          {
            title: "Transaction",
            sheetName: "Transaction",
            table: {
              storageKey: "bulk_payment",
              columns: canonicalTransactionHeaders(
                appData.BankExport?.headers?.length
                  ? appData.BankExport.headers
                  : [...BANK_TRANSACTION_EXPORT_HEADERS],
              )
                .filter(
                  (key) =>
                    !key.startsWith("_") &&
                    !/^(id|uuid|rowid|recordid)$/i.test(key),
                )
                .map((key) => ({
                  key,
                  label: key,
                  type: key === "Payment Amount" ? "currency" : "text",
                })),
              rows: transactionExportRows,
              cards: [
                { label: "Reporting Month", value: reportingMonth },
                { label: "Transactions", value: transactionRows.length },
                { label: "Total Transaction", value: actualTotal },
              ],
            },
          },
          {
            title: "Reconciliation by BU",
            sheetName: "Reconciliation by BU",
            table: {
              rows: reconciliation.byBusinessRows,
              cards: [
                { label: "Expected Payment", value: expectedTotal },
                { label: "Actual Transaction", value: actualTotal },
                { label: "Variance", value: actualTotal - expectedTotal },
              ],
            },
            children: [
              {
                title: "Reconciliation Details",
                sheetName: "Reconciliation Details",
                table: {
                  storageKey: "master-reconcile",
                  headers: ["No.", "ID NUMBER", "FULL NAME", "Bank Acc No. from AE", "Bank Acc No. from ACC", "TOTAL BANK AE", "TOTAL BANK ACC", "Diff", "Process Sync", "Problems"],
                  rows: reconciliation.details.map(row => ({
                    "No.": row["No."], "ID NUMBER": row["ID Number"], "FULL NAME": row["Full Name"],
                    "Bank Acc No. from AE": row["Bank Account Number"], "Bank Acc No. from ACC": row["Bank Account Number"],
                    "TOTAL BANK AE": row["Actual Transaction"], "TOTAL BANK ACC": row["Expected Payment"],
                    Diff: row.Variance, "Process Sync": "", Problems: row.Status === "MATCHED" ? "" : "Chênh lệch số tiền",
                  })),
                  cards: [
                    {
                      label: "Employees",
                      value: reconciliation.details.length,
                    },
                    {
                      label: "Variance Rows",
                      value: reconciliation.details.filter(
                        (row) => Math.abs(Number(row.Variance || 0)) >= 1,
                      ).length,
                    },
                  ],
                },
              },
            ],
          },
          {
            title: "HOLD Lifecycle Analysis",
            sheetName: "HOLD Analysis",
            table: {
              headers: ["No.", "BU", "Tháng HOLD", "Tổng số dư HOLD", "Số dư HOLD đầu kỳ", "HOLD phát sinh", "Thanh toán HOLD tại kỳ", "CANCEL tại kỳ", "Các tháng đã thanh toán", "Số dư HOLD còn lại", "Trạng thái HOLD"],
              storageKey: "master-analysis",
              rows: analytics.summaryRows as unknown as DataRow[],
              cards: [
                { label: "Reporting Month", value: analytics.currentPeriod },
                { label: "Business Units", value: analytics.businessUnits.length },
                { label: "Lifecycle Rows", value: analytics.summaryRows.length },
                {
                  label: "Total Original HOLD",
                  value: sumField(
                    analytics.summaryRows as unknown as DataRow[],
                    "Tổng số dư HOLD",
                  ),
                },
                {
                  label: "Opening HOLD",
                  value: sumField(
                    analytics.summaryRows as unknown as DataRow[],
                    "Số dư HOLD đầu kỳ",
                  ),
                },
                {
                  label: "HOLD Incurred",
                  value: sumField(
                    analytics.summaryRows as unknown as DataRow[],
                    "HOLD phát sinh",
                  ),
                },
                {
                  label: "HOLD Paid",
                  value: sumField(
                    analytics.summaryRows as unknown as DataRow[],
                    "Thanh toán HOLD tại kỳ",
                  ),
                },
                {
                  label: "CANCEL",
                  value: sumField(
                    analytics.summaryRows as unknown as DataRow[],
                    "CANCEL tại kỳ",
                  ),
                },
                {
                  label: "Remaining HOLD",
                  value: sumField(
                    analytics.summaryRows as unknown as DataRow[],
                    "Số dư HOLD còn lại",
                  ),
                },
              ],
            },
            children: [
              {
                title: "HOLD Employee Details",
                sheetName: "HOLD Employee Details",
                table: {
                  rows: lifecycle.employeeRows,
                  cards: [
                    { label: "Employees", value: lifecycle.employeeRows.length },
                    {
                      label: "Remaining HOLD",
                      value: sumField(
                        lifecycle.employeeRows,
                        "Số dư HOLD còn lại",
                      ),
                    },
                  ],
                },
              },
              {
                title: "HOLD Transaction Details",
                sheetName: "HOLD Transaction Details",
                table: {
                  rows: lifecycle.details,
                  cards: [
                    { label: "Transactions", value: lifecycle.details.length },
                    {
                      label: "Amount",
                      value: sumField(lifecycle.details, "Amount"),
                    },
                  ],
                },
              },
            ],
          },
        ],
      },
      {
        title: "Pivot Master",
        children: [
          {
            title: "Cost Allocation by BU, L07 & Task Type",
            sheetName: "Pivot Master",
            table: {
              storageKey: "master-pivot",
              rows: pivot.rows,
              cards: [
                { label: "Reporting Month", value: reportingMonth },
                { label: "Allocation Rows", value: Math.max(pivot.rows.length - 1, 0) },
                { label: "Grand Total", value: pivot.grandTotal },
              ],
            },
          },
        ],
      },
    ],
  };
}
