import assert from "node:assert/strict";
import test, { type TestContext } from "node:test";
import * as XLSX from "xlsx";
import { INITIAL_APP_DATA } from "../src/app/constants/initial-data";
import { processMasterAEData, type MasterAEProcessingContext } from "../src/app/lib/utils/master-ae-processing";
import { parseMasterWorkbook } from "../src/app/workers/masterImport.worker";
import type { AppData } from "../src/app/types";
import { captureSavedDeductions, protectSavedDeductions, restoreSavedDeductions } from "../src/app/lib/utils/deductions-lock";
import { reopenMasterDeductionsMonth } from "../src/app/lib/utils/master-source-refresh";
import { calculateReconciliationTotals, syncReportingMonthReconciliation } from "../src/app/lib/utils/reconciliation-sync";
import { applyLocalTransactionVersions } from "../src/app/lib/utils/transaction-local-history";
import { restoreSavedTransactionMonths } from "../src/app/lib/utils/transaction-cloud-restore";
import { getTransactionMonthOriginal } from "../src/app/lib/utils/transaction-originals";
import { restoreTableOriginals, trackTableOriginals } from "../src/app/lib/utils/table-originals";

function importHarness(t: TestContext, file: File) {
  const storage = new Map<string, string>();
  const localStorage = {
    getItem: (key: string) => storage.get(key) ?? null,
    setItem: (key: string, value: string) => storage.set(key, value),
  };
  const events = new EventTarget();
  const descriptors = new Map(["window", "localStorage"].map((key) => [key, Object.getOwnPropertyDescriptor(globalThis, key)]));
  Object.defineProperty(globalThis, "window", { configurable: true, value: events });
  Object.defineProperty(globalThis, "localStorage", { configurable: true, value: localStorage });
  t.after(() => {
    for (const [key, descriptor] of descriptors) {
      if (descriptor) Object.defineProperty(globalThis, key, descriptor);
      else Reflect.deleteProperty(globalThis, key);
    }
  });
  let data = structuredClone(INITIAL_APP_DATA);
  data.globalMonth = "08.2026";
  data.Ae_Global_Inputs = [{ id: "import-1", name: file.name, fileObj: file, status: "ready", bank: "NORTH", month: "08.2026" }];
  const processing: boolean[] = [];
  let completions = 0;
  const context: MasterAEProcessingContext = {
    appData: data,
    updateAppData: (updater, save, _persist, sources) => {
      data = trackTableOriginals(data, protectSavedDeductions(data, updater(data)), Boolean(save), sources);
    },
    preparedMasterFiles: new Map(),
    parseMasterFileInWorker: parseMasterWorkbook,
    masterAeFields: [],
    setIsProcessing: (value) => { processing.push(value); },
    setProgress: () => {},
    setProcessingMessage: () => {},
    onComplete: () => { completions++; },
  };
  return { context, storage, processing, data: () => data, completions: () => completions,
    setData(next: AppData) { data = next; context.appData = next; } };
}

const anna = "000000000001";
const bao = "000000000002";

function revisionWorkbook(name: string, bankIds: string[], holds: [string, number][]) {
  const workbook = XLSX.utils.book_new();
  const person = (id: string) => [id, id === anna ? "Nguyen Van An" : "Nguyen Van Bao", "Ocean Park"];
  const append = (sheet: string, rows: (string | number)[][]) =>
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(rows), sheet);
  append("Sheet 1", [["ID Number", "Full name", "Center", "TOTAL PAYMENT"],
    [...person(anna), 120000], [...person(bao), 200000]]);
  append("Bank North AE", [["ID Number", "Full name", "Center", "Bank Account Number", "TOTAL PAYMENT"],
    ...bankIds.map(id => [...person(id), id === anna ? "0012345678" : "0087654321", id === anna ? 100000 : 200000])]);
  append("HOLD T8", [["ID Number", "Full name", "Center", "TOTAL PAYMENT", "Note"],
    ...holds.map(([id, amount]) => [...person(id), amount, "Hold tháng 08.2026"])]);
  return new File([XLSX.write(workbook, {bookType: "xlsx", type: "array"})], name);
}

function historyVersion(period: string, id: string, rows: AppData["BankExport"]["data"]) {
  return {period: `${period}-01`, id, rows, row_count: rows.length, created_at: "2026-09-03T00:00:00Z"};
}

test("a corrected Master applies to the saved month only after unlock and survives old cloud history", async (t) => {
  const original = revisionWorkbook("NORTH 08.2026.xlsx", [anna, bao], [[anna, 20000]]);
  const corrected = revisionWorkbook("NORTH 08.2026 revised.xlsx", [anna], [[anna, 20000], [bao, 200000]]);
  const harness = importHarness(t, original);
  await processMasterAEData(harness.context);
  const august = historyVersion("2026-08", "1", harness.data().BankExport.data);
  const september = historyVersion("2026-09", "2", [{"Tháng báo cáo": "2026-09", "Document ID": "SEP", "Payment Amount": 91000,
    "Beneficiary Account No.": "", "Beneficiary Name": ""}]);
  const saved = captureSavedDeductions(syncReportingMonthReconciliation({
    ...applyLocalTransactionVersions(harness.data(), [august, september]),
    HoldCarrySnapshots: {"08.2026": {savedAt: august.created_at, eligibleCount: 1, carriedCount: 0}},
  }, "09.2026"));
  // Existing cloud imports predate the stable input identifier.
  for (const table of [saved.Sheet1_AE, saved.Bank_North_AE, saved.Hold_AE, saved.Hold_AE_Source]) {
    for (const row of table?.data || []) delete row._masterInputId;
  }
  const input = {...saved.Ae_Global_Inputs[0], name: corrected.name, previousSourceName: original.name, fileObj: corrected};
  harness.setData({...saved, Ae_Global_Inputs: [input]});
  await processMasterAEData(harness.context);

  const staged = harness.data();
  assert.equal(staged.globalMonth, "09.2026");
  assert.deepEqual(staged.BankExport.data, september.rows);
  assert.equal(staged.Hold_AE.data.length, 1);
  assert.equal(staged.Hold_AE_Source?.data.length, 2);
  assert.ok(staged.TransactionMonthCache?.pendingMaster?.["2026-08"]);
  assert.equal(staged.TransactionMonthCache?.months["2026-08"].table.data.length, 2);
  assert.equal(staged.Bank_North_AE.data.filter(row => row._fileMonth === "08.2026").length, 1);
  assert.equal(staged.Sheet1_AE.data.filter(row => row._fileMonth === "08.2026").length, 2);

  // An older build could store the revised source while leaving both working tables frozen.
  const legacy = structuredClone(staged);
  delete legacy.TransactionMonthCache?.pendingMaster;
  for (const row of legacy.Hold_AE_Source?.data || []) row._uploadTimestamp = "2026-09-04T00:00:00Z";
  harness.setData(legacy);
  await processMasterAEData(harness.context);
  assert.ok(harness.data().TransactionMonthCache?.pendingMaster?.["2026-08"]);

  const pulled = restoreSavedDeductions(restoreSavedTransactionMonths(structuredClone(harness.data()), [august, september]));
  const augustView = syncReportingMonthReconciliation(pulled, "08.2026");
  const reopened = protectSavedDeductions(augustView, reopenMasterDeductionsMonth(augustView, "08.2026"), {unlockDeductionsMonth: "08.2026"});
  assert.equal(reopened.HoldCarrySnapshots?.["08.2026"], undefined);
  assert.equal(reopened.TransactionMonthCache?.pendingMaster?.["2026-08"], undefined);
  assert.deepEqual(reopened.Hold_AE.data.map(row => row["ID Number"]).sort(), [anna, bao]);
  assert.deepEqual(reopened.BankExport.data.map(row => row["Document ID"]), [anna]);
  assert.equal(reopened.TransactionMonthCache?.months["2026-08"].masterDraft, true);
  assert.deepEqual(reopened.TransactionMonthCache?.months["2026-09"].table.data, september.rows);
  assert.equal(calculateReconciliationTotals(reopened, "08.2026").variance, 0);
  assert.equal(getTransactionMonthOriginal(reopened).data.length, 1);

  const stale = historyVersion("2026-08", "3", august.rows);
  const restored = restoreSavedTransactionMonths(structuredClone(reopened), [stale, september]);
  const returned = syncReportingMonthReconciliation(syncReportingMonthReconciliation(restored, "09.2026"), "08.2026");
  assert.deepEqual(returned.BankExport.data, reopened.BankExport.data);
  assert.deepEqual(getTransactionMonthOriginal(returned).data, reopened.BankExport.data);
  const reset = restoreTableOriginals(returned, ["BankExport"]);
  assert.equal(reset.TransactionMonthCache?.months["2026-08"].masterDraft, true);
  assert.deepEqual(restoreSavedTransactionMonths(reset, [stale]).BankExport.data, reopened.BankExport.data);
  const revisedVersion = historyVersion("2026-08", "4", reset.BankExport.data);
  const published = restoreSavedTransactionMonths(reset, [revisedVersion, september]);
  assert.equal(published.TransactionMonthCache?.months["2026-08"].cloudVersionId, "4");
  assert.equal(published.TransactionMonthCache?.months["2026-08"].masterDraft, undefined);

  harness.setData({...published, Ae_Global_Inputs: [input]});
  await processMasterAEData(harness.context);
  assert.equal(harness.data().TransactionMonthCache?.months["2026-08"].cloudVersionId, "4");
  assert.deepEqual(harness.data().BankExport.data, revisedVersion.rows);
});

test("an open month can replace stale payments with a valid empty Bank sheet when everyone is on HOLD", async (t) => {
  const file = revisionWorkbook("NORTH 08.2026.xlsx", [anna, bao], [[anna, 20000]]);
  const harness = importHarness(t, file);
  await processMasterAEData(harness.context);
  const oldVersion = historyVersion("2026-08", "1", harness.data().BankExport.data);
  const saved = applyLocalTransactionVersions(harness.data(), [oldVersion]);
  harness.setData({...saved, Ae_Global_Inputs: [{...saved.Ae_Global_Inputs[0],
    fileObj: revisionWorkbook(file.name, [], [[anna, 120000], [bao, 200000]])}]});
  await processMasterAEData(harness.context);

  const result = restoreSavedTransactionMonths(harness.data(), [oldVersion]);
  assert.deepEqual(result.BankExport.data, []);
  assert.deepEqual(getTransactionMonthOriginal(result).data, []);
  assert.equal(result.Hold_AE.data.length, 2);
  assert.equal(calculateReconciliationTotals(result, "08.2026").variance, 0);
  assert.equal(calculateReconciliationTotals(result, "08.2026").hasActual, true);
});

test("Master import retains salary totals, complete Deductions and Pivot cache after processing extraction", async (t) => {
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([
    ["ID Number", "Full name", "Center", "CHARGE TO LXO", "TOTAL PAYMENT"],
    ["001090627040", "Nguyen Van A", "Ocean Park", 120000, 120000],
  ]), "Sheet 1");
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([
    ["ID Number", "Full name", "Center", "TOTAL PAYMENT", "Note"],
    ["001090627040", "Nguyen Van A", "Ocean Park", 20000, "Hold tháng 08.2026"],
    ["", "Tổng cộng", "", 20000, ""],
  ]), "HOLD T8");
  const file = new File([XLSX.write(workbook, { bookType: "xlsx", type: "array" })], "NORTH 08.2026.xlsx");
  const harness = importHarness(t, file);
  await processMasterAEData(harness.context);

  const data = harness.data();
  assert.deepEqual(harness.processing, [true, false]);
  assert.equal(harness.completions(), 1);
  assert.equal(data.Ae_Global_Inputs[0].status, "Success");
  assert.equal(data.Sheet1_AE.data.length, 1);
  assert.equal(data.Sheet1_AE.data[0]["TOTAL PAYMENT"], 120000);
  assert.deepEqual(data.TableOriginals?.Sheet1_AE, data.Sheet1_AE);
  assert.equal(data.Sheet1_AE.data[0].L07, "HN0027.OPK");
  assert.equal(data.Hold_AE.data.length, 1);
  assert.equal(data.Hold_AE.data[0]["ID Number"], "001090627040");
  assert.equal(Math.abs(data.Hold_AE.data[0]["TOTAL PAYMENT"]), 20000);
  assert.equal(data.Hold_AE.data[0]["Tháng báo cáo"], "08.2026");
  const cache = JSON.parse(harness.storage.get("pivot_master_processed_data") || "null");
  assert.equal(cache.reportingMonth, "08.2026");
  assert.equal(cache.groupedData.AHN["HN0027.OPK"]["08.2026"].LXO, 120000);
  assert.deepEqual(cache.diagnosticLogs, []);
  assert.equal(harness.context.preparedMasterFiles.size, 0);
});

test("Master import clears processing state and marks failed files without completing", async (t) => {
  const harness = importHarness(t, new File(["invalid"], "NORTH.xlsx"));
  harness.context.parseMasterFileInWorker = async () => { throw new Error("Cannot read workbook"); };
  await processMasterAEData(harness.context);
  assert.deepEqual(harness.processing, [true, false]);
  assert.equal(harness.data().Ae_Global_Inputs[0].status, "Error: Cannot read workbook");
  assert.equal(harness.completions(), 0);
  assert.equal(harness.context.preparedMasterFiles.size, 0);
  assert.equal(harness.data().Sheet1_AE.data.length, 0);
});

test("Gross Pay import keeps equal Cambridge and Contest payments distinct and maps Cambridge HP", async (t) => {
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([
    ["ID Number", "Full name", "Center", "CHARGE TO LXO", "TOTAL PAYMENT"],
    ["001090627040", "Nguyen Van A", "Cambridge", 120000, 120000],
    ["001090627040", "Nguyen Van A", "Contest", 120000, 120000],
    ["001090627040", "Nguyen Van A", "Job Fair", 120000, 120000],
    ["001090627040", "Nguyen Van A", "Cambridge HP", 120000, 120000],
  ]), "Sheet 1");
  const file = new File([XLSX.write(workbook, { bookType: "xlsx", type: "array" })], "NORTH 08.2026.xlsx");
  const harness = importHarness(t, file);
  await processMasterAEData(harness.context);
  assert.equal(harness.completions(), 1);
  assert.deepEqual(harness.data().Sheet1_AE.data.map(row => [row.L07, row.Business, row["TOTAL PAYMENT"]]), [
    ["CAMBRIDGE", "AHN", 120000],
    ["CONTEST", "AHN", 120000],
    ["JOB FAIR", "AHN", 120000],
    ["CAMBRIDGE", "AHP", 120000],
  ]);
  const cache = JSON.parse(harness.storage.get("pivot_master_processed_data") || "null");
  assert.equal(cache.groupedData.AHN.CAMBRIDGE["08.2026"].LXO, 120000);
  assert.equal(cache.groupedData.AHN.CONTEST["08.2026"].LXO, 120000);
  assert.equal(cache.groupedData.AHN["JOB FAIR"]["08.2026"].LXO, 120000);
  assert.equal(cache.groupedData.AHP.CAMBRIDGE["08.2026"].LXO, 120000);
});

test("Master import preserves historical Bank North months for Batch Payment", async (t) => {
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([
    ["No", "ID Number", "Full name", "Center", "Bank Account Number", "TOTAL PAYMENT", "Payment details"],
    [1, "001090627040", "Nguyen Van A", "Ocean Park", "0012345678", 120000, "Salary 08.2026"],
  ]), "Bank North AE");
  const file = new File([XLSX.write(workbook, { bookType: "xlsx", type: "array" })], "NORTH 08.2026.xlsx");
  const harness = importHarness(t, file);

  harness.context.appData.Bank_North_AE = {
    ...harness.context.appData.Bank_North_AE,
    data: [{
      No: 1,
      "ID Number": "OLD-07",
      "Full name": "NGUYEN VAN JULY",
      L07: "HN0027.OPK",
      Business: "AHN",
      "Bank Account Number": "0099001100",
      "TOTAL PAYMENT": 90000,
      "Payment details": "Salary 07.2026",
      "TÊN FILE": "NORTH 07.2026.xlsx",
      _fileMonth: "07.2026",
      "Tháng báo cáo": "07.2026",
    }],
  };

  await processMasterAEData(harness.context);

  const bankRows = harness.data().Bank_North_AE.data;
  assert.equal(bankRows.length, 2);
  assert.deepEqual(
    bankRows.map((row) => row["Tháng báo cáo"]).sort(),
    ["07.2026", "08.2026"],
  );
  assert.equal(
    bankRows.find((row) => row["Tháng báo cáo"] === "07.2026")?.["ID Number"],
    "OLD-07",
  );
  assert.equal(
    bankRows.find((row) => row["Tháng báo cáo"] === "08.2026")?.["ID Number"],
    "001090627040",
  );
});
