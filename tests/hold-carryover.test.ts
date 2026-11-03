import assert from "node:assert/strict";
import test from "node:test";
import {
  carryEligibleHoldsToNextMonth,
  collapseMergedHoldSourceRows,
  getEligibleHoldRowsForReport,
  getMergedHoldOriginalIndexes,
  mergeDuplicateHoldRows,
  parsePayrollMonth,
  reconcileHoldTransactionRows,
} from "../src/app/lib/utils/hold-carryover";

test("cloud ISO months retain October through December and reject invalid month prefixes", () => {
  for (const month of [10, 11, 12]) {
    assert.equal(parsePayrollMonth(`2026-${month}`)?.dot, `${month}.2026`);
    assert.equal(parsePayrollMonth(`2026-${month}-01`)?.dot, `${month}.2026`);
  }
  assert.equal(parsePayrollMonth("2026-13"), null);
  assert.equal(parsePayrollMonth("13.2026"), null);
});

const transaction = {
  "Tháng phát sinh": "01.2026",
  "ID Number": "AE001",
  "Full name": "NGUYEN VAN A",
  L07: "HN0001.TEST",
  BU: "AHN",
  "Bank Account Number": "123456789",
  "TOTAL PAYMENT": -1_250_000,
  "Sheet Source": "Hold T1",
};

function row(
  reportMonth: string,
  operation: "Hold" | "Cancel" | "Add",
  extra: Record<string, unknown> = {},
) {
  return {
    ...transaction,
    "Tháng báo cáo": reportMonth,
    _fileMonth: reportMonth,
    "Nghiệp vụ": operation,
    "TOTAL PAYMENT": operation === "Add" ? 1_250_000 : -1_250_000,
    ...extra,
  };
}

test("merges carried and imported HOLD rows before an operation change", () => {
  const carried = row("02.2026", "Hold", {
    id: "carry",
    _holdCarryKey: "carry",
    _originalIndex: 0,
  });
  const imported = row("02.2026", "Hold", {
    id: "source",
    _originalIndex: 1,
  });

  const merged = mergeDuplicateHoldRows([carried, imported]);

  assert.equal(merged.length, 1);
  assert.equal(merged[0]["Nghiệp vụ"], "Hold");
  assert.equal(merged[0]["TOTAL PAYMENT"], -1_250_000);
  assert.deepEqual(getMergedHoldOriginalIndexes(merged[0]), [0, 1]);
});

test("normalizes a legacy HOLD-sheet row before eligibility checks", () => {
  const legacy = {
    ...transaction,
    "Tháng báo cáo": "02.2026",
    _fileMonth: "02.2026",
    "TOTAL PAYMENT": 1_250_000,
  };

  const reconciled = reconcileHoldTransactionRows([legacy]);

  assert.equal(reconciled[0]["Nghiệp vụ"], "Hold");
  assert.equal(reconciled[0]["TOTAL PAYMENT"], -1_250_000);
  assert.equal(getEligibleHoldRowsForReport([legacy], "02.2026").length, 1);
});

test("CANCEL wins over its duplicate HOLD and both source rows collapse", () => {
  const carried = row("02.2026", "Hold", {
    id: "carry",
    _holdCarryKey: "carry",
    _originalIndex: 0,
  });
  const cancelled = row("02.2026", "Cancel", {
    id: "source",
    _holdStatusBeforeSave: "Hold",
    _originalIndex: 1,
  });
  const merged = mergeDuplicateHoldRows([carried, cancelled]);

  assert.equal(merged.length, 1);
  assert.equal(merged[0]["Nghiệp vụ"], "Cancel");
  assert.equal(getEligibleHoldRowsForReport([carried, cancelled], "02.2026").length, 0);

  const canonicalIndex = Number(merged[0]._originalIndex);
  const collapsed = collapseMergedHoldSourceRows({
    rows: [carried, cancelled],
    mergedRow: merged[0],
    canonicalIndex,
    updatedRow: { ...cancelled, "Nghiệp vụ": "Cancel" },
  });

  assert.equal(collapsed.length, 1);
  assert.equal(collapsed[0]["Nghiệp vụ"], "Cancel");
});

test("a February CANCEL removes stale March copies and is never carried again", () => {
  const januaryHold = row("01.2026", "Hold", { id: "jan" });
  const februaryHoldCopy = row("02.2026", "Hold", {
    id: "feb-carry",
    _holdCarryKey: "feb-carry",
    _holdCarryFromReportMonth: "01.2026",
  });
  const februaryCancel = row("02.2026", "Cancel", {
    id: "feb-cancel",
    _holdStatusBeforeSave: "Hold",
  });
  const staleMarchCarry = row("03.2026", "Hold", {
    id: "mar-carry",
    _holdCarryKey: "mar-carry",
    _holdCarryFromReportMonth: "02.2026",
  });
  const staleMarchSource = row("03.2026", "Hold", { id: "mar-source" });

  const reconciled = reconcileHoldTransactionRows([
    januaryHold,
    februaryHoldCopy,
    februaryCancel,
    staleMarchCarry,
    staleMarchSource,
  ]);

  assert.deepEqual(
    reconciled.map((item) => [item["Tháng báo cáo"], item["Nghiệp vụ"]]),
    [
      ["01.2026", "Hold"],
      ["02.2026", "Cancel"],
    ],
  );

  const carryResult = carryEligibleHoldsToNextMonth({
    sourceRows: [februaryHoldCopy, februaryCancel],
    existingRows: reconciled,
    reportMonth: "02.2026",
    createdAt: "2026-02-28T00:00:00.000Z",
  });

  assert.equal(carryResult.carriedCount, 0);
  assert.equal(
    carryResult.rows.some((item) => item["Tháng báo cáo"] === "03.2026"),
    false,
  );
});

test("ADD also resolves the single HOLD transaction without leaving a copy", () => {
  const hold = row("02.2026", "Hold", {
    id: "hold",
    _holdCarryKey: "hold",
  });
  const add = row("02.2026", "Add", {
    id: "add",
    _holdStatusBeforeSave: "Hold",
  });

  const merged = mergeDuplicateHoldRows([hold, add]);

  assert.equal(merged.length, 1);
  assert.equal(merged[0]["Nghiệp vụ"], "Add");
  assert.equal(merged[0]["TOTAL PAYMENT"], 1_250_000);
  assert.equal(getEligibleHoldRowsForReport([hold, add], "02.2026").length, 0);
});

test("the latest explicit operation change stays authoritative", () => {
  const cancelled = row("02.2026", "Cancel", {
    id: "cancel",
    _holdStatusBeforeSave: "Hold",
    _holdOperationUpdatedAt: "2026-02-20T00:00:00.000Z",
  });
  const changedToAdd = row("02.2026", "Add", {
    id: "add",
    _holdStatusBeforeSave: "Hold",
    _holdOperationUpdatedAt: "2026-02-21T00:00:00.000Z",
  });
  const repeatedWorkbookHold = row("02.2026", "Hold", {
    id: "workbook",
    _uploadTimestamp: "2026-02-22T00:00:00.000Z",
  });

  const merged = mergeDuplicateHoldRows([
    cancelled,
    changedToAdd,
    repeatedWorkbookHold,
  ]);

  assert.equal(merged.length, 1);
  assert.equal(merged[0]["Nghiệp vụ"], "Add");
  assert.equal(merged[0]["TOTAL PAYMENT"], 1_250_000);
});

test("does not merge transactions when an exact identity field differs", () => {
  const original = row("02.2026", "Hold", { id: "one" });
  const differentAccount = row("02.2026", "Hold", {
    id: "two",
    "Bank Account Number": "999999999",
  });

  assert.equal(
    mergeDuplicateHoldRows([original, differentAccount]).length,
    2,
  );
});


test("merges the same HOLD even when carried and workbook source labels differ", () => {
  const carried = row("02.2026", "Hold", {
    id: "carry",
    _holdCarryKey: "carry",
    _originalIndex: 0,
    "Sheet Source": "Hold T1",
    BU: "AHN",
    L07: "HN0001.TEST",
  });
  const imported = row("02.2026", "Hold", {
    id: "source",
    _originalIndex: 1,
    "Sheet Source": "HOLD",
    BU: "AHN-ENRICHED",
    L07: "HN0001.UPDATED",
  });

  const merged = mergeDuplicateHoldRows([carried, imported]);
  assert.equal(merged.length, 1);
  assert.deepEqual(getMergedHoldOriginalIndexes(merged[0]), [0, 1]);
});

test("operation change collapses semantic duplicates even without original indexes", () => {
  const carried = row("02.2026", "Hold", { id: "carry", _holdCarryKey: "carry" });
  const imported = row("02.2026", "Hold", { id: "source" });
  const merged = mergeDuplicateHoldRows([carried, imported]);

  const collapsed = collapseMergedHoldSourceRows({
    rows: [carried, imported],
    mergedRow: merged[0],
    canonicalIndex: 1,
    updatedRow: { ...imported, "Nghiệp vụ": "Cancel" },
  });

  assert.equal(collapsed.length, 1);
  assert.equal(collapsed[0]["Nghiệp vụ"], "Cancel");
});

test("a resolved HOLD suppresses a later workbook copy despite source metadata changes", () => {
  const februaryCancel = row("02.2026", "Cancel", {
    id: "feb-cancel",
    _holdStatusBeforeSave: "Hold",
    "Sheet Source": "Hold T1",
    BU: "AHN",
    L07: "HN0001.TEST",
  });
  const marchWorkbookHold = row("03.2026", "Hold", {
    id: "mar-source",
    "Sheet Source": "HOLD",
    BU: "AHN-ENRICHED",
    L07: "HN0001.UPDATED",
  });

  const reconciled = reconcileHoldTransactionRows([
    februaryCancel,
    marchWorkbookHold,
  ]);

  assert.deepEqual(
    reconciled.map((item) => [item["Tháng báo cáo"], item["Nghiệp vụ"]]),
    [["02.2026", "Cancel"]],
  );
});

test("resolved ADD and CANCEL history rows are never discarded as stale HOLD copies", () => {
  const januaryHold = row("01.2026", "Hold", { id: "jan-hold" });
  const februaryAdd = row("02.2026", "Add", {
    id: "feb-add",
    _holdStatusBeforeSave: "Hold",
    _holdOperationUpdatedAt: "2026-02-20T00:00:00.000Z",
  });
  const marchCancel = row("03.2026", "Cancel", {
    id: "mar-cancel",
    _holdStatusBeforeSave: "Hold",
    _holdOperationUpdatedAt: "2026-03-20T00:00:00.000Z",
  });
  const aprilStaleHold = row("04.2026", "Hold", { id: "apr-stale" });

  const reconciled = reconcileHoldTransactionRows([
    januaryHold,
    februaryAdd,
    marchCancel,
    aprilStaleHold,
  ]);

  assert.deepEqual(
    reconciled.map((item) => [item["Tháng báo cáo"], item["Nghiệp vụ"]]),
    [
      ["01.2026", "Hold"],
      ["02.2026", "Add"],
      ["03.2026", "Cancel"],
    ],
  );
});
