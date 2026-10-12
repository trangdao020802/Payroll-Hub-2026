import assert from "node:assert/strict";
import test from "node:test";
import { INITIAL_APP_DATA } from "../src/app/constants/initial-data";
import {
  calculateReconciliationTotals,
  createReconciliationTotalsCalculator,
} from "../src/app/lib/utils/reconciliation-sync";
import { createDeductionsRowProjector } from "../src/app/lib/utils/deductions-view";
import { encode, decode } from "../src/app/lib/cloud/codec";
import type { AppData, TableData } from "../src/app/types";
const table = (data: Record<string, unknown>[]): TableData => ({
  headers: [],
  data,
});
const saved = { lastAction: "saved" as const, saveVersion: 1, editCount: 0 };
function january(): AppData {
  return {
    ...INITIAL_APP_DATA,
    globalMonth: "01.2026",
    Sheet1_AE: table([
      {
        "ID Number": "TEST-1",
        "Tháng báo cáo": "01.2026",
        "TOTAL PAYMENT": 851_460_974,
      },
    ]),
    Hold_AE: table([
      {
        "ID Number": "TEST-2",
        "Full name": "Test employee",
        BU: "AHN",
        L07: "HN01",
        "TOTAL PAYMENT": 12_895_000,
        "Nghiệp vụ": "Add",
        Tháng: "01.2026",
        "Sheet Source": "Hold T12",
        "Tháng phát sinh": "12.2025",
      },
    ]),
    // A second machine can restore a Transaction-shaped Bank AE, without TOTAL PAYMENT.
    Bank_North_AE: table([
      { "Tháng báo cáo": "2026-01", "Payment Amount": 864_355_974 },
    ]),
    BankExport: table([
      { "Tháng báo cáo": "2026-01", "Payment Amount": 864_355_974 },
    ]),
    TransactionActivity: saved,
    TransactionMonthCache: { activePeriod: "2026-01", months: {} },
  };
}
const displayed = (data: AppData): AppData => ({
  ...data,
  Hold_AE: {
    ...data.Hold_AE,
    data: createDeductionsRowProjector(data.globalMonth!)(data.Hold_AE.data),
  },
});
test("January Deductions remain reconciled after cloud serialization to a new computer", async () => {
  const original = january();
  const restored = (await decode(await encode(original))) as AppData;
  for (const data of [original, restored]) {
    const totals = calculateReconciliationTotals(displayed(data));
    assert.equal(totals.actual, 864_355_974);
    assert.equal(totals.deductionsTotal, 12_895_000);
    assert.equal(totals.variance, 0);
    assert.equal(displayed(data).Hold_AE.data[0]["Tháng phát sinh"], "12.2025");
  }
  assert.deepEqual(restored.BankExport, original.BankExport);
});
test("confirmed Batch Payment takes precedence over an obsolete Bank AE source", () => {
  const data = january();
  data.Bank_North_AE = table([
    { "Tháng báo cáo": "01.2026", "TOTAL PAYMENT": 5 },
  ]);
  assert.equal(
    calculateReconciliationTotals(displayed(data)).actual,
    864_355_974,
  );
});
test("cached confirmed historical month is used without mixing another current month", () => {
  const data = january();
  const januaryPayment = data.BankExport;
  data.globalMonth = "02.2026";
  data.BankExport = table([
    { "Tháng báo cáo": "2026-02", "Payment Amount": 10 },
  ]);
  data.TransactionMonthCache = {
    activePeriod: "2026-02",
    months: { "2026-01": { table: januaryPayment, activity: saved } },
  };
  assert.equal(
    calculateReconciliationTotals(displayed(data), "01.2026").actual,
    864_355_974,
  );
});
test("an intentionally empty confirmed month stays zero instead of falling back to imported bank data", () => {
  const data = january();
  data.BankExport = table([]);
  data.Bank_North_AE = table([
    { "Tháng báo cáo": "01.2026", "TOTAL PAYMENT": 864_355_974 },
  ]);
  assert.equal(calculateReconciliationTotals(displayed(data)).actual, 0);
});
test("draft and generated payment previews are not treated as confirmed amounts", () => {
  for (const action of ["generated", "edited"] as const) {
    const data = january();
    data.TransactionActivity = { ...saved, lastAction: action };
    data.Bank_North_AE = table([
      { "Tháng báo cáo": "01.2026", "TOTAL PAYMENT": 120 },
    ]);
    assert.equal(calculateReconciliationTotals(displayed(data)).actual, 120);
  }
});
test("cached totals recalculate when the approved payment changes, and expose real discrepancies", () => {
  const data = displayed(january());
  const calculate = createReconciliationTotalsCalculator("01.2026");
  assert.equal(calculate(data).variance, 0);
  const changed = {
    ...data,
    BankExport: table([
      { "Tháng báo cáo": "2026-01", "Payment Amount": 864_355_984 },
    ]),
  };
  assert.equal(calculate(changed).variance, 10);
  assert.deepEqual(calculate(changed), calculateReconciliationTotals(changed));
});
