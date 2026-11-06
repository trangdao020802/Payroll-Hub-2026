import assert from "node:assert/strict";
import test from "node:test";

import { calculateTrialBalanceHeaderTotals, calculateTrialBalanceInternCost } from "../src/app/lib/utils/trial-balance-header-totals";

test("Trial Balance header totals use business columns, not row labels", () => {
  const totals = calculateTrialBalanceHeaderTotals([
    {
      id: "02.2026_AHN_adjustment_01.2026",
      customMonthDisplay: "Hold lương tháng 01.2026",
      thu: 0,
      hold: 400_000,
      add: 125_000,
      cancel: 75_000,
    },
    {
      id: "02.2026_AHN",
      thu: "839.461.357",
      hold: 0,
      add: 0,
      cancel: 0,
    },
  ]);

  assert.deepEqual(totals, {
    payrollCost: 839_461_357,
    hold: 400_000,
    add: 125_000,
    cancel: 75_000,
  });
});

test("Trial Balance header totals ignore rows excluded from totals", () => {
  const totals = calculateTrialBalanceHeaderTotals([
    { thu: 10, hold: 20, add: 30, cancel: 40 },
    {
      thu: 1_000,
      hold: 2_000,
      add: 3_000,
      cancel: 4_000,
      _excludeFromTotals: true,
    },
  ]);

  assert.deepEqual(totals, {
    payrollCost: 10,
    hold: 20,
    add: 30,
    cancel: 40,
  });
});

test("header counts current HOLD, past ADD and current CANCEL after auto approval", () => {
  const rows = [
    {month: "Tháng 2/2026", reportMonth: "02.2026", thu: 10_000_000, chi: -500_000, rawHold: 0},
    {month: "01.2026", reportMonth: "02.2026", customMonthDisplay: "Hold", thu: 3_575_833, rawAdd: -3_575_833, rawHold: 900_000, add: 0, hold: 0},
    {month: "12.2025", reportMonth: "02.2026", customMonthDisplay: "Cancel", rawCancel: -542_500, cancel: 0},
    {month: "02.2026", reportMonth: "02.2026", customMonthDisplay: "Cancel", rawCancel: -100_000, chi: 100_000},
    {month: "02.2026", reportMonth: "02.2026", customMonthDisplay: "Add", rawAdd: 200_000, thu: 200_000},
    {month: "03.2026", reportMonth: "03.2026", thu: 9_999_999, hold: 9_999_999},
  ];
  assert.deepEqual(calculateTrialBalanceHeaderTotals(rows, "02.2026"), {
    payrollCost: 13_775_833, hold: 500_000, add: 3_575_833, cancel: 100_000,
  });
});

test("Intern cost uses earned payroll for the selected month, excluding hold and release movements", () => {
  const september = "Tháng 9/2026";
  const rows = [
    { month: september, thu: "753.821.753" },
    { month: september, displayMonth: "05.2026", _isOpeningHold: true, thu: 815_500 },
    { month: september, customMonthDisplay: "Hold lương tháng 09.2026", chi: 1_979_833 },
    { month: "08.2026", reportMonth: september, customMonthDisplay: "Add lương tháng 08.2026", thu: 2_927_166 },
    { month: september, id: "09.2026_AHN_cancel", chi: 80_000 },
    { month: september, id: "09.2026_AHN_add", thu: 200_000 },
    { month: "09.2025", thu: 200_000 },
    { month: "10.2026", thu: 300_000 },
    { month: september, _excludeFromTotals: true, thu: 500_000 },
  ];
  assert.equal(calculateTrialBalanceInternCost(rows, "09.2026"), 753_821_753);
  assert.equal(calculateTrialBalanceInternCost(rows, "10.2026"), 300_000);
  assert.equal(calculateTrialBalanceInternCost(rows.slice(1, 6), september), 0);

  const legacyBonus = { month: "09.2026", id: "09.2026_AHN_bonus", customMonthDisplay: "Bonus lương tháng 09.2026", thu: 35_000 };
  assert.equal(calculateTrialBalanceInternCost([...rows, legacyBonus], september), 753_856_753);
});
