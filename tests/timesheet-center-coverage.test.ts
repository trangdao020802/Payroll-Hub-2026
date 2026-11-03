import assert from "node:assert/strict";
import test from "node:test";
import { getMissingTimesheetCenters, getMissingTimesheetCentersInBatches } from "../src/app/lib/utils/timesheet-center-coverage";
import { applyTimesheetLinkResult } from "../src/app/lib/utils/timesheet-link-sync";
import { INITIAL_APP_DATA } from "../src/app/constants/initial-data";

test("missing centers are unique configured L07s with a file or link, normalized across aliases", () => {
  const inputs = [
    { id: "a", l07: "TH0001.TPU", url: "sheet-a" },
    { id: "b", l07: "HN0027.OPK", fileName: "roster.xlsx" },
    { id: "b-link", l07: " hn0027.opk ", url: "sheet-b" },
    { id: "c", l07: "VP0001.PCT", url: "  " },
    { id: "blank", l07: "", url: "not-a-center" },
  ];
  const coverage = getMissingTimesheetCenters(inputs, [{ l07: " th0001.tpu " }, { l07: "TH0001.TPU" }]);
  assert.deepEqual(coverage.centers, ["HN0027.OPK"]);
  assert.deepEqual(coverage.inputs.map((row) => row.id), ["b-link"]);
  assert.equal(coverage.expectedCount, 2);
  assert.deepEqual(getMissingTimesheetCenters(inputs, [{ center: "TH0001.TPU" }, { center: "HN27.OPK" }]).centers, []);
});

test("empty tables report every eligible center; subtotal labels never count as data", () => {
  const inputs = [{ id: "a", l07: "TH0001.TPU", url: "sheet-a" }];
  assert.deepEqual(getMissingTimesheetCenters(inputs, []).centers, ["TH0001.TPU"]);
  assert.deepEqual(getMissingTimesheetCenters(inputs, [{ l07: "TH0001.TPU", _isSubtotal: true }]).centers, ["TH0001.TPU"]);
});

test("Pivot compares its destination L07 instead of the source marketing center", () => {
  const inputs = [
    { id: "a", l07: "TH0001.TPU", url: "sheet-a" },
    { id: "b", l07: "HN0027.OPK", url: "sheet-b" },
  ];
  const rows = [{ l07: "MKT LOCAL NORTH", chargeToCenterMkt: "TH0001.TPU" }];
  assert.deepEqual(getMissingTimesheetCenters(inputs, rows, "allocation").centers, ["HN0027.OPK"]);
});

test("Pivot only expects Hai Phong for AHP, while the other tables still check each HP center", () => {
  const inputs = [
    { id: "hp1", l07: "HP0001.LHP", url: "hp1-link" },
    { id: "hp2", l07: "HP0002.HBT", url: "hp2-link" },
    { id: "hp3", l07: "HP0003.VIN", url: "hp3-link" },
    { id: "hp", l07: "Hai Phong", url: "hp-link" },
    { id: "vp", l07: "VP0001.PCT", url: "vp-link" },
  ];
  const rows = [{ l07: "MKT LOCAL NORTH", chargeToCenterMkt: "Hai Phong" }];
  const pivot = getMissingTimesheetCenters(inputs, rows, "allocation");
  assert.deepEqual(pivot.centers, ["VP0001.PCT"]);
  assert.deepEqual(pivot.inputs.map((row) => row.id), ["vp"]);
  assert.equal(pivot.expectedCount, 2, "HP1/HP2/HP3 are not expected destinations in Pivot");
  assert.deepEqual(getMissingTimesheetCenters(inputs.slice(0, 3), [], "allocation").centers, []);
  assert.deepEqual(getMissingTimesheetCenters(inputs.slice(0, 3), [], "l07").centers,
    ["HP0001.LHP", "HP0002.HBT", "HP0003.VIN"]);
  assert.deepEqual(getMissingTimesheetCenters([inputs[3]], [], "allocation").inputs, [inputs[3]],
    "an explicit Hai Phong link can still be refreshed when missing");
});

test("successfully synced sources missing from the table are skipped; failed and untried sources remain eligible", () => {
  const inputs = [
    { id: "synced-empty", l07: "VIN001.CTG", url: "vin-link", status: "success", count: 0 },
    { id: "synced-filtered", l07: "VP0001.PCT", url: "vp-link", status: "success", count: 12 },
    { id: "failed", l07: "PT0001.HVG", url: "pt-link", status: "error" },
    { id: "untried", l07: "TH0001.TPU", url: "th-link", status: "ready" },
  ];
  for (const mode of ["l07", "allocation"] as const) {
    const coverage = getMissingTimesheetCenters(inputs, [], mode);
    assert.deepEqual(coverage.centers, ["PT0001.HVG", "TH0001.TPU"]);
    assert.deepEqual(coverage.inputs.map((row) => row.id), ["failed", "untried"]);
    assert.equal(coverage.expectedCount, 4);
  }
});

test("September excludes April-only centers from missing counts and refresh targets", () => {
  const inputs = [
    { id: "april", l07: "HN0027.OPK", url: "april-link" },
    { id: "september", l07: "TH0001.TPU", url: "september-link" },
    { id: "never-imported", l07: "VP0001.PCT", url: "missing-link" },
  ];
  const source = {
    rows: [
      { _rowId: "april", l07: "HN27.OPK", DATE: "30/04/2026" },
      { _rowId: "september", l07: "TH0001.TPU", ngay: "2026-09-15" },
    ],
    from: "2026-09-01", to: "2026-09-30",
  };
  const coverage = getMissingTimesheetCenters(inputs, [], "l07", source);
  assert.deepEqual(coverage.centers, ["TH0001.TPU", "VP0001.PCT"]);
  assert.deepEqual(coverage.inputs.map((row) => row.id), ["september", "never-imported"]);
  assert.equal(coverage.expectedCount, 3);
  assert.deepEqual(getMissingTimesheetCenters(inputs, [], "l07", { rows: source.rows }).centers,
    ["HN0027.OPK", "TH0001.TPU", "VP0001.PCT"], "clearing the time filter restores the usual table comparison");
});

test("time exclusion requires every known date to be outside the inclusive filter", () => {
  const inputs = [{ id: "a", l07: "TH0001.TPU", url: "sheet-a" }];
  const oldRow = { _rowId: "a", ngay: "30/04/2026" };
  const range = { from: "2026-09-01", to: "2026-09-30" };
  assert.deepEqual(getMissingTimesheetCenters(inputs, [], "l07", { ...range, rows: [oldRow] }).centers, []);
  for (const date of ["2026-09-01", "30/09/2026", "", "invalid date"]) {
    const rows = [oldRow, { _rowId: "a", ngay: date }];
    assert.deepEqual(getMissingTimesheetCenters(inputs, [], "l07", { ...range, rows }).centers, ["TH0001.TPU"]);
  }
  assert.deepEqual(getMissingTimesheetCenters(inputs, [], "l07", { from: range.from, rows: [oldRow] }).centers, []);
  assert.deepEqual(getMissingTimesheetCenters(inputs, [], "l07", { to: range.to, rows: [{ _rowId: "a", ngay: "01/10/2026" }] }).centers, []);
});

test("historical marketing allocations and configured source IDs obey the same time exclusion", () => {
  const inputs = [
    { id: "mkt", l07: "MKT LOCAL NORTH", url: "mkt-link" },
    { id: "target", l07: "TH0001.TPU", url: "target-link" },
  ];
  const source = { from: "2026-09-01", to: "2026-09-30", rows: [
    { _rowId: "mkt", chargeToCenterMkt: "TH0001.TPU", "Session Date": "15/04/2026" },
  ] };
  assert.deepEqual(getMissingTimesheetCenters(inputs, [], "l07", source).centers, []);
  assert.deepEqual(getMissingTimesheetCenters(inputs, [], "allocation", source).centers, []);
});

test("link refresh replaces the center and its legacy rows without duplicating or changing other centers", () => {
  const input = { id: "a", l07: "TH0001.TPU", url: "sheet-a", legacyRowIds: ["old-a"] };
  const other = { _rowId: "b", l07: "HN0027.OPK", ma_nv: "002", duration: 7 };
  const before = {
    ...INITIAL_APP_DATA,
    Timesheet_InputList: [input, { id: "b", l07: "HN0027.OPK", url: "sheet-b" }],
    Timesheet_Roster: [{ _rowId: "old-a", l07: "TH0001.TPU", ma_nv: "001", duration: 1 }, other],
  };
  const imported = [{ _rowId: "a", l07: "TH0001.TPU", ma_nv: "001", duration: 3 }];
  const metadata = { url: "sheet-a", fileName: "updated.xlsx", date: "20/09/26" };
  const after = applyTimesheetLinkResult(before, input, imported, metadata);
  const twice = applyTimesheetLinkResult(after, input, imported, metadata);
  assert.equal(twice.Timesheet_Roster.length, 2);
  assert.deepEqual(twice.Timesheet_Roster.find((row) => row._rowId === "b"), other);
  assert.equal(twice.Timesheet_Roster.find((row) => row._rowId === "a")?.duration, 3);
  assert.equal(twice.Timesheet_InputList[0].status, "success");
  assert.equal(twice.Timesheet_InputList[0].count, 1);
  assert.deepEqual(twice.Timesheet_InputList[0].legacyRowIds, []);
  assert.equal(before.Timesheet_Roster[0].duration, 1);
});

test("large Center scans yield, preserve date filtering, and abandon an outdated query", async () => {
  const inputs = [
    { id: "outside", l07: "TH0001.TPU", url: "old-sheet" },
    { id: "missing", l07: "VP0001.PCT", url: "missing-sheet" },
  ];
  const range = {
    from: "2026-09-21", to: "2026-10-20",
    rows: Array.from({ length: 2400 }, () => ({ _rowId: "outside", ngay: "20/09/2026" })),
  };
  const controller = new AbortController();
  const obsolete = getMissingTimesheetCentersInBatches(controller.signal, inputs, [], "l07", range);
  controller.abort();
  assert.equal(await obsolete, undefined, "an outdated scan must yield and stop before publishing results");

  const current = await getMissingTimesheetCentersInBatches(new AbortController().signal, inputs, [], "l07", range);
  assert.deepEqual(current?.centers, ["VP0001.PCT"]);
  assert.deepEqual(current, getMissingTimesheetCenters(inputs, [], "l07", range));
});
