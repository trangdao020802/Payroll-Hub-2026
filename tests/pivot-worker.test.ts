import assert from "node:assert/strict";
import test from "node:test";
import * as XLSX from "xlsx";
import {
  applyPivotMktTypeCache, buildPivotFromAppData, extractPivotMktTypeData,
  omitPivotMonths, updatePivotMktTypeCache,
} from "../src/app/lib/utils/pivot-utils";
import { processExcelData } from "../src/app/workers/pivot.worker";

function workbookFile(name: string, sheetName: string, rows: (string | number)[][]) {
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(rows), sheetName);
  return { name, buffer: XLSX.write(workbook, { bookType: "xlsx", type: "array" }) as ArrayBuffer };
}

test("Pivot worker aggregates MKT roster task aliases and excludes adjustments", () => {
  const result = processExcelData([workbookFile("NORTH.MKT.xlsx", "ROSTER", [
    ["CENTER", "TYPE", "DURATION"],
    ["Ocean Park", "LPAR", 2],
    ["Ocean Park", "LPAR01 - support", 1],
    ["Ocean Park", "ADD", 10],
    ["Ocean Park", "CANCEL", 10],
  ])]);
  assert.deepEqual(result.logs, []);
  assert.deepEqual(result.typeColumns, ["LPAR01"]);
  assert.deepEqual(result.groupedData, { AHN: { "HN0027.OPK": { LPAR01: 60000 } } });
});

test("Pivot worker normalizes charge columns including bonus", () => {
  const result = processExcelData([workbookFile("NORTH.xlsx", "Sheet1", [
    ["CENTER", "CHARGE TO LXO", "CHARGE TO BONUS", "CHARGE TO ADD", "TOTAL PAYMENT"],
    ["Ocean Park", 100000, 20000, 999999, 120000],
    ["Ocean Park", 50000, 10000, 999999, 60000],
  ])]);
  assert.deepEqual(result.logs, []);
  assert.deepEqual(result.typeColumns, ["EXTRA SUMMER INSTRUCTORS", "LXO"]);
  assert.deepEqual(result.groupedData, { AHN: { "HN0027.OPK": { LXO: 150000, "EXTRA SUMMER INSTRUCTORS": 30000 } } });
});

test("Pivot worker groups Total Payment by normalized type without charge columns", () => {
  const result = processExcelData([workbookFile("NORTH.xlsx", "Sheet1", [
    ["CENTER", "TYPE", "TOTAL PAYMENT"],
    ["Ocean Park", "LDEM", 50000],
    ["Ocean Park", "LDEM01", 30000],
    ["Ocean Park", "CANCEL", 90000],
  ])]);
  assert.deepEqual(result.logs, []);
  assert.deepEqual(result.typeColumns, ["LDEM01"]);
  assert.deepEqual(result.groupedData, { AHN: { "HN0027.OPK": { LDEM01: 80000 } } });
});

test("Pivot replacement and deletion match month and year without adding prior snapshots", () => {
  const prior = { AHN: { "HN0027.OPK": {
    "08/2026": { LXO: 150000, LPAR01: 90000 },
    "09.2026": { LXO: 120000, LPAR01: 10000 },
    "08.2025": { LXO: 110000, LPAR01: 30000 },
  } } };
  const replacement = { AHN: { "HN0027.OPK": { "2026-08": { LPAR01: 60000 } } } };
  const first = updatePivotMktTypeCache(extractPivotMktTypeData(prior), replacement, ["LPAR01"], ["08.2026"], true);
  const retry = updatePivotMktTypeCache(first, replacement, ["LPAR01"], ["8-2026"], true);
  const result = applyPivotMktTypeCache(prior, retry, ["08.2026"]);
  const months = result.AHN["HN0027.OPK"];
  assert.deepEqual(retry.groupedData, first.groupedData);
  assert.deepEqual(retry.manualMonths, ["08.2026"]);
  assert.equal(months["08/2026"].LXO, 150000);
  assert.equal(months["08/2026"].LPAR01, 60000);
  assert.equal(months["08.2026"], undefined);
  assert.deepEqual(omitPivotMonths(result, ["2026-08"]), { AHN: { "HN0027.OPK": {
    "09.2026": { LXO: 120000, LPAR01: 10000 },
    "08.2025": { LXO: 110000, LPAR01: 30000 },
  } } });
  assert.equal(prior.AHN["HN0027.OPK"]["08/2026"].LPAR01, 90000);

  const roster = (source: string, month: string, durationHours: number) => ({
    _sourceFile: source, month, l07: "HN0027.OPK", type: "LPAR01", durationHours,
  });
  const legacy = buildPivotFromAppData([], [], [
    roster("older.xlsx", "8/2026", 5),
    roster("other year.xlsx", "08.2025", 3),
    roster("replacement.xlsx", "08.2026", 1),
    roster("replacement.xlsx", "08.2026", 2),
  ], "09.2026");
  assert.equal(legacy.groupedData.AHN["HN0027.OPK"]["08.2026"].LPAR01, 60000);
  assert.equal(legacy.groupedData.AHN["HN0027.OPK"]["08.2025"].LPAR01, 60000);
});
