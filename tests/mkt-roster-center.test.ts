import assert from "node:assert/strict";
import test from "node:test";
import * as XLSX from "xlsx";
import {
  getL07FromChargeToCenterMkt,
  resolveMktRosterCenter,
} from "../src/app/lib/utils/center-utils";
import {
  parseExcelData,
  prepareExcelResult,
} from "../src/app/workers/excelParser.worker";
import { calculateTimesheet } from "../src/app/workers/timesheet.worker";
import { TASK_COLUMNS } from "../src/app/constants/timesheet-logic";
import { applyTimesheetRosterCellEdit, getTimesheetChargeToCenter } from '../src/app/lib/utils/timesheet-edits';

function createWorkbookBuffer(
  sheetName: string,
  rows: Array<Array<string | number>>,
): ArrayBuffer {
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(
    workbook,
    XLSX.utils.aoa_to_sheet(rows),
    sheetName,
  );
  return XLSX.write(workbook, { bookType: "xlsx", type: "array" });
}

function importMktRoster(
  sheetName: string,
  header: "CENTER" | "CHARGE TO CENTER",
  centerNumber: number,
  fileName = "MKT LOCAL NORTH.xlsx",
) {
  const fileBuffer = createWorkbookBuffer(sheetName, [
    [header, "ID NUMBER", "FULL NAME", "DATE", "TYPE", "DURATION"],
    [centerNumber, "001090627040", "NGUYEN PHUNG MANH", "15/08/2026", "LPAR01", 2],
  ]);
  const rawRows = parseExcelData(fileBuffer, fileName);
  return prepareExcelResult(rawRows, fileName, "mkt-roster", "roster").rows[0];
}

test("maps the numeric part of a North MKT center to its full L07", () => {
  assert.equal(getL07FromChargeToCenterMkt("1"), "HN0001.PHY");
  assert.equal(getL07FromChargeToCenterMkt("16"), "HN0016.PDP");
  assert.equal(getL07FromChargeToCenterMkt("27.0"), "HN0027.OPK");
  assert.deepEqual(resolveMktRosterCenter("27"), {
    chargeToCenterMkt: "HN0027.OPK",
    l07: "HN0027.OPK",
    business: "AHN",
  });
});

test("reads CENTER from an MKT Local North ROSTER sheet as the L07 allocation code", () => {
  const row = importMktRoster("ROSTER", "CENTER", 27);

  assert.equal(row.l07, "MKT LOCAL NORTH");
  assert.equal(row.chargeToCenterMkt, "HN0027.OPK");
  assert.equal(row.business, "AHN");
});

test("reads CHARGE TO CENTER from a Q_ROSTER sheet and converts it to L07", () => {
  const row = importMktRoster(
    "Q_ROSTER",
    "CHARGE TO CENTER",
    16,
    "NORTH.MKT ROSTER AUG.xlsx",
  );

  assert.equal(row.l07, "MKT LOCAL NORTH");
  assert.equal(row.chargeToCenterMkt, "HN0016.PDP");
  assert.equal(row.business, "AHN");
});

function updatedRoster(fileName = 'North.MKT Roster 21092025.xlsx', allocationHeader?: string, allocationValue = '27') {
  const headers = ['', 'ID', 'Name', 'Date', 'Code', 'Class', 'From', 'To', 'Duration', 'Notes', 'Hour', 'Total'];
  const row: Array<string | number> = ['MKT North', '001234567890', 'TEST PERSON', '21/09/2026', 'LPAR01', 'Vu Tong Phan', '08:00', '12:00', '4:00', 'Test notes', 4, 80000];
  if (allocationHeader) {headers.push(allocationHeader); row.push(allocationValue);}
  const buffer = createWorkbookBuffer('Roster', [headers, row]);
  return prepareExcelResult(parseExcelData(buffer, fileName), fileName, 'mkt-file', 'roster').rows;
}

test('updated North MKT Roster reads Class as charge destination while retaining Class, Code/Type and duration', () => {
  const rows = updatedRoster();
  assert.equal(rows.length, 1);
  assert.equal(rows[0].l07, 'MKT LOCAL NORTH');
  assert.equal(rows[0].chargeToCenterMkt, 'HN0018.VTP');
  assert.equal(rows[0].class, 'Vu Tong Phan');
  assert.equal(rows[0].Class, 'Vu Tong Phan');
  assert.equal(rows[0].type, 'LPAR01');
  assert.equal(rows[0].duration, 4);
  assert.equal(rows[0].ma_nv, '001234567890');
});

test('existing CENTER and charge columns keep priority over the new Class fallback', () => {
  for (const header of ['CENTER', 'CHARGE TO CENTER', 'Charge to Center MKT', 'CHARGETOCENTER']) {
    const row = updatedRoster(undefined, header)[0];
    assert.equal(row.chargeToCenterMkt, 'HN0027.OPK', header);
    assert.equal(row.class, 'Vu Tong Phan', header);
  }
});

test('new Class allocations flow through the actual Timesheet calculation into destination salary totals', () => {
  const result = calculateTimesheet({
    rosterData: updatedRoster(), salaryScaleData: [], staffData: [], cacheData: [],
    fromDateStr: '2026-09-01', toDateStr: '2026-09-30', preferredYear: 2026,
    appData: {Timesheet_InputList: [{id: 'mkt-file', l07: 'MKT LOCAL NORTH'}]},
    checkTAsMap: {}, classSizeMap: {}, TASK_COLUMNS,
  });
  assert.equal(result.processedRosterData[0].chargeToCenterMkt, 'HN0018.VTP');
  assert.equal(result.processedRosterData[0].taskType, 'LPAR01');
  assert.equal(result.centerSummary.length, 1);
  assert.equal(result.centerSummary[0].l07, 'HN0018.VTP');
  assert.equal(result.centerSummary[0].chargeMktLocal, 80000);
});

test('configured L07 enables Class allocation even for a generically named imported Roster', () => {
  const rows = updatedRoster('Updated Roster.xlsx');
  assert.equal(rows[0].chargeToCenterMkt, '');
  const result = calculateTimesheet({
    rosterData: rows, salaryScaleData: [], staffData: [], cacheData: [],
    fromDateStr: '2026-09-01', toDateStr: '2026-09-30', preferredYear: 2026,
    appData: {Timesheet_InputList: [{id: 'mkt-file', l07: 'MKT LOCAL NORTH'}]},
    checkTAsMap: {}, classSizeMap: {}, TASK_COLUMNS,
  });
  assert.equal(result.processedRosterData[0].l07, 'MKT LOCAL NORTH');
  assert.equal(result.processedRosterData[0].chargeToCenterMkt, 'HN0018.VTP');
  assert.equal(result.centerSummary[0].chargeMktLocal, 80000);
});

test('Class is not a global charge alias and explicit allocation edits, including clearing, survive reload', () => {
  const row = {l07: 'MKT LOCAL NORTH', Class: 'An Hung'};
  assert.equal(getTimesheetChargeToCenter(row), 'HN0031.AHG');
  assert.equal(getTimesheetChargeToCenter({...row, l07: 'HN0001.PHY'}), '');
  assert.equal(getTimesheetChargeToCenter({...row, l07: 'MKT LOCAL SOUTH'}), '');
  const edited = applyTimesheetRosterCellEdit(row, 'chargeToCenterMkt', 'Hai Phong');
  assert.equal(getTimesheetChargeToCenter(JSON.parse(JSON.stringify(edited))), 'Hai Phong');
  const cleared = applyTimesheetRosterCellEdit(edited, 'chargeToCenterMkt', '');
  assert.equal(getTimesheetChargeToCenter(JSON.parse(JSON.stringify(cleared))), '');
  assert.equal(getTimesheetChargeToCenter({...row, chargeToCenterMkt: '', charge_to_center_mkt: ''}), '');
});

test('legacy North rows continue to use Center and explicit charge columns before the equivalent Class alias', () => {
  const row = {l07: 'MKT LOCAL NORTH', Class: 'An Hung', Center: 'Vu Tong Phan'};
  assert.equal(getTimesheetChargeToCenter(row), 'HN0018.VTP');
  assert.equal(getTimesheetChargeToCenter({...row, 'CHARGE TO CENTER': 'Dong Anh'}), 'HN0033.DAH');
  assert.equal(getTimesheetChargeToCenter({...row, Center: '', Class: '27'}), 'HN0027.OPK');
});

test('Center or L07 equal to the MKT LOCAL NORTH group leaves Class as the allocation destination', () => {
  for (const fileName of ['North.MKT Roster 21092025.xlsx', 'Updated Roster.xlsx']) {
    for (const groupHeader of ['CENTER', 'L07']) {
      const row = updatedRoster(fileName, groupHeader, 'MKT LOCAL NORTH')[0];
      assert.equal(row.l07, 'MKT LOCAL NORTH', `${fileName}: ${groupHeader}`);
      assert.equal(row.chargeToCenterMkt, 'HN0018.VTP', `${fileName}: ${groupHeader}`);
      assert.equal(row.class, 'Vu Tong Phan');
    }
  }
  assert.equal(getTimesheetChargeToCenter({l07: 'MKT LOCAL NORTH', center: 'MKT LOCAL NORTH', chargeToCenterMkt: 'MKT LOCAL NORTH', Class: 'An Hung'}), 'HN0031.AHG');
  assert.equal(getTimesheetChargeToCenter({l07: 'MKT LOCAL NORTH', center: 'MKT LOCAL NORTH', Class: 'An Hung', 'Charge to Center': 'Dong Anh'}), 'HN0033.DAH');
});

test("numeric charge destinations are normalized even in ordinary Raw Data files", () => {
  const buffer = createWorkbookBuffer("ROSTER", [
    ["CENTER", "CHARGE TO CENTER", "ID NUMBER", "FULL NAME", "DATE", "TYPE", "DURATION"],
    ["TH0001.TPU", 27, "001090627040", "TEST PERSON", "15/09/2026", "IN-CLASS", 2],
  ]);
  const rows = prepareExcelResult(parseExcelData(buffer, "TH0001.TPU.xlsx"), "TH0001.TPU.xlsx", "th", "roster").rows;
  assert.equal(rows[0].chargeToCenterMkt, "HN0027.OPK");
});

test("recalculation prefers edited allocations and preserves individual source identities", () => {
  const rows = [
    { _rowId: "mkt-file", _recordId: "source-1", chargeToCenterMkt: "TH0001.TPU" },
    { _rowId: "mkt-file", _recordId: "source-2", chargeToCenterMkt: "Hai Phong" },
  ].map((row, index) => ({
    ...row, _sourceFile: "MKT LOCAL NORTH.xlsx", center: "27", "charge to center": "27",
    "ID NUMBER": `00109062704${index}`, "FULL NAME": "TEST PERSON", DATE: "15/09/2026", TYPE: "LPAR01", DURATION: 2,
  }));
  const result = calculateTimesheet({
    rosterData: JSON.parse(JSON.stringify(rows)), salaryScaleData: [], staffData: [], cacheData: [],
    fromDateStr: "2026-09-01", toDateStr: "2026-09-30", preferredYear: 2026,
    appData: { Timesheet_InputList: [{ id: "mkt-file", l07: "MKT LOCAL NORTH" }] },
    checkTAsMap: {}, classSizeMap: {}, TASK_COLUMNS,
  });
  assert.deepEqual(result.processedRosterData.map((row: Record<string, unknown>) => row.chargeToCenterMkt), ["TH0001.TPU", "Hai Phong"]);
  assert.deepEqual(result.processedRosterData.map((row: Record<string, unknown>) => row._sourceKey), ["source-1", "source-2"]);
  assert.deepEqual(result.processedRosterData.map((row: Record<string, unknown>) => row.business), ["ATH", "AHP"]);
});

test("Roster Center assigns non-zero MKT Local cost to its destination L07", () => {
  const result = calculateTimesheet({
    rosterData: [
      {
        _rowId: "mkt-source",
        _sourceFile: "MKT LOCAL NORTH.xlsx",
        center: "27",
        chargeToCenterMkt: "HN0027.OPK",
        "ID NUMBER": "001090627040",
        "FULL NAME": "NGUYEN PHUNG MANH",
        DATE: "15/08/2026",
        TYPE: "LPAR01",
        FROM: "09:00",
        TO: "11:00",
      },
    ],
    salaryScaleData: [],
    staffData: [],
    cacheData: [],
    fromDateStr: "2026-08-01",
    toDateStr: "2026-08-31",
    appData: {
      Timesheet_InputList: [
        {
          id: "mkt-source",
          l07: "MKT LOCAL NORTH",
          aeCode: "MKT LOCAL NORTH",
          bus: "AHN",
        },
      ],
    },
    preferredYear: 2026,
    checkTAsMap: {},
    classSizeMap: {},
    TASK_COLUMNS,
  });

  assert.equal(result.centerSummary.length, 1);
  assert.equal(result.centerSummary[0].l07, "HN0027.OPK");
  assert.equal(result.centerSummary[0].business, "AHN");
  assert.equal(result.centerSummary[0].chargeMktLocal, 40_000);
  assert.equal(
    result.centerSummary.some(
      (row: Record<string, unknown>) =>
        row.l07 === "MKT LOCAL NORTH" &&
        Number(row.chargeMktLocal || 0) !== 0,
    ),
    false,
  );
});

test("Roster Center removes duplicate MKT columns outside and nests them under MKT LOCAL NORTH_TIMESHEET group", async () => {
  const { buildCenterTable, MKT_LOCAL_NORTH_GROUP } = await import(
    "../src/app/lib/utils/center-table"
  );
  const { CENTER_COLUMNS } = await import(
    "../src/app/constants/timesheet-columns"
  );

  const duplicateKeys = [
    "chargeLdem01",
    "chargeLdec01",
    "chargeLpar01",
    "chargeLret01",
    "chargeMoth01",
    "chargeMktLocal",
  ];

  // 1. CENTER_COLUMNS should not contain standalone duplicate MKT columns
  duplicateKeys.forEach((key) => {
    assert.equal(
      CENTER_COLUMNS.some((col) => col.key === key),
      false,
      `CENTER_COLUMNS should not contain duplicate key ${key}`,
    );
  });

  const centerData = [
    {
      business: "AHN",
      l07: "HN0027.OPK",
      totalSalary: 5000000,
      chargeLpar01: 1170000,
      chargeLret01: 210000,
    },
  ];

  const mktLocalNorthData = [
    {
      taskType: "LPAR01",
      workingHours: 58.5,
      chargeToCenterMkt: "HN0027.OPK",
    },
    {
      taskType: "LRET01",
      workingHours: 10.5,
      chargeToCenterMkt: "HN0027.OPK",
    },
  ];

  const { centerColumns, centerRows } = buildCenterTable(
    centerData,
    mktLocalNorthData,
  );

  // 2. MKT columns must belong to the MKT_LOCAL_NORTH_GROUP group
  const groupedCols = centerColumns.filter(
    (col) => col.group === MKT_LOCAL_NORTH_GROUP,
  );
  assert.equal(groupedCols.length, 2);
  assert.deepEqual(
    groupedCols.map((c) => c.label),
    ["LPAR01", "LRET01"],
  );

  // 3. Columns outside the group should have no duplicate MKT columns
  const outsideCols = centerColumns.filter(
    (col) => col.group !== MKT_LOCAL_NORTH_GROUP,
  );
  duplicateKeys.forEach((key) => {
    assert.equal(
      outsideCols.some((col) => col.key === key),
      false,
    );
  });

  // 4. Values are properly populated
  const mainRow = centerRows.find((r) => !r._isSubtotal);
  assert.ok(mainRow);
  assert.equal(mainRow["mktLocalNorth::LPAR01"], 58.5 * 20000);
  assert.equal(mainRow["mktLocalNorth::LRET01"], 10.5 * 20000);
});
