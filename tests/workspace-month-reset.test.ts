import assert from 'node:assert/strict';
import test from 'node:test';
import { INITIAL_APP_DATA } from '../src/app/constants/initial-data';
import type { AppData, TableData } from '../src/app/types';
import { grossPayTotalsByMonth } from '../src/app/lib/utils/gross-pay';
import { replacePivotCacheMonth } from '../src/app/lib/utils/pivot-utils';
import { restoreWorkspaceMonth } from '../src/app/lib/utils/workspace-month-reset';

const month = '09.2026';
const otherMonth = '08.2026';
const now = '2026-10-01T00:00:00Z';
const row = (period: string, amount: number, business = 'AHN') => ({
  'Tháng báo cáo': period, 'ID Number': 'TEST-1', Business: business, 'TOTAL PAYMENT': amount,
});
const table = (...data: TableData['data']): TableData => ({headers: Object.keys(data[0] || {}), data});

test('Gross Pay groups the same report period by BU, honoring true zero and excluding total rows', () => {
  const totals = grossPayTotalsByMonth([
    row('Tháng 9/2026', 120),
    row('2026-09', 150, ' ahn_hp '),
    {...row(month, 0, 'AHP'), 'Grand Total': 50},
    {...row('', 20), _fileMonth: month},
    {...row(otherMonth, 7), _fileMonth: month},
    {...row(month, 1000), _isTotalRow: true},
    row('09.2025', 9),
    row('T9', 500),
  ]);
  assert.deepEqual(Object.fromEntries(totals.get(month)!), {AHN: 140, AHP: 150});
  assert.equal(totals.get(otherMonth)?.get('AHN'), 7);
  assert.equal(totals.get('09.2025')?.get('AHN'), 9);
});

test('global Reset replaces one month of each source and keeps saved Deductions frozen', () => {
  const augustGross = row(otherMonth, 800);
  const augustRoster = {date: '2026-08-31', duration: 4};
  const septemberRoster = {date: '02/09/2026', duration: 2};
  const deductions = table({...row(month, -100), 'Nghiệp vụ': 'Hold', 'Tháng phát sinh': month});
  const bank = (period: string, amount: number) => ({'Tháng báo cáo': period, 'Document ID': 'TEST-1', 'Payment Amount': amount});
  const augustBatch = {table: table(bank(otherMonth, 800))};
  const data: AppData = {...structuredClone(INITIAL_APP_DATA), globalMonth: month,
    Sheet1_AE: table(augustGross, row(month, 900), row(month, 901)),
    SavedGrossPayByMonth: {[month]: table(row(month, 1200))},
    Hold_AE: deductions,
    HoldCarrySnapshots: {[month]: {savedAt: now, eligibleCount: 1, carriedCount: 0, rows: deductions.data}},
    Bank_North_AE: table(row(otherMonth, 800), row(month, 700)),
    BankExport: table(bank(month, 700)),
    Timesheet_Roster: [augustRoster, {...septemberRoster, duration: 99}],
    PivotMasterClearedMonths: [otherMonth, month], AuditClearedTables: {main: true, detail: true},
    TransactionMonthCache: {activePeriod: '2026-09', months: {'2026-08': augustBatch}, cloudMonths: {
      '2026-09': {id: '9', period: '2026-09-01', created_at: now, rows: [bank('2026-09', 1100)]},
    }},
    TableOriginals: {Sheet1_AE: table(row(month, 1000)), Hold_AE: table(row(month, 0)),
      Bank_North_AE: table(row(month, 1000)), Timesheet_Roster: [septemberRoster], BankExport: table(bank(otherMonth, 10))},
  };

  const first = restoreWorkspaceMonth(data, month, now);
  const {data: reset, batchRestored} = restoreWorkspaceMonth(first.data, month, now);
  assert.equal(batchRestored, true);
  assert.equal(reset.globalMonth, month);
  assert.deepEqual(reset.Sheet1_AE.data, [augustGross, row(month, 1200)]);
  assert.equal(reset.Sheet1_AE.data[0], augustGross);
  assert.equal(reset.Hold_AE, deductions);
  assert.equal(reset.HoldCarrySnapshots, data.HoldCarrySnapshots);
  assert.deepEqual(reset.Bank_North_AE.data, [row(otherMonth, 800), row(month, 1000)]);
  assert.deepEqual(reset.Timesheet_Roster, [augustRoster, septemberRoster]);
  assert.deepEqual(reset.PivotMasterClearedMonths, [otherMonth]);
  assert.deepEqual(reset.AuditClearedTables, {main: false, detail: false});
  assert.equal(reset.TransactionMonthCache?.months['2026-08'], augustBatch);
  assert.equal(reset.BankExport.data.length, 1);
  assert.equal(reset.ReconciliationByMonth?.[month].variance, 0);
  assert.equal(restoreWorkspaceMonth(data, otherMonth).data, data);
});

test('global Reset keeps working Batch Payment when its only baseline belongs to another month', () => {
  const data: AppData = {...structuredClone(INITIAL_APP_DATA), globalMonth: month,
    BankExport: table(row(month, 100)), Sheet1_AE: table(row(month, 500)),
    TableOriginals: {BankExport: table(row(otherMonth, 500)), Sheet1_AE: table(row(month, 200))},
  };
  const reset = restoreWorkspaceMonth(data, month);
  assert.equal(reset.batchRestored, false);
  assert.equal(reset.data.BankExport, data.BankExport);
  assert.deepEqual(reset.data.Sheet1_AE.data, [row(month, 200)]);
});

test('rebuilding one Pivot Master month preserves manual edits in other months', () => {
  const cached = {AHN: {HN01: {[otherMonth]: {LXO: 800}, [month]: {LXO: 999}}}};
  const rebuilt = {AHN: {HN01: {[otherMonth]: {LXO: 700}, [month]: {LXO: 1000}}}};
  assert.deepEqual(replacePivotCacheMonth(cached, rebuilt, month), {AHN: {HN01: {[otherMonth]: {LXO: 800}, [month]: {LXO: 1000}}}});
});
