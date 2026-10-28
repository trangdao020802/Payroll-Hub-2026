import assert from 'node:assert/strict';
import test from 'node:test';
import { INITIAL_APP_DATA } from '../src/app/constants/initial-data';
import type { AppData, TableData } from '../src/app/types';
import { restoreTableOriginals, trackTableOriginals } from '../src/app/lib/utils/table-originals';
import { syncReportingMonthReconciliation } from '../src/app/lib/utils/reconciliation-sync';
import { applyLocalTransactionVersions } from '../src/app/lib/utils/transaction-local-history';
import { replaceTransactionSnapshotsInAppData } from '../src/app/lib/utils/transaction-snapshot';
import type { TransactionVersion } from '../src/app/lib/transaction-history-store';

const table = (month: string, amount: number, documentId = `ID-${month}`): TableData => ({
  headers: ['Tháng báo cáo', 'Document ID', 'Beneficiary Account No.', 'Payment Amount'],
  data: [{id: `row-${month}`, 'Tháng báo cáo': month, 'Document ID': documentId,
    'Beneficiary Account No.': '0012345678', 'Beneficiary Name': 'NGUYEN VAN A', 'Payment Amount': amount}],
});

function workspace(): AppData {
  const app = structuredClone(INITIAL_APP_DATA);
  app.globalMonth = '05.2026';
  app.BankExport = table('05.2026', 500);
  app.TransactionActivity = {lastAction: 'saved', editCount: 4, saveVersion: 7};
  app.TransactionMonthCache = {activePeriod: '2026-05', months: {
    '2026-04': {table: table('04.2026', 400), activity: {lastAction: 'saved', editCount: 2, saveVersion: 3}},
    '2026-05': {table: app.BankExport, activity: app.TransactionActivity},
  }};
  app.TableOriginals = {BankExport: table('04.2026', 40)};
  return app;
}

test('Reset rejects a previous-month global original without changing payroll or its cache', () => {
  const app = workspace();
  const before = structuredClone(app);
  assert.throws(() => restoreTableOriginals(app, ['BankExport']), /đúng tháng 05\.2026/);
  assert.deepEqual(app, before);
});

test('first Batch Payment edit records its own month despite a stale older global original', () => {
  const app = workspace();
  const edited = trackTableOriginals(app, {...app, BankExport: table('05.2026', 999)}, true);
  assert.equal(edited.TableOriginals?.BankExportByMonth?.['2026-05'].data[0]['Payment Amount'], 500);
  const restored = restoreTableOriginals(edited, ['BankExport']);
  assert.equal(restored.globalMonth, '05.2026');
  assert.equal(restored.BankExport.data.length, 1);
  assert.equal(restored.BankExport.data[0]['Tháng báo cáo'], '2026-05');
  assert.equal(restored.BankExport.data[0]['Payment Amount'], 500);
  assert.equal(restored.TransactionMonthCache?.activePeriod, '2026-05');
  assert.equal(restored.TransactionMonthCache?.months['2026-05'].table.data[0]['Payment Amount'], 500);
  assert.equal(restored.TransactionMonthCache?.months['2026-04'], app.TransactionMonthCache?.months['2026-04']);
  assert.equal(restored.TransactionActivity?.saveVersion, 8);
  assert.equal(restored.TransactionActivity?.lastAction, 'saved');
  assert.equal(restored.ReconciliationByMonth?.['05.2026'].actual, 500);
  assert.equal(restored.Sheet1_AE, app.Sheet1_AE);
  assert.equal(restored.Hold_AE, app.Hold_AE);
  assert.equal(restored.Bank_North_AE, app.Bank_North_AE);
});

test('navigation never captures outgoing April rows as the May original', () => {
  const app = workspace();
  const april = {...app, globalMonth: '04.2026', BankExport: table('04.2026', 400)};
  const may = trackTableOriginals(april, {...april, globalMonth: '05.2026', BankExport: table('05.2026', 500)}, true);
  assert.equal(may.TableOriginals?.BankExportByMonth?.['2026-05'], undefined);
  const changedSelectorFirst = {...april, globalMonth: '05.2026'};
  const reconciled = trackTableOriginals(changedSelectorFirst, {...changedSelectorFirst, BankExport: table('05.2026', 500)}, true);
  assert.equal(reconciled.TableOriginals?.BankExportByMonth?.['2026-05'], undefined);
});

test('per-month originals survive edits, reload and navigation across December and January', () => {
  let app = workspace();
  app.globalMonth = '12.2025';
  app.Bank_North_AE.data = ['12.2025', '01.2026'].map((month, index) => ({
    'Tháng báo cáo': month, 'ID Number': month, 'TOTAL PAYMENT': (index + 1) * 100,
  }));
  app = syncReportingMonthReconciliation(app, '12.2025');
  app = trackTableOriginals(app, {...app, BankExport: table('12.2025', 999),
    TransactionActivity: {...app.TransactionActivity!, lastAction: 'saved'}}, true);
  app = syncReportingMonthReconciliation(app, '01.2026');
  app = trackTableOriginals(app, {...app, BankExport: table('01.2026', 888),
    TransactionActivity: {...app.TransactionActivity!, lastAction: 'saved'}}, true);
  const restored = restoreTableOriginals(structuredClone(app), ['BankExport']);
  assert.equal(restored.globalMonth, '01.2026');
  assert.equal(restored.BankExport.data[0]['Payment Amount'], 200);
  const december = restoreTableOriginals(syncReportingMonthReconciliation(restored, '12.2025'), ['BankExport']);
  assert.equal(december.BankExport.data[0]['Payment Amount'], 100);
  assert.equal(december.TransactionMonthCache?.months['2026-01'].table.data[0]['Payment Amount'], 200);
});

test('a mixed legacy original restores only the selected month and keeps other current rows', () => {
  const app = workspace();
  const aprilRows = app.TransactionMonthCache!.months['2026-04'].table.data;
  app.BankExport.data.push(...aprilRows);
  app.TableOriginals = {BankExport: {headers: app.BankExport.headers,
    data: [...table('04.2026', 40).data, ...table('05.2026', 50).data]}};
  const restored = restoreTableOriginals(app, ['BankExport']);
  assert.deepEqual(restored.BankExport.data.map(row => row['Payment Amount']), [50, 400]);
  assert.equal(restored.BankExport.data[1], aprilRows[0]);
  assert.equal(restored.TransactionMonthCache?.months['2026-04'], app.TransactionMonthCache?.months['2026-04']);
  assert.equal(restored.TransactionMonthCache?.months['2026-05'].table.data.length, 1);
});

test('an original filed under the wrong month and an unlabelled legacy original fail safely', () => {
  const app = workspace();
  app.TableOriginals = {BankExportByMonth: {'2026-05': table('04.2026', 40)}};
  assert.throws(() => restoreTableOriginals(app, ['BankExport']), /không đúng tháng/);
  app.TableOriginals = {BankExport: {headers: [], data: [{'Payment Amount': 40}]}};
  assert.throws(() => restoreTableOriginals(app, ['BankExport']), /đúng tháng/);
});

test('an intentionally empty month original remains empty without restoring previous-month payroll', () => {
  const app = workspace();
  app.TableOriginals = {BankExportByMonth: {'2026-05': {headers: app.BankExport.headers, data: []}}};
  const restored = restoreTableOriginals(app, ['BankExport']);
  assert.deepEqual(restored.BankExport.data, []);
  assert.deepEqual(restored.TransactionMonthCache?.months['2026-05'].table.data, []);
  assert.equal(restored.ReconciliationByMonth?.['05.2026'].actual, 0);
});

test('verified ID/account corrections become the reset baseline for every affected month', () => {
  const app = workspace();
  app.TableOriginals = {BankExportByMonth: {'2026-04': table('04.2026', 40), '2026-05': table('05.2026', 50)}};
  const versions: TransactionVersion[] = ['2026-04', '2026-05'].map((period, index) => ({
    id: `${index + 1}`, period: `${period}-01`, created_at: '2026-10-28T03:00:00Z',
    rows: table(period, (index + 4) * 100, 'CORRECT-ID').data.map(row => ({...row, 'Beneficiary Account No.': '0099887766'})),
  }));
  for (const loaded of [applyLocalTransactionVersions(app, versions), replaceTransactionSnapshotsInAppData(app, versions)]) {
    const restored = restoreTableOriginals(loaded, ['BankExport']);
    assert.equal(restored.BankExport.data[0]['Document ID'], 'CORRECT-ID');
    assert.equal(restored.BankExport.data[0]['Beneficiary Account No.'], '0099887766');
    const april = restoreTableOriginals(syncReportingMonthReconciliation(restored, '04.2026'), ['BankExport']);
    assert.equal(april.BankExport.data[0]['Document ID'], 'CORRECT-ID');
    assert.equal(april.BankExport.data[0]['Payment Amount'], 400);
  }
});

test('explicit regeneration updates only its month original; reopening and repeated Reset never append rows', () => {
  const app = workspace();
  app.TableOriginals = {BankExportByMonth: {'2026-04': table('04.2026', 40), '2026-05': table('05.2026', 50)}};
  const imported = trackTableOriginals(app, {...app, BankExport: table('05.2026', 550)}, false, ['BankExport']);
  assert.equal(imported.TableOriginals?.BankExportByMonth?.['2026-04'], app.TableOriginals.BankExportByMonth?.['2026-04']);
  let restored = restoreTableOriginals(imported, ['BankExport']);
  for (let index = 0; index < 3; index++) {
    restored = restoreTableOriginals(syncReportingMonthReconciliation(structuredClone(restored), '05.2026'), ['BankExport']);
    assert.equal(restored.BankExport.data.length, 1);
    assert.equal(restored.BankExport.data[0]['Payment Amount'], 550);
  }
});
