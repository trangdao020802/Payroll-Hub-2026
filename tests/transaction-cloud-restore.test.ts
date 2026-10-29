import assert from 'node:assert/strict';
import test from 'node:test';
import { INITIAL_APP_DATA } from '../src/app/constants/initial-data';
import type { AppData, TableData } from '../src/app/types';
import type { TransactionVersion } from '../src/app/lib/transaction-history-store';
import { restoreSavedTransactionMonths } from '../src/app/lib/utils/transaction-cloud-restore';
import { restoreTableOriginals } from '../src/app/lib/utils/table-originals';
import { calculateReconciliationTotals, createReconciliationTotalsCalculator, syncReportingMonthReconciliation } from '../src/app/lib/utils/reconciliation-sync';
import { cacheActiveTransactionMonth } from '../src/app/lib/utils/reconciliation-sync';

const saved = {lastAction: 'saved' as const, editCount: 0, saveVersion: 1};
const table = (data: TableData['data']): TableData => ({headers: ['Tháng báo cáo', 'Payment Amount'], data});
function version(period = '2026-02', amount = 700_595_349, id = '181'): TransactionVersion {
  return {id, period: `${period}-01`, created_at: '2026-10-29T05:24:30Z',
    rows: [{'Tháng báo cáo': period, 'Document ID': '00123', 'Payment Amount': amount,
      'Beneficiary Account No.': '00998877', 'Beneficiary Name': 'TEST'}]};
}
function workspace(): AppData {
  return {...structuredClone(INITIAL_APP_DATA), globalMonth: '02.2026',
    Sheet1_AE: table([{'ID Number': '00123', 'Tháng báo cáo': '02.2026', 'TOTAL PAYMENT': 700_595_349}]),
    BankExport: table([]), Bank_North_AE: table([]),
    TransactionMonthCache: {activePeriod: '2026-02', months: {}}};
}

test('missing Batch Payment restores from its saved month, and Reset cannot bring back an obsolete generated baseline', () => {
  const data = workspace();
  data.TransactionActivity = saved; // An old global marker, without a retained empty month.
  data.TableOriginals = {BankExportByMonth: {'2026-02': table([])}, BankExport: table(version('2026-01', 10).rows)};
  const restored = restoreSavedTransactionMonths(data, [version()]);
  assert.equal(restored.BankExport.data.length, 1);
  assert.equal(restored.TransactionMonthCache?.months['2026-02'].cloudVersionId, '181');
  assert.equal(calculateReconciliationTotals(restored).variance, 0);
  // Even if an older build regenerates TableOriginals, saved history is authoritative for Reset.
  const reset = restoreTableOriginals({...restored, TableOriginals: data.TableOriginals}, ['BankExport']);
  assert.equal(reset.BankExport.data.length, 1);
  assert.equal(reset.BankExport.data[0]['Beneficiary Account No.'], '00998877');
  assert.equal(reset.globalMonth, '02.2026');
  assert.equal(calculateReconciliationTotals(reset).variance, 0);
  assert.equal(reset.Sheet1_AE, data.Sheet1_AE);
  assert.equal(reset.Hold_AE, data.Hold_AE);
  assert.equal(reset.Bank_North_AE, data.Bank_North_AE);
  assert.equal(restoreSavedTransactionMonths(restored, [version()]), restored);
});

test('all nine saved months survive serialization, navigation and repeated Reset without duplicating rows or hiding real differences', () => {
  const versions = Array.from({length: 9}, (_, index) => version(`2026-${String(index + 1).padStart(2, '0')}`, 700_595_349 + (index === 8 ? 7 : 0), String(index + 1)));
  const data = workspace();
  data.Sheet1_AE = table(versions.map(item => ({'ID Number': '00123', 'Tháng báo cáo': item.period.slice(0, 7), 'TOTAL PAYMENT': 700_595_349})));
  let restored = structuredClone(restoreSavedTransactionMonths(data, versions));
  for (const [index, item] of versions.entries()) {
    restored = syncReportingMonthReconciliation(restored, item.period.slice(0, 7));
    for (let repeat = 0; repeat < 2; repeat++) {
      restored = restoreTableOriginals(restored, ['BankExport']);
      assert.equal(restored.BankExport.data.length, 1);
      assert.equal(calculateReconciliationTotals(restored).variance, index === 8 ? 7 : 0);
    }
  }
  assert.equal(Object.keys(restored.TransactionMonthCache!.months).length, 9);
  assert.equal(Object.keys(restored.TransactionMonthCache!.cloudMonths!).length, 9);
});

test('a newer history version replaces a clean working month and its Reset baseline', () => {
  const first = restoreSavedTransactionMonths(workspace(), [version()]);
  const second = restoreSavedTransactionMonths(first, [version('2026-02', 700_595_361, '182')]);
  assert.equal(second.BankExport.data[0]['Payment Amount'], 700_595_361);
  assert.equal(calculateReconciliationTotals(second).variance, 12);
  assert.equal(restoreTableOriginals(second, ['BankExport']).BankExport.data[0]['Payment Amount'], 700_595_361);
  assert.equal(restoreSavedTransactionMonths(second, [version('2026-02', 700_595_361, '182')]), second);
});

test('history refresh preserves working edits, while Reset explicitly restores the latest saved version', () => {
  const first = restoreSavedTransactionMonths(workspace(), [version()]);
  const edited = cacheActiveTransactionMonth({...first, BankExport: table(version('2026-02', 123).rows),
    TransactionActivity: {...saved, lastAction: 'edited'}});
  const refreshed = restoreSavedTransactionMonths(edited, [version('2026-02', 456, '182')]);
  assert.equal(refreshed.BankExport.data[0]['Payment Amount'], 123);
  assert.equal(refreshed.TransactionMonthCache?.months['2026-02'].table.data[0]['Payment Amount'], 123);
  assert.equal(refreshed.TransactionMonthCache?.cloudMonths?.['2026-02'].rows[0]['Payment Amount'], 456);
  assert.equal(restoreTableOriginals(refreshed, ['BankExport']).BankExport.data[0]['Payment Amount'], 456);
  const uncachedEdit = {...first, BankExport: edited.BankExport, TransactionActivity: edited.TransactionActivity};
  assert.equal(restoreSavedTransactionMonths(uncachedEdit, [version('2026-02', 456, '182')]).BankExport, edited.BankExport);
});

test('explicit empty months and intentional clears remain empty after history refresh', () => {
  const first = restoreSavedTransactionMonths(workspace(), [version()]);
  const emptied = cacheActiveTransactionMonth({...first, BankExport: table([]), TransactionActivity: saved});
  const refreshed = restoreSavedTransactionMonths(emptied, [version('2026-02', 456, '182')]);
  assert.deepEqual(refreshed.BankExport.data, []);
  assert.equal(calculateReconciliationTotals(refreshed).hasActual, true);
  assert.equal(calculateReconciliationTotals(refreshed).variance, -700_595_349);
  const cleared = {...workspace(), TransactionHistoryRestoreDisabled: true};
  assert.equal(restoreSavedTransactionMonths(cleared, [version()]), cleared);
});

test('wrong-month history fails without mutating the working payroll', () => {
  const data = workspace();
  const before = structuredClone(data);
  assert.throws(() => restoreSavedTransactionMonths(data, [{...version(), rows: version('2026-01').rows}]), /tháng/);
  assert.deepEqual(data, before);
});

test('missing actual data is distinct from an intentionally saved zero or a real nonzero variance', () => {
  const data = workspace();
  const calculate = createReconciliationTotalsCalculator('02.2026');
  for (const totals of [calculate(data), calculateReconciliationTotals(data)]) {
    assert.equal(totals.hasActual, false);
    assert.equal(totals.actualSource, 'missing');
    assert.equal(totals.expected, 700_595_349);
  }
  const loaded = restoreSavedTransactionMonths(data, [version('2026-02', 700_595_359)]);
  assert.deepEqual(calculate(loaded), calculateReconciliationTotals(loaded));
  assert.equal(calculate(loaded).hasActual, true);
  assert.equal(calculate(loaded).variance, 10);
  const wrongMonth = {...data, TransactionActivity: saved, BankExport: table(version('2026-01', 10).rows)};
  assert.equal(calculateReconciliationTotals(wrongMonth).hasActual, false);
});
