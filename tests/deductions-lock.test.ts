import assert from 'node:assert/strict';
import test from 'node:test';
import { INITIAL_APP_DATA } from '../src/app/constants/initial-data';
import type { AppData } from '../src/app/types';
import { captureSavedDeductions, protectSavedDeductions, restoreSavedDeductions, isDeductionsMonthLocked } from '../src/app/lib/utils/deductions-lock';
import { applyBulkReconciliationReferenceSync, applyTransactionReferenceSync, hasPendingReconciliationReferenceSync } from '../src/app/lib/utils/transaction-reference-sync';
import { carryEligibleHoldsToNextMonth } from '../src/app/lib/utils/hold-carryover';
import { encode, decode } from '../src/app/lib/cloud/codec';
import { build } from 'esbuild';
import { createRequire } from 'node:module';

function fixture(): AppData {
  const row = {id: 'feb', 'Tháng báo cáo': '02.2026', 'Tháng phát sinh': '02.2026',
    'ID Number': '00123', 'Full name': 'TEST EMPLOYEE', 'Bank Account Number': '00111111',
    'TOTAL PAYMENT': -100, 'Sheet Source': 'Hold T2', 'Nghiệp vụ': 'Hold', BU: 'AHN', L07: 'HN001'};
  return {...structuredClone(INITIAL_APP_DATA), globalMonth: '02.2026',
    Hold_AE: {headers: Object.keys(row), data: [row, {...row, id: 'march', 'Tháng báo cáo': '03.2026'}]},
    HoldCarrySnapshots: {'02.2026': {savedAt: '2026-10-30T01:00:00Z', eligibleCount: 1, carriedCount: 1}},
    BankExport: {headers: [], data: [{'Tháng báo cáo': '02.2026', 'Document ID': '00123', 'Beneficiary Name': 'TEST EMPLOYEE', 'Beneficiary Account No.': '00999999', 'Payment Amount': 900}]}};
}

test('bulk and per-row Reconcile cannot modify saved Deductions; only editable Gross Pay is counted', () => {
  const data = captureSavedDeductions(fixture());
  const args = {grossRows: [], deductionRows: data.Hold_AE.data, transactionRows: data.BankExport.data,
    reportMonth: data.globalMonth, lockedDeductionMonths: data.HoldCarrySnapshots};
  assert.equal(hasPendingReconciliationReferenceSync(args), false);
  for (const sync of [applyBulkReconciliationReferenceSync, applyTransactionReferenceSync]) {
    const result = sync(args);
    assert.deepEqual(result.deductionRows, data.Hold_AE.data);
    assert.equal(result.correctedCells, 0);
    const gross = {...data.Hold_AE.data[0], 'TOTAL PAYMENT': 1000};
    const mixed = sync({...args, grossRows: [gross]});
    assert.equal(mixed.grossRows[0]['Bank Account Number'], '00999999');
    assert.equal(mixed.deductionRows[0]['Bank Account Number'], '00111111');
    assert.equal(mixed.correctedRows, 1);
  }
});

test('local write protection preserves every saved field and lock while allowing other months to change', () => {
  const data = restoreSavedDeductions(fixture());
  const candidate = {...data, HoldCarrySnapshots: {}, Hold_AE: {...data.Hold_AE,
    data: data.Hold_AE.data.map(row => ({...row, 'TOTAL PAYMENT': 999, 'ID Number': 'OTHER', 'Nghiệp vụ': 'Add'}))}};
  const guarded = protectSavedDeductions(data, candidate);
  assert.deepEqual(guarded.Hold_AE.data[0], data.Hold_AE.data[0]);
  assert.equal(guarded.Hold_AE.data[1]['TOTAL PAYMENT'], 999);
  assert.equal(guarded.HoldCarrySnapshots?.['02.2026'], data.HoldCarrySnapshots?.['02.2026']);
  assert.equal(isDeductionsMonthLocked(guarded.HoldCarrySnapshots, '2026-02'), true);
});

test('deletion, insertion, Reset and Undo-style lock removal cannot bypass saved month protection', () => {
  const data = restoreSavedDeductions(fixture());
  for (const rows of [[], [{...data.Hold_AE.data[0], id: 'new', 'TOTAL PAYMENT': 555}], [...data.Hold_AE.data, {...data.Hold_AE.data[0], id: 'duplicate'}]]) {
    const guarded = protectSavedDeductions(data, {...data, HoldCarrySnapshots: {}, Hold_AE: {...data.Hold_AE, data: rows}});
    assert.deepEqual(guarded.Hold_AE.data.filter(row => row['Tháng báo cáo'] === '02.2026'), data.HoldCarrySnapshots!['02.2026'].rows);
    assert.equal(isDeductionsMonthLocked(guarded.HoldCarrySnapshots, '02.2026'), true);
  }
});

test('a legacy in-place writer cannot alter the separately captured saved rows', () => {
  const data = restoreSavedDeductions(fixture());
  const before = structuredClone(data.HoldCarrySnapshots!['02.2026'].rows);
  data.Hold_AE.data[0]['TOTAL PAYMENT'] = 777;
  const guarded = protectSavedDeductions(data, {...data, Hold_AE: {...data.Hold_AE}});
  assert.deepEqual(guarded.Hold_AE.data.filter(row => row['Tháng báo cáo'] === '02.2026'), before);
  assert.deepEqual(guarded.HoldCarrySnapshots!['02.2026'].rows, before);
});

test('only explicit Delete saved unlocks its month and Reconcile works again afterward', () => {
  const data = restoreSavedDeductions(fixture());
  const withTwoLocks = captureSavedDeductions({...data, HoldCarrySnapshots: {...data.HoldCarrySnapshots,
    '03.2026': {savedAt: '2026-10-30T02:00:00Z', eligibleCount: 1, carriedCount: 0}}});
  const unlocked = protectSavedDeductions(withTwoLocks, {...withTwoLocks, HoldCarrySnapshots: {}}, {unlockDeductionsMonth: '2026-02'});
  assert.equal(isDeductionsMonthLocked(unlocked.HoldCarrySnapshots, '02.2026'), false);
  assert.equal(isDeductionsMonthLocked(unlocked.HoldCarrySnapshots, '03.2026'), true);
  const synced = applyBulkReconciliationReferenceSync({grossRows: [], deductionRows: unlocked.Hold_AE.data,
    transactionRows: unlocked.BankExport.data, reportMonth: '02.2026', lockedDeductionMonths: unlocked.HoldCarrySnapshots});
  assert.equal(synced.deductionRows[0]['Bank Account Number'], '00999999');
  assert.equal(synced.deductionRows[1]['Bank Account Number'], '00111111');
});

test('saving an earlier month cannot add or remove carried rows in an already saved next month', () => {
  const data = restoreSavedDeductions(fixture());
  const january = {...data.Hold_AE.data[0], id: 'jan', 'Tháng báo cáo': '01.2026', 'Tháng phát sinh': '01.2026'};
  const carry = carryEligibleHoldsToNextMonth({sourceRows: [january], existingRows: [january, ...data.Hold_AE.data], reportMonth: '01.2026'});
  const guarded = protectSavedDeductions(data, {...data, Hold_AE: {...data.Hold_AE, data: carry.rows}});
  assert.deepEqual(guarded.Hold_AE.data.filter(row => row['Tháng báo cáo'] === '02.2026'), data.HoldCarrySnapshots!['02.2026'].rows);
});

test('saved rows, display and totals survive cloud serialization and new source data without drift or repeated writes', async () => {
  const data = restoreSavedDeductions(fixture());
  const restored = restoreSavedDeductions(await decode(await encode(data)) as AppData);
  const snapshot = structuredClone(restored.HoldCarrySnapshots!['02.2026']);
  const changed = protectSavedDeductions(restored, {...restored, BankExport: {headers: [], data: []},
    Sheet1_AE: {headers: [], data: [{'Tháng báo cáo': '02.2026', 'ID Number': '00123', 'TOTAL PAYMENT': 123456789}]}});
  assert.deepEqual(changed.HoldCarrySnapshots!['02.2026'], snapshot);
  assert.equal(restoreSavedDeductions(changed), changed);
  assert.deepEqual(protectSavedDeductions(changed, {...changed, Hold_AE: {...changed.Hold_AE, data: changed.Hold_AE.data.map(row => ({...row}))}}), changed);
  const remoteUnlock = {...changed, HoldCarrySnapshots: {}};
  assert.equal(restoreSavedDeductions(remoteUnlock), remoteUnlock, 'another computer can publish an explicit unlock');
});

test('the real Deductions Save and Delete saved callbacks freeze the table and block a stale Reconcile click until explicit unlock', async () => {
  const initial = fixture();
  initial.HoldCarrySnapshots = {};
  const bundle = await build({
    stdin: {resolveDir: process.cwd(), contents: `
      import React from 'react';
      import {renderToString} from 'react-dom/server';
      import {HoldAETable} from './src/app/pages/03-master/components/HoldAETable';
      import {buttons} from 'capture-react';
      import {state} from 'lock-app-data';
      import {table} from 'lock-table';
      import {applyBulkReconciliationReferenceSync} from './src/app/lib/utils/transaction-reference-sync';
      export async function run() {
        globalThis.React = React;
        globalThis.localStorage = {getItem(){return null;},setItem(){},removeItem(){}};
        state.data = ${JSON.stringify(initial)};
        const render = () => {buttons.length=0; renderToString(React.createElement(HoldAETable,{searchTerm:''}));};
        const action = label => {const button=buttons.find(button=>button.label===label);if(!button)throw new Error('Missing '+label+': '+JSON.stringify(buttons.map(button=>button.label)));return button;};
        render();
        const staleSync=action('Đồng bộ từ Reconcile').onClick;
        await action('SAVE HOLD').onClick();
        render();
        const saved=JSON.stringify(table.props.data);
        const savedTotals=JSON.stringify(state.data.HoldCarrySnapshots['02.2026'].reconciliationTotals);
        const locked=!table.props.isEditable && action('Đồng bộ từ Reconcile').disabled;
        staleSync();
        render();
        const unchanged=saved===JSON.stringify(table.props.data);
        state.update(prev=>({...prev,Sheet1_AE:{headers:[],data:[{'Tháng báo cáo':'02.2026','ID Number':'00123','TOTAL PAYMENT':999999999}]}}));
        render();
        const totalsFixed=savedTotals===JSON.stringify(state.data.HoldCarrySnapshots['02.2026'].reconciliationTotals);
        action('Xóa bản lưu').onClick();
        render();
        const unlocked=table.props.isEditable && !state.data.HoldCarrySnapshots['02.2026'];
        action('Đồng bộ từ Reconcile').onClick();
        return {locked,unchanged,totalsFixed,unlocked,account:state.data.Hold_AE.data.find(row=>row.id==='feb')['Bank Account Number']};
      }
    `}, bundle:true, write:false, platform:'node', format:'cjs', packages:'external',
    jsx:'transform', jsxFactory:'React.createElement', jsxFragment:'React.Fragment',
    tsconfigRaw:{compilerOptions:{jsx:'react'}},
    plugins:[{name:'deductions-ui-boundaries',setup(builder){
      builder.onResolve({filter:/^react$|^capture-react$/},args=>args.namespace==='lock-fixture'?{path:'react',external:true}:{path:'capture-react',namespace:'lock-fixture'});
      builder.onResolve({filter:/AppDataContext$|^lock-app-data$/},()=>({path:'lock-app-data',namespace:'lock-fixture'}));
      builder.onResolve({filter:/components\/DataTable$|^lock-table$/},()=>({path:'lock-table',namespace:'lock-fixture'}));
      builder.onResolve({filter:/CloudWorkspaceContext$/},()=>({path:'lock-cloud',namespace:'lock-fixture'}));
      builder.onLoad({filter:/.*/,namespace:'lock-fixture'},args=>({loader:'js',resolveDir:process.cwd(),contents:args.path==='capture-react' ? `
        import * as Actual from 'react'; export * from 'react';
        export const buttons=[];
        function label(node){return Array.isArray(node)?node.map(label).join(''):typeof node==='string'?node:node?.props?label(node.props.children):'';}
        export default {...Actual,createElement(type,props,...children){if(props?.onClick)buttons.push({...props,label:label(children).trim()});return Actual.createElement(type,props,...children);}};
      `:args.path==='lock-cloud'?`export function useCloudWorkspace(){return null;}`:args.path==='lock-app-data'?`
        import {captureSavedDeductions,protectSavedDeductions} from './src/app/lib/utils/deductions-lock';
        export const state={data:null,update(updater,_history,_persist,_sources,options){const previous=captureSavedDeductions(state.data);state.data=protectSavedDeductions(previous,updater(previous),options);}};
        export function useAppData(){return {appData:state.data,updateAppData:state.update};}
      `:`import React from 'react';export const table={};export const DataTable=React.forwardRef((props,ref)=>{table.props=props;return React.createElement('div',null,props.footerStatusContent);});`}));
    }}],
  });
  const module = {exports: {} as {run: () => Promise<unknown>}};
  new Function('require','module','exports',bundle.outputFiles[0].text)(createRequire(import.meta.url),module,module.exports);
  assert.deepEqual(await module.exports.run(), {locked:true,unchanged:true,totalsFixed:true,unlocked:true,account:'00999999'});
});
