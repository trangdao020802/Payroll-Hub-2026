import assert from "node:assert/strict";
import test from "node:test";
import { createRequire } from "node:module";
import { build } from "esbuild";

test("Batch Payment clears drilldown searches and column filters without resetting saved payroll", async () => {
  const result = await build({
    stdin: {contents: `
      import React from 'react';
      import { renderToString } from 'react-dom/server';
      import { BulkPayment } from './src/app/pages/04-balance/BulkPayment';
      import { INITIAL_APP_DATA } from './src/app/constants/initial-data';
      import { buttons } from 'test-react-capture';
      import { setData, getData } from 'test-app-data';
      import { tableState } from 'test-data-table';
      import { historyState } from 'test-history-panel';
      export function run() {
        const storage = {getItem() { return null; }, setItem() {}, removeItem() {}};
        globalThis.localStorage = storage;
        globalThis.sessionStorage = storage;
        setData({...structuredClone(INITIAL_APP_DATA), globalMonth: '05.2026',
          BankExport: {headers: ['Document ID', 'Payment Amount'], data: [
            {id: 'a', 'Tháng báo cáo': '05.2026', 'Document ID': 'FIXED-ID', 'Beneficiary Account No.': '0012345678', 'Payment Amount': 500},
            {id: 'b', 'Tháng báo cáo': '05.2026', 'Document ID': 'OTHER-ID', 'Beneficiary Account No.': '0099887766', 'Payment Amount': 600},
          ]}});
        const before = JSON.stringify(getData());
        let query = 'OLD-ACCOUNT';
        function render() {
          buttons.length = 0;
          return renderToString(React.createElement(BulkPayment, {searchTerm: query, onSearchTermChange: value => {query = value;}}));
        }
        const html = render();
        const clearSearch = buttons.find(button => button.label === 'Bỏ tìm kiếm');
        if (!clearSearch) throw new Error('Missing search action: ' + JSON.stringify(buttons.map(button => button.label)) + ' / ' + html.includes('Bỏ tìm kiếm'));
        clearSearch.onClick();
        const afterSearch = {query, clears: tableState.clears, unchanged: before === JSON.stringify(getData())};
        query = 'OLD-ID';
        render();
        buttons.find(button => button.label === 'Bỏ lọc').onClick();
        const afterFilters = {query, clears: tableState.clears, unchanged: before === JSON.stringify(getData())};
        query = 'WRONG-ID-FROM-RECONCILE';
        render();
        historyState.props.onApplyVersions([{id: '12', period: '2026-05-01', created_at: '2026-10-28T03:00:00Z', rows: getData().BankExport.data}]);
        return {html, afterSearch, afterFilters, afterCorrection: {query, clears: tableState.clears,
          month: getData().globalMonth, rows: getData().BankExport.data.length,
          originalId: getData().TableOriginals.BankExportByMonth['2026-05'].data[0]['Document ID']}};
      }
    `, resolveDir: process.cwd()},
    bundle: true, write: false, platform: "node", format: "cjs", packages: "external",
    jsx: "transform", jsxFactory: "React.createElement", jsxFragment: "React.Fragment",
    plugins: [{name: "batch-ui-boundaries", setup(builder) {
      builder.onResolve({filter: /^react$|^test-react-capture$/}, args => args.namespace === "fixture"
        ? {path: "react", external: true} : {path: "test-react-capture", namespace: "fixture"});
      builder.onResolve({filter: /^react\/jsx-runtime$/}, args => args.namespace === "fixture"
        ? {path: "react/jsx-runtime", external: true} : {path: "test-jsx-capture", namespace: "fixture"});
      builder.onResolve({filter: /AppDataContext$|^test-app-data$/}, () => ({path: "test-app-data", namespace: "fixture"}));
      builder.onResolve({filter: /components\/DataTable$|^test-data-table$/}, () => ({path: "test-data-table", namespace: "fixture"}));
      builder.onResolve({filter: /TransactionHistoryPanel$|^test-history-panel$/}, () => ({path: "test-history-panel", namespace: "fixture"}));
      builder.onLoad({filter: /.*/, namespace: "fixture"}, args => {
        const contents = args.path === "test-jsx-capture" ? `
          import * as Actual from 'react/jsx-runtime';
          import {buttons, label} from 'capture-state';
          export const Fragment = Actual.Fragment;
          export function jsx(type, props, key) {
            if (type === 'button') buttons.push({...props, label: label(props?.children).trim()});
            return Actual.jsx(type, props, key);
          }
          export function jsxs(type, props, key) {
            if (type === 'button') buttons.push({...props, label: label(props?.children).trim()});
            return Actual.jsxs(type, props, key);
          }
        ` : args.path === "test-react-capture" ? `
          import * as Actual from 'react';
          export * from 'react';
          export const buttons = [];
          export function label(node) { return Array.isArray(node) ? node.map(label).join('') : typeof node === 'string' ? node : node?.props ? label(node.props.children) : ''; }
          export default {...Actual, createElement(type, props, ...children) {
            if (type === 'button') buttons.push({...props, label: label(children).trim()});
            return Actual.createElement(type, props, ...children);
          }};
        ` : args.path === "test-app-data" ? `
          let appData;
          export function setData(value) {appData = value;}
          export function getData() {return appData;}
          export function useAppData() {return {appData, updateAppData(updater) {appData = typeof updater === 'function' ? updater(appData) : {...appData, ...updater};}};}
        ` : args.path === "test-data-table" ? `
          import React from 'react';
          export const tableState = {clears: 0};
          export const isAhpBuValue = value => value === 'AHP';
          export const DataTable = React.forwardRef((props, ref) => {if (ref) ref.current = {clearAllFilters() {tableState.clears++; props.onExternalSearchChange('');}}; return null;});
        ` : `export const historyState = {}; export function TransactionHistoryPanel(props) {historyState.props = props; return null;}`;
        return {contents, loader: "js"};
      });
      builder.onResolve({filter: /^capture-state$/}, () => ({path: "test-react-capture", namespace: "fixture"}));
    }}],
  });
  const module = {exports: {} as {run: () => {
    html: string; afterSearch: {query: string; clears: number; unchanged: boolean};
    afterFilters: {query: string; clears: number; unchanged: boolean};
    afterCorrection: {query: string; clears: number; month: string; rows: number; originalId: string};
  }}};
  new Function("require", "module", "exports", result.outputFiles[0].text)(createRequire(import.meta.url), module, module.exports);
  const actual = module.exports.run();
  assert.match(actual.html, /Tìm kiếm Batch Payment/);
  assert.match(actual.html, /value="OLD-ACCOUNT"/);
  assert.deepEqual(actual.afterSearch, {query: "", clears: 0, unchanged: true});
  assert.deepEqual(actual.afterFilters, {query: "", clears: 1, unchanged: true});
  assert.deepEqual(actual.afterCorrection, {query: "", clears: 2, month: "05.2026", rows: 2, originalId: "FIXED-ID"});
});

test("BulkPayment renders adjustment totals for HOLD, ADD, CANCEL and bonus rows", async () => {
  const result = await build({
    stdin: {
      contents: `
        import React from 'react';
        import { renderToString } from 'react-dom/server';
        import { useBulkPaymentLogic } from './src/app/hooks/useBulkPaymentLogic';
        import { setData } from 'test-app-data';
        import { INITIAL_APP_DATA } from './src/app/constants/initial-data';
        export function calculate(rows) {
          setData({ ...INITIAL_APP_DATA, globalMonth: '09.2026', Hold_AE: { headers: [], data: rows },
            Sheet1_AE: { headers: [], data: [] }, Bank_North_AE: { headers: [], data: [] } });
          let value;
          function Probe() { value = useBulkPaymentLogic(); return null; }
          renderToString(React.createElement(Probe));
          return { ...value.dynamicReportStats, calculationSummary: value.calculationSummary,
            holdPaymentDetails: value.holdPaymentDetails };
        }
      `,
      resolveDir: process.cwd(),
    },
    bundle: true,
    write: false,
    platform: "node",
    format: "cjs",
    packages: "external",
    plugins: [{
      name: "app-data-boundary",
      setup(builder) {
        builder.onResolve({ filter: /AppDataContext$|^test-app-data$/ }, () => ({ path: "test-app-data", namespace: "fixture" }));
        builder.onLoad({ filter: /.*/, namespace: "fixture" }, () => ({ contents: `
          let appData;
          export function setData(value) { appData = value; }
          export function useAppData() { return { appData, updateAppData() {} }; }
        ` }));
      },
    }],
  });
  const module = { exports: {} as { calculate: (rows: Record<string, unknown>[]) => {
    sameMonthHoldTotal: number; sameMonthAddTotal: number; diffMonthAddTotal: number;
    holdAddItems: { type: string; amount: number }[]; finalTotals: Record<string, number>;
    calculationSummary: { holdTotal: number; calculatedTotal: number };
    holdPaymentDetails: { holdAddTotal: number };
  } } };
  new Function("require", "module", "exports", result.outputFiles[0].text)(createRequire(import.meta.url), module, module.exports);
  const row = (fields: Record<string, unknown>) => ({
    "Tháng báo cáo": "09.2026", "Tháng phát sinh": "09.2026", BU: "AHN", ...fields,
  });
  const totals = module.exports.calculate([
    row({ "Nghiệp vụ": "HOLD", "Sheet Source": "HOLD", "TOTAL PAYMENT": -100 }),
    row({ "Nghiệp vụ": "ADD", "Sheet Source": "ADD", "TOTAL PAYMENT": 20 }),
    row({ "Nghiệp vụ": "ADD", "Sheet Source": "ADD", "Tháng phát sinh": "08.2026", "TOTAL PAYMENT": 40 }),
    row({ "Nghiệp vụ": "CANCEL", "TOTAL PAYMENT": -50 }),
    row({ "Trạng thái": "CANCEL", "Tháng phát sinh": "", "TOTAL PAYMENT": -60 }),
    row({ "Tình trạng thanh toán": "CANCEL", "TOTAL PAYMENT": -70 }),
    row({ "Nghiệp vụ": "BONUS", "TOTAL PAYMENT": 500 }),
  ]);
  assert.equal(totals.sameMonthHoldTotal, 100);
  assert.equal(totals.sameMonthAddTotal, 20);
  assert.equal(totals.diffMonthAddTotal, 40);
  assert.equal(totals.holdAddItems.filter(item => item.type === "ADD").reduce((sum, item) => sum + item.amount, 0), 60);
  assert.equal(totals.holdAddItems.filter(item => item.type === "CANCEL").reduce((sum, item) => sum + item.amount, 0), -180);
  assert.equal(totals.finalTotals.AHN, -40);
  assert.equal(totals.calculationSummary.holdTotal, -40);
  assert.equal(totals.holdPaymentDetails.holdAddTotal, -40);

  const edited = module.exports.calculate([
    row({ "Nghiệp vụ": "ADD", "Sheet Source": "HOLD", "Trạng thái": "CANCEL", "TOTAL PAYMENT": -25 }),
    row({ "Nghiệp vụ": "A", "Sheet Source": "CANCEL", "Tháng phát sinh": "08.2026", "TOTAL PAYMENT": -35 }),
    row({ "Nghiệp vụ": "C", "Sheet Source": "ADD", "Tháng phát sinh": "08.2026", "TOTAL PAYMENT": 900 }),
  ]);
  assert.deepEqual(edited.holdAddItems.map(item => [item.type, item.amount]), [["ADD", 25], ["ADD", 35], ["CANCEL", -900]]);
  assert.equal(edited.sameMonthAddTotal, 25);
  assert.equal(edited.diffMonthAddTotal, 35);
  assert.equal(edited.finalTotals.AHN, 60);
  assert.equal(edited.calculationSummary.calculatedTotal, 60);
  assert.equal(edited.holdPaymentDetails.holdAddTotal, 60);
});
