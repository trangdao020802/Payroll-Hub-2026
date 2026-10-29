import assert from 'node:assert/strict';
import test from 'node:test';
import {createElement} from 'react';
import {renderToStaticMarkup} from 'react-dom/server';
import {ReconciliationTotalsSummary} from '../src/app/pages/03-master/components/ReconciliationTotalsSummary';

test('the visible reconciliation total distinguishes missing Batch Payment from an actual zero or real negative variance', () => {
  const totals = {actual: 0, expected: 700_595_349, variance: -700_595_349,
    grossPayTotal: 700_595_349, deductionsTotal: 0, hasActual: false, actualSource: 'missing' as const};
  const render = (value = totals) => renderToStaticMarkup(createElement(ReconciliationTotalsSummary, {month: '02.2026', totals: value}));
  const missing = render();
  assert.match(missing, /Chưa có Batch/);
  assert.doesNotMatch(missing, /text-emerald-600|text-rose-600/);
  const zero = render({...totals, actual: 700_595_349, variance: 0, hasActual: true});
  assert.match(zero, /text-emerald-600/);
  assert.doesNotMatch(zero, /Chưa có Batch/);
  assert.match(render({...totals, hasActual: true}), /text-rose-600/);
});
