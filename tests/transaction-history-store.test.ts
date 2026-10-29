import assert from 'node:assert/strict';
import test from 'node:test';
import type { SupabaseClient } from '@supabase/supabase-js';
import { loadLatestTransactionMonths, loadLatestVersion, replaceVersionIfCurrent, saveVersion } from '../src/app/lib/transaction-history-store';

function clientFixture({member = true, signedIn = true, historyError = false} = {}) {
  const calls: unknown[][] = [];
  const client = {
    auth: { getUser: async () => ({data: {user: signedIn ? {id: 'member-id'} : null}, error: null}) },
    from: (table: string) => {
      calls.push(['from', table]);
      const query = {
        select: (columns: string) => { calls.push(['select', columns]); return query; },
        eq: (column: string, value: string) => {calls.push(['eq', column, value]); return query;},
        order: (column: string, options: unknown) => {calls.push(['order', column, options]); return query;},
        limit: (count: number) => {calls.push(['limit', count]); return query;},
        maybeSingle: async () => ({data: table === 'transaction_history_members' && member ? {user_id: 'member-id'} : null,
          error: table !== 'transaction_history_members' && historyError ? {message: 'History network error'} : null}),
      };
      return query;
    },
    rpc: async (fn: string, args: unknown) => {calls.push(['rpc', fn, args]); return {data: 42, error: null};},
  } as unknown as SupabaseClient;
  return {client, calls};
}

test('history reads exactly the requested month and latest version, never an earlier month', async () => {
  const {client, calls} = clientFixture();
  assert.equal(await loadLatestVersion(client, '12.2025'), null);
  assert.ok(calls.some(call => JSON.stringify(call) === JSON.stringify(['eq', 'period', '2025-12-01'])));
  assert.ok(calls.some(call => JSON.stringify(call) === JSON.stringify(['order', 'id', {ascending: false}])));
  assert.ok(calls.some(call => JSON.stringify(call) === JSON.stringify(['limit', 1])));
});

test('missing membership, login and transport failures cannot masquerade as missing history', async () => {
  await assert.rejects(loadLatestVersion(clientFixture({member: false}).client, '01.2026'), /chưa được cấp quyền/);
  await assert.rejects(loadLatestVersion(clientFixture({signedIn: false}).client, '01.2026'), /đăng nhập/);
  await assert.rejects(loadLatestVersion(clientFixture({historyError: true}).client, '01.2026'), /network/);
});

test('saving calls the atomic RPC with unchanged request id and whole snapshot', async () => {
  const {client, calls} = clientFixture();
  const rows = [{'Document ID': '001', 'Beneficiary Account No.': '00123'}];
  assert.equal(await saveVersion(client, '01.2026', rows, 'request-uuid'), '42');
  assert.deepEqual(calls.at(-1), ['rpc', 'append_transaction_version', {
    p_period: '2026-01-01', p_rows: rows, p_request_id: 'request-uuid',
  }]);
});

function replacementClient(latestId = '7') {
  const calls: unknown[][] = [];
  const client = {
    auth: {getUser: async () => ({data: {user: {id: 'member-id'}}, error: null})},
    from(table: string) {
      const query = {
        select: () => query,
        eq: () => query,
        order: () => query,
        limit: () => query,
        maybeSingle: async () => ({
          data: table === 'transaction_history_members'
            ? {user_id: 'member-id'}
            : {id: latestId, period: '2026-01-01', created_at: '2026-09-08T00:00:00Z', rows: [{old: true}]},
          error: null,
        }),
      };
      return query;
    },
    rpc: async (fn: string, args: unknown) => {
      calls.push([fn, args]);
      return {data: 8, error: null};
    },
  } as unknown as SupabaseClient;
  return {client, calls};
}

test('resolution replaces a historical month only while its checked version is still latest', async () => {
  const expected = {
    id: '7', period: '2026-01-01', created_at: '2026-09-07T00:00:00Z', rows: [{old: true}],
  };
  const fresh = replacementClient();
  assert.equal(await replaceVersionIfCurrent(
    fresh.client,
    expected,
    [{'Document ID': '001'}],
    'request-uuid',
  ), '8');
  assert.equal(fresh.calls.length, 1);

  const stale = replacementClient('9');
  await assert.rejects(
    replaceVersionIfCurrent(stale.client, expected, [{'Document ID': '001'}], 'request-uuid'),
    /đã thay đổi/,
  );
  assert.equal(stale.calls.length, 0);
});

import { loadAllPriorVersions } from '../src/app/lib/transaction-history-store';
function historyClient(failId?: string, {allMonths = false, changeOnRecheck = false} = {}) {
  const records = [
    {id: '1', period: '2026-01-01', rows: [{old: true}]},
    {id: '2', period: '2026-01-01', rows: [{old: false}]},
    {id: '3', period: '2026-07-01', rows: []},
    {id: '4', period: '2026-08-01', rows: []},
    {id: '5', period: '2026-09-01', rows: []},
    {id: '6', period: '2025-12-01', rows: []},
    {id: '7', period: '2026-01-01', rows: [{latest: true}]},
  ].map(item => ({...item, created_at: '2026-09-07T00:00:00Z'}));
  const reads: string[] = [];
  let pages = 0;
  let rounds = 0;
  const client = {auth: {getUser: async () => ({data:{user:{id:'member'}},error:null})},
    from(table: string) {
      let before = ''; let cursor = Infinity; let id = '';
      const q = {
        select: () => q, eq: (_key: string, value: string) => {id=value; return q;},
        lt: (key: string, value: string) => {if(key==='period') before=value; else cursor=Number(value); return q;},
        order: () => q, limit: () => q,
        maybeSingle: async () => ({data:{user_id:'member'},error:null}),
        single: async () => {reads.push(id); return {data:records.find(r=>r.id===id),error: id===failId ? {message:'snapshot read failed'} : null};},
        then(resolve: (value: unknown) => unknown, reject: (reason: unknown) => unknown) {
          pages++;
          assert.equal(table, 'transaction_monthly_versions');
          assert.equal(before, allMonths ? '' : '2026-08-01');
          if (cursor === Infinity && ++rounds === 2 && changeOnRecheck) {
            records.push({...records[4], id: '8'});
          }
          // Simulate server cap smaller than requested page size, forcing keyset continuation.
          return Promise.resolve({data:records.filter(r=>(!before || r.period<before) && Number(r.id)<cursor).sort((a,b)=>Number(b.id)-Number(a.id)).slice(0,2),error:null}).then(resolve,reject);
        },
      }; return q;
    },
  } as unknown as SupabaseClient;
  return {client,reads,get pages(){return pages;}};
}
test('all prior months load their newest version across every metadata page, excluding current and future months', async () => {
  const fixture=historyClient();
  const result=await loadAllPriorVersions(fixture.client,'08.2026');
  assert.deepEqual(result.map(v=>[v.period,v.id]), [['2025-12-01','6'],['2026-01-01','7'],['2026-07-01','3']]);
  assert.deepEqual([...fixture.reads].sort(),['3','6','7']);
  assert.equal(fixture.pages,4);
});
test('all-month check fails instead of returning partial history when any snapshot fails', async () => {
  await assert.rejects(loadAllPriorVersions(historyClient('3').client,'08.2026'),/snapshot read failed/);
  await assert.rejects(loadAllPriorVersions(clientFixture({member:false}).client,'08.2026'),/chưa được cấp quyền/);
});

test('workspace restore reads the latest version of every month and reuses only unchanged pinned versions', async () => {
  const fixture = historyClient(undefined, {allMonths: true});
  const versions = await loadLatestTransactionMonths(fixture.client);
  assert.deepEqual(versions.map(item => [item.period, item.id]), [
    ['2025-12-01', '6'], ['2026-01-01', '7'], ['2026-07-01', '3'], ['2026-08-01', '4'], ['2026-09-01', '5'],
  ]);
  assert.equal(fixture.pages, 10, 'all metadata pages are read before and after snapshot downloads');
  const known = Object.fromEntries(versions.map(item => [item.period.slice(0, 7), item]));
  const second = historyClient(undefined, {allMonths: true});
  assert.deepEqual(await loadLatestTransactionMonths(second.client, known), versions);
  assert.deepEqual(second.reads, [], 'unchanged verified snapshots do not download again');
  known['2026-01'] = {...known['2026-01'], id: '2'};
  const changed = historyClient(undefined, {allMonths: true});
  await loadLatestTransactionMonths(changed.client, known);
  assert.deepEqual(changed.reads, ['7'], 'an old version ID must download the latest rows');
});

test('workspace history restore rejects partial reads, missing membership and a version changed during download', async () => {
  await assert.rejects(loadLatestTransactionMonths(historyClient('3', {allMonths: true}).client), /snapshot read failed/);
  await assert.rejects(loadLatestTransactionMonths(clientFixture({member: false}).client), /chưa được cấp quyền/);
  await assert.rejects(loadLatestTransactionMonths(historyClient(undefined, {allMonths: true, changeOnRecheck: true}).client), /vừa được cập nhật/);
});
