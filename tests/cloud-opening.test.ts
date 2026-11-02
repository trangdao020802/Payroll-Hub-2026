import assert from 'node:assert/strict';
import { createRequire } from 'node:module';
import { Worker as NodeWorker } from 'node:worker_threads';
import test from 'node:test';
import { build } from 'esbuild';
import { encode, decode, digest } from '../src/app/lib/cloud/codec';
import { createCloudValueReceiver, streamCloudValue } from '../src/app/lib/cloud/value-chunks';
import { withAppStorageLock } from '../src/app/lib/cloud/field-cache';
import type { CloudRepository } from '../src/app/lib/cloud/repository';

test('nested cloud tables arrive in bounded ordered pieces with dates, files and imported property names intact', async () => {
  const source = {
    ...JSON.parse('{"__proto__":{"polluted":"data only"}}'),
    TableOriginals: { Timesheet_Roster: Array.from({ length: 2_505 }, (_, i) => ({ id: i, name: 'Dữ liệu gốc', amount: i * 1_000 })) },
    BankExport: { headers: ['ID', 'Tháng'], data: [{ ID: '00123', Tháng: '02.2026' }] },
    file: new File(['Nội dung nguồn'], 'nguon.csv', { type: 'text/csv' }),
    date: new Date('2026-11-02T00:00:00Z'), empty: undefined,
  };
  const expected = await decode(await encode(source)) as typeof source;
  const receiver = createCloudValueReceiver();
  let count = 0;
  for (const part of streamCloudValue(await decode(await encode(source)))) {
    if (part.kind === 'items') assert.ok(part.values.length <= 500);
    receiver.accept(part); count++;
  }
  const result = receiver.finish() as typeof source;
  assert.ok(count > 5);
  assert.deepEqual(result.TableOriginals, expected.TableOriginals);
  assert.deepEqual(result.BankExport, expected.BankExport);
  assert.equal(result.file.name, 'nguon.csv');
  assert.equal(await result.file.text(), 'Nội dung nguồn');
  assert.equal(result.date.toISOString(), expected.date.toISOString());
  assert.equal(Object.getPrototypeOf(result), Object.prototype);
  assert.equal(Object.hasOwn(result, '__proto__'), true);
  assert.equal((Object.prototype as Record<string, unknown>).polluted, undefined);
  const incomplete = createCloudValueReceiver();
  incomplete.accept({ kind: 'array', path: [], length: 2 });
  incomplete.accept({ kind: 'items', path: [], offset: 0, values: [1] });
  assert.throws(() => incomplete.finish(), /chưa đầy đủ/);
  assert.throws(() => incomplete.accept({ kind: 'items', path: [], offset: 0, values: [2] }), /thứ tự/);
});

test('cloud parts download concurrently in order and decode through a real worker while input turns keep running', async () => {
  const workerBundle = await build({ entryPoints: ['src/app/lib/cloud/codec.worker.ts'], bundle: true,
    write: false, platform: 'node', format: 'cjs', packages: 'external',
    plugins: [{ name: 'worker-indexeddb-boundary', setup(builder) {
      builder.onResolve({ filter: /^localforage$/ }, () => ({ path: 'storage', namespace: 'fixture' }));
      builder.onLoad({ filter: /.*/, namespace: 'fixture' }, () => ({ contents: `
        export default {config(){},getItem:async key=>key==='fixture'?globalThis.cacheFixture:null};
      ` }));
    } }],
  });
  const workers: Bridge[] = [];
  class Bridge {
    onmessage: ((event: { data: Record<string, unknown> }) => void) | null = null;
    onerror: (() => void) | null = null;
    onmessageerror: (() => void) | null = null;
    chunks = 0;
    terminated = false;
    thread = new NodeWorker(`const {parentPort}=require('node:worker_threads');
      global.cacheFixture=${JSON.stringify(source)};
      global.self={postMessage:message=>parentPort.postMessage(message)};
      ${workerBundle.outputFiles[0].text}
      parentPort.on('message',data=>self.onmessage({data}));`, { eval: true, execArgv: [] });
    constructor() {
      workers.push(this);
      this.thread.on('message', data => {
        if (data.type === 'codec-chunk') this.chunks++;
        this.onmessage?.({ data });
      });
      this.thread.on('error', () => this.onerror?.());
    }
    postMessage(message: unknown) { this.thread.postMessage(message); }
    terminate() { this.terminated = true; void this.thread.terminate(); }
  }
  const source = { TableOriginals: { roster: Array.from({ length: 5_005 }, (_, i) => ({ id: i, amount: i * 20_000, name: 'TEST NHÂN VIÊN' })) },
    globalMonth: '02.2026', other: { rows: [{ id: '00123', amount: 700_595_349 }] } };
  const blob = await encode(source);
  const hash = await digest(blob);
  // Small synthetic parts exercise ordering and concurrency without a large fixture.
  const partSize = Math.ceil(blob.size / 6);
  const pieces = new Map<string, Blob>();
  for (let offset = 0; offset < blob.size; offset += partSize) {
    pieces.set(`owner/${hash}/${pieces.size}.part`, blob.slice(offset, offset + partSize));
  }
  let active = 0, maxActive = 0;
  let corrupt = false;
  const testSupabase = { storage: { from: () => ({ download: async (path: string) => {
    active++; maxActive = Math.max(maxActive, active);
    await new Promise(resolve => setTimeout(resolve, path.endsWith('/0.part') ? 8 : 1));
    active--;
    const data = pieces.get(path)!;
    if (corrupt && path.endsWith('/0.part')) {
      const bytes = new Uint8Array(await data.arrayBuffer()); bytes[0] ^= 1;
      return { data: new Blob([bytes]), error: null };
    }
    return { data, error: null };
  } }) } };
  const bundle = await build({ entryPoints: ['src/app/lib/cloud/repository.ts'], bundle: true,
    write: false, platform: 'node', format: 'cjs', packages: 'external',
    define: { 'import.meta.url': JSON.stringify(new URL('../src/app/lib/cloud/repository.ts', import.meta.url).href) },
    plugins: [{ name: 'cloud-service-boundary', setup(builder) {
      builder.onResolve({ filter: /supabaseClient$/ }, () => ({ path: 'client', namespace: 'fixture' }));
      builder.onLoad({ filter: /.*/, namespace: 'fixture' }, () => ({ contents: 'export const supabase = testSupabase;' }));
    } }],
  });
  const module = { exports: {} as { CloudRepository: new (owner: string) => CloudRepository } };
  new Function('require', 'module', 'exports', 'Worker', 'testSupabase', bundle.outputFiles[0].text)(createRequire(import.meta.url), module, module.exports, Bridge, testSupabase);
  const repository = new module.exports.CloudRepository('owner');
  const asset = { hash, bytes: blob.size, parts: [...pieces.keys()] };
  let turns = 0;
  const heartbeat = setInterval(() => turns++, 0);
  try {
    assert.deepEqual(await repository.download(asset), source);
    assert.ok(maxActive > 1 && maxActive <= 4);
    assert.ok(turns > 5, 'input/event-loop tasks must run during a large cloud restore');
    assert.ok(workers[0].chunks > 10);
    assert.equal(workers[0].terminated, true, 'the decoded worker heap must be released');
    assert.deepEqual(await repository.readCachedField('fixture'), source, 'large IndexedDB reads use the same yielding worker transport');
    corrupt = true;
    await assert.rejects(repository.download(asset), /chưa đầy đủ/);
    corrupt = false;
    assert.deepEqual(await repository.download(asset), source, 'a failed read must not poison a retry');
  } finally {
    clearInterval(heartbeat); repository.dispose();
    await Promise.all(workers.map(worker => worker.thread.terminate()));
  }
});

test('opening storage has a bounded lock wait and permits retry without running a timed-out reader', async t => {
  let signal: AbortSignal | undefined;
  let called = false;
  Object.defineProperty(globalThis, 'navigator', { configurable: true, value: { locks: { request: (_name: string, options: { signal: AbortSignal }) => {
    signal = options.signal;
    return new Promise((_resolve, reject) => options.signal.addEventListener('abort', () => reject(new Error('aborted'))));
  } } } });
  try {
    await assert.rejects(withAppStorageLock(async () => { called = true; }, 5), /tab khác/);
    assert.equal(signal?.aborted, true);
    assert.equal(called, false);
    t.mock.method(navigator.locks, 'request', async (_name, _options, task) => task());
    await withAppStorageLock(async () => { called = true; }, 5);
    assert.equal(called, true);
  } finally { Reflect.deleteProperty(globalThis, 'navigator'); }
});
