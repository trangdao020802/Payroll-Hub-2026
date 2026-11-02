import assert from "node:assert/strict";
import test from "node:test";
import { createRequire } from "node:module";
import { build } from "esbuild";
import { calculateSnapshotStats } from "../src/app/lib/utils/snapshot-manager";
import type { AppData } from "../src/app/types";

test("snapshot size estimation does not serialize every row of a large Timesheet", () => {
  let visited = 0;
  const row = { ma_nv: "EMP1", duration: 2, toJSON() { visited++; return { ma_nv: "EMP1", duration: 2 }; } };
  const stats = calculateSnapshotStats({ Timesheet_Roster: Array(100_000).fill(row) } as AppData);
  assert.equal(stats.timesheetRows, 100_000);
  assert.ok(stats.approxSizeKb > 100);
  assert.ok(visited <= 64, `size estimate serialized ${visited} rows`);
});

test('a required pre-login backup rejects a storage failure instead of allowing cloud to replace unsynced data', async t => {
  t.mock.method(console, 'error', () => {});
  const result = await build({ entryPoints: ['src/app/lib/utils/snapshot-manager.ts'], bundle: true,
    write: false, platform: 'node', format: 'cjs', packages: 'external',
    plugins: [{ name: 'failed-browser-storage', setup(builder) {
      builder.onResolve({ filter: /^localforage$/ }, () => ({ path: 'storage', namespace: 'fixture' }));
      builder.onLoad({ filter: /.*/, namespace: 'fixture' }, () => ({ contents: `
        export default { createInstance() { return { setItem: async () => { throw new Error('storage full'); } }; } };
      ` }));
    } }],
  });
  const module = { exports: {} as { saveSnapshot: (data: AppData, options?: {required?: boolean}) => Promise<unknown> } };
  new Function('require', 'module', 'exports', 'window', result.outputFiles[0].text)(createRequire(import.meta.url), module, module.exports, {});
  await assert.rejects(module.exports.saveSnapshot({ Timesheet_Roster: [] } as unknown as AppData, {required: true}), /storage full/);
});

test("a successfully persisted snapshot is not retained as a second in-memory backup", async () => {
  const result = await build({
    stdin: { contents: `
      export { saveSnapshot, getSnapshotById } from './src/app/lib/utils/snapshot-manager';
      export { removePersistedSnapshot } from 'localforage';
    `, resolveDir: process.cwd() },
    bundle: true, write: false, platform: "node", format: "cjs", packages: "external",
    plugins: [{ name: "persisted-storage-boundary", setup(builder) {
      builder.onResolve({ filter: /^localforage$/ }, () => ({ path: "storage", namespace: "fixture" }));
      builder.onLoad({ filter: /.*/, namespace: "fixture" }, () => ({ contents: `
        const stores = new Map();
        export function removePersistedSnapshot(id) { stores.get('snapshots_data').delete(id); }
        export default { createInstance({ storeName }) {
          const rows = new Map(); stores.set(storeName, rows);
          return { getItem: async key => structuredClone(rows.get(key) ?? null),
            setItem: async (key, value) => { rows.set(key, structuredClone(value)); return value; },
            removeItem: async key => rows.delete(key) };
        } };
      ` }));
    } }],
  });
  const module = { exports: {} as {
    saveSnapshot: (data: AppData) => Promise<{ id: string }>;
    getSnapshotById: (id: string) => Promise<{ data: AppData } | null>;
    removePersistedSnapshot: (id: string) => void;
  } };
  new Function("require", "module", "exports", result.outputFiles[0].text)(createRequire(import.meta.url), module, module.exports);
  const api = module.exports;
  const data = { Timesheet_Roster: [{ ma_nv: "ORIGINAL" }] } as AppData;
  const saved = await api.saveSnapshot(data);
  data.Timesheet_Roster[0].ma_nv = "EDITED";
  assert.equal((await api.getSnapshotById(saved.id))?.data.Timesheet_Roster[0].ma_nv, "ORIGINAL");
  api.removePersistedSnapshot(saved.id);
  assert.equal(await api.getSnapshotById(saved.id), null, "deleted storage data must not reappear from an unbounded RAM mirror");
});
