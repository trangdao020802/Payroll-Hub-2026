import assert from "node:assert/strict";
import test from "node:test";
import { createRequire } from "node:module";
import { build } from "esbuild";
import { createLatestTaskQueue } from "../src/app/lib/utils/latest-task-queue";
import { isTimesheetImporting, runTimesheetImport, subscribeTimesheetImports, withTimesheetImportBatch } from "../src/app/lib/utils/timesheet-import-queue";
import { normalizeTimesheetAllocation } from "../src/app/lib/utils/timesheet-edits";
import { prepareTimesheetRosterReplacement, replaceTimesheetRosterRows } from "../src/app/lib/utils/timesheet-roster-utils";
import { applyTimesheetLinkResult, fetchTimesheetFolderFiles } from "../src/app/lib/utils/timesheet-link-sync";
import { fetchGoogleSheetAsFile } from "../src/app/lib/utils/data-utils";
import { INITIAL_APP_DATA } from "../src/app/constants/initial-data";
import { calculateTimesheet } from "../src/app/workers/timesheet.worker";

test("rapid link requests run one at a time, stay busy through the queue and recover after a failed link", async () => {
  const states: boolean[] = [];
  const stop = subscribeTimesheetImports(() => states.push(isTimesheetImporting()));
  let release!: () => void;
  const gate = new Promise<void>(resolve => { release = resolve; });
  const started: number[] = [];
  const first = runTimesheetImport(async () => { started.push(1); await gate; return 1; });
  const failed = runTimesheetImport(async () => { started.push(2); throw new Error("bad link"); });
  const failure = assert.rejects(failed, /bad link/);
  const last = runTimesheetImport(async () => { started.push(3); return 3; });
  await Promise.resolve();
  assert.deepEqual(started, [1]);
  assert.equal(isTimesheetImporting(), true);
  release();
  assert.equal(await first, 1);
  await failure;
  assert.equal(await last, 3);
  assert.deepEqual(started, [1, 2, 3]);
  assert.deepEqual(states, [true, false]);
  stop();
});

test("slow storage keeps only the active snapshot and latest complete state", async () => {
  const saved: number[] = [];
  const errors: unknown[] = [];
  const queue = createLatestTaskQueue(error => errors.push(error));
  let release!: () => void;
  const gate = new Promise<void>(resolve => { release = resolve; });
  const first = queue.enqueue(async () => { saved.push(0); await gate; throw new Error("storage retry"); });
  await Promise.resolve();
  let last = first;
  for (let i = 1; i <= 100; i++) last = queue.enqueue(async () => { saved.push(i); });
  release();
  await last;
  assert.deepEqual(saved, [0, 100]);
  assert.equal(errors.length, 1);
  await queue.enqueue(async () => { saved.push(101); });
  assert.deepEqual(saved, [0, 100, 101]);
});

test("a folder batch suspends calculations between links, does not deadlock, and releases on failure", async () => {
  const states: boolean[] = [];
  const stop = subscribeTimesheetImports(() => states.push(isTimesheetImporting()));
  try {
    await assert.rejects(withTimesheetImportBatch(async () => {
      await runTimesheetImport(async () => 1);
      assert.equal(isTimesheetImporting(), true, "no intermediate calculation between two links");
      await new Promise(resolve => setTimeout(resolve, 0));
      await runTimesheetImport(async () => 2);
      throw new Error("folder stopped");
    }), /folder stopped/);
    assert.deepEqual(states, [true, false]);
    assert.equal(await runTimesheetImport(async () => 3), 3);
  } finally { stop(); }
});

test("large roster replacement yields to input and preserves duplicate, legacy, concurrent and stale-source rules", async () => {
  const input = { id: "north", l07: "MKT LOCAL NORTH", url: "sheet", legacyRowIds: ["legacy"] };
  const row = (index: number, l07 = "HN0027.OPK") => ({
    _rowId: "other", l07, ma_nv: `EMP${index}`, ngay: "15/09/2026", type: "LPAR01",
    gio_vao: "09:00", gio_ra: "11:00", duration: 2,
  });
  const existing = Array.from({ length: 10_000 }, (_, index) => row(index));
  existing.push({ ...row(10_001), _rowId: "legacy" }, row(10_002, "MKT LOCAL NORTH"));
  const incoming = Array.from({ length: 1_200 }, (_, index) => ({ ...row(index, "MKT LOCAL NORTH"), _rowId: "north" }));
  incoming.push({ ...incoming[0] });
  const options = { sourceRowIds: ["north", "legacy"], targetL07: input.l07 };
  const expected = replaceTimesheetRosterRows(structuredClone(existing), structuredClone(incoming), options);
  let turns = 0;
  const heartbeat = setInterval(() => turns++, 0);
  let prepared;
  try { prepared = await prepareTimesheetRosterReplacement(existing, incoming, options); }
  finally { clearInterval(heartbeat); }
  assert.ok(turns > 10, "UI tasks must run during cold normalization, not just after the complete roster");
  assert.deepEqual(prepared.rows, expected);
  assert.equal(prepared.rows[0], existing[0], "unrelated immutable rows stay shared");
  const data = { ...INITIAL_APP_DATA, Timesheet_Roster: existing, Timesheet_InputList: [input] };
  const source = { url: "sheet", fileName: "North.MKT Roster.csv", date: "now" };
  const imported = applyTimesheetLinkResult(data, input, incoming, source, prepared);
  assert.equal(imported.Timesheet_Roster, prepared.rows, "commit uses the prepared replacement without recomputing it");
  const concurrent = row(20_000);
  const changed = { ...data, Timesheet_Roster: [...existing, concurrent] };
  assert.ok(applyTimesheetLinkResult(changed, input, incoming, source, prepared).Timesheet_Roster.includes(concurrent));
  const deleted = { ...data, Timesheet_InputList: [] };
  assert.equal(applyTimesheetLinkResult(deleted, input, incoming, source, prepared), deleted);
});

test("Google downloads and folder listings time out even if a network request ignores abort, then recover", async (t) => {
  const signals: AbortSignal[] = [];
  t.mock.method(globalThis, "fetch", (_url: string, init: RequestInit) => {
    signals.push(init.signal as AbortSignal);
    return new Promise<Response>(() => {});
  });
  await assert.rejects(fetchGoogleSheetAsFile("sheet", "roster", 5), /quá lâu/);
  await assert.rejects(fetchTimesheetFolderFiles("folder", 5), /quá lâu/);
  assert.ok(signals.every(signal => signal.aborted));
  t.mock.method(globalThis, "fetch", async () => new Response("CENTER,ID\nHN0027.OPK,00123", { headers: { "content-type": "text/csv" } }));
  assert.match(await (await fetchGoogleSheetAsFile("sheet", "roster", 100)).text(), /00123/);
});

test("already normalized raw rows are shared instead of copied on every sync", () => {
  const row = { _uuid: "source-1", chargeToCenterMkt: "HN0027.OPK" };
  assert.equal(normalizeTimesheetAllocation(row), row);
});

test("Excel parsing releases its worker after the last result and recovers after a worker failure", async () => {
  type Reply = { requestId: string; success?: boolean; result?: unknown; type?: string; kind?: string; rows?: unknown[]; chunkId?: number };
  class ParserWorker {
    onmessage: ((event: { data: Reply }) => void) | null = null;
    onerror: ((event: { message: string }) => void) | null = null;
    requests: { requestId: string }[] = [];
    terminated = false;
    constructor() { workers.push(this); }
    postMessage(message: { requestId: string }) { this.requests.push(message); }
    terminate() { this.terminated = true; }
  }
  const workers: ParserWorker[] = [];
  const result = await build({ entryPoints: ["src/app/lib/utils/excel-worker-client.ts"],
    bundle: true, write: false, platform: "node", format: "cjs", packages: "external",
    plugins: [{ name: "parser-worker-boundary", setup(builder) {
      builder.onResolve({ filter: /^(\.\/data-utils|.*excelParser\.worker\?worker)$/ }, args => ({ path: args.path, namespace: "fixture" }));
      builder.onLoad({ filter: /.*/, namespace: "fixture" }, args => ({ contents: args.path === "./data-utils" ? `
        let nextId = 0;
        export const generateUUID = () => String(++nextId);
        export const getExcelFileBuffer = async file => ({ buffer: new ArrayBuffer(8), name: file.name });
      ` : "export default ParserWorker;" }));
    } }],
  });
  const module = { exports: {} as { parseExcelInWorker: (file: { name: string }, options?: { timeoutMs: number }) => Promise<unknown> } };
  new Function("require", "module", "exports", "ParserWorker", result.outputFiles[0].text)(createRequire(import.meta.url), module, module.exports, ParserWorker);
  const first = module.exports.parseExcelInWorker({ name: "first.xlsx" });
  const second = module.exports.parseExcelInWorker({ name: "second.xlsx" });
  await Promise.resolve();
  const worker = workers[0];
  assert.equal(workers.length, 1);
  worker.onmessage!({ data: { requestId: worker.requests[0].requestId, success: true, result: { rows: [1] } } });
  assert.deepEqual(await first, { rows: [1] });
  assert.equal(worker.terminated, false, "a queued parse still needs this worker");
  worker.onmessage!({ data: { requestId: worker.requests[1].requestId, success: true, result: { rows: [2] } } });
  assert.deepEqual(await second, { rows: [2] });
  assert.equal(worker.terminated, true);
  assert.equal(worker.onmessage, null);
  const failed = module.exports.parseExcelInWorker({ name: "broken.xlsx" });
  const rejected = assert.rejects(failed, /parse failed/);
  await Promise.resolve();
  workers[1].onerror!({ message: "parse failed" });
  await rejected;
  assert.equal(workers[1].terminated, true);
  const recovered = module.exports.parseExcelInWorker({ name: "retry.xlsx" });
  await Promise.resolve();
  workers[2].onmessage!({ data: { requestId: workers[2].requests[0].requestId, success: true, result: { rows: [3] } } });
  assert.deepEqual(await recovered, { rows: [3] });
  assert.equal(workers[2].terminated, true);

  const chunked = module.exports.parseExcelInWorker({ name: "large.csv" });
  await Promise.resolve();
  const chunkWorker = workers[3];
  const requestId = chunkWorker.requests[0].requestId;
  chunkWorker.onmessage!({ data: { type: "excel-result-start", requestId, kind: "roster" } });
  chunkWorker.onmessage!({ data: { type: "excel-result-chunk", requestId, chunkId: 1, rows: ["first"] } });
  assert.equal(chunkWorker.requests.length, 1, "acknowledgement must wait for an event-loop turn");
  await new Promise(resolve => setTimeout(resolve, 0));
  assert.equal(chunkWorker.requests.length, 2);
  chunkWorker.onmessage!({ data: { type: "excel-result-chunk", requestId, chunkId: 2, rows: ["last"] } });
  chunkWorker.onmessage!({ data: { type: "excel-result-complete", requestId, success: true } });
  assert.deepEqual(await chunked, { kind: "roster", rows: ["first", "last"] });
  assert.equal(chunkWorker.terminated, true);

  const stalled = module.exports.parseExcelInWorker({ name: "stalled.csv" }, { timeoutMs: 5 });
  await assert.rejects(stalled, /quá lâu/);
  assert.equal(workers[4].terminated, true);
  const retried = module.exports.parseExcelInWorker({ name: "retry.csv" });
  await Promise.resolve();
  workers[5].onmessage!({ data: { requestId: workers[5].requests[0].requestId, success: true, result: { rows: [4] } } });
  assert.deepEqual(await retried, { rows: [4] });
});

test("Excel producer waits for each acknowledgement and delivers every row in bounded chunks", async () => {
  const result = await build({ entryPoints: ["src/app/workers/excelParser.worker.ts"], bundle: true,
    write: false, platform: "node", format: "cjs", packages: "external" });
  type Message = { type: string; requestId: string; rows?: Record<string, unknown>[]; chunkId?: number; success?: boolean };
  const messages: Message[] = [];
  const scope = { onmessage: (event: { data: unknown }) => { void event; }, postMessage: (message: Message) => messages.push(structuredClone(message)) };
  const module = { exports: {} };
  new Function("require", "module", "exports", "self", result.outputFiles[0].text)(createRequire(import.meta.url), module, module.exports, scope);
  const csv = "CENTER,ID NUMBER,FULL NAME,DATE,TYPE,FROM,TO,DURATION\n" + Array.from({ length: 2_505 }, (_, i) =>
    `HN0027.OPK,EMP${i},Test Person,15/09/2026,IN-CLASS,09:00,11:00,2`).join("\n");
  scope.onmessage({ data: { requestId: "parse", fileName: "roster.csv", fileBuffer: new TextEncoder().encode(csv).buffer,
    fileId: "source", mode: "roster", resultBackpressure: true } });
  assert.equal(messages.length, 2);
  assert.equal(messages.shift()!.type, "excel-result-start");
  const first = messages[0];
  scope.onmessage({ data: { type: "excel-result-ack", requestId: "old", chunkId: first.chunkId } });
  assert.equal(messages.length, 1);
  const received: Record<string, unknown>[] = [];
  let completed = false;
  while (messages.length) {
    const message = messages.shift()!;
    if (message.type === "excel-result-chunk") {
      assert.ok(message.rows!.length <= 1_000);
      received.push(...message.rows!);
      assert.equal(messages.length, 0);
      scope.onmessage({ data: { type: "excel-result-ack", requestId: "parse", chunkId: message.chunkId } });
    } else if (message.type === "excel-result-complete") completed = true;
  }
  assert.equal(completed, true);
  assert.equal(received.length, 2_505);
  assert.equal(new Set(received.map(row => row.ma_nv)).size, 2_505);
  assert.ok(received.every(row => row._rowId === "source"));
});

test("Settings calculates identical payroll totals without returning unused detailed roster rows", () => {
  const params = { rosterData: [{ _rowId: "mkt", _sourceFile: "North.MKT Roster.xlsx", center: "MKT LOCAL NORTH",
    Class: "Vu Tong Phan", "ID NUMBER": "00123", "FULL NAME": "TEST PERSON", DATE: "15/09/2026",
    TYPE: "LPAR01", FROM: "09:00", TO: "13:00" }], appData: { Timesheet_InputList: [{ id: "mkt", l07: "MKT LOCAL NORTH" }] },
    preferredYear: 2026, checkTAsMap: {}, classSizeMap: {}, TASK_COLUMNS: [] };
  const detailed = calculateTimesheet(params);
  const settings = calculateTimesheet({ ...params, includeProcessedRosterData: false });
  assert.equal(detailed.error, undefined);
  assert.equal(settings.error, undefined);
  assert.equal(detailed.processedRosterData.length, 1);
  assert.deepEqual(settings.processedRosterData, []);
  assert.deepEqual(settings.centerSummary, detailed.centerSummary);
  assert.equal(settings.employeeSummary[0].totalSalary, detailed.employeeSummary[0].totalSalary);
  assert.equal(settings.centerSummary[0].chargeMktLocal, 80_000);
});

test("the Timesheet worker waits for each chunk acknowledgement and returns every row and total", async () => {
  const result = await build({ entryPoints: ["src/app/workers/timesheet.worker.ts"], bundle: true,
    write: false, platform: "node", format: "cjs", packages: "external" });
  type Message = { type: string; requestId?: string; field?: string; rows?: Record<string, unknown>[]; chunkId?: number; error?: string };
  const messages: Message[] = [];
  const workerScope = { onmessage: (event: { data: unknown }) => { void event; }, postMessage: (message: Message) => messages.push(structuredClone(message)) };
  const module = { exports: {} };
  new Function("require", "module", "exports", "self", result.outputFiles[0].text)(createRequire(import.meta.url), module, module.exports, workerScope);
  const count = 4500;
  const rows = Array.from({ length: count }, (_, i) => ({
    _uuid: `source-${i}`, _rowId: "mkt", _sourceFile: "MKT LOCAL NORTH.xlsx", center: "MKT LOCAL NORTH",
    chargeToCenterMkt: "HN0027.OPK", "ID NUMBER": `EMP${i}`, "FULL NAME": `Person ${i}`,
    DATE: "15/09/2026", TYPE: "LPAR01", FROM: "09:00", TO: "11:00",
  }));
  workerScope.onmessage({ data: { type: "timesheet-input-start", requestId: "run-1", params: {
    appData: { Timesheet_InputList: [{ id: "mkt", l07: "MKT LOCAL NORTH" }] },
    preferredYear: 2026, fromDateStr: "2026-09-01", toDateStr: "2026-09-30",
    checkTAsMap: {}, classSizeMap: {}, TASK_COLUMNS: [], resultBackpressure: true,
  } } });
  workerScope.onmessage({ data: { type: "timesheet-input-chunk", requestId: "run-1", field: "rosterData", rows } });
  workerScope.onmessage({ data: { type: "timesheet-input-complete", requestId: "run-1" } });
  assert.equal(messages.length, 2, "only start and the first chunk may be queued");
  assert.equal(messages[0].error, undefined);
  workerScope.onmessage({ data: { type: "timesheet-result-ack", requestId: "stale-run", chunkId: 1 } });
  assert.equal(messages.length, 2);
  const collected: Record<string, Record<string, unknown>[]> = {};
  let receivedComplete = false;
  while (messages.length) {
    const message = messages.shift()!;
    if (message.type === "timesheet-result-chunk") {
      assert.ok(message.rows && message.rows.length <= 2000);
      (collected[message.field!] ||= []).push(...message.rows!);
      assert.equal(messages.length, 0, "producer must wait until the UI has consumed this batch");
      workerScope.onmessage({ data: { type: "timesheet-result-ack", requestId: "run-1", chunkId: message.chunkId } });
    }
    if (message.type === "timesheet-result-complete") receivedComplete = true;
  }
  assert.equal(receivedComplete, true);
  assert.equal(collected.processedRosterData.length, count);
  assert.equal(new Set(collected.processedRosterData.map(row => row._sourceKey)).size, count);
  assert.equal(collected.centerSummary[0].chargeMktLocal, count * 2 * 20000);
});
