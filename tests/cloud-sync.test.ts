import test from "node:test";
import assert from "node:assert/strict";
import { encode, decode, digest } from "../src/app/lib/cloud/codec";
import {
  conflictingFields,
  mergeManifest,
  type Manifest,
  type Workspace,
} from "../src/app/lib/cloud/protocol";
import { SyncEngine, OWNER_KEY } from "../src/app/lib/cloud/engine";
import type { CloudRepository } from "../src/app/lib/cloud/repository";
import type { AppData } from "../src/app/types";
import { rememberCloudField, cachedFieldHash, cachedFieldHashes, restoreFieldHashes, persistedFieldHashes, withAppStorageLock } from "../src/app/lib/cloud/field-cache";
import { withDeadline } from "../src/app/lib/cloud/deadline";
import { INITIAL_APP_DATA } from "../src/app/constants/initial-data";
import { calculateReconciliationTotals, syncReportingMonthReconciliation } from "../src/app/lib/utils/reconciliation-sync";
const asset = (hash: string) => ({ hash, bytes: 10, parts: [hash] });
const manifest = (fields: Record<string, string>): Manifest => ({
  version: 1,
  fields: Object.fromEntries(
    Object.entries(fields).map(([key, hash]) => [key, asset(hash)]),
  ),
});
test("gzip preserves Vietnamese, undefined, uploaded files and date values", async () => {
  const source = {
    month: "Tháng phát sinh",
    data: [{ amount: 1000 }],
    file: new File(["source data"], "payroll.xlsx", {
      type: "test/plain",
      lastModified: 123,
    }),
    date: new Date("2026-10-01"),
  };
  const compressed = await encode(source);
  const restored = (await decode(compressed)) as typeof source;
  assert.equal(restored.month, source.month);
  assert.deepEqual(restored.data, source.data);
  assert.equal(restored.file.name, "payroll.xlsx");
  assert.equal(await restored.file.text(), "source data");
  assert.equal(restored.date.toISOString(), source.date.toISOString());
  assert.equal(await decode(await encode(undefined)), undefined);
  assert.equal((await digest(compressed)).length, 64);
});
test("Deductions changes merge without overwriting another machine confirmed Batch Payment", () => {
  const base = manifest({ Hold_AE: "hold1", Bank_North_AE: "confirmed1" });
  const local = manifest({ Hold_AE: "hold2", Bank_North_AE: "confirmed1" });
  const remote = manifest({ Hold_AE: "hold1", Bank_North_AE: "confirmed2" });
  assert.deepEqual(conflictingFields(base, local, remote), []);
  assert.equal(
    mergeManifest(base, local, remote).fields.Bank_North_AE.hash,
    "confirmed2",
  );
  assert.equal(mergeManifest(base, local, remote).fields.Hold_AE.hash, "hold2");
});
test("same table edited on two machines conflicts; identical changes do not", () => {
  assert.deepEqual(
    conflictingFields(
      manifest({ Q_Staff: "a" }),
      manifest({ Q_Staff: "b" }),
      manifest({ Q_Staff: "c" }),
    ),
    ["Q_Staff"],
  );
  assert.deepEqual(
    conflictingFields(
      manifest({ Q_Staff: "a" }),
      manifest({ Q_Staff: "b" }),
      manifest({ Q_Staff: "b" }),
    ),
    [],
  );
});
class MemoryStorage {
  values = new Map<string, string>();
  getItem(key: string) {
    return this.values.get(key) || null;
  }
  setItem(key: string, value: string) {
    this.values.set(key, value);
  }
}
function harness() {
  Object.defineProperty(globalThis, "localStorage", {
    configurable: true,
    value: new MemoryStorage(),
  });
  Object.defineProperty(globalThis, "navigator", {
    configurable: true,
    value: { onLine: true },
  });
  let server: Workspace | null = null;
  const values = new Map<string, unknown>();
  let serial = 0;
  let commits = 0;
  let downloads = 0;
  let online = true;
  let blockedUpload: { gate: Promise<void>; started: () => void } | null = null;
  let blockedDownload: { gate: Promise<void>; started: () => void } | null = null;
  let rejectCommit = false;
  const repository = (owner = "owner") =>
    ({
      owner,
      dispose() {},
      async cleanup() {},
      async checkMembership() {},
      async read() {
        if (!online) throw new Error("offline test");
        return server ? structuredClone(server) : null;
      },
      async upload(value: unknown) {
        if (blockedUpload) {
          const blocked = blockedUpload;
          blockedUpload = null;
          blocked.started();
          await blocked.gate;
        }
        const hash = JSON.stringify(value) || "undefined";
        values.set(hash, structuredClone(value));
        return asset(hash);
      },
      async download(ref: { hash: string }) {
        downloads++;
        if (blockedDownload) {
          const blocked = blockedDownload;
          blockedDownload = null;
          blocked.started();
          await blocked.gate;
        }
        return structuredClone(values.get(ref.hash));
      },
      async commit(revision: number, next: Manifest) {
        if (rejectCommit) {
          rejectCommit = false;
          throw new Error("commit test failed");
        }
        if (revision !== (server?.revision || 0)) throw { code: "P0002" };
        commits++;
        server = {
          owner_id: owner,
          revision: ++serial,
          manifest: structuredClone(next),
          updated_at: new Date().toISOString(),
        };
        return structuredClone(server);
      },
    }) as unknown as CloudRepository;
  function client(initial: Record<string, unknown>, deferApply = false) {
    let data = initial;
    const updates: Array<{ updater: (data: AppData) => AppData; resolve: () => void }> = [];
    let notifyApply!: () => void;
    const applyStarted = new Promise<void>(resolve => { notifyApply = resolve; });
    const states: unknown[] = [];
    const engine = new SyncEngine(
      repository(),
      () => data as unknown as AppData,
      (updater) => {
        if (deferApply) {
          return new Promise<void>(resolve => {
            updates.push({ updater, resolve });
            notifyApply();
          });
        }
        data = updater(data as unknown as AppData) as unknown as Record<
          string,
          unknown
        >;
      },
      (state) => states.push(state),
      async () => {},
    );
    return {
      engine,
      states,
      applyStarted,
      get data() {
        return data;
      },
      publish() {
        for (const { updater, resolve } of updates.splice(0)) {
          data = updater(data as unknown as AppData) as unknown as Record<string, unknown>;
          resolve();
        }
        engine.changed();
      },
      edit(key: string, value: unknown, notify = true) {
        data = { ...data, [key]: value };
        if (notify) engine.changed();
      },
    };
  }
  return {
    client,
    failCommit() {
      rejectCommit = true;
    },
    pauseUpload() {
      let release!: () => void, started!: () => void;
      const startedPromise = new Promise<void>((resolve) => {
        started = resolve;
      });
      const gate = new Promise<void>((resolve) => {
        release = resolve;
      });
      blockedUpload = { gate, started };
      return { release, started: startedPromise };
    },
    pauseDownload() {
      let release!: () => void, started!: () => void;
      const startedPromise = new Promise<void>(resolve => { started = resolve; });
      const gate = new Promise<void>(resolve => { release = resolve; });
      blockedDownload = { gate, started };
      return { release, started: startedPromise };
    },
    setOnline(value: boolean) {
      online = value;
      Object.defineProperty(globalThis, "navigator", {
        configurable: true,
        value: { onLine: value },
      });
    },
    get commits() {
      return commits;
    },
    get downloads() {
      return downloads;
    },
    get server() {
      return server;
    },
    resetDevice() {
      Object.defineProperty(globalThis, "localStorage", {
        configurable: true,
        value: new MemoryStorage(),
      });
    },
  };
}
test("second device restores cloud instead of uploading its empty defaults, then receives changes", async () => {
  const h = harness();
  const first = h.client({
    Q_Staff: ["teacher"],
    Bank_North_AE: { data: ["confirmed"] },
  });
  await first.engine.start();
  assert.equal(h.commits, 1);
  h.resetDevice();
  const second = h.client({ Q_Staff: [], Bank_North_AE: { data: [] } });
  await second.engine.start();
  assert.equal(h.commits, 1);
  assert.deepEqual(second.data.Q_Staff, ["teacher"]);
  first.edit("Q_Staff", ["teacher", "new"]);
  await first.engine.sync();
  await second.engine.sync();
  assert.deepEqual(second.data.Q_Staff, ["teacher", "new"]);
  first.engine.stop();
  second.engine.stop();
});

test("a fresh device restores every retained payroll month and keeps real discrepancies visible", async () => {
  const h = harness();
  const periods = Array.from({ length: 9 }, (_, index) => `${String(index + 1).padStart(2, "0")}.2026`);
  const source: AppData = {
    ...structuredClone(INITIAL_APP_DATA),
    globalMonth: "10.2026",
    Sheet1_AE: { headers: [], data: periods.map(month => ({
      "Tháng báo cáo": month, "ID Number": `TEST-${month}`, "TOTAL PAYMENT": 700_595_349,
    })) },
    BankExport: { headers: [], data: [] },
    TransactionMonthCache: {
      activePeriod: "2026-10",
      months: Object.fromEntries(periods.map((month, index) => [`2026-${month.slice(0, 2)}`, {
        table: { headers: [], data: [{ "Tháng báo cáo": month, "Payment Amount": 700_595_349 + (index === 8 ? 7 : 0) }] },
        // Older cloud saves have no activity marker; later ones retain generated previews.
        activity: index % 2 ? { lastAction: "generated", saveVersion: 0, editCount: 0 } : undefined,
      }])),
    },
  };
  const first = h.client({ ...source });
  await first.engine.start();
  await first.engine.flush();
  first.engine.stop();
  h.resetDevice();
  const second = h.client({ ...structuredClone(INITIAL_APP_DATA) });
  try {
    await second.engine.start();
    assert.equal(h.commits, 1, "a fresh computer must not upload its defaults");
    const restored = await decode(await encode(second.data)) as AppData;
    for (const [index, month] of periods.entries()) {
      const expectedVariance = index === 8 ? 7 : 0;
      assert.equal(calculateReconciliationTotals(restored, month).variance, expectedVariance);
      const opened = syncReportingMonthReconciliation(restored, month);
      assert.equal(calculateReconciliationTotals(opened).variance, expectedVariance);
    }
    await second.engine.flush();
    assert.equal(h.commits, 1, "viewing restored totals does not rewrite cloud payroll");
  } finally {
    second.engine.stop();
  }
});
test("two devices preserve nonoverlapping tables and pause overlapping edits", async () => {
  const h = harness();
  const a = h.client({ Hold_AE: ["initial"], Bank_North_AE: ["confirmed"] });
  await a.engine.start();
  h.resetDevice();
  const b = h.client({ Hold_AE: [], Bank_North_AE: [] });
  await b.engine.start();
  a.edit("Hold_AE", ["deduction"]);
  b.edit("Bank_North_AE", ["confirmed next"]);
  await b.engine.sync();
  await a.engine.sync();
  assert.deepEqual(a.data.Bank_North_AE, ["confirmed next"]);
  await b.engine.sync();
  a.edit("Hold_AE", ["machine a"]);
  b.edit("Hold_AE", ["machine b"]);
  await a.engine.sync();
  await b.engine.sync();
  assert.equal((b.states.at(-1) as { status: string }).status, "conflict");
  await b.engine.resolve(false);
  assert.deepEqual(b.data.Hold_AE, ["machine a"]);
  a.engine.stop();
  b.engine.stop();
});
test("different account cannot upload another account cached data", async () => {
  const h = harness();
  localStorage.setItem(OWNER_KEY, "other");
  const a = h.client({ Q_Staff: ["private payroll"] });
  await a.engine.start();
  assert.equal(h.commits, 0);
  assert.equal((a.states.at(-1) as { status: string }).status, "error");
  a.engine.stop();
});

test("offline edits survive reopening and upload when connection returns", async () => {
  const h = harness();
  const first = h.client({ Q_Staff: ["saved"], Bank_North_AE: ["confirmed"] });
  await first.engine.start();
  h.setOnline(false);
  first.edit("Q_Staff", ["offline edit"]);
  await first.engine.sync();
  first.engine.stop();
  const reopened = h.client({
    Q_Staff: ["offline edit"],
    Bank_North_AE: ["confirmed"],
  });
  await reopened.engine.start();
  assert.equal(
    (reopened.states.at(-1) as { status: string }).status,
    "offline",
  );
  assert.equal(h.commits, 1);
  h.setOnline(true);
  await reopened.engine.retry();
  assert.equal(h.commits, 2);
  assert.equal(
    h.server?.manifest.fields.Q_Staff.hash,
    JSON.stringify(["offline edit"]),
  );
  reopened.engine.stop();
});

test("explicit Save flushes edits before the debounce/effect notifies the engine", async () => {
  const h = harness(),
    client = h.client({ Hold_AE: ["old"], BankExport: ["confirmed"] });
  await client.engine.start();
  client.edit("Hold_AE", ["saved deduction"], false);
  await client.engine.flush();
  assert.equal(
    h.server?.manifest.fields.Hold_AE.hash,
    JSON.stringify(["saved deduction"]),
  );
  client.engine.stop();
});
test("Save waits for an existing upload and then persists newer deductions", async () => {
  const h = harness(),
    client = h.client({ Hold_AE: ["old"], BankExport: ["confirmed"] });
  await client.engine.start();
  const blocked = h.pauseUpload();
  client.edit("Hold_AE", ["first"]);
  const inflight = client.engine.sync();
  await blocked.started;
  client.edit("Hold_AE", ["latest"]);
  let completed = false;
  const saving = client.engine.flush().then(() => {
    completed = true;
  });
  await Promise.resolve();
  assert.equal(completed, false);
  blocked.release();
  await Promise.all([inflight, saving]);
  assert.equal(
    h.server?.manifest.fields.Hold_AE.hash,
    JSON.stringify(["latest"]),
  );
  assert.equal(h.commits, 3);
  client.engine.stop();
});
test("a rejected cloud commit cannot be acknowledged as saved", async () => {
  const h = harness(),
    client = h.client({ Hold_AE: ["old"], BankExport: ["confirmed"] });
  await client.engine.start();
  h.failCommit();
  client.edit("Hold_AE", ["new"]);
  await assert.rejects(client.engine.flush(), /commit test failed/);
  assert.equal(h.commits, 1);
  assert.deepEqual(client.data.Hold_AE, ["new"]);
  await client.engine.flush();
  assert.equal(h.commits, 2);
  client.engine.stop();
});

test("recreating an unchanged table does not keep committing or report pending changes", async () => {
  const h = harness();
  const client = h.client({ Hold_AE: { data: [{ amount: 1200, status: "HOLD" }] } });
  try {
    await client.engine.start();
    for (let i = 0; i < 3; i++) {
      client.edit("Hold_AE", structuredClone(client.data.Hold_AE));
      await client.engine.flush();
    }
    assert.equal(h.commits, 1, "identical content must not create more cloud versions");
    assert.equal((client.states.at(-1) as { status: string }).status, "synced");
  } finally {
    client.engine.stop();
  }
});

test("a delayed React cloud update is not mistaken for a new local payroll edit", async () => {
  const h = harness();
  const first = h.client({ Hold_AE: { data: [{ amount: 1200, status: "HOLD" }] } });
  await first.engine.start();
  h.resetDevice();
  const second = h.client({ Hold_AE: { data: [] } }, true);
  try {
    const starting = second.engine.start();
    await second.applyStarted;
    second.publish();
    await starting;
    await second.engine.flush();
    assert.equal(h.commits, 1, "restoring cloud is a read, even when React defers the updater");
    assert.deepEqual(second.data.Hold_AE, { data: [{ amount: 1200, status: "HOLD" }] });
  } finally {
    first.engine.stop();
    second.engine.stop();
  }
});

test("a cached restart skips unchanged tables but still receives the latest saved numbers", async () => {
  const h = harness();
  const first = h.client({ Hold_AE: { data: [{ amount: 1200 }] }, BankExport: { data: [{ amount: 600 }] } });
  await first.engine.start();
  const persisted = structuredClone(first.data);
  const hashes = cachedFieldHashes("owner", first.data);
  first.engine.stop();
  restoreFieldHashes("owner", persisted, hashes);
  const reopened = h.client(persisted);
  try {
    await reopened.engine.start();
    assert.equal(h.downloads, 0, "durably saved unchanged tables are reused");
    const other = h.client(structuredClone(persisted));
    await other.engine.start();
    other.edit("Hold_AE", { data: [{ amount: 3400 }] });
    await other.engine.flush();
    other.engine.stop();
    await reopened.engine.retry();
    assert.deepEqual(reopened.data.Hold_AE, { data: [{ amount: 3400 }] });
    assert.deepEqual(reopened.data.BankExport, { data: [{ amount: 600 }] });
    assert.equal(h.commits, 2);
  } finally { reopened.engine.stop(); }
});

test("a checkpoint without matching durable hashes restores data instead of trusting stale local numbers", async () => {
  const h = harness();
  const first = h.client({ Hold_AE: { data: [{ amount: 1200 }] } });
  await first.engine.start();
  first.engine.stop();
  const reopened = h.client({ Hold_AE: { data: [{ amount: 9999 }] } });
  try {
    await reopened.engine.start();
    assert.deepEqual(reopened.data.Hold_AE, { data: [{ amount: 1200 }] });
    assert.equal(h.downloads, 1);
    assert.equal(h.commits, 1);
  } finally { reopened.engine.stop(); }
});

test("an edit during a background pull is retained and a same-table cloud conflict stays visible", async () => {
  const h = harness();
  const a = h.client({ Hold_AE: { data: [{ amount: 1200 }] } });
  await a.engine.start();
  h.resetDevice();
  const b = h.client({ Hold_AE: { data: [] } });
  await b.engine.start();
  a.edit("Hold_AE", { data: [{ amount: 3400 }] });
  await a.engine.flush();
  const blocked = h.pauseDownload();
  const pulling = b.engine.retry();
  await blocked.started;
  b.edit("Hold_AE", { data: [{ amount: 5600 }] });
  blocked.release();
  await pulling;
  await b.engine.sync();
  try {
    assert.deepEqual(b.data.Hold_AE, { data: [{ amount: 5600 }] });
    assert.equal((b.states.at(-1) as { status: string }).status, "conflict");
    assert.equal(h.commits, 2, "neither concurrent edit is silently overwritten");
    await b.engine.resolve(false);
    assert.deepEqual(b.data.Hold_AE, { data: [{ amount: 3400 }] });
  } finally { a.engine.stop(); b.engine.stop(); }
});

test("editing the owned cache during startup survives closing before the restore finishes", async () => {
  const h = harness();
  const first = h.client({ Hold_AE: { data: [{ amount: 1200 }] } });
  await first.engine.start();
  first.engine.stop();
  const opening = h.client(structuredClone(first.data));
  const blocked = h.pauseDownload();
  const starting = opening.engine.start();
  await blocked.started;
  opening.edit("Hold_AE", { data: [{ amount: 5600 }] });
  opening.engine.stop();
  blocked.release();
  await starting;
  const reopened = h.client(structuredClone(opening.data));
  try {
    await reopened.engine.start();
    await reopened.engine.flush();
    assert.deepEqual(reopened.data.Hold_AE, { data: [{ amount: 5600 }] });
    assert.equal(h.server?.manifest.fields.Hold_AE.hash, JSON.stringify({ data: [{ amount: 5600 }] }));
  } finally { reopened.engine.stop(); }
});

test("saved hashes certify only the original value and account", () => {
  const original = { data: [{ amount: 1200 }] };
  rememberCloudField("owner", "Hold_AE", original, "saved-hash");
  assert.equal(cachedFieldHash("owner", "Hold_AE", original), "saved-hash");
  assert.equal(cachedFieldHash("other", "Hold_AE", original), undefined);
  assert.equal(cachedFieldHash("owner", "Hold_AE", structuredClone(original)), undefined);
  rememberCloudField("owner", "globalMonth", "2026-10", "month-hash");
  assert.equal(cachedFieldHash("owner", "globalMonth", "2026-09"), undefined);
});

test("a partial save or another tab's split data cannot be certified by stale hashes", () => {
  const roster = [{ amount: 1200 }];
  rememberCloudField("owner", "Timesheet_Roster", roster, "roster-hash");
  const data = { Timesheet_Roster: roster };
  assert.deepEqual(persistedFieldHashes("owner", data, ["Timesheet_Roster"], new Set(), null), {});
  assert.deepEqual(persistedFieldHashes("owner", data, ["Timesheet_Roster"], new Set(), { owner: "owner", hashes: { Timesheet_Roster: "other-tab-hash" } }), {});
  assert.deepEqual(persistedFieldHashes("owner", data, ["Timesheet_Roster"], new Set(["Timesheet_Roster"]), null), { Timesheet_Roster: "roster-hash" });
  assert.deepEqual(persistedFieldHashes("owner", data, ["Timesheet_Roster"], new Set(), { owner: "owner", hashes: { Timesheet_Roster: "roster-hash" } }), { Timesheet_Roster: "roster-hash" });
});

test("storage locks serialize readers and writers across tabs", async () => {
  const h = harness();
  let tail = Promise.resolve();
  Object.defineProperty(globalThis, "navigator", { configurable: true, value: {
    onLine: true,
    locks: { request(_name: string, task: () => Promise<void>) {
      const result = tail.then(task);
      tail = result.catch(() => {});
      return result;
    } },
  } });
  const events: string[] = [];
  let release!: () => void;
  const gate = new Promise<void>(resolve => { release = resolve; });
  const write = withAppStorageLock(async () => { events.push("write-start"); await gate; events.push("write-end"); });
  const read = withAppStorageLock(async () => { events.push("read"); });
  await Promise.resolve();
  assert.deepEqual(events, ["write-start"]);
  release();
  await Promise.all([write, read]);
  assert.deepEqual(events, ["write-start", "write-end", "read"]);
  h.resetDevice();
});

test("a hanging cloud request times out and closing the session aborts a pending request", async () => {
  let signal!: AbortSignal;
  await assert.rejects(withDeadline(input => { signal = input; return new Promise(() => {}); }, "timed out", 5), /timed out/);
  assert.equal(signal.aborted, true);
  const session = new AbortController();
  const request = withDeadline(() => new Promise(() => {}), "timed out", 1000, session.signal);
  session.abort();
  await assert.rejects(request, /Phiên đồng bộ đã đóng/);
});
