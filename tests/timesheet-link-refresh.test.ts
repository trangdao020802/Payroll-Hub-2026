import assert from "node:assert/strict";
import test from "node:test";
import { createRequire } from "node:module";
import { readFileSync } from "node:fs";
import { build } from "esbuild";
import { INITIAL_APP_DATA } from "../src/app/constants/initial-data";
import type { AppData } from "../src/app/types";
import type { MissingTimesheetCenters } from "../src/app/hooks/useMissingTimesheetCenters";
import type { TimesheetCoverageDateRange } from "../src/app/lib/utils/timesheet-center-coverage";

test("refresh uses Settings sync for only missing links, handles file-only sources, and retains successful data after a failure", async (t) => {
  const result = await build({
    stdin: {
      contents: `
        import React from 'react';
        import { renderToString } from 'react-dom/server';
        import { useMissingTimesheetCenters } from './src/app/hooks/useMissingTimesheetCenters';
        import { useTimesheetLinkSync } from './src/app/hooks/useTimesheetLinkSync';
        import { getData } from 'test-app-data';
        export { setData, getData } from 'test-app-data';
        export function render(rows = getData().Timesheet_Roster, range) {
          let value;
          function Probe() {
            const coverage = useMissingTimesheetCenters(rows, 'l07', range);
            const settings = useTimesheetLinkSync();
            value = { ...coverage, ...settings }; return null;
          }
          renderToString(React.createElement(Probe)); return value;
        }
      `,
      resolveDir: process.cwd(),
    },
    bundle: true, write: false, platform: "node", format: "cjs", packages: "external",
    plugins: [{
      name: "storage-and-worker-boundaries",
      setup(builder) {
        builder.onResolve({ filter: /AppDataContext$|^test-app-data$/ }, () => ({ path: "app-data", namespace: "fixture" }));
        builder.onResolve({ filter: /excel-worker-client$/ }, () => ({ path: "worker", namespace: "fixture" }));
        builder.onLoad({ filter: /.*/, namespace: "fixture" }, ({ path }) => ({
          resolveDir: process.cwd(),
          contents: path === "app-data" ? `
            let appData;
            export function setData(value) { appData = value; }
            export function getData() { return appData; }
            export function useAppData() { return { appData, updateAppData(update) { appData = update(appData); } }; }
          ` : `
            import { parseExcelData, prepareExcelResult } from './src/app/workers/excelParser.worker';
            export async function parseExcelInWorker(file, options) {
              const rows = parseExcelData(await file.arrayBuffer(), file.name);
              return prepareExcelResult(rows, file.name, options.fileId, options.mode);
            }
          `,
        }));
      },
    }],
  });
  type Harness = {
    setData(data: AppData): void;
    getData(): AppData;
    render(rows?: Record<string, unknown>[], range?: TimesheetCoverageDateRange): MissingTimesheetCenters & { syncRow(id: string): Promise<number | null> };
  };
  const module = { exports: {} as Harness };
  new Function("require", "module", "exports", result.outputFiles[0].text)(createRequire(import.meta.url), module, module.exports);
  const harness = module.exports;
  const requests: string[] = [];
  let failing = false;
  let unmatched = false;
  t.mock.method(globalThis, "fetch", async (url: string, init?: RequestInit) => {
    if (String(url).startsWith("/api/drive-folder-files?")) return Response.json({ success: true, files: [
      { id: "old-file", name: "VP0001.PCT.xlsx", modifiedTime: "2026-09-01T00:00:00Z" },
      { id: "new-file", name: "VP0001.PCT.xlsx", modifiedTime: "2026-09-20T00:00:00Z" },
    ] });
    assert.equal(url, "/api/gs-export");
    const source = JSON.parse(String(init?.body)).url;
    requests.push(source);
    if (failing && source === "missing-link") return Response.json({ error: "Access denied" }, { status: 403 });
    const l07 = unmatched ? "VP0001.PCT" : source === "missing-link" ? "TH0001.TPU" : source === "present-link" ? "HN0027.OPK" : "VP0001.PCT";
    return new Response(`CENTER,ID NUMBER,FULL NAME,DATE,TYPE,FROM,TO,DURATION\n${l07},TEST001,Test Person,15/09/2026,IN-CLASS,09:00,11:00,2`, {
      headers: { "content-type": "text/csv", "x-spreadsheet-name": `${l07}.csv` },
    });
  });
  const initial = () => ({
    ...INITIAL_APP_DATA,
    Timesheet_InputList: [
      { id: "present", l07: "HN0027.OPK", url: "present-link" },
      { id: "missing", l07: "TH0001.TPU", url: "missing-link" },
      { id: "file", l07: "VP0001.PCT", fileName: "VP0001.PCT.xlsx", url: "" },
      { id: "no-source", l07: "BN0001.LTT", url: "" },
    ],
    Timesheet_Roster: [{ _rowId: "present", l07: "HN0027.OPK", ma_nv: "UNCHANGED", duration: 7 }],
  });
  harness.setData(initial());
  const before = harness.render();
  assert.deepEqual(before.centers, ["TH0001.TPU", "VP0001.PCT"]);
  assert.deepEqual(requests, [], "showing missing centers never starts a refresh");
  await Promise.all([before.refreshMissing(), before.refreshMissing()]);
  assert.deepEqual(requests, ["missing-link", "https://docs.google.com/spreadsheets/d/new-file/edit"]);
  const missingSyncedAt = String(harness.getData().Timesheet_InputList.find((row) => row.id === "missing")?.lastSyncedAt || "");
  const fileSyncedAt = String(harness.getData().Timesheet_InputList.find((row) => row.id === "file")?.lastSyncedAt || "");
  assert.match(missingSyncedAt, /^\d{2}:\d{2} \d{2}\/\d{2}\/\d{4}$/);
  assert.match(fileSyncedAt, /^\d{2}:\d{2} \d{2}\/\d{2}\/\d{4}$/);
  assert.ok([missingSyncedAt, fileSyncedAt].includes(harness.render().latestSyncAt));
  assert.equal(harness.getData().Timesheet_Roster.length, 3);
  assert.equal(harness.getData().Timesheet_Roster[0].ma_nv, "UNCHANGED");
  assert.deepEqual(harness.render().centers, []);
  await harness.render().refreshMissing();
  assert.equal(requests.length, 2, "no fetch is needed when all centers are present");

  // Settings Actions and missing-center refresh call the same import workflow.
  await harness.render().syncRow("missing");
  assert.equal(harness.getData().Timesheet_Roster.length, 3);
  const storedRows = harness.getData().Timesheet_Roster;
  const successfulSyncAt = harness.getData().Timesheet_InputList.find((row) => row.id === "missing")?.lastSyncedAt;
  failing = true;
  assert.equal(await harness.render().syncRow("missing"), null);
  assert.deepEqual(harness.getData().Timesheet_Roster, storedRows, "failed refresh retains that source's previously imported rows");
  assert.equal(
    harness.getData().Timesheet_InputList.find((row) => row.id === "missing")?.lastSyncedAt,
    successfulSyncAt,
    "a failed retry must not pretend that a new successful sync happened",
  );

  requests.length = 0;
  harness.setData(initial());
  await harness.render().refreshMissing();
  assert.deepEqual(harness.render().centers, ["TH0001.TPU"]);
  assert.equal(harness.getData().Timesheet_Roster.length, 2);
  assert.equal(harness.getData().Timesheet_Roster[0].duration, 7);
  assert.equal(harness.getData().Timesheet_InputList.find((row) => row.id === "missing").status, "error");
  assert.equal(harness.getData().Timesheet_InputList.find((row) => row.id === "file").status, "success");
  assert.deepEqual(requests, ["missing-link", "https://docs.google.com/spreadsheets/d/new-file/edit"]);

  // A blank September table does not imply that its April-only source is lost.
  requests.length = 0;
  failing = false;
  const historical = initial();
  historical.Timesheet_Roster = historical.Timesheet_Roster.map((row) => ({ ...row, ngay: "30/04/2026" }));
  harness.setData(historical);
  const september = { from: "2026-09-01", to: "2026-09-30" };
  const filtered = harness.render([], september);
  assert.deepEqual(filtered.centers, ["TH0001.TPU", "VP0001.PCT"]);
  await filtered.refreshMissing();
  assert.deepEqual(requests, ["missing-link", "https://docs.google.com/spreadsheets/d/new-file/edit"]);
  assert.equal(harness.getData().Timesheet_Roster.find((row) => row._rowId === "present")?.duration, 7);

  requests.length = 0;
  harness.setData({ ...historical, Timesheet_InputList: historical.Timesheet_InputList.slice(0, 1) });
  const onlyHistorical = harness.render([], september);
  assert.deepEqual(onlyHistorical.centers, []);
  await onlyHistorical.refreshMissing();
  assert.deepEqual(requests, [], "do not sync an April-only center when filtering September");

  // Persisted or unrelated processing flags do not mean this button was clicked.
  harness.setData({
    ...initial(),
    Timesheet_InputList: initial().Timesheet_InputList.map((row) => ({ ...row, status: "processing" })),
  });
  const idle = harness.render();
  assert.equal(idle.isRefreshing, false, "only this control's explicit run shows its spinner");
  harness.render([], september);
  assert.deepEqual(requests, [], "rendering and changing filters require no network requests");
  await idle.refreshMissing();
  assert.equal(requests.length, 2, "a user click can retry stale persisted processing rows");
  assert.deepEqual(harness.render([]).centers, ["HN0027.OPK"],
    "successfully synced missing centers are skipped even if the table still excludes their rows");

  // A successful import with no matching L07 is final for this button.
  requests.length = 0;
  unmatched = true;
  harness.setData({ ...INITIAL_APP_DATA, Timesheet_Roster: [], Timesheet_InputList: [initial().Timesheet_InputList[1]] });
  await harness.render().refreshMissing();
  assert.deepEqual(requests, ["missing-link"]);
  assert.equal(harness.getData().Timesheet_InputList[0].status, "success");
  assert.equal(harness.getData().Timesheet_Roster.length, 1);
  assert.equal(harness.getData().Timesheet_Roster[0].l07, "VP0001.PCT");
  assert.deepEqual(harness.render().centers, []);
  await harness.render().refreshMissing();
  assert.deepEqual(requests, ["missing-link"], "do not repeatedly fetch a successful source with no matching L07");
  await harness.render().syncRow("missing");
  assert.equal(requests.length, 2, "Settings Actions remains available for an explicit full resync");
});


test("Settings Timesheet displays the latest successful link-sync time in both Upload Date and table refresh menu", () => {
  const inputTable = readFileSync(new URL("../src/app/pages/01-timesheet/components/TimesheetInputTable.tsx", import.meta.url), "utf8");
  const coverageMenu = readFileSync(new URL("../src/app/pages/01-timesheet/components/TimesheetCenterCoverage.tsx", import.meta.url), "utf8");
  assert.match(inputTable, /row\.lastSyncedAt \|\| row\.date/);
  assert.match(inputTable, /Upload Date: lần đồng bộ\/nạp dữ liệu thành công gần nhất/);
  assert.match(coverageMenu, /Cập nhật gần nhất/);
  assert.match(coverageMenu, /coverage\.latestSyncAt/);
});
