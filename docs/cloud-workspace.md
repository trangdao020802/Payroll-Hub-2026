# Cross-device payroll workspace

The same Payroll account owns one cloud workspace. Each login, including reopening an owned cache or signing out and back into the same account, waits for the latest saved cloud snapshot before opening the tables. Unsynced local edits are kept in a pinned local backup and replaced by cloud on login; they are never automatically replayed into that workspace. A confirmed Save must finish the cloud commit before it reports success. In an already opened session, concurrent changes to different tables merge, while changes to the same table require a choice.

## Opening and saving

- Every authenticated opening waits for a complete restore, including computers with cached payroll. The manifest is read again after download; if another computer committed meanwhile, the newer snapshot is installed before reporting success. Repeated changes have a bounded retry and an explicit retry action.
- Unchanged local tables are reused only when their persisted hashes match the current cloud manifest. A checkpoint alone cannot certify local values because IndexedDB writes can lag behind a cloud pull.
- A changed table with identical content creates no new cloud revision. Encoding stays in a worker and avoids allocating a Promise for every ordinary payroll cell.
- Save waits for existing uploads and newer edits. A failed commit is an error, not a successful save. Returning to the active tab checks for changes from the other computer; periodic checks run only on visible tabs.
- After a session has opened, a table edited during a background download remains local and retains its previous baseline, so a conflicting cloud edit is not silently overwritten.
- A login without a working cloud connection stays blocked and does not certify cached values as current. An already opened session can keep pending edits offline and save on reconnect. Reopening before those edits were saved backs them up and restores cloud instead.
- Restoring replaces table values and month maps, without appending cached rows. Open/retry/relogin with unchanged cloud content creates no cloud revision. A cloud account with an established checkpoint cannot be silently recreated from a stale cache if its workspace disappears.

Do not bypass membership checks or force a fresh computer to upload its defaults to resolve a loading error. Sign in with the same account, check the sync status and last cloud save time, and use **Đồng bộ ngay** after an error. Use **Ctrl+Shift+R** to load a newly deployed build; the app handles Ctrl+R as a table refresh.

## Persistence boundary

The cache stamp is invalidated before local writes and restored only after they finish. Readers and writers use the same browser storage lock. Skipped split values retain a hash only if the previous on-disk stamp proves the same value and owner; partial saves or another tab's writes cause a safe cache miss. Existing storage keys, asset format, owner isolation, the 200 MB cloud limit and database policies remain compatible. No schema migration is needed. Older builds can read the payroll data; they simply ignore the cache stamp.

Legacy hydration can mutate payroll table containers and arrays in place. Those four tables are detached after registering durable hashes and before migration, so changed values cannot inherit an old certified hash. Large split collections retain their verified immutable references.

Implementation: [engine](../src/app/lib/cloud/engine.ts), [repository](../src/app/lib/cloud/repository.ts), [field cache](../src/app/lib/cloud/field-cache.ts), [AppData provider](../src/app/lib/contexts/AppDataContext.tsx), [cloud provider](../src/app/lib/contexts/CloudWorkspaceContext.tsx).

## Verification and change record

2026-10-13, Codex, requested by the owner after the 2026-10-12 report of an unusable continuously syncing website. Before the fix, unchanged content could repeatedly create cloud revisions, reopening downloaded all 54 fields, and React could acknowledge a restore before publishing it. The fix keeps the account's cloud data authoritative for clean fields while retaining pending edits.

Run `node --import tsx --test tests/cloud-sync.test.ts tests/cloud-sync-db.test.ts` for independent client restores, saved-number propagation, unchanged revisions, delayed React publication, offline recovery, edits during a pull/startup, stale/partial cache protection, request deadlines and database asset/RLS/CAS protection. These 21 tests pass. Typecheck and production build pass, as does ESLint on all changed source and tests.

The complete suite has the same 7 failures as baseline `dd67ee9`. Global lint has the same existing 2 errors and 3 warnings in both versions. The pre-existing failures concern Document ID/export source assertions, layout/font assertions, the syntax of `timesheet-link-refresh.test.ts`, and Trial Balance cancellation totals. They are outside this sync change. Direct verification of the owner's private account on two physical computers remains a user acceptance check; automated tests use independent clients and a test database, never write test payroll into production.

Review this page when the workspace manifest, cache/storage adapter, authentication owner model or Save contract changes.

### 2026-10-28: retained-month Deduction reconciliation

The owner reported that months reconciled on one computer showed a large negative variance after cloud restore on another. A retained generated/legacy Batch Payment month was ignored when its Bank AE import was no longer available. An empty active payment table then made the historical actual total zero. Unlabelled active payment rows could also leak into another reporting month.

Reconciliation now uses that month's retained payment as a fallback after confirmed payments and a matching Bank AE source. Generated/edited and legacy snapshots keep their original activity; this fallback does not approve them. Unlabelled current rows count only in their active month. The memoized calculator includes the month-scope flag. Automatic month initialization waits for a successful cloud sync, so an older local cache cannot generate a payment edit while a restore is in flight or has failed. The cloud indicator also tolerates an absent provider.

Verification: all 341 tests pass, including a fresh independent client restoring nine retained months (generated and legacy activity), a genuine nonzero variance, month navigation, source precedence, intentional empty confirmed months, and the database/RLS tests. Typecheck, the production build and ESLint on the changed source/tests pass. Read-only checks of the production workspace found all manifest parts present with matching stored sizes; no production payroll data or schema was changed. The owner's signed-in session on two physical computers remains an acceptance check; the browser smoke test uses a signed-out session.

### 2026-10-28: one current cloud snapshot on login

The owner requested that different computers always open the same latest saved copy. The login restore now takes cloud precedence over every stale/dirty cache, preserves tracked unsynced work in a pinned local backup, reuses only hash-verified cached fields, replaces complete tables and retained month maps, and rechecks the remote revision before unlocking. The UI waits on cached computers too, and a new sign-in to the same account cannot reuse the previous session's unlocked state. Old screen callbacks, Undo/Redo and automatic Hold bank repair cannot mutate data during this gate. A late session lookup cannot undo a newer auth event. The cloud dialog shows the saved revision for cross-device comparison.

The server format, membership checks, owner isolation, private bucket, 200 MB limit, normal in-session edits and compare-and-swap conflict protection remain compatible. No database migration or production payroll rewrite is needed. This change does not deduplicate legitimate rows already present in a source workbook; login and restore never concatenate them.

Verification: all 350 tests pass, including dirty duplicate cache replacement with backup, repeated logins across three independent device stores without new cloud revisions, a cloud save during login download, failed downloads/backups, offline reopening and reconnect, bounded revision churn, a missing established workspace, and hydration hash integrity. Typecheck, production build and ESLint on changed files pass. Automated cases use test clients and a local test database; verification of the owner's signed-in session on physical computers remains an acceptance check.

### 2026-10-28: Batch Payment filters and month-safe Reset

A Reconcile drilldown left a hidden search active in Batch Payment, leading the
owner to use a data reset to show every row. The old Reset restored a single
global BankExport original, which could belong to the previous month, without
updating the active month cache or reconciliation activity.

Batch Payment now exposes search, search clearing and complete filter clearing.
The complete clear also resets debounced column/search inputs and pagination.
Verified monthly identity corrections and month navigation clear stale queries.
Display-only BU enrichment clones rows instead of mutating saved payroll during
render, protecting immutable cache/hash ownership.

Batch Payment originals are stored by normalized year/month inside the existing
TableOriginals cloud field. Generated months and verified monthly cloud loads
establish their own baselines; subsequent edits retain them. Legacy originals
are usable only for explicitly matching rows. Reset cannot fall back to another
month or guess the month of unlabelled payroll. Successful Reset changes the
selected working month, its activity/cache and reconciliation total together,
preserves other month snapshots/source tables, and discards its stale cell draft.
The separate monthly history still requires **Lưu tháng**. No production data
repair, database migration or automatic historical payroll rewrite is included.

Verification: 360 tests pass with `TZ=Asia/Ho_Chi_Minh npm test`, including the
actual Batch Payment search/clear callbacks with a table boundary, immutable
display rows, legacy wrong-month rejection, month/year navigation, empty months,
verified ID/account baselines, and repeated Reset without appended rows.
Typecheck, production build and focused ESLint pass (three existing warnings).
The host default `America/Sao_Paulo` exposes one existing Timesheet coverage
date-boundary test failure; it reproduces unchanged at parent `1c47e66`.
Automated cases use synthetic payroll and the local test database, without
changing the owner's production payroll.
