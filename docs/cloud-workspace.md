# Cross-device payroll workspace

The same Payroll account owns one cloud workspace. A confirmed Save must finish the cloud commit before it reports success. Another computer restores that account's saved payroll tables and retained months; empty defaults must never replace them. Concurrent changes to different tables merge, while changes to the same table require a choice.

## Opening and saving

- Reopen a computer with its own successfully loaded local data: the workspace stays usable while cloud changes are checked in the background. A fresh computer waits for the complete restore, with progress and bounded failures.
- Unchanged local tables are reused only when their persisted hashes match the current cloud manifest. A checkpoint alone cannot certify local values because IndexedDB writes can lag behind a cloud pull.
- A changed table with identical content creates no new cloud revision. Encoding stays in a worker and avoids allocating a Promise for every ordinary payroll cell.
- Save waits for existing uploads and newer edits. A failed commit is an error, not a successful save. Returning to the active tab checks for changes from the other computer; periodic checks run only on visible tabs.
- A table edited during a download remains local and retains its previous baseline, so a conflicting cloud edit is not silently overwritten.

Do not bypass membership checks or force a fresh computer to upload its defaults to resolve a loading error. Sign in with the same account, check the sync status and last cloud save time, and use **Đồng bộ ngay** after an error. Use **Ctrl+Shift+R** to load a newly deployed build; the app handles Ctrl+R as a table refresh.

## Persistence boundary

The cache stamp is invalidated before local writes and restored only after they finish. Readers and writers use the same browser storage lock. Skipped split values retain a hash only if the previous on-disk stamp proves the same value and owner; partial saves or another tab's writes cause a safe cache miss. Existing storage keys, asset format, owner isolation, the 200 MB cloud limit and database policies remain compatible. No schema migration is needed. Older builds can read the payroll data; they simply ignore the cache stamp.

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
