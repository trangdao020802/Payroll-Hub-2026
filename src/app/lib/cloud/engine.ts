import { INITIAL_APP_DATA } from "../../constants/initial-data";
import type { AppData } from "../../types";
import {
  type Manifest,
  type Workspace,
  conflictingFields,
  mergeManifest,
} from "./protocol";
import { CloudRepository } from "./repository";
import { cachedFieldHash, rememberCloudField, sameManifest, CLOUD_CACHE_EVENT } from "./field-cache";
export type SyncStatus =
  | "starting"
  | "saving"
  | "synced"
  | "offline"
  | "conflict"
  | "error";
export interface SyncState {
  status: SyncStatus;
  message?: string;
  conflicts?: string[];
  updatedAt?: string;
  revision?: number;
}
export const OWNER_KEY = "payroll_cloud_cache_owner";
const EMPTY: Manifest = { version: 1, fields: {} };
const checkpointKey = (owner: string) => `payroll_cloud_checkpoint:${owner}`;
export class SyncEngine {
  private base: Workspace | null = null;
  private refs: Record<string, unknown> = {};
  private dirty = new Set<string>();
  private running = false;
  private syncTask: Promise<void> | null = null;
  private lastError: unknown = null;
  private stopped = false;
  private ready = false;
  private applying = false;
  private initializing = false;
  private startTask: Promise<void> | null = null;
  private timer?: ReturnType<typeof setTimeout>;
  private conflictRemote: Workspace | null = null;
  private staged: Manifest | null = null;
  // Bind checkpoints to this browser/device for the lifetime of the session.
  private readonly storage = localStorage;
  constructor(
    private repository: CloudRepository,
    private getData: () => AppData,
    private apply: (updater: (data: AppData) => AppData) => void | Promise<void>,
    private report: (state: SyncState) => void,
    private backup: () => Promise<void>,
    private refreshData?: (data: AppData) => Promise<AppData>,
  ) {}
  private async refresh() {
    if (!this.refreshData) return;
    for (let attempt = 0; attempt < 3; attempt++) {
      const before = this.getData();
      const refreshed = await this.refreshData(before);
      if (this.stopped) return;
      if (this.getData() !== before) continue;
      if (refreshed === before) return;
      // The refreshed fields form one monthly snapshot. A month switch or edit
      // during the read must not install only half of that snapshot.
      await this.apply(current => current === before ? refreshed : current);
      this.changed();
      if (this.getData() === refreshed) return;
    }
    throw new Error('Dữ liệu thay đổi trong lúc tải Batch Payment. Bấm Đồng bộ ngay để thử lại.');
  }
  private checkpoint() {
    this.storage.setItem(
      checkpointKey(this.repository.owner),
      JSON.stringify({ base: this.base, dirty: [...this.dirty] }),
    );
  }
  private async install(remote: Workspace, fields: string[], replace = false) {
    const patch: Record<string, unknown> = {};
    const before = { ...this.getData() } as unknown as Record<string, unknown>;
    const ordered = [...fields].sort((a, b) => remote.manifest.fields[a].bytes - remote.manifest.fields[b].bytes);
    let completed = 0;
    for (let offset = 0; offset < ordered.length;) {
      // Fetch small tables together, but decode large rosters one at a time.
      const count = remote.manifest.fields[ordered[offset]].bytes > 4 * 1024 * 1024 ? 1 : 3;
      const batch = ordered.slice(offset, offset + count).filter(key => count === 1 || remote.manifest.fields[key].bytes <= 4 * 1024 * 1024);
      await Promise.all(batch.map(async key => {
        patch[key] = await this.repository.download(remote.manifest.fields[key]);
      }));
      if (this.stopped) return;
      completed += batch.length;
      offset += batch.length;
      this.report({ status: this.ready ? "saving" : "starting", message: `Đang tải dữ liệu đã lưu: ${completed}/${fields.length} phần.` });
    }
    if (this.stopped) return;
    this.applying = true;
    // Login installs an exact cloud snapshot; background pulls preserve live edits.
    try {
      await this.apply((current) => {
        const currentFields = current as unknown as Record<string, unknown>;
        const next = {
          ...(replace ? { ...INITIAL_APP_DATA, Timesheet_Roster: [] } : current),
        } as Record<string, unknown>;
        if (replace) {
          for (const key of Object.keys(remote.manifest.fields)) {
            // fields omitted from the download have a verified on-disk hash.
            next[key] = Object.hasOwn(patch, key) ? patch[key] : before[key];
            rememberCloudField(this.repository.owner, key, next[key], remote.manifest.fields[key].hash);
          }
          this.dirty.clear();
          this.refs = { ...next };
          return next as unknown as AppData;
        }
        for (const key of fields) {
          if (!this.dirty.has(key) && currentFields[key] === before[key]) {
            next[key] = patch[key];
            this.refs[key] = patch[key];
            rememberCloudField(this.repository.owner, key, patch[key], remote.manifest.fields[key].hash);
          } else {
            // Keep edits made after the download started, even before changed()
            // has had a chance to run in React's passive effect.
            next[key] = currentFields[key];
            this.dirty.add(key);
            this.refs[key] = before[key];
          }
        }
        return next as unknown as AppData;
      });
    } finally {
      this.applying = false;
    }
    this.changed();
  }
  async retry() {
    if (this.ready) await this.sync();
    else await this.start();
  }
  start(): Promise<void> {
    if (this.startTask) return this.startTask;
    if (this.stopped || this.ready) return Promise.resolve();
    const task = this.initialize();
    this.startTask = task;
    return task.finally(() => {
      if (this.startTask === task) this.startTask = null;
    });
  }
  private async initialize() {
    this.initializing = true;
    this.lastError = null;
    this.report({ status: "starting" });
    try {
      const owner = this.storage.getItem(OWNER_KEY);
      if (owner && owner !== this.repository.owner)
        throw new Error(
          "Máy này có dữ liệu của tài khoản khác. Đăng nhập lại tài khoản đã lưu dữ liệu để tránh trộn dữ liệu.",
        );
      const stored =
        owner === this.repository.owner
          ? this.storage.getItem(checkpointKey(owner))
          : null;
      const checkpoint = stored ? JSON.parse(stored) : null;
      this.refs = { ...this.getData() };
      this.base = checkpoint?.base || null;
      this.dirty = new Set(checkpoint?.dirty || []);
      await this.repository.checkMembership();
      let remote = await this.repository.read();
      if (this.stopped) return;
      if (remote) {
        // Every login takes the saved cloud version. Keep unsynced local work
        // in a pinned backup; never replay an old device's dirty checkpoint.
        if (this.dirty.size || !checkpoint?.base) await this.backup();
        if (this.stopped) return;
        for (let attempt = 0; attempt < 5; attempt++) {
          const current = this.getData() as unknown as Record<string, unknown>;
          const missing = Object.keys(remote.manifest.fields).filter(
            key => cachedFieldHash(this.repository.owner, key, current[key]) !== remote!.manifest.fields[key].hash,
          );
          await this.install(remote, missing, true);
          if (this.stopped) return;
          // Another computer may have saved while large tables downloaded.
          // Do not unlock a stale or partially newer workspace as "synced".
          const latest = await this.repository.read();
          if (this.stopped) return;
          if (!latest) throw new Error("Bản cloud vừa thay đổi. Bấm Đồng bộ ngay để tải lại.");
          this.changed();
          if (!this.dirty.size && latest.revision === remote.revision && sameManifest(latest.manifest, remote.manifest)) {
            this.base = latest;
            await this.refresh();
            if (this.stopped) return;
            const checked = await this.repository.read();
            if (this.stopped) return;
            if (!checked) throw new Error('Bản cloud vừa thay đổi. Bấm Đồng bộ ngay để tải lại.');
            if (checked.revision !== latest.revision || !sameManifest(checked.manifest, latest.manifest)) {
              remote = checked;
              continue;
            }
            this.storage.setItem(OWNER_KEY, this.repository.owner);
            this.ready = true;
            this.checkpoint();
            if (this.dirty.size) await this.sync();
            else this.report({ status: "synced", updatedAt: latest.updated_at, revision: latest.revision });
            return;
          }
          remote = latest;
        }
        throw new Error("Dữ liệu cloud đang được cập nhật liên tục. Bấm Đồng bộ ngay để lấy bản mới nhất.");
      } else if (checkpoint?.base) {
        throw new Error(
          "Không tìm thấy bản cloud đã lưu. Bản trên máy được giữ nguyên; bấm Đồng bộ ngay để kiểm tra lại.",
        );
      } else {
        // A genuinely new workspace can be seeded only after a successful read.
        Object.keys(this.getData()).forEach((key) => this.dirty.add(key));
        await this.refresh();
      }
      if (this.stopped) return;
      this.storage.setItem(OWNER_KEY, this.repository.owner);
      this.ready = true;
      this.checkpoint();
      await this.sync();
    } catch (error) {
      this.fail(error);
    } finally {
      this.initializing = false;
    }
  }
  changed() {
    if ((!this.ready && !this.initializing) || this.stopped || this.applying) return;
    const data = this.getData() as unknown as Record<string, unknown>;
    for (const key of new Set([
      ...Object.keys(data),
      ...Object.keys(this.refs),
    ])) {
      if (data[key] !== this.refs[key]) this.dirty.add(key);
    }
    if (!this.ready) {
      if (this.base) this.checkpoint();
      return;
    }
    this.checkpoint();
    if (this.dirty.size && !this.conflictRemote) {
      this.report({
        status: navigator.onLine ? "saving" : "offline",
        message: "Thay đổi trên máy đang chờ đồng bộ.",
      });
      clearTimeout(this.timer);
      this.timer = setTimeout(() => void this.sync(), 3500);
    }
  }
  private fail(error: unknown) {
    this.lastError = error;
    if (this.stopped) return;
    this.report({
      status: navigator.onLine ? "error" : "offline",
      message:
        error instanceof Error
          ? error.message
          : String(
              (error as { message?: string })?.message ||
                "Không thể đồng bộ. Dữ liệu trên máy vẫn được giữ.",
            ),
    });
  }
  /** Explicit Save waits for the active upload and any newer local changes. */
  async flush() {
    this.changed();
    if (!this.ready) await this.retry();
    for (let attempt = 0; attempt < 3; attempt++) {
      await this.sync();
      if (this.stopped)
        throw new Error(
          "Phiên đồng bộ đã đóng. Dữ liệu trên máy vẫn được giữ.",
        );
      if (this.conflictRemote)
        throw new Error(
          "Có xung đột dữ liệu giữa các máy. Chọn bản cần giữ trước khi xác nhận lưu cloud.",
        );
      if (this.lastError)
        throw new Error(
          (this.lastError as { message?: string })?.message ||
            "Chưa lưu được dữ liệu lên cloud.",
        );
      if (this.ready && this.dirty.size === 0) return;
    }
    throw new Error(
      "Dữ liệu vẫn đang thay đổi. Chờ đồng bộ hoàn tất rồi lưu lại.",
    );
  }
  sync(): Promise<void> {
    if (this.syncTask) return this.syncTask;
    if (!this.ready || this.stopped || this.conflictRemote)
      return Promise.resolve();
    const task = this.runSync();
    this.syncTask = task;
    return task.finally(() => {
      if (this.syncTask === task) this.syncTask = null;
    });
  }
  private async runSync() {
    if (!this.ready || this.running || this.stopped || this.conflictRemote)
      return;
    this.running = true;
    this.lastError = null;
    try {
      let remote = await this.repository.read();
      if (this.stopped) return;
      const snapshot = { ...this.getData() } as unknown as Record<
        string,
        unknown
      >;
      const pending = [...this.dirty];
      const local: Manifest = {
        version: 1,
        fields: { ...(this.base?.manifest.fields || {}) },
      };
      if (pending.length) {
        this.report({ status: "saving" });
        await this.repository.cleanup().catch(() => undefined);
      }
      for (const key of pending) {
        local.fields[key] = await this.repository.upload(snapshot[key], this.base?.manifest.fields[key]);
        if (this.stopped) return;
      }
      // Refresh after uploads: another computer may have committed meanwhile.
      if (pending.length) remote = await this.repository.read();
      if (this.stopped) return;
      const conflicts = conflictingFields(
        this.base?.manifest || EMPTY,
        local,
        remote?.manifest || EMPTY,
      );
      if (conflicts.length && remote) {
        this.conflictRemote = remote;
        this.staged = local;
        this.report({
          status: "conflict",
          conflicts,
          message: "Hai máy đã sửa cùng phần dữ liệu. Chọn bản cần giữ.",
        });
        return;
      }
      const incoming = Object.keys(remote?.manifest.fields || {}).filter(
        (key) =>
          !this.dirty.has(key) &&
          remote?.manifest.fields[key]?.hash !==
            this.base?.manifest.fields[key]?.hash,
      );
      if (remote && incoming.length) await this.install(remote, incoming);
      if (this.stopped) return;
      let committed = remote;
      const merged = mergeManifest(this.base?.manifest || EMPTY, local, remote?.manifest || EMPTY);
      if (pending.length && (!remote || !sameManifest(merged, remote.manifest)))
        committed = await this.repository.commit(
          remote?.revision || 0,
          merged,
        );
      if (this.stopped) return;
      for (const key of pending) {
        this.refs[key] = snapshot[key];
        rememberCloudField(this.repository.owner, key, snapshot[key], local.fields[key].hash);
        if (
          (this.getData() as unknown as Record<string, unknown>)[key] ===
          snapshot[key]
        )
          this.dirty.delete(key);
      }
      if (committed) {
        const fields = { ...committed.manifest.fields };
        // Edits made during a remote pull still refer to the old baseline.
        // Advancing it would hide a same-table conflict on the next upload.
        for (const key of incoming) {
          if (!this.dirty.has(key)) continue;
          const previous = this.base?.manifest.fields[key];
          if (previous) fields[key] = previous;
          else delete fields[key];
        }
        this.base = { ...committed, manifest: { version: 1, fields } };
      } else this.base = null;
      // Read history after installing incoming workspace fields: an older
      // workspace must not replace a newer monthly snapshot just fetched.
      if (!this.initializing) await this.refresh();
      if (this.stopped) return;
      this.checkpoint();
      if (pending.length && typeof window !== "undefined") window.dispatchEvent(new Event(CLOUD_CACHE_EVENT));
      if (pending.length) void this.repository.cleanup().catch(() => undefined);
      this.report({
        status: this.dirty.size ? "saving" : "synced",
        updatedAt: committed?.updated_at,
        revision: committed?.revision,
      });
    } catch (error) {
      if ((error as { code?: string })?.code === "P0002")
        this.timer = setTimeout(() => void this.sync(), 1000);
      else this.fail(error);
    } finally {
      this.running = false;
      if (this.dirty.size && !this.conflictRemote && !this.stopped) {
        clearTimeout(this.timer);
        this.timer = setTimeout(() => void this.sync(), 10000);
      }
    }
  }
  async resolve(keepLocal: boolean) {
    const remote = this.conflictRemote;
    if (!remote || this.running) return;
    await this.backup();
    if (this.stopped) return;
    const conflicts = conflictingFields(
      this.base?.manifest || EMPTY,
      this.staged!,
      remote.manifest,
    );
    if (!keepLocal) {
      for (const key of conflicts) this.dirty.delete(key);
      await this.install(remote, conflicts);
    }
    // Rebase only the conflict keys; unrelated remote and local updates still merge normally.
    const fields = { ...(this.base?.manifest.fields || {}) };
    for (const key of conflicts) fields[key] = remote.manifest.fields[key];
    this.base = { ...remote, manifest: { version: 1, fields } };
    this.conflictRemote = null;
    this.staged = null;
    this.checkpoint();
    await this.sync();
  }
  stop() {
    this.stopped = true;
    clearTimeout(this.timer);
    this.repository.dispose();
  }
}
