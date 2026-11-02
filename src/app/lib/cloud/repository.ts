import { supabase } from "../../../lib/supabaseClient";
import {
  BUCKET,
  PART_SIZE,
  type Asset,
  type Manifest,
  type Workspace,
} from "./protocol";
import { digest } from "./codec";
import { withDeadline } from "./deadline";
import { createCloudValueReceiver } from './value-chunks';
import localforage from 'localforage';
export class CloudRepository {
  private worker: Worker | null = null;
  private session = new AbortController();
  private counter = 0;
  private pending = new Map<
    number,
    {
      resolve: (result: { result: unknown; hash?: string }) => void;
      reject: (error: Error) => void;
      timer: ReturnType<typeof setTimeout>;
      receiver?: ReturnType<typeof createCloudValueReceiver>;
      sequence?: number;
    }
  >();
  constructor(readonly owner: string) {}
  private resetWorker(message: string) {
    this.worker?.terminate();
    this.worker = null;
    for (const task of this.pending.values()) {
      clearTimeout(task.timer);
      task.reject(new Error(message));
    }
    this.pending.clear();
  }
  private getWorker() {
    if (this.worker) return this.worker;
    const worker = new Worker(new URL("./codec.worker.ts", import.meta.url), { type: "module" });
    this.worker = worker;
    worker.onerror = () => this.resetWorker("Không thể đọc/nén dữ liệu. Bấm Đồng bộ ngay để thử lại.");
    worker.onmessageerror = () => this.resetWorker("Không thể nhận dữ liệu đã đọc. Bấm Đồng bộ ngay để thử lại.");
    worker.onmessage = ({ data }) => {
      if (this.worker !== worker) return;
      const task = this.pending.get(data.id);
      if (!task) return;
      if (data.type === 'codec-chunk') {
        if (data.sequence !== (task.sequence || 0) + 1) {
          this.resetWorker('Dữ liệu cloud truyền không đầy đủ. Bấm Đồng bộ ngay để thử lại.');
          return;
        }
        try {
          task.receiver ||= createCloudValueReceiver();
          task.receiver.accept(data.part);
          task.sequence = data.sequence;
        } catch (error) {
          this.resetWorker(error instanceof Error ? error.message : 'Không thể mở dữ liệu cloud.');
          return;
        }
        clearTimeout(task.timer);
        task.timer = setTimeout(() => this.resetWorker('Mở dữ liệu cloud quá lâu. Bấm Đồng bộ ngay để thử lại.'), 120000);
        setTimeout(() => {
          if (this.worker === worker && this.pending.has(data.id)) worker.postMessage({ type: 'codec-ack', id: data.id, sequence: data.sequence });
        }, 0);
        return;
      }
      this.pending.delete(data.id);
      clearTimeout(task.timer);
      try {
        if (data.error) task.reject(Object.assign(new Error(data.error), { code: data.code }));
        else {
          if (data.type === 'codec-complete' && !task.receiver) throw new Error('Dữ liệu cloud chưa đầy đủ.');
          task.resolve(data.type === 'codec-complete' ? { result: task.receiver!.finish() } : data);
        }
      } catch (error) { task.reject(error instanceof Error ? error : new Error(String(error))); }
      // A decoded full workspace must not keep its worker heap alive alongside
      // the reconstructed tables and IndexedDB's durable copy.
      if (!this.pending.size) { worker.terminate(); this.worker = null; }
    };
    return worker;
  }
  dispose() {
    this.session.abort();
    this.resetWorker("Phiên đồng bộ đã đóng.");
  }
  private codec(operation: "encode" | "decode" | "cache-read", value: unknown) {
    return new Promise<{ result: unknown; hash?: string }>(
      (resolve, reject) => {
        if (this.session.signal.aborted) return reject(new Error("Phiên đồng bộ đã đóng."));
        const id = ++this.counter;
        const timer = setTimeout(() => this.resetWorker("Đọc/nén dữ liệu quá lâu. Dữ liệu trên máy vẫn được giữ; bấm Đồng bộ ngay để thử lại."), 120000);
        this.pending.set(id, { resolve, reject, timer });
        try {
          this.getWorker().postMessage({ id, operation, value, chunked: operation !== 'encode' });
        } catch (error) {
          this.pending.delete(id);
          clearTimeout(timer);
          reject(error instanceof Error ? error : new Error(String(error)));
        }
      },
    );
  }
  /** IndexedDB materializes large snapshots in a worker, then yields chunks to the UI. */
  async readCachedField<T>(key: string): Promise<T | null> {
    try { return (await this.codec('cache-read', key)).result as T | null; }
    catch (error) {
      // localforage may use localStorage on browsers without IndexedDB. That
      // small legacy fallback is available only on the window thread.
      if ((error as { code?: string }).code === 'CACHE_WORKER_UNAVAILABLE') return localforage.getItem<T>(key);
      throw error;
    }
  }
  private request<T>(operation: (signal: AbortSignal) => PromiseLike<T>) {
    return withDeadline(operation, "Kết nối cloud quá lâu. Dữ liệu trên máy vẫn được giữ; bấm Đồng bộ ngay để thử lại.", 30000, this.session.signal);
  }
  async read(): Promise<Workspace | null> {
    const { data, error } = await this.request(signal => supabase
      .from("payroll_workspaces")
      .select("owner_id,revision,manifest,updated_at")
      .eq("owner_id", this.owner)
      .abortSignal(signal).maybeSingle());
    if (error) throw error;
    return data;
  }
  async checkMembership() {
    const { data, error } = await this.request(signal => supabase
      .from("transaction_history_members")
      .select("user_id")
      .eq("user_id", this.owner)
      .abortSignal(signal).maybeSingle());
    if (error) throw error;
    if (!data)
      throw new Error("Tài khoản chưa được cấp quyền dữ liệu Payroll.");
  }
  async upload(value: unknown, previous?: Asset): Promise<Asset> {
    const { result, hash } = await this.codec("encode", value);
    if (previous && previous.hash === hash) return previous;
    const blob = result as Blob;
    if (blob.size > 200 * 1024 * 1024)
      throw new Error(
        "Một bảng vượt 200 MB sau nén. Dữ liệu máy này vẫn được giữ.",
      );
    const parts: string[] = [];
    for (let offset = 0; offset < blob.size; offset += PART_SIZE) {
      const path = `${this.owner}/${hash}/${String(parts.length).padStart(6, "0")}.part`;
      const { error } = await this.request(() => supabase.storage
        .from(BUCKET)
        .upload(path, blob.slice(offset, offset + PART_SIZE), {
          contentType: "application/octet-stream",
          upsert: false,
        }));
      if (
        error &&
        !("statusCode" in error && String(error.statusCode) === "409") &&
        !/already exists|duplicate/i.test(error.message)
      )
        throw error;
      parts.push(path);
    }
    return { hash: hash!, bytes: blob.size, parts };
  }
  async download(asset: Asset, progress?: (stage: 'download' | 'decode', received: number) => void): Promise<unknown> {
    const blobs: Blob[] = [];
    let received = 0;
    for (let offset = 0; offset < asset.parts.length; offset += 4) {
      const batch = await Promise.all(asset.parts.slice(offset, offset + 4).map(async path => {
        if (!path.startsWith(`${this.owner}/${asset.hash}/`)) throw new Error("Đường dẫn cloud không hợp lệ");
        const { data, error } = await this.request(() => supabase.storage.from(BUCKET).download(path));
        if (error) throw error;
        received += data.size;
        progress?.('download', received);
        return data;
      }));
      // Promise.all retains source order even when parts finish out of order.
      blobs.push(...batch);
    }
    const blob = new Blob(blobs);
    if (blob.size !== asset.bytes || (await digest(blob)) !== asset.hash)
      throw new Error("Dữ liệu cloud chưa đầy đủ. Bản trên máy vẫn được giữ.");
    progress?.('decode', received);
    return (await this.codec("decode", blob)).result;
  }
  async cleanup() {
    const { data, error } = await this.request(signal => supabase.rpc("payroll_orphan_assets").abortSignal(signal));
    if (error || !data?.length) return;
    const paths = (data as { path: string }[])
      .map((item) => item.path)
      .filter((path) => path.startsWith(`${this.owner}/`));
    for (let offset = 0; offset < paths.length; offset += 100) {
      const result = await this.request(() => supabase.storage
        .from(BUCKET)
        .remove(paths.slice(offset, offset + 100)));
      if (result.error) return;
    }
  }
  async commit(revision: number, manifest: Manifest): Promise<Workspace> {
    const { data, error } = await this.request(signal => supabase.rpc("commit_payroll_workspace", {
      expected_revision: revision,
      new_manifest: manifest,
    }).abortSignal(signal));
    if (error) throw error;
    return data;
  }
}
