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
      const task = this.pending.get(data.id);
      this.pending.delete(data.id);
      if (task) clearTimeout(task.timer);
      if (data.error) task?.reject(new Error(data.error));
      else task?.resolve(data);
    };
    return worker;
  }
  dispose() {
    this.session.abort();
    this.resetWorker("Phiên đồng bộ đã đóng.");
  }
  private codec(operation: "encode" | "decode", value: unknown) {
    return new Promise<{ result: unknown; hash?: string }>(
      (resolve, reject) => {
        if (this.session.signal.aborted) return reject(new Error("Phiên đồng bộ đã đóng."));
        const id = ++this.counter;
        const timer = setTimeout(() => this.resetWorker("Đọc/nén dữ liệu quá lâu. Dữ liệu trên máy vẫn được giữ; bấm Đồng bộ ngay để thử lại."), 120000);
        this.pending.set(id, { resolve, reject, timer });
        try {
          this.getWorker().postMessage({ id, operation, value });
        } catch (error) {
          this.pending.delete(id);
          clearTimeout(timer);
          reject(error instanceof Error ? error : new Error(String(error)));
        }
      },
    );
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
  async download(asset: Asset): Promise<unknown> {
    const blobs: Blob[] = [];
    for (const path of asset.parts) {
      if (!path.startsWith(`${this.owner}/${asset.hash}/`))
        throw new Error("Đường dẫn cloud không hợp lệ");
      const { data, error } = await this.request(() => supabase.storage
        .from(BUCKET)
        .download(path));
      if (error) throw error;
      blobs.push(data);
    }
    const blob = new Blob(blobs);
    if (blob.size !== asset.bytes || (await digest(blob)) !== asset.hash)
      throw new Error("Dữ liệu cloud chưa đầy đủ. Bản trên máy vẫn được giữ.");
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
