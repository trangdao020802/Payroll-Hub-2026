import type { Manifest } from "./protocol";
import type { AppData } from "../../types";

export const CLOUD_HASHES_KEY = "PayrollApp_Data:cloudHashes";
export const CLOUD_CACHE_EVENT = "payroll-cloud-cache-updated";
export interface StoredCloudHashes {
  owner: string;
  hashes: Record<string, string>;
}

/** Legacy hydration mutates these table containers/arrays in place. Detach
 * them after registering cached hashes so a migration cannot retain old proof.
 * Large split collections keep their immutable references and cache stamps. */
export function prepareHydrationTables(data: AppData): AppData {
  const prepared = { ...data };
  for (const field of ["Hold_AE", "Sheet1_AE", "Bank_North_AE", "SoSanh_AE"] as const) {
    const table = data[field];
    if (table) prepared[field] = {
      ...table,
      headers: [...(table.headers || [])],
      data: [...(table.data || [])],
    };
  }
  return prepared;
}

// Serialize readers/writers from other tabs too. The hash stamp is invalidated
// before writes so an interrupted save remains a cache miss, never stale proof.
export async function withAppStorageLock<T>(task: () => Promise<T>): Promise<T> {
  if (typeof navigator !== "undefined" && navigator.locks) {
    return await navigator.locks.request("PayrollApp:app-data", task);
  }
  return task();
}

export function persistedFieldHashes(
  owner: string,
  data: object,
  splitFields: readonly string[],
  written: ReadonlySet<string>,
  previous: StoredCloudHashes | null,
): Record<string, string> {
  const hashes = cachedFieldHashes(owner, data);
  for (const field of splitFields) {
    // A skipped split value may have been changed on disk by another tab.
    if (!written.has(field) && (previous?.owner !== owner || previous.hashes[field] !== hashes[field])) {
      delete hashes[field];
    }
  }
  return hashes;
}
interface FieldVersion {
  owner: string;
  hash: string;
  value: unknown;
}
const objects = new WeakMap<object, Omit<FieldVersion, "value">>();
const primitives = new Map<string, FieldVersion>();

// A hash belongs to the exact immutable value that was downloaded/committed,
// never just to the checkpoint revision (local persistence may lag behind it).
export function rememberCloudField(owner: string, key: string, value: unknown, hash: string) {
  if (value !== null && typeof value === "object") objects.set(value, { owner, hash });
  else primitives.set(key, { owner, value, hash });
}

export function cachedFieldHash(owner: string, key: string, value: unknown): string | undefined {
  const version = value !== null && typeof value === "object"
    ? objects.get(value) : primitives.get(key);
  if (version?.owner !== owner) return;
  if (value === null || typeof value !== "object") {
    if (!Object.is((version as FieldVersion).value, value)) return;
  }
  return version.hash;
}

export function cachedFieldHashes(owner: string, data: object): Record<string, string> {
  const hashes: Record<string, string> = {};
  for (const [key, value] of Object.entries(data)) {
    const hash = cachedFieldHash(owner, key, value);
    if (hash) hashes[key] = hash;
  }
  return hashes;
}

export function restoreFieldHashes(owner: string, data: object, hashes: Record<string, string>) {
  for (const [key, value] of Object.entries(data)) {
    if (typeof hashes[key] === "string") rememberCloudField(owner, key, value, hashes[key]);
  }
}

export function sameManifest(left: Manifest, right: Manifest): boolean {
  const keys = Object.keys(left.fields);
  return keys.length === Object.keys(right.fields).length
    && keys.every(key => left.fields[key]?.hash === right.fields[key]?.hash);
}
