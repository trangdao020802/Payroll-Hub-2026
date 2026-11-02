import localforage from "localforage";
import { formatTimestamp, getRelativeTime, recordDataChange } from "./data-change-tracker";
import type { AppData } from "../../types";

export type SnapshotTrigger = "auto" | "manual" | "import" | "restore" | "edit";

export interface SnapshotMetadata {
  id: string;
  timestamp: string; // ISO 8601 string
  createdAt: number; // Unix timestamp ms
  timeString: string; // e.g. "16:45:10"
  dateString: string; // e.g. "19/09/2026"
  relativeTime: string; // e.g. "5 phút trước"
  title: string; // e.g. "Tự động lưu sau khi nhập file Roster"
  note?: string; // Optional user note
  trigger: SnapshotTrigger;
  globalMonth: string; // e.g. "2026-03"
  isPinned?: boolean;
  stats: {
    timesheetRows: number;
    masterRows: number;
    staffRows: number;
    salaryScaleRows: number;
    totalRows: number;
    approxSizeKb: number;
  };
}

export interface DataSnapshot extends SnapshotMetadata {
  data: AppData;
}

const metaStore = localforage.createInstance({
  name: "PayrollApp",
  storeName: "snapshots_meta",
});

const dataStore = localforage.createInstance({
  name: "PayrollApp",
  storeName: "snapshots_data",
});

const INDEX_KEY = "snapshots_index";
const MAX_SNAPSHOTS = 30;

// In-memory fallback for environments without IndexedDB (e.g. Node/SSR/Tests)
const memoryStore = new Map<string, unknown>();

async function getStorageItem<T>(store: LocalForage, key: string): Promise<T | null> {
  try {
    return await store.getItem<T>(key);
  } catch (error) {
    if (typeof window !== "undefined") throw error;
    return structuredClone((memoryStore.get(key) as T) ?? null);
  }
}

async function setStorageItem<T>(store: LocalForage, key: string, value: T): Promise<T> {
  try {
    await store.setItem(key, value);
    // IndexedDB already owns a durable structured clone. Do not also pin
    // every full snapshot in the tab's heap (up to 30 copies of the roster).
    memoryStore.delete(key);
  } catch (error) {
    if (typeof window !== "undefined") throw error;
    memoryStore.set(key, structuredClone(value));
  }
  return value;
}

async function removeStorageItem(store: LocalForage, key: string): Promise<void> {
  memoryStore.delete(key);
  try {
    await store.removeItem(key);
  } catch {
    // Memory fallback updated
  }
}

function dispatchSnapshotEvent(eventName: string, detail?: unknown) {
  if (typeof window !== "undefined" && typeof window.dispatchEvent === "function") {
    window.dispatchEvent(new CustomEvent(eventName, { detail }));
  }
}

function estimateSnapshotSize(value: unknown, ancestors = new Set<object>()): number {
  if (typeof value === "string") return value.length + 2;
  if (value === null || typeof value !== "object") return 8;
  if (ancestors.has(value)) return 4;
  ancestors.add(value);
  let size = 2;
  if (Array.isArray(value)) {
    const count = Math.min(value.length, 32);
    let sampleSize = 0;
    for (let i = 0; i < count; i++) {
      sampleSize += estimateSnapshotSize(value[Math.floor(i * value.length / count)], ancestors) + 1;
    }
    size += count ? sampleSize * value.length / count : 0;
  } else {
    for (const [key, item] of Object.entries(value)) {
      size += key.length + 4 + estimateSnapshotSize(item, ancestors);
    }
  }
  ancestors.delete(value);
  return size;
}

export function calculateSnapshotStats(appData: AppData): SnapshotMetadata["stats"] {
  const timesheetRows = Array.isArray(appData?.Timesheet_Roster)
    ? appData.Timesheet_Roster.length
    : 0;
  const masterRows = Array.isArray(appData?.Master_Roster)
    ? appData.Master_Roster.length
    : 0;
  const staffRows = Array.isArray(appData?.Q_Staff)
    ? appData.Q_Staff.length
    : 0;
  const salaryScaleRows = Array.isArray(appData?.Q_Salary_Scale)
    ? appData.Q_Salary_Scale.length
    : 0;
  const totalRows = timesheetRows + masterRows + staffRows + salaryScaleRows;

  let approxSizeKb = 0;
  try {
    approxSizeKb = Math.round(estimateSnapshotSize(appData) / 1024);
  } catch {
    approxSizeKb = 50;
  }

  return {
    timesheetRows,
    masterRows,
    staffRows,
    salaryScaleRows,
    totalRows,
    approxSizeKb,
  };
}

/**
 * Save a new snapshot into storage
 */
export async function saveSnapshot(
  appData: AppData,
  options: {
    title?: string;
    trigger?: SnapshotTrigger;
    note?: string;
    isPinned?: boolean;
    required?: boolean;
  } = {},
): Promise<SnapshotMetadata> {
  const now = new Date();
  const timeInfo = formatTimestamp(now);
  const id = `snap_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`;
  const trigger = options.trigger || "manual";
  const title =
    options.title ||
    (trigger === "auto"
      ? `Tự động lưu (${timeInfo.timeString})`
      : trigger === "import"
        ? `Sau khi nạp dữ liệu (${timeInfo.timeString})`
        : trigger === "restore"
          ? `Bản sao lưu trước khôi phục (${timeInfo.timeString})`
          : `Điểm khôi phục (${timeInfo.timeString})`);

  const stats = calculateSnapshotStats(appData);

  const meta: SnapshotMetadata = {
    id,
    timestamp: timeInfo.iso,
    createdAt: Date.now(),
    timeString: timeInfo.timeString,
    dateString: timeInfo.dateString,
    relativeTime: "Vừa xong",
    title,
    note: options.note || "",
    trigger,
    globalMonth: appData?.globalMonth || "2026-03",
    isPinned: Boolean(options.isPinned),
    stats,
  };

  const fullSnapshot: DataSnapshot = {
    ...meta,
    // App state is immutable; the storage adapter performs the one required
    // clone. A JSON round trip here doubles peak memory for large imports.
    data: appData,
  };

  try {
    // 1. Save data payload
    await setStorageItem(dataStore, id, fullSnapshot);

    // 2. Update metadata index
    const currentList = (await getStorageItem<SnapshotMetadata[]>(metaStore, INDEX_KEY)) || [];
    const updatedList = [meta, ...currentList];

    // 3. Prune old unpinned snapshots if exceeding MAX_SNAPSHOTS
    if (updatedList.length > MAX_SNAPSHOTS) {
      const pinned = updatedList.filter((s) => s.isPinned);
      const unpinned = updatedList.filter((s) => !s.isPinned);
      const toKeepUnpinned = unpinned.slice(0, MAX_SNAPSHOTS - pinned.length);
      const toRemove = unpinned.slice(MAX_SNAPSHOTS - pinned.length);

      for (const item of toRemove) {
        try {
          await removeStorageItem(dataStore, item.id);
        } catch {
          // Ignore
        }
      }

      const finalList = [...pinned, ...toKeepUnpinned].sort(
        (a, b) => b.createdAt - a.createdAt,
      );
      await setStorageItem(metaStore, INDEX_KEY, finalList);
    } else {
      await setStorageItem(metaStore, INDEX_KEY, updatedList);
    }

    // 4. Notify listeners
    dispatchSnapshotEvent("payroll-snapshots-updated", { detail: { newId: id } });
  } catch (err) {
    console.error("Failed to save snapshot", err);
    if (options.required) throw err;
  }

  return meta;
}

/**
 * Get all snapshot metadata items, sorted newest first
 */
export async function getSnapshotsList(): Promise<SnapshotMetadata[]> {
  try {
    const list = (await getStorageItem<SnapshotMetadata[]>(metaStore, INDEX_KEY)) || [];
    return list.map((item) => ({
      ...item,
      relativeTime: getRelativeTime(item.timestamp),
    }));
  } catch (err) {
    console.error("Failed to get snapshots list", err);
    return [];
  }
}

/**
 * Load complete snapshot with data for preview or restoration
 */
export async function getSnapshotById(id: string): Promise<DataSnapshot | null> {
  try {
    const snapshot = await getStorageItem<DataSnapshot>(dataStore, id);
    if (!snapshot) return null;
    return {
      ...snapshot,
      relativeTime: getRelativeTime(snapshot.timestamp),
    };
  } catch (err) {
    console.error(`Failed to get snapshot with id ${id}`, err);
    return null;
  }
}

/**
 * Restore a snapshot:
 * 1. Creates an automatic safety snapshot of current data before overwriting
 * 2. Returns the restored AppData
 * 3. Logs an audit event
 */
export async function restoreSnapshot(
  id: string,
  currentAppData: AppData,
): Promise<{ restoredData: AppData; meta: SnapshotMetadata }> {
  // 1. Create safety backup of current state
  try {
    await saveSnapshot(currentAppData, {
      title: "Bảo vệ: Tự động sao lưu trước khi khôi phục",
      trigger: "restore",
      note: `Tạo trước khi khôi phục snapshot ${id}`,
      isPinned: true, // pin safety backups so user won't lose them
    });
  } catch (e) {
    console.warn("Could not create pre-restore safety snapshot", e);
  }

  // 2. Fetch target snapshot
  const targetSnapshot = await getSnapshotById(id);
  if (!targetSnapshot || !targetSnapshot.data) {
    throw new Error("Không tìm thấy dữ liệu điểm khôi phục này.");
  }

  // 3. Record in audit tracker
  try {
    recordDataChange({
      actionType: "restore",
      entity: "Điểm khôi phục",
      summary: `Khôi phục dữ liệu: ${targetSnapshot.title}`,
      details: `Khôi phục về trạng thái lúc ${targetSnapshot.timeString} ngày ${targetSnapshot.dateString} (${targetSnapshot.globalMonth})`,
    });
  } catch {
    // Ignore
  }

  dispatchSnapshotEvent("payroll-snapshots-updated", { detail: { restoredId: id } });

  return {
    restoredData: targetSnapshot.data,
    meta: targetSnapshot,
  };
}

/**
 * Delete a snapshot by ID
 */
export async function deleteSnapshot(id: string): Promise<void> {
  try {
    await removeStorageItem(dataStore, id);
    const currentList = (await getStorageItem<SnapshotMetadata[]>(metaStore, INDEX_KEY)) || [];
    const updated = currentList.filter((s) => s.id !== id);
    await setStorageItem(metaStore, INDEX_KEY, updated);

    dispatchSnapshotEvent("payroll-snapshots-updated", { detail: { deletedId: id } });
  } catch (err) {
    console.error(`Failed to delete snapshot ${id}`, err);
  }
}

/**
 * Toggle pin status
 */
export async function togglePinSnapshot(id: string): Promise<boolean> {
  try {
    const currentList = (await getStorageItem<SnapshotMetadata[]>(metaStore, INDEX_KEY)) || [];
    let nextPinned = false;
    const updated = currentList.map((s) => {
      if (s.id === id) {
        nextPinned = !s.isPinned;
        return { ...s, isPinned: nextPinned };
      }
      return s;
    });
    await setStorageItem(metaStore, INDEX_KEY, updated);

    // Also update data store if present
    const dataItem = await getStorageItem<DataSnapshot>(dataStore, id);
    if (dataItem) {
      dataItem.isPinned = nextPinned;
      await setStorageItem(dataStore, id, dataItem);
    }

    dispatchSnapshotEvent("payroll-snapshots-updated");
    return nextPinned;
  } catch (err) {
    console.error(`Failed to toggle pin for snapshot ${id}`, err);
    return false;
  }
}

/**
 * Rename a snapshot
 */
export async function renameSnapshot(id: string, newTitle: string): Promise<void> {
  if (!newTitle.trim()) return;
  try {
    const currentList = (await getStorageItem<SnapshotMetadata[]>(metaStore, INDEX_KEY)) || [];
    const updated = currentList.map((s) => {
      if (s.id === id) {
        return { ...s, title: newTitle.trim() };
      }
      return s;
    });
    await setStorageItem(metaStore, INDEX_KEY, updated);

    const dataItem = await getStorageItem<DataSnapshot>(dataStore, id);
    if (dataItem) {
      dataItem.title = newTitle.trim();
      await setStorageItem(dataStore, id, dataItem);
    }

    dispatchSnapshotEvent("payroll-snapshots-updated");
  } catch (err) {
    console.error(`Failed to rename snapshot ${id}`, err);
  }
}

/**
 * Export snapshot as JSON file for download
 */
export async function exportSnapshotJson(id: string): Promise<void> {
  const snapshot = await getSnapshotById(id);
  if (!snapshot) {
    throw new Error("Không tìm thấy snapshot để xuất");
  }

  const cleanName = (snapshot.title || "snapshot")
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "-")
    .replace(/-+/g, "-");
  const fileName = `payroll-${cleanName}-${snapshot.dateString.replace(/\//g, "")}-${snapshot.timeString.replace(/:/g, "")}.json`;

  const blob = new Blob([JSON.stringify(snapshot, null, 2)], {
    type: "application/json",
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = fileName;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}
