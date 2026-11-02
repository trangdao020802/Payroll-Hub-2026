import type { AppData } from "../../types";
import { getCenterInfoByL07, getL07FromFileName } from "./center-utils";
import { normalizeTimesheetL07, type TimesheetLinkInput } from "./timesheet-center-coverage";
import { shouldSkipTimesheetSource } from "./timesheet-input-resolver";
import { replaceTimesheetRosterRows } from "./timesheet-roster-utils";
import { formatTimesheetSyncDate } from "./timesheet-sync-date";
import { withDeadline } from "../cloud/deadline";

export const DEFAULT_TIMESHEET_FOLDER_URL = "https://drive.google.com/drive/folders/1gU6Hcrv94Bx_yv1qNTqH0vQNy7ElKzXJ";
export const MKT_LOCAL_NORTH_URL = "https://docs.google.com/spreadsheets/d/1z7DJYJAyWqBw8IXNYbEIHhGXBMumsRA4rUHT1prBsFo/edit?gid=1119129159#gid=1119129159";

type DriveFile = { id: string; name: string; modifiedTime?: string; createdTime?: string; webViewLink?: string; url?: string };
export function fetchTimesheetFolderFiles(folderId: string, timeoutMs = 60_000): Promise<DriveFile[]> {
  return withDeadline(async signal => {
    const response = await fetch(`/api/drive-folder-files?folderId=${encodeURIComponent(folderId)}`, { signal });
    const data = await response.json().catch(() => ({}));
    if (!response.ok) throw new Error(data.error || "Không thể lấy danh sách file từ thư mục. Vui lòng kiểm tra quyền truy cập.");
    if (!data.success || !Array.isArray(data.files)) throw new Error("Không tìm thấy file nào trong thư mục.");
    return data.files;
  }, "Tải danh sách Google Drive quá lâu. Vui lòng thử lại.", timeoutMs);
}

export async function resolveTimesheetLink(input: TimesheetLinkInput) {
  if (input.url?.trim()) return { url: input.url.trim(), date: input.date };
  const l07 = normalizeTimesheetL07(input.l07);
  if (l07.replace(/_/g, " ") === "MKT LOCAL NORTH") return { url: MKT_LOCAL_NORTH_URL, date: input.date };
  if (!l07) throw new Error("Không có mã L07 để tìm kiếm.");

  const folderId = DEFAULT_TIMESHEET_FOLDER_URL.split("/").pop()!;
  const files = await fetchTimesheetFolderFiles(folderId);
  const file = files
    .filter((candidate) => !candidate.name.toLowerCase().includes("copy")
      && !shouldSkipTimesheetSource(candidate.name, candidate.webViewLink, candidate.url)
      && normalizeTimesheetL07(getL07FromFileName(candidate.name)) === l07)
    .sort((a, b) => (Date.parse(b.modifiedTime || b.createdTime || "") || 0) - (Date.parse(a.modifiedTime || a.createdTime || "") || 0))[0];
  if (!file) throw new Error(`Không tìm thấy file nào cho trung tâm ${l07} trong thư mục GDrive.`);
  return {
    url: `https://docs.google.com/spreadsheets/d/${file.id}/edit`,
    date: formatTimesheetSyncDate(new Date(file.modifiedTime || file.createdTime || Date.now())),
  };
}

export function applyTimesheetLinkResult(
  previous: AppData,
  input: TimesheetLinkInput,
  rows: Record<string, unknown>[],
  source: { url: string; fileName: string; date: string; lastSyncedAt?: string },
  prepared?: { original: Record<string, unknown>[]; rows: Record<string, unknown>[] },
): AppData {
  const currentInput = previous.Timesheet_InputList.find((row) => row.id === input.id);
  // A deleted or reassigned source must not reappear when its request finishes.
  if (!currentInput || normalizeTimesheetL07(currentInput.l07) !== normalizeTimesheetL07(input.l07)) return previous;
  if (currentInput.url !== input.url && currentInput.url !== source.url) return previous;
  const targetL07 = input.l07.trim();
  const aeCode = input.aeCode || getCenterInfoByL07(targetL07)?.aeCode || "";
  return {
    ...previous,
    Timesheet_Roster: prepared && previous.Timesheet_Roster === prepared.original
      ? prepared.rows : replaceTimesheetRosterRows(previous.Timesheet_Roster || [], rows, {
      sourceRowIds: new Set([input.id, ...(currentInput.legacyRowIds || [])]),
      targetL07,
      targetAeCode: aeCode,
    }),
    Timesheet_InputList: previous.Timesheet_InputList.map((row) => row.id === input.id ? {
      ...row, ...source, status: "success", count: rows.length, legacyRowIds: [],
    } : row),
  };
}
