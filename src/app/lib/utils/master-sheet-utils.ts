/**
 * Normalize an Excel sheet name before matching it.
 *
 * The normalization is intentionally limited to matching only: the original
 * sheet name is still kept everywhere it is displayed or exported.
 */
export function normalizeMasterSheetName(sheetName: unknown): string {
  return String(sheetName ?? "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[Đđ]/g, "D")
    .toUpperCase()
    .replace(/\s+/g, " ")
    .trim();
}

const SHEET_ONE_PATTERN = /SHEET +1(?!\d)/;

function isNormalizedBankSheetName(normalizedName: string): boolean {
  return normalizedName.includes("BANK") || normalizedName.includes("NGAN HANG");
}

function isNormalizedSheetOneName(normalizedName: string): boolean {
  return SHEET_ONE_PATTERN.test(normalizedName);
}

export function isRosterMasterSheetName(sheetName: unknown): boolean {
  return normalizeMasterSheetName(sheetName).includes("ROSTER");
}

export function findMasterRosterSheetName(sheetNames: readonly string[]): string | undefined {
  return sheetNames.find(name => normalizeMasterSheetName(name) === "ROSTER")
    ?? sheetNames.find(name => /^Q[ _-]+ROSTER$/.test(normalizeMasterSheetName(name)))
    ?? sheetNames.find(isRosterMasterSheetName);
}

export function isBankMasterSheetName(sheetName: unknown): boolean {
  return isNormalizedBankSheetName(normalizeMasterSheetName(sheetName));
}

export function isSheetOneMasterSheetName(sheetName: unknown): boolean {
  const normalizedName = normalizeMasterSheetName(sheetName);

  // A real whitespace gap between SHEET and 1 is required. This accepts
  // "SHEET 1" and "SHEET    1", but rejects "SHEET1", "SHEET-1" and
  // longer sheet numbers such as "SHEET 10".
  return isNormalizedSheetOneName(normalizedName);
}

export function isHoldMasterSheetName(sheetName: unknown): boolean {
  return normalizeMasterSheetName(sheetName).includes("HOLD");
}

export function isRelevantMasterSheetName(
  sheetName: unknown,
  isMktFile: boolean,
): boolean {
  if (isMktFile) {
    return isRosterMasterSheetName(sheetName);
  }

  const normalizedName = normalizeMasterSheetName(sheetName);
  return (
    isNormalizedBankSheetName(normalizedName) ||
    isNormalizedSheetOneName(normalizedName) ||
    normalizedName.includes("HOLD") ||
    normalizedName.includes("ADD") ||
    normalizedName.includes("SUMMER") ||
    normalizedName.includes("BONUS") ||
    normalizedName.includes("SO SANH AE")
  );
}
