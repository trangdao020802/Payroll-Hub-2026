/* eslint-disable @typescript-eslint/no-explicit-any */
import { parseMoneyToNumber, removeVietnameseTones } from "./data-utils";
import { parseMonthPeriod } from "./bulk-payment-analytics";
import { isSavedTransactionField } from './transaction-saved-fields';
import type { AppData } from '../../types';
import { deductionsRowMonth, isDeductionsMonthLocked } from './deductions-lock';

export const TRANSACTION_REFERENCE_AUDIT_KEY =
  "_transactionReferenceAudit" as const;

export type TransactionReferenceTable = "Sheet1_AE" | "Hold_AE";
export type TransactionReferenceField =
  | "idNumber"
  | "fullName"
  | "bankAccountNumber";

export interface TransactionReferenceAuditEntry {
  field: TransactionReferenceField;
  fieldLabel: string;
  oldValue: unknown;
  newValue: unknown;
  correctedAt: string;
  targetTable: TransactionReferenceTable;
  transactionKey: string;
  transactionIndex: number;
  transactionRowId: string;
  transactionSerial: string;
  transactionId: string;
  transactionName: string;
  transactionAccount: string;
}

export interface TransactionReferenceCorrection {
  table: TransactionReferenceTable;
  rowIndex: number;
  field: TransactionReferenceField;
  fieldKey: string;
  oldValue: unknown;
  newValue: unknown;
}

export interface TransactionRawTimesheetCorrection {
  transactionIndex: number;
  field: TransactionReferenceField;
  fieldKey: string;
  oldValue: unknown;
  newValue: unknown;
  source: "RAWDATA_TIMESHEET";
}

export interface TransactionReferenceMatch {
  transactionKey: string;
  transactionIndex: number;
  transactionRow: any;
  grossRowIndexes: number[];
  deductionRowIndexes: number[];
  corrections: TransactionReferenceCorrection[];
  transactionCorrections: TransactionRawTimesheetCorrection[];
  reason:
    | "name-account"
    | "id"
    | "two-fields"
    | "raw-timesheet"
    | "unmatched";
}

export interface TransactionReferencePlan {
  matches: TransactionReferenceMatch[];
  byTransactionKey: Map<string, TransactionReferenceMatch>;
  byTransactionIndex: Map<number, TransactionReferenceMatch>;
  effectiveTransactionRows: any[];
  transactionCorrections: TransactionRawTimesheetCorrection[];
}

interface ReferenceIdentity {
  id: string;
  name: string;
  account: string;
  rawId: unknown;
  rawName: unknown;
  rawAccount: unknown;
}

const ID_KEYS = [
  "ID Number",
  "ID NUMBER",
  "Document ID",
  "Document ID / CCCD",
  "Mã AE",
  "Mã ae",
  "CCCD",
  "National ID",
  "Số CCCD",
  "Số CCCD Instructor",
  "CMND",
  "employeeId",
  "employee_id",
  "ma_nv",
  "Mã NV",
];
const NAME_KEYS = [
  "Full name",
  "Full Name",
  "FULL NAME",
  "Beneficiary Name",
  "Họ tên",
  "Họ và tên",
  "Tên nhân viên",
  "Họ và tên Instructor",
  "fullName",
  "full_name",
  "employeeName",
  "employee_name",
];
const ACCOUNT_KEYS = [
  "Bank Account Number",
  "BANK ACCOUNT NUMBER",
  "Beneficiary Account No.",
  "STK AE",
  "STK",
  "Số tài khoản",
  "Account Number",
  "bankAccountNumber",
  "bank_account_number",
  "accountNumber",
  "account_number",
];

function readField(row: any, keys: string[]): { key: string; value: unknown } {
  for (const key of keys) {
    const value = row?.[key];
    if (value !== undefined && value !== null && String(value).trim() !== "") {
      return { key, value };
    }
  }
  return { key: keys[0], value: "" };
}

function readWritableField(
  row: any,
  keys: string[],
): { key: string; value: unknown } {
  const populated = readField(row, keys);
  if (String(populated.value ?? "").trim()) return populated;
  const existingKey = keys.find((key) =>
    Object.prototype.hasOwnProperty.call(row || {}, key),
  );
  return existingKey
    ? { key: existingKey, value: row?.[existingKey] }
    : populated;
}

function normalizeCompact(value: unknown): string {
  return String(value ?? "")
    .replace(/\s+/g, "")
    .trim()
    .toUpperCase();
}

function normalizeName(value: unknown): string {
  return removeVietnameseTones(String(value ?? ""))
    .replace(/\s+/g, " ")
    .trim()
    .toUpperCase();
}

function identityOf(row: any): ReferenceIdentity {
  const id = readField(row, ID_KEYS);
  const name = readField(row, NAME_KEYS);
  const account = readField(row, ACCOUNT_KEYS);
  return {
    id: normalizeCompact(id.value),
    name: normalizeName(name.value),
    account: normalizeCompact(account.value),
    rawId: id.value,
    rawName: name.value,
    rawAccount: account.value,
  };
}

function rowMatchesMonth(row: any, reportMonth?: string): boolean {
  if (!reportMonth) return true;
  const target = parseMonthPeriod(reportMonth);
  if (!target) return true;
  const raw = String(
    row?.["Tháng báo cáo"] ??
      row?._fileMonth ??
      row?.["Tháng"] ??
      row?.Month ??
      "",
  ).trim();
  if (!raw) return true;
  return parseMonthPeriod(raw, target)?.key === target.key;
}

function transactionKey(row: any, index: number): string {
  const identity = identityOf(row);
  const serial = String(row?.["Payment Serial Number"] ?? "").trim();
  const explicitId = String(row?.id ?? row?._id ?? "").trim();
  return [
    explicitId || serial || index,
    identity.id,
    identity.name,
    identity.account,
  ].join("|");
}

function targetField(
  row: any,
  field: TransactionReferenceField,
): { key: string; value: unknown } {
  if (field === "idNumber") return readField(row, ID_KEYS);
  if (field === "fullName") return readField(row, NAME_KEYS);
  return readField(row, ACCOUNT_KEYS);
}

function writableTargetField(
  row: any,
  field: TransactionReferenceField,
): { key: string; value: unknown } {
  if (field === "idNumber") return readWritableField(row, ID_KEYS);
  if (field === "fullName") return readWritableField(row, NAME_KEYS);
  return readWritableField(row, ACCOUNT_KEYS);
}

function fieldDiffers(
  field: TransactionReferenceField,
  currentValue: unknown,
  referenceValue: unknown,
): boolean {
  if (referenceValue === undefined || referenceValue === null) return false;
  if (String(referenceValue).trim() === "") return false;
  return field === "fullName"
    ? normalizeName(currentValue) !== normalizeName(referenceValue)
    : normalizeCompact(currentValue) !== normalizeCompact(referenceValue);
}

function pairKey(identity: ReferenceIdentity): string {
  return identity.name && identity.account
    ? `${identity.name}|${identity.account}`
    : "";
}

function identityKey(identity: ReferenceIdentity): string {
  return `${identity.id}|${identity.name}|${identity.account}`;
}

function identityCompleteness(identity: ReferenceIdentity): number {
  return [identity.id, identity.name, identity.account].filter(Boolean).length;
}

function countIdentityValues(
  identities: ReferenceIdentity[],
  field: "id" | "account",
): Map<string, number> {
  const counts = new Map<string, number>();
  identities.forEach((identity) => {
    const value = identity[field];
    if (value) counts.set(value, (counts.get(value) || 0) + 1);
  });
  return counts;
}

function uniqueBestRawTimesheetReference(
  candidates: Array<{ row: any; identity: ReferenceIdentity }>,
): ReferenceIdentity | null {
  if (candidates.length === 0) return null;
  const maxCompleteness = Math.max(
    ...candidates.map(({ identity }) => identityCompleteness(identity)),
  );
  const bestByIdentity = new Map<string, ReferenceIdentity>();
  candidates.forEach(({ identity }) => {
    if (identityCompleteness(identity) !== maxCompleteness) return;
    bestByIdentity.set(identityKey(identity), identity);
  });
  return bestByIdentity.size === 1
    ? Array.from(bestByIdentity.values())[0]
    : null;
}

function resolveRawTimesheetReference(
  transactionIdentity: ReferenceIdentity,
  rawReferences: Array<{ row: any; identity: ReferenceIdentity }>,
  duplicateId: boolean,
  duplicateAccount: boolean,
): ReferenceIdentity | null {
  const matchingFields = (candidate: ReferenceIdentity) =>
    [
      !duplicateId &&
        transactionIdentity.id &&
        candidate.id === transactionIdentity.id,
      transactionIdentity.name && candidate.name === transactionIdentity.name,
      !duplicateAccount &&
        transactionIdentity.account &&
        candidate.account === transactionIdentity.account,
    ].filter(Boolean).length;

  const twoFieldMatch = uniqueBestRawTimesheetReference(
    rawReferences.filter(({ identity }) => matchingFields(identity) >= 2),
  );
  if (twoFieldMatch) return twoFieldMatch;

  // A unique Full Name in RAWDATA_TIMESHEET is the deterministic tiebreaker
  // when Transaction reused an ID Number or bank account for another person.
  if (transactionIdentity.name) {
    const byName = uniqueBestRawTimesheetReference(
      rawReferences.filter(
        ({ identity }) => identity.name === transactionIdentity.name,
      ),
    );
    if (byName) return byName;
  }

  if (transactionIdentity.id && !duplicateId) {
    const byId = uniqueBestRawTimesheetReference(
      rawReferences.filter(({ identity }) => identity.id === transactionIdentity.id),
    );
    if (byId) return byId;
  }

  if (transactionIdentity.account && !duplicateAccount) {
    return uniqueBestRawTimesheetReference(
      rawReferences.filter(
        ({ identity }) => identity.account === transactionIdentity.account,
      ),
    );
  }
  return null;
}

function planTransactionRepairsFromRawTimesheet(
  transactionRows: any[],
  rawTimesheetRows: any[],
): {
  effectiveRows: any[];
  corrections: TransactionRawTimesheetCorrection[];
} {
  if (rawTimesheetRows.length === 0 || transactionRows.length === 0) {
    return { effectiveRows: [...transactionRows], corrections: [] };
  }

  const transactionIdentities = transactionRows.map(identityOf);
  const idCounts = countIdentityValues(transactionIdentities, "id");
  const accountCounts = countIdentityValues(transactionIdentities, "account");
  const rawReferences = rawTimesheetRows
    .filter(Boolean)
    .map((row) => ({ row, identity: identityOf(row) }))
    .filter(({ identity }) => identityCompleteness(identity) >= 2);
  const effectiveRows = [...transactionRows];
  const corrections: TransactionRawTimesheetCorrection[] = [];

  transactionRows.forEach((row, transactionIndex) => {
    const current = transactionIdentities[transactionIndex];
    const duplicateId = Boolean(current.id && (idCounts.get(current.id) || 0) > 1);
    const duplicateAccount = Boolean(
      current.account && (accountCounts.get(current.account) || 0) > 1,
    );
    const hasMissingIdentity = !current.id || !current.name || !current.account;
    if (!duplicateId && !duplicateAccount && !hasMissingIdentity) return;

    const reference = resolveRawTimesheetReference(
      current,
      rawReferences,
      duplicateId,
      duplicateAccount,
    );
    if (!reference) return;

    const referenceValues: Record<TransactionReferenceField, unknown> = {
      idNumber: reference.rawId,
      fullName: reference.rawName,
      bankAccountNumber: reference.rawAccount,
    };
    let nextRow = row;
    (["idNumber", "fullName", "bankAccountNumber"] as const).forEach(
      (field) => {
        const currentField = writableTargetField(nextRow, field);
        if (isSavedTransactionField(nextRow, currentField.key)) return;
        const nextValue = referenceValues[field];
        if (!fieldDiffers(field, currentField.value, nextValue)) return;
        corrections.push({
          transactionIndex,
          field,
          fieldKey: currentField.key,
          oldValue: currentField.value,
          newValue: nextValue,
          source: "RAWDATA_TIMESHEET",
        });
        nextRow = { ...nextRow, [currentField.key]: nextValue };
      },
    );
    effectiveRows[transactionIndex] = nextRow;
  });

  return { effectiveRows, corrections };
}

function amountOf(row: any, table: TransactionReferenceTable): number {
  const value = parseMoneyToNumber(
    row?.["Payment Amount"] ??
      row?.["TOTAL PAYMENT"] ??
      row?.["Grand Total"] ??
      row?.["GRAND TOTAL"] ??
      row?.["Total Payment"] ??
      0,
  );
  if (table === "Sheet1_AE") return value;

  const operation = String(row?.["Nghiệp vụ"] ?? "").trim().toUpperCase();
  if (operation === "H" || operation.includes("HOLD")) return -Math.abs(value);
  if (operation === "C" || operation.includes("CANCEL")) return -Math.abs(value);
  if (operation === "A" || operation.includes("ADD")) return Math.abs(value);
  return value;
}

function buildCorrections(
  match: Pick<
    TransactionReferenceMatch,
    "transactionRow" | "grossRowIndexes" | "deductionRowIndexes"
  >,
  grossRows: any[],
  deductionRows: any[],
): TransactionReferenceCorrection[] {
  const reference = identityOf(match.transactionRow);
  const values: Record<TransactionReferenceField, unknown> = {
    idNumber: reference.rawId,
    fullName: reference.rawName,
    bankAccountNumber: reference.rawAccount,
  };
  const corrections: TransactionReferenceCorrection[] = [];

  const addForRows = (
    table: TransactionReferenceTable,
    rows: any[],
    indexes: number[],
  ) => {
    indexes.forEach((rowIndex) => {
      const row = rows[rowIndex];
      (["idNumber", "fullName", "bankAccountNumber"] as const).forEach(
        (field) => {
          const current = targetField(row, field);
          const nextValue = values[field];
          if (!fieldDiffers(field, current.value, nextValue)) return;
          corrections.push({
            table,
            rowIndex,
            field,
            fieldKey: current.key,
            oldValue: current.value,
            newValue: nextValue,
          });
        },
      );
    });
  };

  addForRows("Sheet1_AE", grossRows, match.grossRowIndexes);
  addForRows("Hold_AE", deductionRows, match.deductionRowIndexes);
  return corrections;
}

export function buildTransactionReferenceSyncPlan({
  grossRows,
  deductionRows,
  transactionRows,
  rawTimesheetRows = [],
  reportMonth,
}: {
  grossRows: any[];
  deductionRows: any[];
  transactionRows: any[];
  rawTimesheetRows?: any[];
  reportMonth?: string;
}): TransactionReferencePlan {
  const transactionRepairPlan = planTransactionRepairsFromRawTimesheet(
    transactionRows,
    rawTimesheetRows,
  );
  const targets = [
    ...grossRows.map((row, rowIndex) => ({
      table: "Sheet1_AE" as const,
      row,
      rowIndex,
      identity: identityOf(row),
    })),
    ...deductionRows.map((row, rowIndex) => ({
      table: "Hold_AE" as const,
      row,
      rowIndex,
      identity: identityOf(row),
    })),
  ].filter(({ row }) => rowMatchesMonth(row, reportMonth));

  const transactions = transactionRepairPlan.effectiveRows
    .map((row, transactionIndex) => ({
      row,
      transactionIndex,
      transactionKey: transactionKey(
        transactionRows[transactionIndex],
        transactionIndex,
      ),
      identity: identityOf(row),
      originalIdentity: identityOf(transactionRows[transactionIndex]),
    }))
    .filter(({ row }) => rowMatchesMonth(row, reportMonth));

  const transactionIdentityCount = new Map<string, number>();
  const transactionPairIdentityCount = new Map<string, Set<string>>();
  const transactionIdIdentityCount = new Map<string, Set<string>>();
  const transactionNameIdentityCount = new Map<string, Set<string>>();
  transactions.forEach(({ identity }) => {
    const fullKey = identityKey(identity);
    transactionIdentityCount.set(
      fullKey,
      (transactionIdentityCount.get(fullKey) || 0) + 1,
    );
    const pair = pairKey(identity);
    if (pair) {
      if (!transactionPairIdentityCount.has(pair)) {
        transactionPairIdentityCount.set(pair, new Set());
      }
      transactionPairIdentityCount.get(pair)!.add(fullKey);
    }
    if (identity.id) {
      if (!transactionIdIdentityCount.has(identity.id)) {
        transactionIdIdentityCount.set(identity.id, new Set());
      }
      transactionIdIdentityCount.get(identity.id)!.add(fullKey);
    }
    if (identity.name) {
      if (!transactionNameIdentityCount.has(identity.name)) {
        transactionNameIdentityCount.set(identity.name, new Set());
      }
      transactionNameIdentityCount.get(identity.name)!.add(fullKey);
    }
  });

  const assignedTargets = new Set<string>();
  const draftMatches = transactions.map((transaction) => ({
    ...transaction,
    targetIndexes: [] as number[],
    reason: "unmatched" as TransactionReferenceMatch["reason"],
  }));
  const targetToken = (targetIndex: number) => String(targetIndex);

  // Name + bank account is the strongest discriminator for the reported case:
  // multiple Gross Pay rows can carry different IDs while belonging to one
  // actual Transaction beneficiary. Resolve these before considering ID-only
  // matches so the incorrect ID cannot steal the row from the right person.
  draftMatches.forEach((match) => {
    const pair = pairKey(match.identity);
    if (!pair || transactionPairIdentityCount.get(pair)?.size !== 1) return;
    targets.forEach((target, targetIndex) => {
      if (assignedTargets.has(targetToken(targetIndex))) return;
      if (pairKey(target.identity) !== pair) return;
      match.targetIndexes.push(targetIndex);
      assignedTargets.add(targetToken(targetIndex));
    });
    if (match.targetIndexes.length > 0) match.reason = "name-account";
  });

  // Fallback for a wrong/missing name or account: a unique Transaction ID can
  // authoritatively repair the other common fields.
  draftMatches.forEach((match) => {
    if (!match.identity.id) return;
    if (transactionIdIdentityCount.get(match.identity.id)?.size !== 1) return;
    targets.forEach((target, targetIndex) => {
      if (assignedTargets.has(targetToken(targetIndex))) return;
      if (target.identity.id !== match.identity.id) return;
      match.targetIndexes.push(targetIndex);
      assignedTargets.add(targetToken(targetIndex));
    });
    if (match.reason === "unmatched" && match.targetIndexes.length > 0) {
      match.reason = "id";
    }
  });

  // When RAWDATA_TIMESHEET repaired duplicate Transaction keys, use the
  // unique beneficiary name to connect the still-unrepaired source row in the
  // same click. Without this bridge the user would need a second sync pass.
  const transactionIndexesRepairedFromRaw = new Set(
    transactionRepairPlan.corrections.map(
      (correction) => correction.transactionIndex,
    ),
  );
  draftMatches.forEach((match) => {
    if (!transactionIndexesRepairedFromRaw.has(match.transactionIndex)) return;
    if (
      !match.identity.name ||
      transactionNameIdentityCount.get(match.identity.name)?.size !== 1
    ) {
      return;
    }
    targets.forEach((target, targetIndex) => {
      if (assignedTargets.has(targetToken(targetIndex))) return;
      const sameName = target.identity.name === match.identity.name;
      const matchesOriginalIdentity = [
        match.originalIdentity.id &&
          match.originalIdentity.id === target.identity.id,
        match.originalIdentity.name &&
          match.originalIdentity.name === target.identity.name,
        match.originalIdentity.account &&
          match.originalIdentity.account === target.identity.account,
      ].filter(Boolean).length >= 2;
      if (!sameName && !matchesOriginalIdentity) return;
      match.targetIndexes.push(targetIndex);
      assignedTargets.add(targetToken(targetIndex));
    });
    if (match.reason === "unmatched" && match.targetIndexes.length > 0) {
      match.reason = "raw-timesheet";
    }
  });

  // Last deterministic fallback: any two matching identity fields.
  draftMatches.forEach((match) => {
    targets.forEach((target, targetIndex) => {
      if (assignedTargets.has(targetToken(targetIndex))) return;
      const fieldsMatched = [
        match.identity.id && match.identity.id === target.identity.id,
        match.identity.name && match.identity.name === target.identity.name,
        match.identity.account &&
          match.identity.account === target.identity.account,
      ].filter(Boolean).length;
      if (fieldsMatched < 2) return;
      match.targetIndexes.push(targetIndex);
      assignedTargets.add(targetToken(targetIndex));
    });
    if (match.reason === "unmatched" && match.targetIndexes.length > 0) {
      match.reason = "two-fields";
    }
  });

  const matches = draftMatches.map((draft) => {
    const grossRowIndexes: number[] = [];
    const deductionRowIndexes: number[] = [];
    draft.targetIndexes.forEach((targetIndex) => {
      const target = targets[targetIndex];
      if (target.table === "Sheet1_AE") grossRowIndexes.push(target.rowIndex);
      else deductionRowIndexes.push(target.rowIndex);
    });

    const withoutCorrections: Omit<
      TransactionReferenceMatch,
      "corrections" | "transactionCorrections"
    > = {
      transactionKey: draft.transactionKey,
      transactionIndex: draft.transactionIndex,
      transactionRow: draft.row,
      grossRowIndexes,
      deductionRowIndexes,
      reason: draft.reason,
    };
    return {
      ...withoutCorrections,
      transactionCorrections: transactionRepairPlan.corrections.filter(
        (correction) => correction.transactionIndex === draft.transactionIndex,
      ),
      corrections: buildCorrections(
        withoutCorrections,
        grossRows,
        deductionRows,
      ),
    };
  });

  return {
    matches,
    byTransactionKey: new Map(matches.map((match) => [match.transactionKey, match])),
    byTransactionIndex: new Map(
      matches.map((match) => [match.transactionIndex, match]),
    ),
    effectiveTransactionRows: transactionRepairPlan.effectiveRows,
    transactionCorrections: transactionRepairPlan.corrections,
  };
}

export function getTransactionReferenceMatchAmounts(
  match: TransactionReferenceMatch | undefined,
  grossRows: any[],
  deductionRows: any[],
): { grossAmount: number; deductionAmount: number; expectedAmount: number } {
  if (!match) return { grossAmount: 0, deductionAmount: 0, expectedAmount: 0 };
  const grossAmount = match.grossRowIndexes.reduce(
    (sum, index) => sum + amountOf(grossRows[index], "Sheet1_AE"),
    0,
  );
  const deductionAmount = match.deductionRowIndexes.reduce(
    (sum, index) => sum + amountOf(deductionRows[index], "Hold_AE"),
    0,
  );
  return {
    grossAmount,
    deductionAmount,
    expectedAmount: grossAmount + deductionAmount,
  };
}

export function applyTransactionReferenceSync({
  grossRows,
  deductionRows,
  transactionRows,
  rawTimesheetRows = [],
  reportMonth,
  transactionKeys,
  targetTable,
  correctedAt = new Date().toISOString(),
  lockedDeductionMonths,
}: {
  grossRows: any[];
  deductionRows: any[];
  transactionRows: any[];
  rawTimesheetRows?: any[];
  reportMonth?: string;
  transactionKeys?: Iterable<string>;
  targetTable?: TransactionReferenceTable;
  correctedAt?: string;
  lockedDeductionMonths?: AppData['HoldCarrySnapshots'];
}): {
  grossRows: any[];
  deductionRows: any[];
  transactionRows: any[];
  correctedCells: number;
  correctedRows: number;
  transactionCorrectedCells: number;
  transactionCorrectedRows: number;
  appliedTransactionKeys: string[];
} {
  const plan = buildTransactionReferenceSyncPlan({
    grossRows,
    deductionRows,
    transactionRows,
    // Scoped actions are downstream-only. When targetTable is set,
    // Transaction/Batch Payment remains the authoritative source and must not
    // be repaired from RAWDATA_TIMESHEET as a side effect.
    rawTimesheetRows: targetTable ? [] : rawTimesheetRows,
    reportMonth,
  });
  const allowedKeys = transactionKeys ? new Set(transactionKeys) : null;
  const selectedMatches = plan.matches.map(match => ({
    ...match,
    corrections: match.corrections.filter(c => (!targetTable || c.table === targetTable) &&
      (c.table !== 'Hold_AE' || !isDeductionsMonthLocked(lockedDeductionMonths,
        deductionsRowMonth(deductionRows[c.rowIndex], reportMonth)))),
    // targetTable is a hard mutation boundary: scoped actions may only write
    // the requested downstream table, never Transaction/Batch Payment.
    transactionCorrections: targetTable ? [] : match.transactionCorrections,
  })).filter(
    (match) =>
      (match.corrections.length > 0 ||
        match.transactionCorrections.length > 0) &&
      (!allowedKeys || allowedKeys.has(match.transactionKey)),
  );
  const nextGross = [...grossRows];
  const nextDeductions = [...deductionRows];
  const nextTransactions = [...transactionRows];
  const correctedRowTokens = new Set<string>();
  const correctedTransactionRows = new Set<number>();
  let correctedCells = 0;
  let transactionCorrectedCells = 0;

  selectedMatches.forEach((match) => {
    const reference = identityOf(match.transactionRow);
    match.transactionCorrections.forEach((correction) => {
      const original = nextTransactions[correction.transactionIndex];
      if (!original) return;
      nextTransactions[correction.transactionIndex] = {
        ...original,
        [correction.fieldKey]: correction.newValue,
        _rawTimesheetReferenceAudit: {
          ...(original._rawTimesheetReferenceAudit || {}),
          [correction.field]: {
            ...correction,
            correctedAt,
          },
        },
      };
      correctedCells += 1;
      transactionCorrectedCells += 1;
      correctedTransactionRows.add(correction.transactionIndex);
      correctedRowTokens.add(`BankExport:${correction.transactionIndex}`);
    });
    match.corrections.forEach((correction) => {
      const rows = correction.table === "Sheet1_AE" ? nextGross : nextDeductions;
      const original = rows[correction.rowIndex];
      if (!original) return;
      const auditEntry: TransactionReferenceAuditEntry = {
        field: correction.field,
        fieldLabel:
          correction.field === "idNumber"
            ? "ID NUMBER"
            : correction.field === "fullName"
              ? "FULL NAME"
              : "BANK ACCOUNT NUMBER",
        oldValue: correction.oldValue,
        newValue: correction.newValue,
        correctedAt,
        targetTable: correction.table,
        transactionKey: match.transactionKey,
        transactionIndex: match.transactionIndex,
        transactionRowId: String(
          match.transactionRow?.id ?? match.transactionRow?._id ?? "",
        ),
        transactionSerial: String(
          match.transactionRow?.["Payment Serial Number"] ??
            match.transactionIndex + 1,
        ),
        transactionId: String(reference.rawId ?? ""),
        transactionName: String(reference.rawName ?? ""),
        transactionAccount: String(reference.rawAccount ?? ""),
      };
      rows[correction.rowIndex] = {
        ...original,
        [correction.fieldKey]: correction.newValue,
        [TRANSACTION_REFERENCE_AUDIT_KEY]: {
          ...(original[TRANSACTION_REFERENCE_AUDIT_KEY] || {}),
          [correction.field]: auditEntry,
        },
      };
      correctedCells += 1;
      correctedRowTokens.add(`${correction.table}:${correction.rowIndex}`);
    });
  });

  return {
    grossRows: nextGross,
    deductionRows: nextDeductions,
    transactionRows: nextTransactions,
    correctedCells,
    correctedRows: correctedRowTokens.size,
    transactionCorrectedCells,
    transactionCorrectedRows: correctedTransactionRows.size,
    appliedTransactionKeys: selectedMatches.map(
      (match) => match.transactionKey,
    ),
  };
}

export interface BulkReconciliationReferenceSyncResult {
  grossRows: any[];
  deductionRows: any[];
  correctedCells: number;
  correctedRows: number;
  pendingTransactionKeys: string[];
}

export function getPendingReconciliationReferenceSyncKeys({
  grossRows,
  deductionRows,
  transactionRows,
  reportMonth,
  lockedDeductionMonths,
}: {
  grossRows: any[];
  deductionRows: any[];
  transactionRows: any[];
  reportMonth?: string;
  lockedDeductionMonths?: AppData['HoldCarrySnapshots'];
}): string[] {
  const plan = buildTransactionReferenceSyncPlan({
    grossRows,
    deductionRows,
    transactionRows,
    // Reconciliation sync always treats confirmed Batch Payment as source.
    rawTimesheetRows: [],
    reportMonth,
  });
  return Array.from(new Set(
    plan.matches
      .filter((match) => match.corrections.some(c => c.table !== 'Hold_AE' ||
        !isDeductionsMonthLocked(lockedDeductionMonths, deductionsRowMonth(deductionRows[c.rowIndex], reportMonth))))
      .map((match) => match.transactionKey),
  ));
}

export function hasPendingReconciliationReferenceSync(args: {
  grossRows: any[];
  deductionRows: any[];
  transactionRows: any[];
  reportMonth?: string;
  lockedDeductionMonths?: AppData['HoldCarrySnapshots'];
}): boolean {
  return getPendingReconciliationReferenceSyncKeys(args).length > 0;
}

/**
 * Canonical all-row Reconciliation sync.
 * This intentionally repeats the same per-Transaction-key Process Sync against
 * the evolving Gross Pay / Deductions state so every UI entry point behaves
 * identically and resolved rows disappear after either action is used.
 */
export function applyBulkReconciliationReferenceSync({
  grossRows,
  deductionRows,
  transactionRows,
  reportMonth,
  lockedDeductionMonths,
}: {
  grossRows: any[];
  deductionRows: any[];
  transactionRows: any[];
  reportMonth?: string;
  lockedDeductionMonths?: AppData['HoldCarrySnapshots'];
}): BulkReconciliationReferenceSyncResult {
  const pendingTransactionKeys = getPendingReconciliationReferenceSyncKeys({
    grossRows,
    deductionRows,
    transactionRows,
    reportMonth,
    lockedDeductionMonths,
  });

  let nextGrossRows = grossRows;
  let nextDeductionRows = deductionRows;
  let correctedCells = 0;
  let correctedRows = 0;

  for (const transactionKey of pendingTransactionKeys) {
    const result = applyTransactionReferenceSync({
      grossRows: nextGrossRows,
      deductionRows: nextDeductionRows,
      transactionRows,
      rawTimesheetRows: [],
      reportMonth,
      transactionKeys: [transactionKey],
      lockedDeductionMonths,
    });
    if (result.correctedCells === 0) continue;
    nextGrossRows = result.grossRows;
    nextDeductionRows = result.deductionRows;
    correctedCells += result.correctedCells;
    correctedRows += result.correctedRows;
  }

  return {
    grossRows: nextGrossRows,
    deductionRows: nextDeductionRows,
    correctedCells,
    correctedRows,
    pendingTransactionKeys,
  };
}

export function getTransactionReferenceAudit(
  row: any,
  field: TransactionReferenceField,
): TransactionReferenceAuditEntry | undefined {
  return row?.[TRANSACTION_REFERENCE_AUDIT_KEY]?.[field];
}

export function getTransactionReferenceField(
  header: string,
): TransactionReferenceField | null {
  const normalized = String(header || "").trim().toUpperCase();
  if (
    normalized === "ID NUMBER" ||
    normalized === "DOCUMENT ID" ||
    normalized === "CCCD" ||
    normalized === "MÃ AE"
  ) {
    return "idNumber";
  }
  if (
    normalized === "FULL NAME" ||
    normalized === "BENEFICIARY NAME" ||
    normalized === "HỌ TÊN"
  ) {
    return "fullName";
  }
  if (
    normalized === "BANK ACCOUNT NUMBER" ||
    normalized === "BENEFICIARY ACCOUNT NO." ||
    normalized === "STK AE" ||
    normalized === "STK" ||
    normalized === "SỐ TÀI KHOẢN"
  ) {
    return "bankAccountNumber";
  }
  return null;
}
