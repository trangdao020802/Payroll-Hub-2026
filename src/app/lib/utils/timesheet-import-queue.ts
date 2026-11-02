// One fetch/parse/import at a time across Settings, folder imports and table refresh.
let tail: Promise<void> = Promise.resolve();
let pending = 0;
const listeners = new Set<() => void>();

export const isTimesheetImporting = () => pending > 0;
export const subscribeTimesheetImports = (listener: () => void) => {
  listeners.add(listener);
  return () => { listeners.delete(listener); };
};
const emit = () => listeners.forEach((listener) => listener());

// A folder holds calculation suspension across sequential imports without
// entering the serial queue itself (which would deadlock its child imports).
export async function withTimesheetImportBatch<T>(task: () => Promise<T>): Promise<T> {
  pending++;
  if (pending === 1) emit();
  try {
    return await task();
  } finally {
    pending--;
    if (pending === 0) emit();
  }
}

export function runTimesheetImport<T>(task: () => Promise<T>): Promise<T> {
  pending++;
  if (pending === 1) emit();
  const result = tail.then(task);
  tail = result.then(() => undefined, () => undefined);
  return result.finally(() => {
    pending--;
    if (pending === 0) emit();
  });
}
