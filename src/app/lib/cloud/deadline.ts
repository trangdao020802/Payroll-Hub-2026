// Race as well as abort: a pending auth lock or worker can otherwise ignore an
// AbortSignal and keep the sync dialog spinning forever.
export async function withDeadline<T>(
  operation: (signal: AbortSignal) => PromiseLike<T>,
  message: string,
  timeoutMs = 30000,
  sessionSignal?: AbortSignal,
): Promise<T> {
  const controller = new AbortController();
  let timer: ReturnType<typeof setTimeout> | undefined;
  let onSessionAbort: (() => void) | undefined;
  const timeout = new Promise<never>((_resolve, reject) => {
    const fail = (text: string) => {
      reject(new Error(text));
      controller.abort();
    };
    onSessionAbort = () => fail("Phiên đồng bộ đã đóng.");
    if (sessionSignal?.aborted) onSessionAbort();
    else sessionSignal?.addEventListener("abort", onSessionAbort, { once: true });
    timer = setTimeout(() => fail(message), timeoutMs);
  });
  try {
    return await Promise.race([Promise.resolve().then(() => {
      if (controller.signal.aborted) throw new Error("Phiên đồng bộ đã đóng.");
      return operation(controller.signal);
    }), timeout]);
  } finally {
    clearTimeout(timer);
    if (onSessionAbort) sessionSignal?.removeEventListener("abort", onSessionAbort);
  }
}
