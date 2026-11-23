// Listener errors: a listener that throws never stops the other listeners or the work of the
// code that notifies them. Each error goes to the host's error report, as an uncaught error
// would, but is not thrown to the code that notifies.

/**
 * Report an error that a listener threw, without throwing it to the caller: through
 * `reportError` when the host has it, else as an uncaught error in a microtask.
 *
 * @internal
 */
export function reportListenerError(error: unknown): void {
  const host = globalThis as {
    reportError?: (e: unknown) => void;
    queueMicrotask?: (task: () => void) => void;
  };
  if (typeof host.reportError === 'function') host.reportError(error);
  else if (typeof host.queueMicrotask === 'function')
    host.queueMicrotask(() => {
      throw error;
    });
  else void Promise.reject(error);
}

/**
 * Call each listener of a copy of `listeners`, also when one throws, as DOCX Editor's event
 * dispatch runs every queued notification. Each error goes to `reportListenerError` when its
 * listener throws, as the store reports the errors of its own listeners; none is thrown to the
 * caller, so the caller's own updates after the call still run.
 *
 * @internal
 */
export function runListeners(listeners: Iterable<() => void>): void {
  for (const listener of [...listeners]) {
    try {
      listener();
    } catch (error) {
      reportListenerError(error);
    }
  }
}
