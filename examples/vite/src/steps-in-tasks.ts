// The demo's loop for the work that it does itself on a workbook: the check before an open. The
// store gives this work as steps (`Generator<void, T>`). The loop runs the steps for a few
// milliseconds at a time, each slice in its own task, so the page stays responsive during a large
// workbook. The editor runs its own load and save in short tasks.

/** The work of one slice, in milliseconds. */
const SLICE_MS = 12;

/** Queue `run` as a new task. A message port does not clamp nested calls to 4 ms, as a timer does. */
function queueTask(run: () => void): void {
  const channel = new MessageChannel();
  channel.port1.onmessage = () => {
    channel.port1.close();
    run();
  };
  channel.port2.postMessage(null);
}

/** Run `steps` to the end in short tasks. The promise rejects with what a step throws. */
export function runInTasks<T>(steps: Generator<void, T>): Promise<T> {
  return new Promise<T>((resolve, reject) => {
    const slice = (): void => {
      const deadline = performance.now() + SLICE_MS;
      try {
        for (;;) {
          const next = steps.next();
          if (next.done) {
            resolve(next.value);
            return;
          }
          if (performance.now() >= deadline) break;
        }
      } catch (error) {
        reject(error);
        return;
      }
      queueTask(slice);
    };
    slice();
  });
}

/** `steps`, which end with null at the first yield after `wanted` turns false. */
export function* whileWanted<T>(
  steps: Generator<void, T>,
  wanted: () => boolean
): Generator<void, T | null> {
  for (;;) {
    const next = steps.next();
    if (next.done) return next.value;
    if (!wanted()) {
      steps.return(null as T);
      return null;
    }
    yield;
  }
}
