// A worker thread of the corpus child (`worker.ts`). It stops the child when the runner ends.
//
// The runner stops the child itself after a signal. After a SIGKILL or a crash of the runner,
// nothing outside the child is left to stop it or to remove its work folder, and the child
// gets another parent. The main thread of the child can be busy in one workbook for minutes,
// so this check runs on its own thread. The child leads its own process group (the runner
// spawns it with `detached: true`), so the signal also stops the processes that the child
// started, such as xmllint.
//
// The guard also stops the child when its memory passes `guardBytes`: by default
// `GUARD_CAP_FACTOR` (in `runner.ts`) times the memory cap. The watchdog of the runner stops it at
// the cap. This second check acts only when the runner cannot watch: a runner that is paused by
// the system or blocked on a terminal write. The runner records a SIGKILL that its watchdog did
// not send as a memory stop.

import { workerData } from 'node:worker_threads';
import { endOrphan } from './orphan.ts';

/**
 * From `worker.ts`: the process id of the runner, which the runner gives in `RUNNER_ENV`, the
 * work folder of the child, which the runner gives in `CORPUS_TMP_ENV`, and the memory at which
 * the guard stops the child, which the runner gives in `CORPUS_GUARD_ENV`.
 */
const { runner, dir, guardBytes } = workerData as {
  runner: number;
  dir: string | undefined;
  guardBytes: number;
};

setInterval(() => {
  // The memory of the whole child process, also when this thread reads it.
  if (process.ppid !== runner || process.memoryUsage.rss() > guardBytes) endOrphan(dir);
}, 250);
