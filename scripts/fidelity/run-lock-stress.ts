// A worker of the stress test of the run lock (`run-lock.test.ts`). It takes and releases the
// lock of the folder `dir` again and again for `ms` milliseconds, and prints its counts as JSON:
// - `got`: the times that it held the lock;
// - `double`: the times that another worker held the lock at the same time, seen as a marker
//   file that only a holder makes and removes, and that names a running worker. A marker of a
//   worker that was killed while it held the lock is no double hold;
// - `crashes`: the errors that the lock threw;
// - `held`: the times that it found the lock held.
// With `die` set to a number, the worker exits while it holds the lock after that many holds, as
// a killed run does, so that the other workers take a stale lock over.
//
//   bun scripts/fidelity/run-lock-stress.ts <dir> <ms> [die]

import { closeSync, openSync, readFileSync, rmSync, writeFileSync, writeSync } from 'node:fs';
import { join } from 'node:path';
import { processRuns, takeRunLock } from './run-lock.ts';

const [dir, ms, die] = process.argv.slice(2);
if (!dir || !ms) {
  console.error('usage: bun scripts/fidelity/run-lock-stress.ts <dir> <ms> [die]');
  process.exit(2);
}
const marker = join(dir, 'holder');
const counts = { got: 0, double: 0, crashes: 0, held: 0 };
const end = Date.now() + Number(ms);
while (Date.now() < end) {
  let lock: ReturnType<typeof takeRunLock>;
  try {
    lock = takeRunLock(dir, process.pid);
  } catch {
    counts.crashes += 1;
    continue;
  }
  if ('held' in lock) {
    counts.held += 1;
    continue;
  }
  counts.got += 1;
  let mine = true;
  try {
    const fd = openSync(marker, 'wx');
    writeSync(fd, String(process.pid));
    closeSync(fd);
  } catch {
    const other = Number(readFileSync(marker, 'utf8'));
    if (other > 0 && other !== process.pid && processRuns(other)) {
      counts.double += 1;
      mine = false;
    } else writeFileSync(marker, String(process.pid));
  }
  for (let i = 0; i < 2000; i += 1);
  if (mine) rmSync(marker, { force: true });
  if (die !== undefined && counts.got >= Number(die)) {
    console.log(JSON.stringify(counts));
    process.exit(0);
  }
  try {
    lock.release();
  } catch {
    counts.crashes += 1;
  }
}
console.log(JSON.stringify(counts));
