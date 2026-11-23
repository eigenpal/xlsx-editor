// A stand-in child for `runner.test.ts`. It runs the job loop of `worker.ts` with a stand-in
// task, and each file names what the task does: `ok`, `slow` (takes 1.5 s,
// and writes its process id to the file that `CORPUS_TEST_PIDS` names, if any), `allocate` (grows
// its memory until the watchdog stops it), `allocate-later` (writes its process id and work
// folder to that file, and grows its memory when the file of that name with `.go` appears),
// `hang`, `throw` (fails with a path in its message),
// `exit` (ends in the middle of its batch), `flood` (writes a line longer than the runner takes),
// `grandchild` (starts a process, writes a work file, writes both process ids and the work
// folder to the file that `CORPUS_TEST_PIDS` names, and hangs), `grandchild-locked` (the same,
// and also makes a subfolder that the runner cannot remove), `orphan` (starts a process
// through a shell that ends at once, so the process stays in the group of the child but leaves
// its tree, writes its id to that file, and hangs), `keep <n>` (writes its process id and its
// name to that file, and keeps 64 MiB until the child ends),
// `after-runner` (writes its process id and work folder to that file, and returns as soon as
// the runner is gone), `forge-id` (returns a result with another `id`), `runner-env` (returns
// the runner id and the guard limit that its environment names), or `refuse-start` (ends before
// its first workbook).
// It never runs outside the tests.

import { spawn, spawnSync } from 'node:child_process';
import { randomFillSync } from 'node:crypto';
import {
  appendFileSync,
  chmodSync,
  existsSync,
  mkdirSync,
  readFileSync,
  writeFileSync,
} from 'node:fs';
import { join } from 'node:path';
import { CORPUS_GUARD_ENV, RUNNER_ENV, taskTempRoot, type CorpusJob } from './runner.ts';
import { runJobs } from './worker.ts';

/** The ceiling of the allocation, so a broken watchdog cannot take the machine. */
const MAX_ALLOCATION = 768 * 1024 * 1024;
const CHUNK = 16 * 1024 * 1024;

const jobsFile = process.argv[3]!;
const jobs = JSON.parse(readFileSync(jobsFile, 'utf8')) as CorpusJob[];
if (jobs.some((job) => readFileSync(job.path, 'utf8').trim() === 'refuse-start')) process.exit(2);
const held: Buffer[] = [];

await runJobs(async (bytes) => {
  const behavior = new TextDecoder().decode(bytes).trim();
  if (behavior === 'slow') {
    if (process.env.CORPUS_TEST_PIDS)
      appendFileSync(process.env.CORPUS_TEST_PIDS, `${process.pid}\n`);
    await Bun.sleep(1500);
  }
  if (behavior === 'allocate-later') {
    // Starts to allocate when the test has paused the runner.
    const pids = process.env.CORPUS_TEST_PIDS!;
    appendFileSync(pids, `${process.pid} 0 ${taskTempRoot()}\n`);
    while (!existsSync(`${pids}.go`)) await Bun.sleep(10);
  }
  if (behavior.startsWith('allocate')) {
    // Random bytes, so the system cannot compress the pages and the process table sees them.
    while (held.length * CHUNK < MAX_ALLOCATION) {
      held.push(randomFillSync(Buffer.allocUnsafe(CHUNK)));
      await Bun.sleep(5);
    }
    await Bun.sleep(60_000);
  }
  if (behavior.startsWith('keep')) {
    appendFileSync(process.env.CORPUS_TEST_PIDS!, `${process.pid} ${behavior}\n`);
    for (let i = 0; i < 4; i += 1) held.push(randomFillSync(Buffer.allocUnsafe(CHUNK)));
    // Long enough for the watchdog to read the memory while this workbook runs.
    await Bun.sleep(300);
  }
  if (behavior === 'forge-id') return { id: 'workbook9', behavior };
  if (behavior === 'runner-env')
    return { behavior, runner: process.env[RUNNER_ENV], guard: process.env[CORPUS_GUARD_ENV] };
  if (behavior === 'hang') await Bun.sleep(60_000);
  if (behavior === 'after-runner') {
    // Returns its result as soon as the runner is gone, before the guard thread looks.
    const parent = process.ppid;
    appendFileSync(process.env.CORPUS_TEST_PIDS!, `${process.pid} 0 ${taskTempRoot()}\n`);
    while (process.ppid === parent) await Bun.sleep(1);
  }
  if (behavior === 'throw') throw new Error('cannot read /Users/someone/corpus/Secret plan.xlsx');
  if (behavior === 'exit') process.exit(3);
  if (behavior === 'flood') {
    process.stdout.write('x'.repeat(5 * 1024 * 1024));
    await Bun.sleep(60_000);
  }
  if (behavior === 'grandchild' || behavior === 'grandchild-locked') {
    const sleeper = spawn('sleep', ['60'], { stdio: 'ignore' });
    writeFileSync(join(taskTempRoot(), 'part.xml'), '<private/>');
    if (behavior === 'grandchild-locked') {
      // A folder with no write permission keeps its file, so the removal of the folder fails.
      const locked = join(taskTempRoot(), 'locked');
      mkdirSync(locked);
      writeFileSync(join(locked, 'part.xml'), '<private/>');
      chmodSync(locked, 0o500);
    }
    const line = `${process.pid} ${sleeper.pid} ${taskTempRoot()}\n`;
    appendFileSync(process.env.CORPUS_TEST_PIDS!, line);
    await Bun.sleep(60_000);
  }
  if (behavior === 'orphan') {
    // A shell with no job control keeps its background process in the group of the child.
    const shell = spawnSync('sh', ['-c', 'sleep 60 >/dev/null 2>&1 & echo $!'], {
      encoding: 'utf8',
    });
    appendFileSync(process.env.CORPUS_TEST_PIDS!, `${shell.stdout.trim()}\n`);
    await Bun.sleep(60_000);
  }
  return { behavior, child: process.pid };
}, jobsFile);
