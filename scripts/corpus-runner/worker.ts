// The child of the corpus runner (`runner.ts`): runs one task over one batch of workbooks.
//
//   bun scripts/corpus-runner/worker.ts <task> <jobs.json>
//
// The runner starts it; nobody runs it by hand. It writes one JSON line to its standard output
// when it starts a workbook and one with the result, and nothing else. A result holds counts,
// codes, and schema messages without values, never a path, a file name, or workbook text.
// A thread (`parent-guard.ts`) stops the child and the processes that it started when the
// runner ends or cannot watch a child that grows far above the cap, also while one workbook
// keeps the main thread busy. A write to a runner that is gone does the same at once
// (`orphan.ts`), and so does a child whose runner ended during its last workbook. The child
// removes `jobs.json` when it has read it.

import { readFileSync, rmSync, statSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { Worker } from 'node:worker_threads';
import { endOrphan } from './orphan.ts';
import type { CorpusJob } from './runner.ts';
import {
  CORPUS_GUARD_ENV,
  CORPUS_TMP_ENV,
  GUARD_CAP_FACTOR,
  MEMORY_CAP_BYTES,
  redactPaths,
  RUNNER_ENV,
} from './runner.ts';

/** A task: the result of one workbook, from its bytes. */
export type CorpusTask = (bytes: Uint8Array) => object | Promise<object>;

/**
 * The largest file that the child reads. It is the default `zip.maxTotalBytes` of the reader, so
 * a larger file is refused before it is read into memory.
 */
export const MAX_FILE_BYTES = 512 * 1024 * 1024;

/** The tasks that the runner can start, by name. */
const TASKS: Record<string, () => Promise<CorpusTask>> = {
  census: async () => (await import('../fidelity/census.ts')).censusWorkbook,
  xsd: async () => (await import('../validate-xsd.ts')).validateCorpusWorkbook,
};

/**
 * A write to the runner fails, as an `error` event, only when the runner is gone. The child then
 * ends at once, as the guard thread does, and does not crash with its work folder left.
 */
const orphaned = () => endOrphan(process.env[CORPUS_TMP_ENV]);

const say = (line: Record<string, unknown>) => process.stdout.write(`${JSON.stringify(line)}\n`);

/**
 * Starts the parent guard, then runs `task` over the jobs of `jobsFile` and writes the protocol
 * lines. The test child (`test-worker.ts`) runs its stand-in task through this loop too.
 */
export async function runJobs(task: CorpusTask, jobsFile: string): Promise<void> {
  // The runner names itself and the work folder. `process.ppid` is the fallback for a child
  // started by hand.
  const runner = Number(process.env[RUNNER_ENV]) || process.ppid;
  const guard = new Worker(new URL('./parent-guard.ts', import.meta.url), {
    workerData: {
      runner,
      dir: process.env[CORPUS_TMP_ENV],
      guardBytes: Number(process.env[CORPUS_GUARD_ENV]) || MEMORY_CAP_BYTES * GUARD_CAP_FACTOR,
    },
  });
  guard.unref();
  process.stdout.on('error', orphaned);
  const jobs = JSON.parse(readFileSync(jobsFile, 'utf8')) as CorpusJob[];
  // The list names the workbooks: it goes at once from the work folder of the runner.
  const work = process.env[CORPUS_TMP_ENV];
  if (work && dirname(resolve(jobsFile)) === resolve(work)) rmSync(jobsFile, { force: true });
  for (const job of jobs) {
    say({ type: 'start', id: job.id });
    let result: object;
    try {
      const size = statSync(job.path).size;
      if (size > MAX_FILE_BYTES) result = { open: 'refused', refusal: 'file-too-large' };
      else {
        const file = readFileSync(job.path);
        result = await task(new Uint8Array(file.buffer, file.byteOffset, file.byteLength));
      }
    } catch (error) {
      result = { error: redactPaths(error instanceof Error ? error.message : String(error)) };
    }
    say({ type: 'result', id: job.id, result });
  }
  // A runner that ended during the last workbook gives no write error when a process that it
  // started still holds the read end of the pipe. The child would then end with its folder left.
  if (process.ppid !== runner) orphaned();
  await guard.terminate();
}

async function main(): Promise<void> {
  const [name, jobsFile] = process.argv.slice(2);
  const load = name ? TASKS[name] : undefined;
  if (!load || !jobsFile)
    throw new Error(`Unknown task. Known tasks: ${Object.keys(TASKS).join(', ')}.`);
  await runJobs(await load(), jobsFile);
}

if (import.meta.main) await main();
