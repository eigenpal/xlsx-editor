// The memory-safe corpus runner: runs one task over a folder of private workbooks, one child
// process at a time, a new child for each batch, with a watchdog outside the child.
//
// After pptx-editor/scripts/fidelity/corpus-list.ts at 49713b43 (EigenPal): opaque ids instead
// of file names, a `results.jsonl` line for each workbook, messages without paths, a run that
// refuses to replace its results unless asked. Differences, each for a reason:
// - PPTX Editor reads its corpus in the process of the command. A read of the private corpus in
//   one process grew to more than 10 GB and stopped the machine, so each batch of
//   `BATCH_SIZE` workbooks runs in a new child, and only one child runs at a time.
// - The watchdog of the hang test (`scripts/hang/watchdog.ts`) reads the process tree of the
//   child. It stops the tree above `MEMORY_CAP_BYTES` or when one workbook passes the time of
//   its task. The workbook that ran is recorded as stopped, with the cause. A memory stop after
//   other workbooks of the child first runs that workbook once more alone, in a new child. The
//   workbooks of the batch that did not start go to the next child, so one bad workbook stops
//   only itself.
// - The corpus folder comes from `--dir` or `XLSX_CORPUS_DIR`. The committed code names no
//   corpus location.
// - PPTX Editor takes a job list with ids from its caller. This runner lists the folder itself
//   and names the workbooks `workbook1` to `workbookN` by sorted path, so no file name reaches
//   the output. An id names another workbook when a file is added or removed before it.
// - The child leads its own process group, so a stop also reaches the processes that it started
//   (xmllint). A thread of the child (`parent-guard.ts`) stops the group when the runner ends,
//   or when the child grows well above the cap while the runner cannot watch it.
// - Ctrl+Z ends the running child and removes its folder before the runner pauses, so a paused
//   run holds no memory. After `fg`, a new child runs the workbooks that did not finish.
// - The runner owns one temporary folder for each child: its job list and the work files of the
//   task. The runner removes it when the child ends, also after a stop or a signal.
// - The output folder must be outside every git work tree or a folder that git ignores
//   (`local/` by default), so no result can be committed by mistake.
// - There is no `--resume` or `--shard` yet.
//
// The child (`worker.ts`) writes one JSON line to its standard output when it starts a workbook
// and one when it has the result. The runner checks each line, writes the result to
// `results.jsonl`, and keeps no workbook data in memory.

import { spawn, spawnSync, type ChildProcess } from 'node:child_process';
import {
  accessSync,
  appendFileSync,
  constants,
  existsSync,
  mkdirSync,
  mkdtempSync,
  readdirSync,
  rmSync,
  statSync,
  writeFileSync,
} from 'node:fs';
import { tmpdir } from 'node:os';
import { basename, dirname, join, relative, resolve } from 'node:path';
import {
  killGroup,
  MEMORY_CAP_BYTES,
  signalGroup,
  watchProcess,
  type Watch,
} from '../hang/watchdog.ts';

export { MEMORY_CAP_BYTES };

export const ROOT = resolve(import.meta.dir, '../..');

/** The file names of the workbooks that a corpus run opens: plain and macro-enabled forms. */
const WORKBOOK_FILE = /\.(xlsx|xltx|xlsm|xltm)$/i;

/** The workbooks of one child. A new child for each batch keeps a leak to one batch. */
export const BATCH_SIZE = 25;

/** The environment variable that names the corpus folder when `--dir` does not. */
export const CORPUS_ENV = 'XLSX_CORPUS_DIR';

/** The default output folder of a task, under `local/`, which git ignores. */
export const defaultOut = (task: string): string => join(ROOT, 'local', 'corpus', task);

/** The version of the `results.jsonl` and `run.json` format. */
export const RESULTS_SCHEMA = 1;

/** The longest line that the runner takes from a child. A longer line stops the child. */
export const MAX_LINE_BYTES = 4 * 1024 * 1024;

/** The child of a real run. Tests pass their own. */
export const WORKER = join(import.meta.dir, 'worker.ts');

/** The files that a corpus command writes to its output folder. Git must ignore each one. */
export const OUTPUT_FILES = ['results.jsonl', 'run.json', 'census.md'] as const;

/** The environment variable that names the temporary folder of the running child. */
export const CORPUS_TMP_ENV = 'XLSX_CORPUS_TMP';

/**
 * The environment variable that gives the child the process id of its runner. The guard thread
 * of the child reads it, so a runner that ends before the child starts is still seen.
 */
export const RUNNER_ENV = 'XLSX_CORPUS_RUNNER';

/**
 * The environment variable that gives the child the memory at which its guard thread stops it,
 * for a runner that cannot watch (see `parent-guard.ts`).
 */
export const CORPUS_GUARD_ENV = 'XLSX_CORPUS_GUARD_BYTES';

/** How far above the cap of the watchdog the guard thread of the child stops it. */
export const GUARD_CAP_FACTOR = 1.25;

/**
 * The folder for the work files of a task: the folder of the child, which the runner removes
 * when the child ends, or the system folder outside a run.
 */
export const taskTempRoot = (): string => process.env[CORPUS_TMP_ENV] ?? tmpdir();

/** One workbook of a run: an opaque id and the absolute path. The path never enters the output. */
export interface CorpusJob {
  readonly id: string;
  readonly path: string;
}

/** Why the runner stopped a workbook. */
export type StopCause = 'memory' | 'time' | 'exit' | 'output';

/** A workbook that has no result, and why. */
export interface Stopped {
  readonly id: string;
  readonly stopped: StopCause;
}

/** A run that did not start, and why. */
export interface CorpusRefusal {
  readonly refused: string;
}

/** What a run did. */
export interface CorpusRun {
  readonly workbooks: number;
  readonly results: number;
  readonly stopped: readonly Stopped[];
  readonly children: number;
  /** The highest memory of one child that the watchdog read, in bytes. */
  readonly peakBytes: number;
  readonly resultsFile: string;
}

/** The options of {@link runCorpus}. */
export interface RunOptions {
  /** The task name that `worker.ts` knows. */
  readonly task: string;
  readonly jobs: readonly CorpusJob[];
  /** The output folder. It must pass {@link checkOut}. */
  readonly out: string;
  /** The time that one workbook may take, in milliseconds. */
  readonly workbookMs: number;
  /** Replace the results of an earlier run in `out`. */
  readonly fresh?: boolean;
  /** The workbooks of one child: a positive integer. */
  readonly batchSize?: number;
  readonly capBytes?: number;
  /**
   * The memory at which the guard thread of the child stops it: {@link GUARD_CAP_FACTOR} times
   * the cap by default. Tests set it high to show that the watchdog alone stops a child.
   */
  readonly guardBytes?: number;
  readonly everyMs?: number;
  /** The child script. Tests only. */
  readonly worker?: string;
  /** Called after each workbook gets its line, for progress. */
  readonly onProgress?: (done: number, total: number) => void;
}

/**
 * The workbook files (`.xlsx`, `.xltx`, `.xlsm`, `.xltm`) under `dir`, sorted by their path inside
 * it, with the ids `workbook1` to `workbookN`. Hidden entries, owner files (`~$name.xlsx`), and
 * symbolic links are left out. The id depends only on the sorted position, as PPTX Editor names its
 * corpus decks, so it holds no file name.
 */
export function listWorkbooks(dir: string): CorpusJob[] {
  const found: string[] = [];
  const walk = (folder: string): void => {
    let entries;
    try {
      entries = readdirSync(folder, { withFileTypes: true });
    } catch (error) {
      // The message of the system names the folder: give only its code.
      const code = (error as NodeJS.ErrnoException).code ?? 'error';
      throw new Error(`A folder of the corpus cannot be read (${code}).`);
    }
    for (const entry of entries) {
      // Hidden entries, and the owner files (`~$name.xlsx`) that a program keeps while it edits.
      if (entry.name.startsWith('.') || entry.name.startsWith('~$')) continue;
      const path = join(folder, entry.name);
      if (entry.isDirectory()) walk(path);
      else if (entry.isFile() && WORKBOOK_FILE.test(entry.name)) found.push(path);
    }
  };
  walk(dir);
  return found
    .map((path) => ({ path, key: relative(dir, path) }))
    .sort((a, b) => (a.key < b.key ? -1 : a.key > b.key ? 1 : 0))
    .map(({ path }, i) => ({ id: `workbook${i + 1}`, path }));
}

/**
 * The corpus folder from `--dir` or `XLSX_CORPUS_DIR`, as an absolute path. Throws with the
 * remedy when neither names an existing folder: a corpus check that ran on nothing is a gap.
 */
export function corpusDir(args: readonly string[], env = process.env): string {
  const at = args.indexOf('--dir');
  const source = at >= 0 ? '--dir' : CORPUS_ENV;
  const raw = at >= 0 ? args[at + 1] : env[CORPUS_ENV];
  if (!raw)
    throw new Error(
      `Name the corpus folder with --dir <folder> or ${CORPUS_ENV}. No folder is assumed.`
    );
  const dir = resolve(raw);
  if (!existsSync(dir) || !statSync(dir).isDirectory())
    throw new Error(`The folder that ${source} names does not exist or is not a folder.`);
  return dir;
}

/** The value of `--<name> <value>`, or undefined. */
export function flag(args: readonly string[], name: string): string | undefined {
  const at = args.indexOf(`--${name}`);
  return at >= 0 ? args[at + 1] : undefined;
}

/**
 * Refuses an output folder that git could commit: inside a git work tree, it must be a folder
 * that git ignores. Git itself answers, from the nearest folder that exists, so a path through a
 * symbolic link or in another letter case gets the same answer as its plain form. The `GIT_*`
 * variables of the caller do not reach git, so a hook cannot point it at another repository. A
 * folder that cannot be made or written, or that git cannot place, is refused. Returns the
 * absolute path.
 */
export function checkOut(out: string): string {
  const abs = resolve(out);
  let probe = abs;
  const tail: string[] = [];
  while (!existsSync(probe) && dirname(probe) !== probe) {
    tail.unshift(basename(probe));
    probe = dirname(probe);
  }
  let writable = statSync(probe).isDirectory();
  try {
    accessSync(probe, constants.W_OK);
  } catch {
    writable = false;
  }
  if (!writable) throw new Error('--out must be a folder that can be made and written.');
  // Git answers in English, so its message can be read in any locale.
  const env = {
    ...Object.fromEntries(Object.entries(process.env).filter(([name]) => !name.startsWith('GIT_'))),
    LC_ALL: 'C',
    LANGUAGE: '',
  };
  const git = (args: string[]) =>
    spawnSync('git', ['-C', probe, ...args], { encoding: 'utf8', env });
  const tree = git(['rev-parse', '--is-inside-work-tree']);
  if (tree.status !== 0 && /not a git repository/i.test(tree.stderr ?? '')) return abs;
  if (tree.status !== 0)
    throw new Error('git could not tell whether --out can be committed, so it is refused.');
  // Inside `.git`, git answers `false`: nothing there can be committed.
  if (tree.stdout.trim() !== 'true') return abs;
  // Each file on its own: a rule such as `*.jsonl` ignores one file and not the others. A file
  // that git tracks is not ignored, so a tracked output file is refused too.
  const ignored = OUTPUT_FILES.every(
    (file) => git(['check-ignore', '-q', join(...tail, file)]).status === 0
  );
  if (!ignored)
    throw new Error(
      '--out must be outside every git work tree or in a folder that git ignores, such as local/.'
    );
  return abs;
}

/**
 * The refusal of a run whose output folder holds the results of an earlier run, unless `fresh`.
 * The commands call it before their other work, so a rerun stops at once.
 */
export function checkFresh(out: string, fresh: boolean): string | undefined {
  const file = join(out, 'results.jsonl');
  if (!fresh && existsSync(file) && statSync(file).size > 0)
    return `The output folder ${out} holds the results of an earlier run. Add --fresh to replace them.`;
  return undefined;
}

/**
 * Removes paths and addresses from a message before it enters the output, which holds no paths
 * and no file names. Keeps the first line, and cuts it at the first path: a path may hold
 * spaces, so nothing after it is safe to keep. From pptx-editor/scripts/fidelity/corpus-list.ts
 * at 49713b43, with one difference: the message is cut at 300 characters, so a long message of
 * the engine cannot carry much workbook text.
 */
export function redactPaths(message: string): string {
  return (message.split('\n')[0] ?? '')
    .replace(/\b(?:https?|file):\/\/\S+/g, '<address>')
    .replace(/(^|[\s'"(=])~?\/.*$/, '$1<path>')
    .slice(0, 300);
}

/** One line of the child: a start or a result of a workbook of its batch. */
type ChildLine =
  | { readonly type: 'start'; readonly id: string }
  | { readonly type: 'result'; readonly id: string; readonly result: Record<string, unknown> };

/** Reads one child line, or undefined when it is not one of the batch. */
export function readChildLine(line: string, batch: ReadonlySet<string>): ChildLine | undefined {
  let value: unknown;
  try {
    value = JSON.parse(line);
  } catch {
    return undefined;
  }
  if (typeof value !== 'object' || value === null) return undefined;
  const v = value as { type?: unknown; id?: unknown; result?: unknown };
  if (typeof v.id !== 'string' || !batch.has(v.id)) return undefined;
  if (v.type === 'start') return { type: 'start', id: v.id };
  if (v.type === 'result' && typeof v.result === 'object' && v.result !== null)
    return { type: 'result', id: v.id, result: v.result as Record<string, unknown> };
  return undefined;
}

/** What one child did with its batch. */
interface ChildOutcome {
  readonly finished: ReadonlySet<string>;
  /** The workbook that ran when the child ended without its result. */
  readonly running: string | undefined;
  readonly cause: StopCause | undefined;
  readonly peakBytes: number;
  /** The runner ended the child for a pause: the workbooks that it did not finish run again. */
  readonly paused: boolean;
}

/** The running child, its work folder and watch, and whether a pause ended it. */
let current: { child: ChildProcess; dir: string; watch?: Watch; paused?: boolean } | undefined;

/** Removes the work folder of a child. The child removed `jobs.json` when it read it. */
function removeWorkFolder(dir: string): void {
  rmSync(dir, { recursive: true, force: true });
}

/** The signals that end a run, with the exit code of each: 128 and the signal number. */
const EXIT_CODES: Partial<Record<NodeJS.Signals, number>> = {
  SIGHUP: 129,
  SIGINT: 130,
  SIGTERM: 143,
};
const SIGNALS = Object.keys(EXIT_CODES) as NodeJS.Signals[];
/**
 * The signal of Ctrl+Z. SIGTTIN and SIGTTOU keep their default: a handler for them makes a
 * blocked terminal write repeat at once, so the runner would spin and could not watch.
 */
const PAUSE_SIGNAL: NodeJS.Signals = 'SIGTSTP';

/**
 * Stops the running child and removes its folder when the command gets SIGHUP, SIGINT, or
 * SIGTERM, and before the command pauses on SIGTSTP (Ctrl+Z). After a SIGKILL of the runner,
 * the guard thread of the child removes the folder and stops the child.
 */
function guardSignals(resultsFile: string): { stopping: () => boolean; unguard: () => void } {
  let stopping = false;
  const onSignal = (signal: NodeJS.Signals) => {
    // The run loop reads this flag: it records nothing more and starts no other child.
    stopping = true;
    const running = current;
    void (running?.child.pid ? killGroup(running.child.pid) : Promise.resolve()).finally(() => {
      try {
        if (running) removeWorkFolder(running.dir);
      } catch {
        console.error(
          `The run could not remove the work folder ${running!.dir}, which holds work files of a workbook. Remove it.`
        );
      } finally {
        console.error(
          `Stopped by ${signal}. ${resultsFile} holds the results so far, and the run wrote no run.json. Add --fresh to run again.`
        );
        process.exit(EXIT_CODES[signal] ?? 1);
      }
    });
  };
  // Ctrl+Z pauses only the group of the runner, and the child leads its own group. So the runner
  // ends the child and removes its folder before it pauses: a paused run holds no memory, and
  // no child or folder stays if the paused runner is killed. The run goes on after `fg`.
  const onPause = () => {
    if (stopping) return;
    const running = current;
    if (running?.child.pid) {
      running.paused = true;
      running.watch?.stop();
      signalGroup(running.child.pid);
      try {
        removeWorkFolder(running.dir);
      } catch {
        // The runner removes the folder again when the child has ended.
      }
    }
    console.error(
      running?.child.pid
        ? 'Paused. The runner ended its child. After fg, a new child runs the unfinished workbooks.'
        : 'Paused. After fg, the run continues.'
    );
    process.kill(process.pid, 'SIGSTOP');
  };
  for (const signal of SIGNALS) process.on(signal, onSignal);
  process.on(PAUSE_SIGNAL, onPause);
  return {
    stopping: () => stopping,
    unguard: () => {
      for (const signal of SIGNALS) process.off(signal, onSignal);
      process.off(PAUSE_SIGNAL, onPause);
    },
  };
}

async function runChild(
  options: RunOptions,
  batch: readonly CorpusJob[],
  write: (id: string, result: Record<string, unknown>) => void
): Promise<ChildOutcome> {
  const dir = mkdtempSync(join(resolve(tmpdir()), 'xlsx-corpus-'));
  const jobsFile = join(dir, 'jobs.json');
  writeFileSync(jobsFile, JSON.stringify(batch));
  const ids = new Set(batch.map((job) => job.id));
  const finished = new Set<string>();
  let running: string | undefined;
  let outputCause: StopCause | undefined;
  // `detached` makes the child lead its own process group, so a stop reaches its processes.
  const child = spawn(process.execPath, [options.worker ?? WORKER, options.task, jobsFile], {
    cwd: ROOT,
    detached: true,
    // An absolute TMPDIR, so the guard thread of the child, in another folder, finds the same one.
    env: {
      ...process.env,
      TMPDIR: dirname(dir),
      [CORPUS_TMP_ENV]: dir,
      [RUNNER_ENV]: String(process.pid),
      [CORPUS_GUARD_ENV]: String(
        options.guardBytes ?? (options.capBytes ?? MEMORY_CAP_BYTES) * GUARD_CAP_FACTOR
      ),
    },
    stdio: ['ignore', 'pipe', 'inherit'],
  });
  // Before any wait: a stop signal removes the folder of the current child.
  const state: NonNullable<typeof current> = { child, dir };
  current = state;
  // A SIGKILL that the watchdog did not send came from the guard thread of the child, which stops
  // a child far above the cap, or from the system when memory ran out.
  let killed = false;
  const exited = new Promise<void>((done) => {
    child.once('close', (_code, signal) => {
      killed = signal === 'SIGKILL';
      done();
    });
    child.once('error', () => done());
  });
  if (child.pid === undefined) {
    await exited;
    current = undefined;
    removeWorkFolder(dir);
    return { finished, running: undefined, cause: 'exit', peakBytes: 0, paused: false };
  }
  const watch = watchProcess(child.pid, {
    deadlineMs: options.workbookMs,
    everyMs: options.everyMs ?? 250,
    capBytes: options.capBytes ?? MEMORY_CAP_BYTES,
    name: 'the child',
    unit: 'the workbook',
    group: true,
  });
  state.watch = watch;
  let pending = '';
  child.stdout!.setEncoding('utf8');
  child.stdout!.on('data', (chunk: string) => {
    pending += chunk;
    let newline = pending.indexOf('\n');
    while (newline >= 0) {
      const line = readChildLine(pending.slice(0, newline), ids);
      pending = pending.slice(newline + 1);
      newline = pending.indexOf('\n');
      if (!line || finished.has(line.id)) continue;
      if (line.type === 'start') {
        running = line.id;
        watch.restart();
      } else {
        finished.add(line.id);
        if (running === line.id) running = undefined;
        // The time until the next start line counts as the time of the next workbook.
        watch.restart();
        write(line.id, line.result);
      }
    }
    if (pending.length > MAX_LINE_BYTES && outputCause === undefined) {
      outputCause = 'output';
      pending = '';
      void killGroup(child.pid!);
    }
  });
  await exited;
  watch.stop();
  // The child is gone, but a process that it started may hold the folder: stop the group. Only
  // the group: the id of the child is free now, and another process can take it.
  signalGroup(child.pid);
  current = undefined;
  removeWorkFolder(dir);
  const tripped = watch.tripped();
  const cause: StopCause | undefined =
    outputCause ??
    (tripped ? (watch.trippedMemory() ? 'memory' : 'time') : killed ? 'memory' : undefined);
  const complete = finished.size === batch.length;
  return {
    finished,
    running: complete ? undefined : running,
    cause: complete ? undefined : (cause ?? 'exit'),
    peakBytes: watch.peakBytes(),
    paused: state.paused ?? false,
  };
}

/**
 * Runs `options.task` over `options.jobs`, one child at a time, and writes `results.jsonl` and
 * `run.json` to `options.out`. A workbook that a child did not finish gets a `stopped` line. As
 * in PPTX Editor, a run that would replace earlier results returns a refusal and does nothing.
 * Throws when the output folder or the batch size is not one that the run takes.
 */
export async function runCorpus(options: RunOptions): Promise<CorpusRun | CorpusRefusal> {
  const out = checkOut(options.out);
  const size = options.batchSize ?? BATCH_SIZE;
  if (!Number.isInteger(size) || size < 1)
    throw new Error('The batch size must be a positive integer.');
  const refused = checkFresh(out, options.fresh ?? false);
  if (refused) return { refused };
  mkdirSync(out, { recursive: true });
  const resultsFile = join(out, 'results.jsonl');
  // The files of an earlier run, of this task or another, would not match the new results.
  for (const file of OUTPUT_FILES) rmSync(join(out, file), { force: true });
  writeFileSync(resultsFile, '');
  const write = (line: Record<string, unknown>) =>
    appendFileSync(resultsFile, `${JSON.stringify(line)}\n`);
  const queue = [...options.jobs];
  const stopped: Stopped[] = [];
  // Workbooks that run alone in their child: see the memory stop below.
  const alone = new Set<string>();
  let results = 0;
  let children = 0;
  let peak = 0;
  const progress = () => options.onProgress?.(results + stopped.length, options.jobs.length);
  const signals = guardSignals(resultsFile);
  try {
    while (queue.length > 0) {
      const batch = alone.has(queue[0]!.id) ? queue.splice(0, 1) : queue.splice(0, size);
      // A batch with no workbook would start children forever.
      if (batch.length === 0) throw new Error('A batch of the corpus runner holds no workbook.');
      children += 1;
      const outcome = await runChild(options, batch, (id, result) => {
        // After a stop signal, the run records nothing more.
        if (signals.stopping()) return;
        results += 1;
        // The id comes first in the line, and no field of the result can replace it.
        write(Object.assign({ id }, result, { id }));
        progress();
      }).catch((error: unknown) => {
        // After a stop signal, the handler reports and ends the process, also when the folder
        // of the child cannot be removed: wait for that.
        if (!signals.stopping()) throw error;
        return new Promise<never>(() => {});
      });
      // After a stop signal, the child ended because the handler stopped it. The handler ends
      // the process, so this run waits for that: no stop line and no next child.
      if (signals.stopping()) await new Promise<never>(() => {});
      peak = Math.max(peak, outcome.peakBytes);
      const left = batch.filter((job) => !outcome.finished.has(job.id));
      // A pause ended the child: its other workbooks run in a new child, with no stop line.
      if (outcome.paused) queue.unshift(...left);
      else if (left.length > 0) {
        const cause = outcome.cause ?? 'exit';
        // The workbook that ran stops. When no workbook started, the child itself failed, so
        // the whole batch stops: a retry would fail the same way.
        const culprits = outcome.running ? left.filter((job) => job.id === outcome.running) : left;
        const again: CorpusJob[] = [];
        for (const job of culprits) {
          // The memory that earlier workbooks of the child kept counts against the workbook that
          // ran. So that workbook runs once more, alone in a new child, and stops only there.
          if (cause === 'memory' && outcome.finished.size > 0 && !alone.has(job.id)) {
            alone.add(job.id);
            again.push(job);
            continue;
          }
          stopped.push({ id: job.id, stopped: cause });
          write({ id: job.id, stopped: cause });
          progress();
        }
        queue.unshift(...again, ...left.filter((job) => !culprits.includes(job)));
      }
    }
  } finally {
    signals.unguard();
  }
  const run: CorpusRun = {
    workbooks: options.jobs.length,
    results,
    stopped,
    children,
    peakBytes: peak,
    resultsFile,
  };
  writeFileSync(
    join(out, 'run.json'),
    `${JSON.stringify(
      {
        schema: RESULTS_SCHEMA,
        task: options.task,
        workbooks: run.workbooks,
        results,
        stopped: stopped.length,
        children,
        peakMiB: Math.round(peak / 1048576),
        batchSize: size,
        capMiB: Math.round((options.capBytes ?? MEMORY_CAP_BYTES) / 1048576),
        workbookMs: options.workbookMs,
      },
      null,
      2
    )}\n`
  );
  return run;
}
