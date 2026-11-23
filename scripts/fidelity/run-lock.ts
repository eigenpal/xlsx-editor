// The lock of `bun run fidelity`: one run at a time in one checkout, because each run clears
// and writes the report folder (`fidelity-report/`). The lock is a folder that holds the process
// id of its run, and it never exists empty:
// - A run writes its id into a folder of its own and renames that folder to the lock. The rename
//   fails while a lock exists, because a lock folder always holds files.
// - A run removes a lock folder only after it renames it aside to a name of its own. So no other
//   run can rename its folder onto an emptied lock.
// - A lock whose process is gone, after a crash or a kill, is stale. To take it over, a run
//   links a takeover file that holds its process id into the lock, which only one run can win.
//   It then checks that the lock still names the stale process, and renames the lock aside. It
//   checks that the moved folder still names the stale process and holds its takeover file, and
//   renames the folder back when it does not.
// - A takeover file whose run is gone is stale. Then the next run tries the takeover file of the
//   next generation, so a run that dies in a takeover never locks out later runs. The age of a
//   takeover file never makes it stale: a run that pauses in a takeover blocks the other runs,
//   and never loses its place to them.
// A run that SIGINT or SIGTERM stops releases its lock (`releaseOnSignals`). When the runs of
// all `TAKEOVER_GENERATIONS` takeover files of a lock are gone, the run reports the lock as held
// and says how to remove it by hand.
//
// Difference: PPTX Editor has no run lock. A run of every deck there clears the report folder of
// its deck set (`rmSync(out, ...)` in pptx-editor/scripts/fidelity/run.ts at 1943d2c9, EigenPal),
// and nothing keeps a second run from writing into that folder at the same time. A run here also
// clears its report folder (`fidelity-report/`), so the lock keeps two runs apart.
// `run-lock.test.ts` runs the lock under contention between processes (`run-lock-stress.ts`).

import {
  closeSync,
  linkSync,
  mkdirSync,
  openSync,
  readdirSync,
  readFileSync,
  renameSync,
  rmSync,
  statSync,
  writeFileSync,
  writeSync,
} from 'node:fs';
import { join } from 'node:path';

/** The name of the lock folder inside the report folder. */
export const LOCK_NAME = '.run.lock';

/** The file of the lock folder that holds the process id of its run. */
const PID_FILE = 'pid';

/**
 * A lock folder with no process id is stale after this time (milliseconds): a lock always
 * appears with its id, so such a folder is left over from an earlier version or a hand.
 */
export const NO_PID_STALE_MS = 5000;

/** The number of takeover generations that a run tries before it reports the lock as held. */
const TAKEOVER_GENERATIONS = 20;

/** Whether the process `pid` runs: signal 0 tests it without a signal. */
export function processRuns(pid: number): boolean {
  try {
    process.kill(pid, 0);
    return true;
  } catch (error) {
    // EPERM: the process runs under another user.
    return (error as NodeJS.ErrnoException).code === 'EPERM';
  }
}

/** The positive whole number in the file `path`, or undefined. */
function pidIn(path: string): number | undefined {
  try {
    const pid = Number(readFileSync(path, 'utf8').trim());
    return Number.isInteger(pid) && pid > 0 ? pid : undefined;
  } catch {
    return undefined;
  }
}

/** The process id in the lock folder `lock`, or undefined when it has none. */
function holder(lock: string): number | undefined {
  return pidIn(join(lock, PID_FILE));
}

/** The time of the last change of `path` in milliseconds, or undefined when it is gone. */
function statMs(path: string): number | undefined {
  try {
    return statSync(path).mtimeMs;
  } catch {
    return undefined;
  }
}

/** Remove `path` and what it holds; a run that removes it at the same time is no error. */
function remove(path: string): void {
  try {
    rmSync(path, { recursive: true, force: true });
  } catch {
    // Another run removes it too; a later run removes what is left (`removeLeftovers`).
  }
}

/** The files of the report folder that a run writes beside the folders of its fixtures. */
export const REPORT_FILES: readonly string[] = ['report.md', 'results.json'];

/**
 * Whether a run of the fixtures `names` writes the entry `entry` of the report folder, and so
 * clears it first: the folder of each fixture that it measures, and the report files. Every
 * other entry stays: the folders of other fixtures, and files that a person put there.
 */
export function clearsEntry(names: readonly string[], entry: string): boolean {
  return names.includes(entry) || REPORT_FILES.includes(entry);
}

/**
 * The file that a clear writes into its aside folder once every entry moved there. Only an aside
 * folder with it, or an empty one, is a leftover that a later run removes (`removeLeftovers`).
 */
export const CLEARED_MARK = '.complete';

/** The error code of `error`, or its first line. */
const errorCode = (error: unknown): string =>
  (error as NodeJS.ErrnoException).code ??
  (error instanceof Error ? error.message : String(error)).split('\n')[0]!;

/** `path` as one word of a POSIX shell: in single quotes, so that `"` and `$` stay as they are. */
export const shellPath = (path: string): string => `'${path.replaceAll("'", "'\\''")}'`;

/**
 * The command that removes the folder `path`, which `code` stopped a removal of. A folder that
 * holds a folder without write or search access needs that access first. Its removal stops with
 * `EACCES` or `EPERM`, or with `ENOTEMPTY` when the removal of that folder fails inside it.
 */
const removeStep = (path: string, code: string): string =>
  ['EACCES', 'EPERM', 'ENOTEMPTY'].includes(code)
    ? `chmod -R u+rwx ${shellPath(path)} && rm -r ${shellPath(path)}`
    : `rm -r ${shellPath(path)}`;

/**
 * Clear the entries of the report folder `dir` that `clear` names, never the lock and its
 * leftovers: all of them, or none. A run that holds the lock calls it. Each entry, in sorted
 * order, first moves into one aside folder beside the lock (`.run.lock.cleared.`). When an entry
 * cannot move, such as a folder without write access, or the clear cannot write `CLEARED_MARK`,
 * the moved entries move back, and the error names the cause and the next step. So a run that
 * stops there keeps the whole earlier report. When an entry cannot move back either, it stays in
 * the aside folder, which has no `CLEARED_MARK`: no run removes it, and the error names it.
 * After every entry moved, the clear writes `CLEARED_MARK`, removes the entries of the folder,
 * and removes the mark and the folder last. When that removal fails, the folder keeps its mark,
 * and the error names it with the command that removes it (`removeStep`). A later run removes
 * the folder when it can (`removeLeftovers`); otherwise the folder stays until someone runs that
 * command. `moves` is a test hook that sees each move.
 */
export function clearReport(
  dir: string,
  clear: (entry: string) => boolean,
  pid: number = process.pid,
  moves?: (entry: string, back: boolean) => void
): void {
  const entries = readdirSync(dir)
    .filter((e) => !e.startsWith(LOCK_NAME) && clear(e))
    .sort();
  if (!entries.length) return;
  const aside = uniqueName(dir, 'cleared', pid);
  mkdirSync(aside);
  const moved: string[] = [];
  // Why the clear stops, and the next step, as two fields: a path can hold any character.
  let blocked: { readonly cause: string; readonly remedy: string } | undefined;
  for (const entry of entries) {
    try {
      renameSync(join(dir, entry), join(aside, entry));
    } catch (error) {
      const path = join(dir, entry);
      blocked = {
        cause: `${path} cannot move aside (${errorCode(error)})`,
        remedy: `give write access to ${path} or remove it, then run again`,
      };
      break;
    }
    moved.push(entry);
    moves?.(entry, false);
  }
  if (!blocked) {
    try {
      writeFileSync(join(aside, CLEARED_MARK), '');
    } catch (error) {
      blocked = {
        cause: `${aside} cannot take its mark ${CLEARED_MARK} (${errorCode(error)})`,
        remedy: `free disk space or give write access to ${dir}, then run again`,
      };
    }
  }
  if (!blocked) {
    try {
      for (const child of readdirSync(aside))
        if (child !== CLEARED_MARK) rmSync(join(aside, child), { recursive: true, force: true });
      rmSync(aside, { recursive: true, force: true });
    } catch (error) {
      const code = errorCode(error);
      throw new Error(
        `${aside} cannot be removed (${code}); it holds only report entries that this run cleared from ${dir}, so remove it, then run again: ${removeStep(aside, code)}`
      );
    }
    return;
  }
  // The run holds the lock, so no other run takes these names back meanwhile.
  const stuck: string[] = [];
  for (const entry of moved.reverse()) {
    try {
      renameSync(join(aside, entry), join(dir, entry));
    } catch {
      stuck.push(entry);
      continue;
    }
    moves?.(entry, true);
  }
  if (!stuck.length) {
    remove(aside);
    throw new Error(`${blocked.cause}; ${blocked.remedy}`);
  }
  throw new Error(
    `${blocked.cause}, and ${stuck.sort().join(', ')} could not move back: they are in ${aside}; move its entries back into ${dir}, or remove ${aside}`
  );
}

/** A unique name beside the lock, for the run `pid`: `.run.lock.<kind>.<pid>.<time>.<random>`. */
function uniqueName(dir: string, kind: 'new' | 'stale' | 'cleared', pid: number): string {
  return join(
    dir,
    `${LOCK_NAME}.${kind}.${pid}.${Date.now()}.${Math.random().toString(36).slice(2)}`
  );
}

/**
 * Rename the lock folder `lock` aside and remove it, when the moved folder passes `check`. False
 * when the lock is gone, or when the moved folder fails `check`: then it goes back to its name.
 * Only the run that renames it removes it, so the lock is never empty at its own name.
 */
function moveAsideAndRemove(
  lock: string,
  dir: string,
  pid: number,
  check: (moved: string) => boolean
): boolean {
  const away = uniqueName(dir, 'stale', pid);
  try {
    renameSync(lock, away);
  } catch {
    return false;
  }
  if (!check(away)) {
    try {
      renameSync(away, lock);
    } catch {
      // A new lock took the name meanwhile. The moved folder stays aside, never removed, so that
      // its run can see it: a folder that names a running process is no leftover.
    }
    return false;
  }
  remove(away);
  return true;
}

/** Whether the aside folder of a clear is empty, or holds `CLEARED_MARK` as a file. */
function clearedLeftover(folder: string): boolean {
  try {
    const held = readdirSync(folder);
    return held.length === 0 || statSync(join(folder, CLEARED_MARK)).isFile();
  } catch {
    return false;
  }
}

/**
 * Remove the folders that runs left beside the lock: a folder of a run that prepared its lock,
 * moved a lock aside, or cleared the report folder (only an empty or marked one,
 * `CLEARED_MARK`), and whose process does not run. A folder of this run's process id is a
 * leftover too, because this run makes its folders after this step. A marked folder that cannot
 * be removed stays, and `notice` names it with the command that removes it (`removeStep`).
 */
function removeLeftovers(
  dir: string,
  pid: number,
  runs: (pid: number) => boolean,
  notice: (text: string) => void
): void {
  let entries: string[];
  try {
    entries = readdirSync(dir);
  } catch {
    return;
  }
  for (const entry of entries) {
    const kind = [`${LOCK_NAME}.new.`, `${LOCK_NAME}.stale.`, `${LOCK_NAME}.cleared.`].find((p) =>
      entry.startsWith(p)
    );
    if (!kind) continue;
    const owner = Number(entry.slice(kind.length).split('.')[0]);
    const live = Number.isInteger(owner) && owner > 0 && owner !== pid && runs(owner);
    // A folder moved aside that names a running process is a live lock that could not go back.
    const moved = kind.endsWith('.stale.') ? holder(join(dir, entry)) : undefined;
    const held = moved !== undefined && moved !== pid && runs(moved);
    if (live || held) continue;
    // An aside folder of a clear without its mark holds entries that could not move back, or that
    // a clear left when a stop ended it: no run removes it, and the notice names it.
    if (kind.endsWith('.cleared.') && !clearedLeftover(join(dir, entry))) {
      notice(
        `${join(dir, entry)} holds report entries that a clear did not finish with; move its entries back into ${dir}, or remove it: ${removeStep(join(dir, entry), 'EACCES')}`
      );
      continue;
    }
    if (!kind.endsWith('.cleared.')) {
      remove(join(dir, entry));
      continue;
    }
    // The mark goes last, so that a folder that cannot be removed keeps it.
    const path = join(dir, entry);
    try {
      for (const child of readdirSync(path))
        if (child !== CLEARED_MARK) rmSync(join(path, child), { recursive: true, force: true });
      rmSync(path, { recursive: true, force: true });
    } catch (error) {
      const code = errorCode(error);
      if (statMs(path) !== undefined)
        notice(
          `${path} is left from a cleared report and cannot be removed (${code}); remove it: ${removeStep(path, code)}`
        );
    }
  }
}

/**
 * Create the file `path` with this run's process id, or false when it exists. The id is written
 * into a file of this run beside the lock first, and then linked to `path`: the link fails when
 * `path` exists, so `path` never exists without the id.
 */
function createExclusive(path: string, dir: string, pid: number): boolean {
  const draft = uniqueName(dir, 'new', pid);
  const fd = openSync(draft, 'wx');
  try {
    writeSync(fd, `${pid}\n`);
  } finally {
    closeSync(fd);
  }
  try {
    linkSync(draft, path);
    return true;
  } catch {
    return false;
  } finally {
    remove(draft);
  }
}

/** Test hooks of `takeRunLock`. */
export interface RunLockHooks {
  /** Called after the holder check of a takeover, before the lock is renamed aside. */
  readonly beforeMoveAside?: () => void;
  /** Prints a notice, such as a leftover that cannot be removed. The default prints one line. */
  readonly notice?: (text: string) => void;
}

/**
 * Take over the stale lock `lock` of the process `stale` (undefined: a lock with no process
 * id). `taken` when this run renamed the lock aside and the caller may take it; `busy` when
 * another run takes it over, or the lock no longer names `stale`; `full` when the runs of every
 * takeover generation are gone, and no generation is left to try.
 */
function takeOver(
  lock: string,
  dir: string,
  stale: number | undefined,
  pid: number,
  runs: (pid: number) => boolean,
  hooks: RunLockHooks
): 'taken' | 'busy' | 'full' {
  for (let generation = 0; generation < TAKEOVER_GENERATIONS; generation += 1) {
    // The name holds the stale process id, so a file that lands in a newer lock is never the
    // takeover file of that lock.
    const name = `takeover.${stale ?? 'none'}.${generation}`;
    const token = join(lock, name);
    if (createExclusive(token, dir, pid)) {
      if (holder(lock) !== stale) {
        // The lock changed before the takeover file: leave it to its run.
        remove(token);
        return 'busy';
      }
      hooks.beforeMoveAside?.();
      const moved = moveAsideAndRemove(
        lock,
        dir,
        pid,
        (away) => holder(away) === stale && pidIn(join(away, name)) === pid
      );
      return moved ? 'taken' : 'busy';
    }
    // Another run holds this generation. Its file is stale only when its run is gone (a file of
    // this run's process id is left from a gone process with the same id); then the next
    // generation decides. A run that is gone can no longer move the lock.
    if (statMs(token) === undefined) return 'busy';
    const writer = pidIn(token);
    const gone = writer === undefined || writer === pid || !runs(writer);
    if (!gone) return 'busy';
  }
  return 'full';
}

/**
 * Take the lock of the report folder `dir`, or return the reason why another run holds it. A
 * lock whose process does not run (`runs`), a lock that names this run's own process id (a
 * process id that the system gave again), and a lock with no process id older than
 * `NO_PID_STALE_MS` are stale and are taken over. The lock clears no entry of the report folder:
 * the run does that next (`clearReport`). Call `release` when the run ends.
 */
export function takeRunLock(
  dir: string,
  pid: number = process.pid,
  runs: (pid: number) => boolean = processRuns,
  now: () => number = Date.now,
  hooks: RunLockHooks = {}
): { release: () => void } | { held: string } {
  const lock = join(dir, LOCK_NAME);
  mkdirSync(dir, { recursive: true });
  removeLeftovers(
    dir,
    pid,
    runs,
    hooks.notice ?? ((text) => console.error(`fidelity: NOTICE: ${text}`))
  );
  const mine = uniqueName(dir, 'new', pid);
  mkdirSync(mine);
  if (!createExclusive(join(mine, PID_FILE), dir, pid)) {
    remove(mine);
    throw new Error(`cannot write the process id into ${mine}`);
  }
  let other: number | undefined;
  for (let attempt = 0; attempt < 3; attempt += 1) {
    try {
      renameSync(mine, lock);
      return {
        release: () => {
          if (holder(lock) === pid)
            moveAsideAndRemove(lock, dir, pid, (moved) => holder(moved) === pid);
        },
      };
    } catch (error) {
      const code = (error as NodeJS.ErrnoException).code ?? '';
      if (!['EEXIST', 'ENOTEMPTY', 'EISDIR'].includes(code)) {
        remove(mine);
        throw error;
      }
    }
    other = holder(lock);
    const time = statMs(lock);
    // The lock went away between the rename and this check: try again.
    if (time === undefined) continue;
    const stale =
      other === undefined ? now() - time > NO_PID_STALE_MS : other === pid || !runs(other);
    const taken = stale ? takeOver(lock, dir, other, pid, runs, hooks) : 'busy';
    if (taken === 'full') {
      remove(mine);
      return {
        held: `${lock} holds ${TAKEOVER_GENERATIONS} takeover files of runs that are gone; when no \`bun run fidelity\` runs, remove that folder by hand: rm -r ${shellPath(lock)}`,
      };
    }
    if (taken !== 'taken') {
      other = holder(lock);
      break;
    }
  }
  remove(mine);
  return {
    held: `another \`bun run fidelity\` (process ${other ?? 'unknown'}) holds ${lock}; wait for that process to end, then run again`,
  };
}

/**
 * Release `lock` and exit when SIGINT (exit code 130) or SIGTERM (143) stops the run. Returns
 * the function that removes the handlers. `exit` is `process.exit`, replaced in tests.
 */
export function releaseOnSignals(
  lock: { release: () => void },
  exit: (code: number) => void = (code) => process.exit(code)
): () => void {
  const stop = (signal: NodeJS.Signals, code: number) => (): void => {
    lock.release();
    console.error(`fidelity: stopped by ${signal}`);
    exit(code);
  };
  const onInt = stop('SIGINT', 130);
  const onTerm = stop('SIGTERM', 143);
  process.once('SIGINT', onInt);
  process.once('SIGTERM', onTerm);
  return () => {
    process.off('SIGINT', onInt);
    process.off('SIGTERM', onTerm);
  };
}
