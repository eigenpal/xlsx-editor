// The lock of `bun run fidelity` (`run-lock.ts`).

import { afterEach, describe, expect, test } from 'bun:test';
import { spawnSync } from 'node:child_process';
import {
  chmodSync,
  mkdirSync,
  mkdtempSync,
  readdirSync,
  readFileSync,
  rmSync,
  writeFileSync,
} from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import {
  CLEARED_MARK,
  clearReport,
  clearsEntry,
  LOCK_NAME,
  NO_PID_STALE_MS,
  processRuns,
  releaseOnSignals,
  shellPath,
  takeRunLock,
} from './run-lock.ts';

const dirs: string[] = [];
afterEach(() => {
  for (const d of dirs.splice(0)) rmSync(d, { recursive: true, force: true });
});

describe('the run lock', () => {
  test('a second run refuses to start while the first holds the report folder', () => {
    const dir = mkdtempSync(join(tmpdir(), 'xlsx-lock-'));
    dirs.push(dir);
    writeFileSync(join(dir, 'old.md'), 'old');
    const first = takeRunLock(dir, 101, () => true);
    expect('release' in first).toBe(true);
    // The first run writes its process id, and the lock clears nothing (`clearReport` does).
    expect(readdirSync(dir).sort()).toEqual([LOCK_NAME, 'old.md']);
    expect(readFileSync(join(dir, LOCK_NAME, 'pid'), 'utf8')).toBe('101\n');
    // The process of the first run runs: the lock is held.
    const second = takeRunLock(dir, 102, () => true);
    expect('held' in second && second.held).toContain(
      'another `bun run fidelity` (process 101) holds'
    );
    if ('release' in first) first.release();
    const third = takeRunLock(dir, 103, () => true);
    expect('release' in third).toBe(true);
    if ('release' in third) third.release();
    expect(readdirSync(dir)).toEqual(['old.md']);
  });

  test('of two runs that find one stale lock, one takes it over', () => {
    const dir = mkdtempSync(join(tmpdir(), 'xlsx-lock-'));
    dirs.push(dir);
    mkdirSync(join(dir, LOCK_NAME));
    writeFileSync(join(dir, LOCK_NAME, 'pid'), '999\n');
    // Run A starts between the check of run B, which finds process 999 gone, and its takeover.
    let a: ReturnType<typeof takeRunLock> | undefined;
    const b = takeRunLock(dir, 2, () => {
      a ??= takeRunLock(dir, 1, () => false);
      return false;
    });
    expect(a && 'release' in a).toBe(true);
    expect('held' in b).toBe(true);
    expect(readFileSync(join(dir, LOCK_NAME, 'pid'), 'utf8')).toBe('1\n');
    // Run B leaves no folder of its own.
    expect(readdirSync(dir)).toEqual([LOCK_NAME]);
  });

  test('SIGTERM and SIGINT release the lock and exit', () => {
    const dir = mkdtempSync(join(tmpdir(), 'xlsx-lock-'));
    dirs.push(dir);
    const lock = takeRunLock(dir, 301, () => true);
    if (!('release' in lock)) throw new Error('not taken');
    const codes: number[] = [];
    const quiet = console.error;
    console.error = () => {};
    try {
      const unregister = releaseOnSignals(lock, (code) => void codes.push(code));
      process.emit('SIGTERM');
      expect(codes).toEqual([143]);
      expect(readdirSync(dir)).toEqual([]);
      unregister();
      // `unregister` removes the handlers that it added.
      const before = process.listenerCount('SIGINT');
      const again = releaseOnSignals(lock, (code) => void codes.push(code));
      expect(process.listenerCount('SIGINT')).toBe(before + 1);
      again();
      expect(process.listenerCount('SIGINT')).toBe(before);
      const next = releaseOnSignals(lock, (code) => void codes.push(code));
      process.emit('SIGINT');
      expect(codes).toEqual([143, 130]);
      next();
    } finally {
      console.error = quiet;
    }
  });

  test('a lock whose process is gone is stale, and the next run takes it over', () => {
    const dir = mkdtempSync(join(tmpdir(), 'xlsx-lock-'));
    dirs.push(dir);
    // A run killed with its lock in place.
    const killed = takeRunLock(dir, 201, () => true);
    expect('release' in killed).toBe(true);
    const next = takeRunLock(dir, 202, (pid) => pid !== 201);
    expect('release' in next).toBe(true);
    expect(readFileSync(join(dir, LOCK_NAME, 'pid'), 'utf8')).toBe('202\n');
    // The release of the stale run leaves the lock of the new run.
    if ('release' in killed) killed.release();
    expect(readdirSync(dir)).toEqual([LOCK_NAME]);
    if ('release' in next) next.release();
    // A lock folder with no process id is held while it is new, and stale after
    // `NO_PID_STALE_MS`.
    mkdirSync(join(dir, LOCK_NAME));
    writeFileSync(join(dir, LOCK_NAME, 'note'), 'no pid');
    expect('held' in takeRunLock(dir, 203, () => false)).toBe(true);
    const later = () => Date.now() + NO_PID_STALE_MS + 1000;
    const after = takeRunLock(dir, 204, () => false, later);
    expect('release' in after).toBe(true);
    expect(readFileSync(join(dir, LOCK_NAME, 'pid'), 'utf8')).toBe('204\n');
    if ('release' in after) after.release();
    // This process runs, and a process id that no process has does not.
    expect(processRuns(process.pid)).toBe(true);
    expect(processRuns(2 ** 22 + 12345)).toBe(false);
  });

  test('of three runs that meet at one stale lock, one holds it', () => {
    const dir = mkdtempSync(join(tmpdir(), 'xlsx-lock-'));
    dirs.push(dir);
    mkdirSync(join(dir, LOCK_NAME));
    writeFileSync(join(dir, LOCK_NAME, 'pid'), '999\n');
    // Run B finds process 999 gone; before B takes the lock over, run A takes it over and run C
    // tries it. Processes 1 (A) and 3 (C) run.
    const live = (pid: number): boolean => pid === 1 || pid === 3;
    let a: ReturnType<typeof takeRunLock> | undefined;
    let c: ReturnType<typeof takeRunLock> | undefined;
    const b = takeRunLock(dir, 2, (pid) => {
      if (pid === 999 && !a) {
        a = takeRunLock(dir, 1, live);
        c = takeRunLock(dir, 3, live);
      }
      return live(pid);
    });
    const holding = [a, b, c].filter((r) => r && 'release' in r);
    expect(holding).toHaveLength(1);
    expect(a && 'release' in a).toBe(true);
    expect(readFileSync(join(dir, LOCK_NAME, 'pid'), 'utf8')).toBe('1\n');
    // The takeover of B leaves the lock of A whole.
    expect(readdirSync(join(dir, LOCK_NAME))).toEqual(['pid']);
  });

  test('a lock with the pid of this run is stale, and leftover folders of gone runs go', () => {
    const dir = mkdtempSync(join(tmpdir(), 'xlsx-lock-'));
    dirs.push(dir);
    mkdirSync(join(dir, LOCK_NAME));
    writeFileSync(join(dir, LOCK_NAME, 'pid'), '401\n');
    mkdirSync(join(dir, `${LOCK_NAME}.new.402.1.x`));
    mkdirSync(join(dir, `${LOCK_NAME}.new.403.1.x`));
    mkdirSync(join(dir, `${LOCK_NAME}.stale.404.1.x`));
    mkdirSync(join(dir, `${LOCK_NAME}.stale.old`));
    mkdirSync(join(dir, `${LOCK_NAME}.cleared.406.1.x`));
    mkdirSync(join(dir, `${LOCK_NAME}.cleared.403.1.x`));
    // Of a gone process 407: a finished clear (marked), and one whose entries could not move back.
    mkdirSync(join(dir, `${LOCK_NAME}.cleared.407.1.done`, 'basic-values'), { recursive: true });
    writeFileSync(join(dir, `${LOCK_NAME}.cleared.407.1.done`, CLEARED_MARK), '');
    mkdirSync(join(dir, `${LOCK_NAME}.cleared.407.1.stuck`, 'basic-values'), { recursive: true });
    // Process 401 runs, but it is this run: the system gave its process id again.
    const own = takeRunLock(dir, 401, () => true);
    expect('release' in own).toBe(true);
    // Every process runs here, so only the folder that names no process goes.
    expect(readdirSync(dir).sort()).toEqual([
      LOCK_NAME,
      `${LOCK_NAME}.cleared.403.1.x`,
      `${LOCK_NAME}.cleared.406.1.x`,
      `${LOCK_NAME}.cleared.407.1.done`,
      `${LOCK_NAME}.cleared.407.1.stuck`,
      `${LOCK_NAME}.new.402.1.x`,
      `${LOCK_NAME}.new.403.1.x`,
      `${LOCK_NAME}.stale.404.1.x`,
    ]);
    if ('release' in own) own.release();
    // Processes 402, 404, 406, and 407 are gone, so the folders that they left go: also an
    // empty or a marked folder of a clear (`clearReport`). A folder of a clear without the mark
    // holds entries that could not move back, so it stays. Process 403 runs, so its folders stay.
    const notices: string[] = [];
    const next = takeRunLock(dir, 405, (pid) => pid === 403, Date.now, {
      notice: (text) => void notices.push(text),
    });
    // The folder that stays is named in a notice, with the command that removes it.
    const stuck = join(dir, `${LOCK_NAME}.cleared.407.1.stuck`);
    expect(notices).toEqual([
      `${stuck} holds report entries that a clear did not finish with; move its entries back into ${dir}, or remove it: chmod -R u+rwx '${stuck}' && rm -r '${stuck}'`,
    ]);
    const stays = [
      `${LOCK_NAME}.cleared.403.1.x`,
      `${LOCK_NAME}.cleared.407.1.stuck`,
      `${LOCK_NAME}.new.403.1.x`,
    ];
    expect(readdirSync(dir).sort()).toEqual([LOCK_NAME, ...stays]);
    if ('release' in next) next.release();
    expect(readdirSync(dir).sort()).toEqual(stays);
  });

  test('a takeover file of a gone run does not lock out later runs; an old one of a live run does', () => {
    const dir = mkdtempSync(join(tmpdir(), 'xlsx-lock-'));
    dirs.push(dir);
    mkdirSync(join(dir, LOCK_NAME));
    writeFileSync(join(dir, LOCK_NAME, 'pid'), '999\n');
    // Run 601 died in its takeover; run 602 still runs, however long ago it stopped.
    writeFileSync(join(dir, LOCK_NAME, 'takeover.999.0'), '601\n');
    writeFileSync(join(dir, LOCK_NAME, 'takeover.999.1'), '602\n');
    const later = () => Date.now() + 3_600_000;
    expect('held' in takeRunLock(dir, 603, (pid) => pid === 602, later)).toBe(true);
    // When run 602 is gone too, the next generation decides.
    const next = takeRunLock(dir, 603, () => false);
    expect('release' in next).toBe(true);
    expect(readdirSync(join(dir, LOCK_NAME))).toEqual(['pid']);
    if ('release' in next) next.release();
  });

  test('a run that pauses in its takeover keeps it, and no other run holds the lock meanwhile', () => {
    const dir = mkdtempSync(join(tmpdir(), 'xlsx-lock-'));
    dirs.push(dir);
    mkdirSync(join(dir, LOCK_NAME));
    writeFileSync(join(dir, LOCK_NAME, 'pid'), '999\n');
    const live = (pid: number): boolean => pid === 1 || pid === 2;
    let b: ReturnType<typeof takeRunLock> | undefined;
    // Run 1 pauses after its holder check, for an hour by the clock of run 2.
    const a = takeRunLock(dir, 1, live, Date.now, {
      beforeMoveAside: () => {
        b = takeRunLock(dir, 2, live, () => Date.now() + 3_600_000);
      },
    });
    expect(b && 'held' in b).toBe(true);
    expect('release' in a).toBe(true);
    expect(readFileSync(join(dir, LOCK_NAME, 'pid'), 'utf8')).toBe('1\n');
    if ('release' in a) a.release();
  });

  test('a takeover puts back a moved folder that no longer names the stale process', () => {
    const dir = mkdtempSync(join(tmpdir(), 'xlsx-lock-'));
    dirs.push(dir);
    mkdirSync(join(dir, LOCK_NAME));
    writeFileSync(join(dir, LOCK_NAME, 'pid'), '999\n');
    // Between the holder check and the rename, the lock changes to the live run 7. The lock
    // cannot change so while it holds the takeover file; the check guards that rule.
    const a = takeRunLock(dir, 1, (pid) => pid === 7 || pid === 1, Date.now, {
      beforeMoveAside: () => writeFileSync(join(dir, LOCK_NAME, 'pid'), '7\n'),
    });
    expect('held' in a).toBe(true);
    expect(readFileSync(join(dir, LOCK_NAME, 'pid'), 'utf8')).toBe('7\n');
    expect(readdirSync(dir)).toEqual([LOCK_NAME]);
  });

  test('a lock whose takeover generations are all used says how to remove it by hand', () => {
    const dir = mkdtempSync(join(tmpdir(), 'xlsx-lock-'));
    dirs.push(dir);
    mkdirSync(join(dir, LOCK_NAME));
    writeFileSync(join(dir, LOCK_NAME, 'pid'), '999\n');
    for (let g = 0; g < 20; g += 1)
      writeFileSync(join(dir, LOCK_NAME, `takeover.999.${g}`), `${700 + g}\n`);
    const held = takeRunLock(dir, 1, (pid) => pid === 1);
    expect('held' in held && held.held).toContain('remove that folder by hand');
    expect(readFileSync(join(dir, LOCK_NAME, 'pid'), 'utf8')).toBe('999\n');
  });

  test('six processes that take and release the lock never hold it at the same time', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'xlsx-lock-'));
    dirs.push(dir);
    // `XLSX_LOCK_STRESS_MS` sets a longer run by hand.
    const ms = String(Number(process.env.XLSX_LOCK_STRESS_MS) || 1500);
    const worker = join(import.meta.dir, 'run-lock-stress.ts');
    // Two workers exit while they hold the lock, as killed runs do.
    const args = [[], [], [], [], ['20'], ['40']];
    const runs = args.map((extra) =>
      Bun.spawn([process.execPath, worker, dir, ms, ...extra], { stdout: 'pipe', stderr: 'pipe' })
    );
    const totals = { got: 0, double: 0, crashes: 0, held: 0, failed: 0 };
    for (const run of runs) {
      const out = await new Response(run.stdout).text();
      if ((await run.exited) !== 0) totals.failed += 1;
      const counts = JSON.parse(out.trim()) as Omit<typeof totals, 'failed'>;
      for (const key of ['got', 'double', 'crashes', 'held'] as const) totals[key] += counts[key];
    }
    if (process.env.XLSX_LOCK_STRESS_MS) console.log(JSON.stringify(totals));
    expect({ double: totals.double, crashes: totals.crashes, failed: totals.failed }).toEqual({
      double: 0,
      crashes: 0,
      failed: 0,
    });
    expect(totals.got).toBeGreaterThan(60);
  }, 60_000);

  test('a subset run clears only the folders that it writes again', () => {
    const dir = mkdtempSync(join(tmpdir(), 'xlsx-lock-'));
    dirs.push(dir);
    mkdirSync(join(dir, 'basic-values'));
    mkdirSync(join(dir, 'formulas'));
    const lock = takeRunLock(dir, 501, () => true);
    clearReport(dir, (entry) => entry === 'formulas');
    expect(readdirSync(dir).sort()).toEqual([LOCK_NAME, 'basic-values']);
    if ('release' in lock) lock.release();
  });
});

describe('the entries that a run clears', () => {
  test('a run clears only the folders of its fixtures and the report files', () => {
    const all = ['basic-values', 'formulas'];
    for (const entry of ['basic-values', 'formulas', 'report.md', 'results.json'])
      expect(clearsEntry(all, entry)).toBe(true);
    for (const entry of ['important.txt', 'src', 'old-fixture', LOCK_NAME])
      expect(clearsEntry(all, entry)).toBe(false);
    // A run of some fixtures keeps the folders of the others.
    expect(clearsEntry(['formulas'], 'basic-values')).toBe(false);
    expect(clearsEntry(['formulas'], 'report.md')).toBe(true);
  });

  test('a lock clears nothing', () => {
    const dir = mkdtempSync(join(tmpdir(), 'xlsx-clear-'));
    dirs.push(dir);
    writeFileSync(join(dir, 'report.md'), 'old');
    const lock = takeRunLock(dir, 101, () => true);
    expect(readdirSync(dir).sort()).toEqual([LOCK_NAME, 'report.md']);
    if ('release' in lock) lock.release();
  });

  test('a foreign file in the report folder survives a run of every fixture', () => {
    const dir = mkdtempSync(join(tmpdir(), 'xlsx-clear-'));
    dirs.push(dir);
    for (const name of ['basic-values', 'src', LOCK_NAME]) mkdirSync(join(dir, name));
    for (const name of ['important.txt', 'report.md', 'results.json'])
      writeFileSync(join(dir, name), name);
    writeFileSync(join(dir, 'src', 'a.ts'), 'source');
    clearReport(dir, (entry) => clearsEntry(['basic-values', 'formulas'], entry));
    expect(readdirSync(dir).sort()).toEqual([LOCK_NAME, 'important.txt', 'src']);
    expect(readFileSync(join(dir, 'src', 'a.ts'), 'utf8')).toBe('source');
  });

  // A root user removes the entry anyway, so the test needs another user.
  test.skipIf(process.getuid?.() === 0)(
    'an entry that cannot move throws (the run then exits with code 2)',
    () => {
      const dir = mkdtempSync(join(tmpdir(), 'xlsx-clear-'));
      dirs.push(dir);
      // The clear moves its entries in sorted order: the folder `formulas` and the file
      // `report.md` move before the blocked folder `zz-blocked`.
      writeFileSync(join(dir, 'report.md'), 'old');
      mkdirSync(join(dir, 'formulas'));
      writeFileSync(join(dir, 'formulas', 'ours.png'), 'old');
      const fixture = join(dir, 'zz-blocked');
      mkdirSync(fixture);
      writeFileSync(join(fixture, 'ours.png'), 'old');
      chmodSync(fixture, 0o500);
      const before = readdirSync(dir).sort();
      const moves: string[] = [];
      try {
        // A folder without write access cannot move into another folder.
        let blocked = '';
        try {
          clearReport(
            dir,
            (entry) => clearsEntry(['formulas', 'zz-blocked'], entry),
            101,
            (entry, back) => void moves.push(`${back ? 'back' : 'aside'} ${entry}`)
          );
        } catch (caught) {
          blocked = (caught as Error).message;
        }
        // The path can hold characters that a regular expression reads: compare strings.
        expect([
          `${fixture} cannot move aside (EACCES); give write access to ${fixture} or remove it, then run again`,
          `${fixture} cannot move aside (EPERM); give write access to ${fixture} or remove it, then run again`,
        ]).toContain(blocked);
        // All or nothing: the entries that moved first moved back, in reverse order.
        expect(moves).toEqual([
          'aside formulas',
          'aside report.md',
          'back report.md',
          'back formulas',
        ]);
        expect(readdirSync(dir).sort()).toEqual(before);
        expect(readFileSync(join(dir, 'report.md'), 'utf8')).toBe('old');
        expect(readFileSync(join(dir, 'formulas', 'ours.png'), 'utf8')).toBe('old');
        expect(readFileSync(join(fixture, 'ours.png'), 'utf8')).toBe('old');
        // An entry that cannot move back stays in a folder that no run removes, and the error
        // names it. Here a file takes the name of `formulas` while it is aside.
        let error: unknown;
        try {
          clearReport(
            dir,
            (entry) => clearsEntry(['formulas', 'zz-blocked'], entry),
            101,
            (entry, back) => {
              if (!back && entry === 'formulas') writeFileSync(join(dir, 'formulas'), 'other');
            }
          );
        } catch (caught) {
          error = caught;
        }
        const kept = readdirSync(dir).filter((e) => e.startsWith(`${LOCK_NAME}.cleared.101.`));
        expect(kept).toHaveLength(1);
        const folder = join(dir, kept[0]!);
        expect(String(error)).toEndWith(
          `, and formulas could not move back: they are in ${folder}; move its entries back into ${dir}, or remove ${folder}`
        );
        expect(readdirSync(folder)).toEqual(['formulas']);
        expect(readFileSync(join(folder, 'formulas', 'ours.png'), 'utf8')).toBe('old');
        expect(readFileSync(join(dir, 'report.md'), 'utf8')).toBe('old');
        // The folder has no mark of a finished clear, so a later run keeps it, also though its
        // process 101 is gone, and names it in a notice.
        const notices: string[] = [];
        const lock = takeRunLock(dir, 102, () => false, Date.now, {
          notice: (text) => void notices.push(text),
        });
        expect(readdirSync(dir)).toContain(kept[0]!);
        expect(notices).toEqual([
          `${folder} holds report entries that a clear did not finish with; move its entries back into ${dir}, or remove it: chmod -R u+rwx '${folder}' && rm -r '${folder}'`,
        ]);
        if ('release' in lock) lock.release();
      } finally {
        chmodSync(fixture, 0o700);
      }
    }
  );

  test.skipIf(process.getuid?.() === 0)(
    'a cleared folder that cannot be removed keeps its mark, and the error gives the command that removes it',
    () => {
      const dir = mkdtempSync(join(tmpdir(), 'xlsx-clear-'));
      dirs.push(dir);
      // The entry moves aside, but a nested folder without any access keeps its file.
      const nested = join(dir, 'formulas', 'nested');
      mkdirSync(nested, { recursive: true });
      writeFileSync(join(nested, 'ours.png'), 'old');
      chmodSync(nested, 0o000);
      let message = '';
      try {
        clearReport(dir, (entry) => entry === 'formulas', 101);
      } catch (caught) {
        message = (caught as Error).message;
      }
      const folders = readdirSync(dir).filter((e) => e.startsWith(`${LOCK_NAME}.cleared.101.`));
      try {
        expect(folders).toHaveLength(1);
        const folder = join(dir, folders[0]!);
        const step = `chmod -R u+rwx '${folder}' && rm -r '${folder}'`;
        // The path can hold characters that a regular expression reads: compare strings.
        expect(
          ['EACCES', 'EPERM', 'ENOTEMPTY'].map(
            (code) =>
              `${folder} cannot be removed (${code}); it holds only report entries that this run cleared from ${dir}, so remove it, then run again: ${step}`
          )
        ).toContain(message);
        // The mark goes last, so the folder keeps it. A later run cannot remove it either: it
        // keeps the folder and its mark, and prints a notice with the same command.
        expect(readdirSync(folder)).toContain(CLEARED_MARK);
        const notices: string[] = [];
        const again = takeRunLock(dir, 102, () => false, Date.now, {
          notice: (text) => void notices.push(text),
        });
        if ('release' in again) again.release();
        expect(readdirSync(folder)).toContain(CLEARED_MARK);
        expect(notices).toHaveLength(1);
        expect(notices[0]).toEndWith(`; remove it: ${step}`);
        // The command removes the folder.
        const removed = spawnSync('sh', ['-c', step]);
        expect(removed.status).toBe(0);
        expect(readdirSync(dir)).toEqual([]);
      } finally {
        for (const f of folders) {
          try {
            chmodSync(join(dir, f, 'formulas', 'nested'), 0o700);
          } catch {
            // Removed already.
          }
        }
      }
    }
  );

  test('a command names a path as one shell word, also with quotes and dollar signs', () => {
    for (const path of ['/tmp/a b', '/tmp/a"b', '/tmp/$HOME', "/tmp/it's"]) {
      const echoed = spawnSync('sh', ['-c', `printf %s ${shellPath(path)}`], { encoding: 'utf8' });
      expect(echoed.stdout).toBe(path);
    }
  });

  test('a clear that cannot write its mark moves every entry back', () => {
    const dir = mkdtempSync(join(tmpdir(), 'xlsx-clear-'));
    dirs.push(dir);
    writeFileSync(join(dir, 'report.md'), 'old');
    mkdirSync(join(dir, 'formulas'));
    writeFileSync(join(dir, 'formulas', 'ours.png'), 'old');
    const clears = (entry: string) => clearsEntry(['formulas'], entry);
    // A folder at the name of the mark makes its write fail, after the last entry moved aside.
    const blockMark = (entry: string, back: boolean): void => {
      if (back || entry !== 'report.md') return;
      const aside = readdirSync(dir).find((e) => e.startsWith(`${LOCK_NAME}.cleared.101.`))!;
      mkdirSync(join(dir, aside, CLEARED_MARK));
    };
    let message = '';
    try {
      clearReport(dir, clears, 101, blockMark);
    } catch (caught) {
      message = (caught as Error).message;
    }
    expect(message).toContain(` cannot take its mark ${CLEARED_MARK} (`);
    expect(message).toEndWith(`; free disk space or give write access to ${dir}, then run again`);
    // All or nothing: both entries are back, and no aside folder is left.
    expect(readdirSync(dir).sort()).toEqual(['formulas', 'report.md']);
    expect(readFileSync(join(dir, 'report.md'), 'utf8')).toBe('old');
    expect(readFileSync(join(dir, 'formulas', 'ours.png'), 'utf8')).toBe('old');
    // When an entry cannot move back either, it stays in the aside folder, which no run removes:
    // its mark is no file. Here a file takes the name of `formulas` while it is aside.
    let stuck = '';
    try {
      clearReport(dir, clears, 101, (entry, back) => {
        if (!back && entry === 'formulas') writeFileSync(join(dir, 'formulas'), 'other');
        blockMark(entry, back);
      });
    } catch (caught) {
      stuck = (caught as Error).message;
    }
    const kept = readdirSync(dir).filter((e) => e.startsWith(`${LOCK_NAME}.cleared.101.`));
    expect(kept).toHaveLength(1);
    const folder = join(dir, kept[0]!);
    expect(stuck).toBe(
      `${folder} cannot take its mark ${CLEARED_MARK} (EISDIR), and formulas could not move back: they are in ${folder}; move its entries back into ${dir}, or remove ${folder}`
    );
    expect(readFileSync(join(folder, 'formulas', 'ours.png'), 'utf8')).toBe('old');
    expect(readFileSync(join(dir, 'report.md'), 'utf8')).toBe('old');
    const notices: string[] = [];
    const lock = takeRunLock(dir, 102, () => false, Date.now, {
      notice: (text) => void notices.push(text),
    });
    expect(readdirSync(dir)).toContain(kept[0]!);
    expect(notices).toEqual([
      `${folder} holds report entries that a clear did not finish with; move its entries back into ${dir}, or remove it: chmod -R u+rwx '${folder}' && rm -r '${folder}'`,
    ]);
    if ('release' in lock) lock.release();
  });
});
