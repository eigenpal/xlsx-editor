import { afterEach, describe, expect, test } from 'bun:test';
import { spawn, spawnSync } from 'node:child_process';
import {
  chmodSync,
  existsSync,
  mkdirSync,
  mkdtempSync,
  readFileSync,
  realpathSync,
  rmSync,
  symlinkSync,
  truncateSync,
  writeFileSync,
} from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join, resolve } from 'node:path';
import {
  checkOut,
  RUNNER_ENV,
  corpusDir,
  CORPUS_ENV,
  CORPUS_TMP_ENV,
  listWorkbooks,
  MEMORY_CAP_BYTES,
  readChildLine,
  redactPaths,
  ROOT,
  runCorpus,
  type CorpusJob,
  type CorpusRefusal,
  type CorpusRun,
  type RunOptions,
} from './runner.ts';
import { MAX_FILE_BYTES } from './worker.ts';

const TEST_WORKER = join(import.meta.dir, 'test-worker.ts');
const RUNNER = join(import.meta.dir, 'runner.ts');
const dirs: string[] = [];
const scratch = (): string => {
  const dir = mkdtempSync(join(tmpdir(), 'xlsx-corpus-test-'));
  dirs.push(dir);
  return dir;
};
afterEach(() => {
  for (const dir of dirs.splice(0)) rmSync(dir, { recursive: true, force: true });
});

/** A folder of stand-in workbooks: each file names what the test child does with it. */
function corpus(behaviors: readonly string[]): { dir: string; jobs: CorpusJob[] } {
  const dir = scratch();
  behaviors.forEach((b, i) => writeFileSync(join(dir, `w${String(i).padStart(2, '0')}.xlsx`), b));
  return { dir, jobs: listWorkbooks(dir) };
}

/** A git work tree that ignores `ignored/`. */
function repository(): string {
  const dir = scratch();
  // The `GIT_*` variables of a hook must not point git at another repository.
  const env = Object.fromEntries(
    Object.entries(process.env).filter(([name]) => !name.startsWith('GIT_'))
  );
  spawnSync('git', ['init', '-q', dir], { env });
  writeFileSync(join(dir, '.gitignore'), 'ignored/\n');
  return dir;
}

async function run(options: RunOptions): Promise<CorpusRun> {
  const result = await runCorpus(options);
  if ('refused' in result) throw new Error(result.refused);
  return result;
}

const lines = (file: string): Record<string, unknown>[] =>
  readFileSync(file, 'utf8')
    .split('\n')
    .filter(Boolean)
    .map((l) => JSON.parse(l) as Record<string, unknown>);

const alive = (pid: number): boolean => {
  try {
    process.kill(pid, 0);
    return true;
  } catch {
    return false;
  }
};

/** Waits up to `ms` for `check`, and returns its last answer. */
async function waitFor(check: () => boolean, ms: number): Promise<boolean> {
  const end = Date.now() + ms;
  while (Date.now() < end) {
    if (check()) return true;
    await Bun.sleep(50);
  }
  return check();
}

/** Runs `body` with `name` set to `value` in the environment, and restores it after. */
async function withEnv<T>(name: string, value: string, body: () => Promise<T> | T): Promise<T> {
  const before = process.env[name];
  process.env[name] = value;
  try {
    return await body();
  } finally {
    if (before === undefined) delete process.env[name];
    else process.env[name] = before;
  }
}

/** Stops a process group that a failed check left running. */
function stopGroup(pid: number | undefined): void {
  if (pid === undefined) return;
  try {
    process.kill(-pid, 'SIGKILL');
  } catch {
    // The group is gone.
  }
}

/** A script that runs the corpus runner over `jobs` in its own process. */
function runnerScript(
  work: string,
  jobs: readonly CorpusJob[],
  batchSize?: number,
  workbookMs = 60_000,
  capBytes?: number
): string {
  const script = join(work, 'run.ts');
  writeFileSync(
    script,
    `import { runCorpus } from ${JSON.stringify(RUNNER)};
await runCorpus(${JSON.stringify({ task: 'test', jobs, out: join(work, 'out'), worker: TEST_WORKER, workbookMs, everyMs: 25, batchSize, capBytes })});`
  );
  return script;
}

/** The exit code of a process, or `'late'` when it has not ended within `ms`. */
const exitWithin = (exited: Promise<number | null>, ms: number) =>
  Promise.race([exited, Bun.sleep(ms).then(() => 'late' as const)]);

/** The first letter of the state of a process in the process table: `T` when it is paused. */
const stateOf = (pid: number): string =>
  spawnSync('ps', ['-o', 'stat=', '-p', String(pid)], { encoding: 'utf8' })
    .stdout.trim()
    .slice(0, 1);

/** The child, grandchild, and work folder that the `grandchild` case wrote. */
function grandchildRecord(file: string): { child: number; grandchild: number; work: string } {
  const [child, grandchild, work] = readFileSync(file, 'utf8').trim().split(' ');
  return { child: Number(child), grandchild: Number(grandchild), work: work! };
}

describe('the workbook list', () => {
  test('takes the .xlsx files of every folder, sorted, and names them by position', () => {
    const dir = scratch();
    mkdirSync(join(dir, 'b'));
    mkdirSync(join(dir, '.hidden'));
    writeFileSync(join(dir, 'b', 'one.xlsx'), '');
    writeFileSync(join(dir, 'a.XLSX'), '');
    writeFileSync(join(dir, 'notes.txt'), '');
    writeFileSync(join(dir, '.hidden', 'x.xlsx'), '');
    // The owner file that a program keeps next to a workbook that it edits.
    writeFileSync(join(dir, '~$a.xlsx'), '');
    symlinkSync(join(dir, 'a.XLSX'), join(dir, 'link.xlsx'));
    const jobs = listWorkbooks(dir);
    expect(jobs.map((j) => j.id)).toEqual(['workbook1', 'workbook2']);
    expect(jobs.map((j) => j.path)).toEqual([join(dir, 'a.XLSX'), join(dir, 'b', 'one.xlsx')]);
  });

  test('names no folder when a folder cannot be read', () => {
    const missing = join(scratch(), 'Secret folder');
    expect(() => listWorkbooks(missing)).toThrow('A folder of the corpus cannot be read (ENOENT).');
  });
});

describe('the corpus folder', () => {
  test('has no default', () => {
    expect(() => corpusDir([], {})).toThrow('No folder is assumed');
  });

  test('comes from --dir before the environment', () => {
    const a = scratch();
    const b = scratch();
    expect(corpusDir(['--dir', a], { [CORPUS_ENV]: b })).toBe(a);
    expect(corpusDir([], { [CORPUS_ENV]: b })).toBe(b);
  });

  test('must exist', () => {
    expect(() => corpusDir(['--dir', join(scratch(), 'missing')], {})).toThrow('does not exist');
  });
});

describe('the output folder', () => {
  test('may be a folder that git ignores, or outside every work tree', () => {
    const repo = repository();
    expect(checkOut(join(repo, 'ignored', 'census'))).toBe(join(repo, 'ignored', 'census'));
    expect(checkOut(join(ROOT, 'local', 'corpus', 'census'))).toBe(
      join(ROOT, 'local', 'corpus', 'census')
    );
    const outside = scratch();
    expect(checkOut(outside)).toBe(outside);
  });

  test('reads the answer of git in any locale', async () => {
    const outside = scratch();
    await withEnv('LC_ALL', 'de_DE.UTF-8', () =>
      withEnv('LANGUAGE', 'de', () => {
        expect(checkOut(outside)).toBe(outside);
      })
    );
  });

  test('may not be a folder that git can commit, in any spelling of its path', () => {
    const repo = repository();
    expect(() => checkOut(repo)).toThrow('git ignores');
    expect(() => checkOut(join(repo, 'docs', 'out'))).toThrow('git ignores');
    // Through a symbolic link, and through the real path of the temporary folder.
    const link = join(scratch(), 'link');
    symlinkSync(repo, link);
    expect(() => checkOut(join(link, 'docs'))).toThrow('git ignores');
    expect(() => checkOut(join(realpathSync(repo), 'docs'))).toThrow('git ignores');
    // In another letter case, on a disk that ignores case.
    const upper = repo.toUpperCase();
    if (existsSync(upper)) expect(() => checkOut(join(upper, 'docs'))).toThrow('git ignores');
  });

  test('asks git about the folder itself, whatever GIT_DIR the caller sets', async () => {
    const repo = repository();
    // With this GIT_DIR, git would call every folder "not a git repository".
    await withEnv('GIT_DIR', join(scratch(), 'none'), () => {
      expect(() => checkOut(join(repo, 'docs'))).toThrow('git ignores');
    });
  });

  test('asks git about each file that a run writes', () => {
    const repo = repository();
    writeFileSync(join(repo, '.gitignore'), '*.jsonl\n');
    expect(() => checkOut(join(repo, 'reports'))).toThrow('git ignores');
  });

  test('may not hold an output file that git tracks, even in an ignored folder', () => {
    const repo = repository();
    mkdirSync(join(repo, 'ignored', 'out'), { recursive: true });
    writeFileSync(join(repo, 'ignored', 'out', 'run.json'), '{}');
    const env = Object.fromEntries(
      Object.entries(process.env).filter(([name]) => !name.startsWith('GIT_'))
    );
    spawnSync('git', ['-C', repo, 'add', '-f', 'ignored/out/run.json'], { env });
    expect(checkOut(join(repo, 'ignored', 'other'))).toBe(join(repo, 'ignored', 'other'));
    expect(() => checkOut(join(repo, 'ignored', 'out'))).toThrow('git ignores');
  });

  test('is refused when git cannot answer', async () => {
    const repo = repository();
    // With no git on the path, the check cannot tell where the folder is.
    await withEnv('PATH', scratch(), () => {
      expect(() => checkOut(join(repo, 'ignored', 'out'))).toThrow('git could not tell');
    });
  });

  test('may not be under a file', () => {
    const file = join(scratch(), 'file');
    writeFileSync(file, '');
    expect(() => checkOut(join(file, 'out'))).toThrow('a folder that can be made and written');
  });
});

describe('the lines of a child', () => {
  const batch = new Set(['workbook1']);

  test('takes a start and a result of its batch', () => {
    expect(readChildLine('{"type":"start","id":"workbook1"}', batch)).toEqual({
      type: 'start',
      id: 'workbook1',
    });
    expect(readChildLine('{"type":"result","id":"workbook1","result":{"a":1}}', batch)).toEqual({
      type: 'result',
      id: 'workbook1',
      result: { a: 1 },
    });
  });

  test('leaves out a line that is not JSON, not of the batch, or has no result', () => {
    expect(readChildLine('workbook1', batch)).toBeUndefined();
    expect(readChildLine('{"type":"start","id":"workbook2"}', batch)).toBeUndefined();
    expect(readChildLine('{"type":"result","id":"workbook1"}', batch)).toBeUndefined();
    expect(readChildLine('null', batch)).toBeUndefined();
  });
});

test('a message loses its paths and addresses', () => {
  expect(redactPaths('cannot open /Users/someone/corpus/Secret plan.xlsx now')).toBe(
    'cannot open <path>'
  );
  expect(redactPaths('fetch https://example.com/a failed\nsecond line')).toBe(
    'fetch <address> failed'
  );
});

describe('a corpus run', () => {
  test('keeps the memory cap of one child at 3 GiB', () => {
    expect(MEMORY_CAP_BYTES).toBe(3 * 1024 ** 3);
  });

  test('stops a child over the memory cap or the time, and goes on with the next workbooks', async () => {
    const { dir, jobs } = corpus(['ok', 'allocate', 'ok', 'hang', 'ok']);
    const out = scratch();
    const done = await run({
      task: 'test',
      jobs,
      out,
      worker: TEST_WORKER,
      batchSize: 3,
      capBytes: 192 * 1024 * 1024,
      // Only the watchdog may stop the child here, not the guard thread of the child.
      guardBytes: Number.MAX_SAFE_INTEGER,
      everyMs: 25,
      workbookMs: 4000,
    });
    expect(done.stopped).toEqual([
      { id: 'workbook2', stopped: 'memory' },
      { id: 'workbook4', stopped: 'time' },
    ]);
    // Workbook 2 runs once more alone, and the workbooks of a batch that did not start run in the
    // next child.
    expect(done.children).toBe(4);
    expect(done.results).toBe(3);
    const written = lines(done.resultsFile);
    expect(written.map((l) => l.id)).toEqual([
      'workbook1',
      'workbook2',
      'workbook3',
      'workbook4',
      'workbook5',
    ]);
    const children = new Set(written.filter((l) => l.child).map((l) => l.child));
    expect(children.size).toBe(3);
    // The output names no file and no folder of the corpus.
    const text =
      readFileSync(done.resultsFile, 'utf8') + readFileSync(join(out, 'run.json'), 'utf8');
    expect(text).not.toContain(dir);
    expect(text).not.toContain('.xlsx');
    expect(JSON.parse(readFileSync(join(out, 'run.json'), 'utf8'))).toMatchObject({
      task: 'test',
      workbooks: 5,
      results: 3,
      stopped: 2,
      children: 4,
    });
  }, 30_000);

  test('runs a workbook alone again when the memory of earlier workbooks tripped the cap', async () => {
    // Each workbook keeps 64 MiB. Alone, one stays under the cap; three in one child pass it.
    const { jobs } = corpus(['keep 1', 'keep 2', 'keep 3', 'keep 4', 'keep 5', 'keep 6']);
    const pids = join(scratch(), 'pids');
    const done = await withEnv('CORPUS_TEST_PIDS', pids, () =>
      run({
        task: 'test',
        jobs,
        out: scratch(),
        worker: TEST_WORKER,
        batchSize: 6,
        capBytes: 200 * 1024 * 1024,
        everyMs: 25,
        workbookMs: 10_000,
      })
    );
    expect(done.stopped).toEqual([]);
    expect(done.results).toBe(6);
    // Each start line names the child and the workbook. A workbook that started twice was retried.
    const starts = readFileSync(pids, 'utf8')
      .trim()
      .split('\n')
      .map((line) => {
        const [pid, ...name] = line.split(' ');
        return { pid: pid!, name: name.join(' ') };
      });
    const again = starts.filter((s, i) => starts.findIndex((t) => t.name === s.name) !== i);
    expect(again.length).toBeGreaterThan(0);
    // The child of each second start ran that workbook alone.
    for (const start of again) expect(starts.filter((s) => s.pid === start.pid)).toEqual([start]);
  }, 30_000);

  test('stops a workbook that crosses the cap first in its child, with no retry', async () => {
    const { jobs } = corpus(['allocate', 'ok']);
    const done = await run({
      task: 'test',
      jobs,
      out: scratch(),
      worker: TEST_WORKER,
      batchSize: 2,
      capBytes: 192 * 1024 * 1024,
      guardBytes: Number.MAX_SAFE_INTEGER,
      everyMs: 25,
      workbookMs: 10_000,
    });
    expect(done.stopped).toEqual([{ id: 'workbook1', stopped: 'memory' }]);
    expect(done).toMatchObject({ results: 1, children: 2 });
  }, 30_000);

  test('keeps the id of the workbook first, whatever the task returns', async () => {
    const { jobs } = corpus(['forge-id']);
    const done = await run({
      task: 'test',
      jobs,
      out: scratch(),
      worker: TEST_WORKER,
      workbookMs: 10_000,
    });
    const text = readFileSync(done.resultsFile, 'utf8');
    expect(text).toStartWith('{"id":"workbook1",');
    expect(lines(done.resultsFile)).toEqual([{ id: 'workbook1', behavior: 'forge-id' }]);
  }, 30_000);

  test.each([
    ['a quarter above the cap by default', undefined, String(1.25 * 64 * 1024 * 1024)],
    ['the limit that the run names', 1000, '1000'],
  ] as const)(
    'names itself to its child, with a guard limit %s',
    async (_, guardBytes, limit) => {
      const { jobs } = corpus(['runner-env']);
      const done = await run({
        task: 'test',
        jobs,
        out: scratch(),
        worker: TEST_WORKER,
        workbookMs: 10_000,
        capBytes: 64 * 1024 * 1024,
        guardBytes,
      });
      expect(lines(done.resultsFile)).toEqual([
        { id: 'workbook1', behavior: 'runner-env', runner: String(process.pid), guard: limit },
      ]);
    },
    30_000
  );

  test('removes the run.json of an earlier run before the first result', async () => {
    const { jobs } = corpus(['ok']);
    const out = scratch();
    writeFileSync(join(out, 'run.json'), '{"old":true}');
    let seen: boolean | undefined;
    await run({
      task: 'test',
      jobs,
      out,
      worker: TEST_WORKER,
      workbookMs: 10_000,
      onProgress: () => {
        seen ??= existsSync(join(out, 'run.json'));
      },
    });
    expect(seen).toBe(false);
    expect(JSON.parse(readFileSync(join(out, 'run.json'), 'utf8'))).toMatchObject({ results: 1 });
  }, 30_000);

  test('pauses only on Ctrl+Z, and keeps the default of the terminal signals', async () => {
    const { jobs } = corpus(['ok']);
    let counts: number[] | undefined;
    await run({
      task: 'test',
      jobs,
      out: scratch(),
      worker: TEST_WORKER,
      workbookMs: 10_000,
      onProgress: () => {
        // A handler for a terminal read or write of a background run makes the blocked write
        // repeat at once: the runner would spin and could not watch its child.
        counts ??= ['SIGTTIN', 'SIGTTOU', 'SIGTSTP'].map((name) => process.listenerCount(name));
      },
    });
    expect(counts).toEqual([0, 0, 1]);
  }, 30_000);

  test('starts the time limit again for each workbook of a batch', async () => {
    // Each workbook takes 1.5 s and the limit is 2.5 s: only a limit for each workbook passes.
    const { jobs } = corpus(['slow', 'slow']);
    const done = await run({
      task: 'test',
      jobs,
      out: scratch(),
      worker: TEST_WORKER,
      everyMs: 25,
      workbookMs: 2500,
    });
    expect(done).toMatchObject({ results: 2, stopped: [], children: 1 });
  }, 30_000);

  test('records a child that ends or floods its output, and a child that never starts', async () => {
    const { jobs } = corpus(['exit', 'ok', 'flood', 'ok', 'refuse-start', 'ok']);
    const done = await run({
      task: 'test',
      jobs,
      out: scratch(),
      worker: TEST_WORKER,
      batchSize: 2,
      everyMs: 25,
      workbookMs: 10_000,
    });
    // Workbook 2 runs again with workbook 3. Batch [4, 5] ends before its first workbook, so both
    // stop: a retry would fail the same way.
    expect(done.stopped).toEqual([
      { id: 'workbook1', stopped: 'exit' },
      { id: 'workbook3', stopped: 'output' },
      { id: 'workbook4', stopped: 'exit' },
      { id: 'workbook5', stopped: 'exit' },
    ]);
    expect(done.results).toBe(2);
    // [1, 2], [2, 3], [4, 5], and [6]: workbook 5 does not run again.
    expect(done.children).toBe(4);
  }, 30_000);

  test('records the error of a task without its paths', async () => {
    const { jobs } = corpus(['throw']);
    const done = await run({
      task: 'test',
      jobs,
      out: scratch(),
      worker: TEST_WORKER,
      workbookMs: 10_000,
    });
    expect(lines(done.resultsFile)).toEqual([{ error: 'cannot read <path>', id: 'workbook1' }]);
  }, 30_000);

  test('refuses a file larger than the reader takes, before it reads it', async () => {
    const { dir } = corpus([]);
    // A sparse file: it takes no space on the disk.
    writeFileSync(join(dir, 'big.xlsx'), '');
    truncateSync(join(dir, 'big.xlsx'), MAX_FILE_BYTES + 1);
    const done = await run({
      task: 'test',
      jobs: listWorkbooks(dir),
      out: scratch(),
      worker: TEST_WORKER,
      workbookMs: 10_000,
    });
    expect(lines(done.resultsFile)).toEqual([
      { id: 'workbook1', open: 'refused', refusal: 'file-too-large' },
    ]);
  }, 30_000);

  test('stops the processes that a child started, and removes its work folder', async () => {
    const { jobs } = corpus(['grandchild']);
    const pids = join(scratch(), 'pids');
    const done = await withEnv('CORPUS_TEST_PIDS', pids, () =>
      run({
        task: 'test',
        jobs,
        out: scratch(),
        worker: TEST_WORKER,
        everyMs: 25,
        workbookMs: 1500,
      })
    );
    const record = grandchildRecord(pids);
    try {
      expect(done.stopped).toEqual([{ id: 'workbook1', stopped: 'time' }]);
      expect(await waitFor(() => !alive(record.child) && !alive(record.grandchild), 2000)).toBe(
        true
      );
      expect(existsSync(record.work)).toBe(false);
    } finally {
      // When a check fails, the processes and the folder of the child may stay.
      stopGroup(record.child);
      if (alive(record.grandchild)) process.kill(record.grandchild, 'SIGKILL');
      if (dirname(record.work) === resolve(tmpdir()))
        rmSync(record.work, { recursive: true, force: true });
    }
  }, 30_000);

  test('stops a process that left the tree of the child but not its group', async () => {
    const { jobs } = corpus(['orphan']);
    const pids = join(scratch(), 'pids');
    const done = await withEnv('CORPUS_TEST_PIDS', pids, () =>
      run({
        task: 'test',
        jobs,
        out: scratch(),
        worker: TEST_WORKER,
        everyMs: 25,
        workbookMs: 1500,
      })
    );
    expect(done.stopped).toEqual([{ id: 'workbook1', stopped: 'time' }]);
    const orphan = Number(readFileSync(pids, 'utf8').trim());
    try {
      expect(orphan).toBeGreaterThan(0);
      expect(await waitFor(() => !alive(orphan), 2000)).toBe(true);
    } finally {
      if (alive(orphan)) process.kill(orphan, 'SIGKILL');
    }
  }, 30_000);

  test.each([
    ['SIGTERM', 'SIGTERM', 143, 'grandchild'],
    ['SIGHUP', 'SIGHUP', 129, 'grandchild'],
    ['SIGINT', 'SIGINT', 130, 'grandchild'],
    // The folder cannot be removed: the runner names it, and still exits.
    ['SIGTERM with a folder that cannot be removed', 'SIGTERM', 143, 'grandchild-locked'],
  ] as const)(
    'stops the child, records nothing more, and exits on %s',
    async (_, signal, code, behavior) => {
      // Two batches: after the signal, the runner must not start the second.
      const { jobs } = corpus([behavior, behavior]);
      const work = scratch();
      const pids = join(work, 'pids');
      const runner = spawn(process.execPath, [runnerScript(work, jobs, 1)], {
        env: { ...process.env, CORPUS_TEST_PIDS: pids },
        stdio: ['ignore', 'ignore', 'pipe'],
      });
      let stderr = '';
      runner.stderr!.setEncoding('utf8');
      runner.stderr!.on('data', (chunk: string) => (stderr += chunk));
      const exited = new Promise<number | null>((done) => runner.on('close', (code) => done(code)));
      let child: number | undefined;
      let folder: string | undefined;
      try {
        expect(await waitFor(() => existsSync(pids), 10_000)).toBe(true);
        const record = grandchildRecord(pids);
        child = record.child;
        folder = dirname(record.work) === resolve(tmpdir()) ? record.work : undefined;
        // The child removed the list of the workbooks when it read it.
        expect(existsSync(join(record.work, 'jobs.json'))).toBe(false);
        runner.kill(signal);
        expect(await exitWithin(exited, 15_000)).toBe(code);
        expect(await waitFor(() => !alive(record.child) && !alive(record.grandchild), 2000)).toBe(
          true
        );
        const locked = behavior === 'grandchild-locked';
        expect(existsSync(record.work)).toBe(locked);
        // The list of the workbooks goes in any case.
        expect(existsSync(join(record.work, 'jobs.json'))).toBe(false);
        // One line that names the results file, after one that names a folder that stays.
        const message = stderr.trim().split('\n');
        expect(message).toHaveLength(locked ? 2 : 1);
        if (locked) expect(message[0]).toContain(`work folder ${record.work}`);
        expect(message.at(-1)).toStartWith(`Stopped by ${signal}. `);
        expect(message.at(-1)).toContain(join('out', 'results.jsonl'));
        // The user stopped the run: no workbook stopped by itself, and no other child started.
        expect(readFileSync(join(work, 'out', 'results.jsonl'), 'utf8')).toBe('');
        await Bun.sleep(500);
        expect(readFileSync(pids, 'utf8').trim().split('\n')).toHaveLength(1);
      } finally {
        runner.kill('SIGKILL');
        stopGroup(child);
        // When the check fails, or the folder is locked, the folder of the child stays.
        if (folder) {
          if (existsSync(join(folder, 'locked'))) chmodSync(join(folder, 'locked'), 0o700);
          rmSync(folder, { recursive: true, force: true });
        }
      }
    },
    30_000
  );

  test('ends the child on a pause, and runs its workbooks again after it', async () => {
    const { jobs } = corpus(['slow', 'ok']);
    const work = scratch();
    const pids = join(work, 'pids');
    const runner = spawn(process.execPath, [runnerScript(work, jobs)], {
      env: { ...process.env, CORPUS_TEST_PIDS: pids },
      stdio: 'ignore',
    });
    const exited = new Promise<number | null>((done) => runner.on('close', (code) => done(code)));
    let child: number | undefined;
    try {
      expect(await waitFor(() => existsSync(pids), 10_000)).toBe(true);
      child = Number(readFileSync(pids, 'utf8').trim());
      runner.kill('SIGTSTP');
      // The paused runner holds no child.
      // The paused runner cannot collect its ended child, so the child stays as a zombie (`Z`).
      expect(
        await waitFor(
          () => stateOf(runner.pid!) === 'T' && (!alive(child!) || stateOf(child!) === 'Z'),
          3000
        )
      ).toBe(true);
      // The ended child recorded no result.
      const results = join(work, 'out', 'results.jsonl');
      expect(existsSync(results) ? lines(results) : []).toEqual([]);
      runner.kill('SIGCONT');
      expect(await exitWithin(exited, 15_000)).toBe(0);
      // The workbook that ran starts again in a new child, and no workbook stops.
      expect(readFileSync(pids, 'utf8').trim().split('\n')).toHaveLength(2);
      expect(lines(join(work, 'out', 'results.jsonl'))).toEqual([
        expect.objectContaining({ id: 'workbook1', behavior: 'slow' }),
        expect.objectContaining({ id: 'workbook2', behavior: 'ok' }),
      ]);
      expect(JSON.parse(readFileSync(join(work, 'out', 'run.json'), 'utf8'))).toMatchObject({
        children: 2,
        results: 2,
        stopped: 0,
      });
    } finally {
      runner.kill('SIGKILL');
      stopGroup(child);
    }
  }, 30_000);

  test('a paused run leaves no child and no work folder when its runner is killed', async () => {
    const { jobs } = corpus(['grandchild']);
    const work = scratch();
    const pids = join(work, 'pids');
    const runner = spawn(process.execPath, [runnerScript(work, jobs)], {
      env: { ...process.env, CORPUS_TEST_PIDS: pids },
      stdio: 'ignore',
    });
    let record: ReturnType<typeof grandchildRecord> | undefined;
    try {
      expect(await waitFor(() => existsSync(pids), 10_000)).toBe(true);
      record = grandchildRecord(pids);
      runner.kill('SIGTSTP');
      expect(await waitFor(() => stateOf(runner.pid!) === 'T', 3000)).toBe(true);
      runner.kill('SIGKILL');
      expect(await waitFor(() => !alive(record!.child) && !alive(record!.grandchild), 3000)).toBe(
        true
      );
      expect(existsSync(record.work)).toBe(false);
    } finally {
      runner.kill('SIGKILL');
      stopGroup(record?.child);
      if (record && dirname(record.work) === resolve(tmpdir()))
        rmSync(record.work, { recursive: true, force: true });
    }
  }, 30_000);

  test('a child that writes after its runner is gone removes its work folder and stops', async () => {
    const { jobs } = corpus(['after-runner']);
    const work = scratch();
    const pids = join(work, 'pids');
    const runner = spawn(process.execPath, [runnerScript(work, jobs)], {
      env: { ...process.env, CORPUS_TEST_PIDS: pids },
      stdio: 'ignore',
    });
    let record: ReturnType<typeof grandchildRecord> | undefined;
    try {
      expect(await waitFor(() => existsSync(pids), 10_000)).toBe(true);
      record = grandchildRecord(pids);
      runner.kill('SIGKILL');
      expect(await waitFor(() => !alive(record!.child), 3000)).toBe(true);
      // A child that crashed on the closed output would leave its folder.
      expect(existsSync(record.work)).toBe(false);
    } finally {
      runner.kill('SIGKILL');
      stopGroup(record?.child);
      if (record && dirname(record.work) === resolve(tmpdir()))
        rmSync(record.work, { recursive: true, force: true });
    }
  }, 30_000);

  test('the child stops itself far above the cap while its runner cannot watch', async () => {
    const { jobs } = corpus(['allocate-later']);
    const work = scratch();
    const pids = join(work, 'pids');
    const runner = spawn(
      process.execPath,
      [runnerScript(work, jobs, undefined, 60_000, 192 * 1024 * 1024)],
      { env: { ...process.env, CORPUS_TEST_PIDS: pids }, stdio: 'ignore' }
    );
    const exited = new Promise<number | null>((done) => runner.on('close', (code) => done(code)));
    let record: ReturnType<typeof grandchildRecord> | undefined;
    try {
      expect(await waitFor(() => existsSync(pids), 10_000)).toBe(true);
      record = grandchildRecord(pids);
      // The runner and its watchdog stop; only the guard thread of the child is left to act.
      runner.kill('SIGSTOP');
      expect(await waitFor(() => stateOf(runner.pid!) === 'T', 3000)).toBe(true);
      writeFileSync(`${pids}.go`, '');
      // The child cannot be collected while its runner is paused, so it stays as a zombie.
      expect(
        await waitFor(() => !alive(record!.child) || stateOf(record!.child) === 'Z', 15_000)
      ).toBe(true);
      expect(existsSync(record.work)).toBe(false);
      // The runner records the stop of the guard as a memory stop.
      runner.kill('SIGCONT');
      expect(await exitWithin(exited, 15_000)).toBe(0);
      expect(lines(join(work, 'out', 'results.jsonl'))).toEqual([
        { id: 'workbook1', stopped: 'memory' },
      ]);
    } finally {
      runner.kill('SIGKILL');
      stopGroup(record?.child);
      if (record && dirname(record.work) === resolve(tmpdir()))
        rmSync(record.work, { recursive: true, force: true });
    }
  }, 30_000);

  test('fails with the error when no stop signal came', async () => {
    const { jobs } = corpus(['ok']);
    const out = scratch();
    const missing = join(scratch(), 'missing');
    // The temporary folder does not exist, so the work folder of the child cannot be made.
    await withEnv('TMPDIR', missing, async () => {
      // A run that waits forever fails here, and does not hold the test file.
      const hung = Bun.sleep(10_000).then(() => 'the run did not end');
      await expect(
        Promise.race([
          runCorpus({ task: 'test', jobs, out, worker: TEST_WORKER, workbookMs: 10_000 }),
          hung,
        ])
      ).rejects.toThrow('ENOENT');
    });
  }, 30_000);

  test('a child whose runner ended during its last workbook removes its work folder and stops', async () => {
    const { jobs } = corpus(['ok']);
    // A work folder of the form that the runner makes, so the child may remove it.
    const dir = mkdtempSync(join(resolve(tmpdir()), 'xlsx-corpus-'));
    const jobsFile = join(dir, 'jobs.json');
    writeFileSync(jobsFile, JSON.stringify(jobs));
    // The output goes nowhere and no write fails, so only the check after the last workbook
    // sees that the runner is gone. A guard thread that looks first gives the same result.
    const child = spawn(process.execPath, [TEST_WORKER, 'test', jobsFile], {
      detached: true,
      env: { ...process.env, TMPDIR: dirname(dir), [RUNNER_ENV]: '1', [CORPUS_TMP_ENV]: dir },
      stdio: 'ignore',
    });
    const ended = new Promise<string | null>((done) =>
      child.on('close', (_, signal) => done(signal))
    );
    try {
      expect(await Promise.race([ended, Bun.sleep(10_000).then(() => 'late')])).toBe('SIGKILL');
      expect(existsSync(dir)).toBe(false);
    } finally {
      stopGroup(child.pid);
      rmSync(dir, { recursive: true, force: true });
    }
  }, 30_000);

  test('a child stops at once when the runner that the environment names is gone', async () => {
    const { jobs } = corpus(['hang']);
    const jobsFile = join(scratch(), 'jobs.json');
    writeFileSync(jobsFile, JSON.stringify(jobs));
    // Its parent is this test process, but the runner that it was given does not exist.
    const child = spawn(process.execPath, [TEST_WORKER, 'test', jobsFile], {
      detached: true,
      env: { ...process.env, [RUNNER_ENV]: '1' },
      stdio: 'ignore',
    });
    try {
      expect(await waitFor(() => child.exitCode !== null || child.signalCode !== null, 5000)).toBe(
        true
      );
    } finally {
      stopGroup(child.pid);
    }
  }, 30_000);

  test.each([
    ['in a folder other than the temporary folder', 'nested'],
    ['a link', 'link'],
    ['of a name that the runner does not make', 'name'],
  ] as const)(
    'a child whose runner is gone keeps a work folder %s',
    async (_, kind) => {
      const { jobs } = corpus(['hang']);
      const jobsFile = join(scratch(), 'jobs.json');
      writeFileSync(jobsFile, JSON.stringify(jobs));
      const target = join(scratch(), 'xlsx-corpus-nested');
      mkdirSync(target);
      // Folders directly in the temporary folder: this test removes them itself.
      const own =
        kind === 'link'
          ? join(tmpdir(), `xlsx-corpus-link-${process.pid}`)
          : kind === 'name'
            ? mkdtempSync(join(tmpdir(), 'xlsx-guard-test-'))
            : undefined;
      if (kind === 'link') symlinkSync(target, own!);
      const dir = own ?? target;
      const child = spawn(process.execPath, [TEST_WORKER, 'test', jobsFile], {
        detached: true,
        env: { ...process.env, [RUNNER_ENV]: '1', [CORPUS_TMP_ENV]: dir },
        stdio: 'ignore',
      });
      try {
        expect(
          await waitFor(() => child.exitCode !== null || child.signalCode !== null, 5000)
        ).toBe(true);
        expect(existsSync(dir)).toBe(true);
        expect(existsSync(target)).toBe(true);
      } finally {
        stopGroup(child.pid);
        if (own) rmSync(own, { recursive: true, force: true });
      }
    },
    30_000
  );

  test.each([
    ['the temporary folder', undefined],
    // The child starts in the repository folder: a relative path would name another folder.
    ['a relative temporary folder', 'tmp'],
  ] as const)(
    'a child removes its work folder and stops when its runner is killed, with %s',
    async (_, relative) => {
      const { jobs } = corpus(['grandchild']);
      const work = scratch();
      const pids = join(work, 'pids');
      if (relative) mkdirSync(join(work, relative));
      const runner = spawn(process.execPath, [runnerScript(work, jobs)], {
        cwd: relative ? work : undefined,
        env: {
          ...process.env,
          CORPUS_TEST_PIDS: pids,
          ...(relative ? { TMPDIR: relative } : {}),
        },
        stdio: 'ignore',
      });
      let child: number | undefined;
      let folder: string | undefined;
      try {
        expect(await waitFor(() => existsSync(pids), 10_000)).toBe(true);
        const record = grandchildRecord(pids);
        child = record.child;
        folder = dirname(record.work) === resolve(tmpdir()) ? record.work : undefined;
        if (relative) expect(dirname(record.work)).toBe(realpathSync(join(work, relative)));
        runner.kill('SIGKILL');
        // The guard thread of the child reads its parent every 250 ms.
        expect(await waitFor(() => !alive(record.child) && !alive(record.grandchild), 3000)).toBe(
          true
        );
        expect(existsSync(record.work)).toBe(false);
      } finally {
        runner.kill('SIGKILL');
        // When the check fails, the group of the child is still running: stop it here.
        stopGroup(child);
        // When the check fails, the folder of the child may stay.
        if (folder) rmSync(folder, { recursive: true, force: true });
      }
    },
    30_000
  );

  test('refuses to replace the results of an earlier run unless asked', async () => {
    const { jobs } = corpus(['ok']);
    const out = scratch();
    const options = { task: 'test', jobs, out, worker: TEST_WORKER, workbookMs: 10_000 };
    await run(options);
    expect((await runCorpus(options)) as CorpusRefusal).toEqual({
      refused: `The output folder ${out} holds the results of an earlier run. Add --fresh to replace them.`,
    });
    // A report of the other command in the same folder does not stay next to new results.
    writeFileSync(join(out, 'census.md'), '# Fixture set census\n');
    expect((await run({ ...options, fresh: true })).results).toBe(1);
    expect(existsSync(join(out, 'census.md'))).toBe(false);
  }, 30_000);

  test('refuses a batch size that is not a positive integer', async () => {
    const { jobs } = corpus(['ok']);
    const options = { task: 'test', jobs, out: scratch(), worker: TEST_WORKER, workbookMs: 1000 };
    await expect(runCorpus({ ...options, batchSize: Number.NaN })).rejects.toThrow('batch size');
    await expect(runCorpus({ ...options, batchSize: 0 })).rejects.toThrow('batch size');
  });

  test('runs the census and the schema task in the real child over committed fixtures', async () => {
    const jobs = listWorkbooks(join(ROOT, 'fixtures')).slice(0, 3);
    const census = await run({
      task: 'census',
      jobs,
      out: scratch(),
      batchSize: 2,
      workbookMs: 60_000,
    });
    expect(census).toMatchObject({ results: 3, stopped: [], children: 2 });
    expect(lines(census.resultsFile).map((l) => [l.id, l.open, l.save])).toEqual([
      ['workbook1', 'ok', 'same'],
      ['workbook2', 'ok', 'same'],
      ['workbook3', 'ok', 'same'],
    ]);
    const xsd = await run({
      task: 'xsd',
      jobs: jobs.slice(0, 1),
      out: scratch(),
      workbookMs: 60_000,
    });
    expect(lines(xsd.resultsFile)).toMatchObject([
      { id: 'workbook1', open: 'ok', saved: 0, edited: 0 },
    ]);
  }, 60_000);

  test('refuses an output folder that git can commit, before it writes anything', async () => {
    const { jobs } = corpus(['ok']);
    const repo = repository();
    await expect(
      runCorpus({
        task: 'test',
        jobs,
        out: join(repo, 'out'),
        worker: TEST_WORKER,
        workbookMs: 1000,
      })
    ).rejects.toThrow('git ignores');
    expect(existsSync(join(repo, 'out'))).toBe(false);
  });
});
