// The work counters of `bun run bench:edit`, pinned on small generated workbooks. A count is exact
// on every machine, a wall-clock limit is not, so these gates run inside `bun run test`.
//
// Same role as DOCX Editor's `scripts/bench/edit-bench-gates.test.ts` at 1c9b6a29b (EigenPal),
// the commit that added it: the expected work of each step, as exact numbers. The scaling test
// follows PPTX Editor's `packages/core/src/editor/__tests__/edit-scaling.test.ts` at d4502552:
// four times the content, and the work of a step that must not grow stays the same.
// As DOCX Editor's gate runs its bench command on its fixture and reads the JSON, the first test
// runs `bun scripts/bench/edit-bench.ts` on small workbooks (the default sizes take 3 GiB), reads
// its `--json` report, and compares a second run with it through `--compare`. The other tests call
// `benchWorkbook` in the test process. The scaling test shows that a keystroke and a scroll step
// do no counted work (row heights, sheet indexes) at four times the rows, and that a commit
// measures as many rows; the rows that a commit takes from the cache grow with the sheet.
// DOCX Editor's gate pins the SHA-256 of its committed fixture, so a changed fixture fails it.
// Here each run builds its workbooks again from `workbooks.ts`, so the gate pins no SHA-256:
// `--compare` refuses a report of other workbook bytes, and `workbooks.test.ts` checks the cell
// count and the part limits of each workbook at its default sizes.

import { expect, test } from 'bun:test';
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import {
  benchWorkbook,
  checkRuns,
  compareReports,
  regressions,
  type Args,
  type Report,
  type ScenarioResult,
} from './edit-bench.ts';

const SMALL: Args = {
  workbooks: ['grid', 'sheets', 'strings'],
  rows: 2000,
  sheets: 20,
  strings: 20_000,
  runs: 3,
  warmup: 1,
  json: true,
};

type Work = ScenarioResult['perOperation'];

const work = (
  rowsMeasured: number,
  rowsCached: number,
  sheetIndexes: number,
  cellsIndexed: number
): Work => ({ rowsMeasured, rowsCached, sheetIndexes, cellsIndexed });

const NONE = work(0, 0, 0, 0);

/**
 * The work of each step. A keystroke, a scroll step, and a switch to a sheet seen before do no
 * work for each row or cell. A commit, of a number or of a new text, measures its row, walks the
 * other rows from the cache, and builds the sheet index again: those pins go down
 * when the issue is fixed. The open of the sheets workbook builds the index of each worksheet,
 * not only of the first one. The counters show only the row heights and the sheet
 * indexes, so these pins do not cover the other work of a commit in proportion to its sheet: the
 * edit state of the sheet (`blockedCells`, `masterCells`, `arrayRangesOf`), the formula scan,
 * `sheetFacts`, the check of the touched parts (`touchedPartViolation`), the sort of the cell
 * keys (`sortedKeys`) and the hidden lines (`hiddenLines`) of each new sheet index, and the copy
 * of the shared string table that a new text makes, and the structure check over
 * the other worksheets. A commit of a new text pins the same counters as a number.
 */
const EXPECTED: Readonly<Record<string, Work>> = {
  'open grid': work(2000, 0, 1, 40_000),
  'keystroke grid': NONE,
  'commit grid': work(1, 1999, 1, 40_000),
  'commit-text grid': work(1, 1999, 1, 40_000),
  'scroll grid': NONE,
  'open sheets': work(100, 0, 20, 20_000),
  'keystroke sheets': NONE,
  'commit sheets': work(1, 99, 1, 1000),
  'commit-text sheets': work(1, 99, 1, 1000),
  'scroll sheets': NONE,
  'sheet-switch sheets': NONE,
  'open strings': work(1000, 0, 1, 20_000),
  'keystroke strings': NONE,
  'commit strings': work(1, 999, 1, 20_000),
  'commit-text strings': work(1, 999, 1, 20_000),
  'scroll strings': NONE,
};

async function run(args: Args): Promise<ScenarioResult[]> {
  const out: ScenarioResult[] = [];
  for (const name of args.workbooks) out.push(...(await benchWorkbook(name, args)));
  return out;
}

/** The flags of `SMALL` on the command line. */
const SMALL_FLAGS = ['--rows', '2000', '--sheets', '20', '--strings', '20000', '--runs', '3'];

/** Runs the bench command; stdout is the `--json` report. */
function bench(...flags: string[]): { exitCode: number; stdout: string; stderr: string } {
  const out = Bun.spawnSync([process.execPath, join(import.meta.dir, 'edit-bench.ts'), ...flags], {
    stdout: 'pipe',
    stderr: 'pipe',
  });
  return { exitCode: out.exitCode, stdout: out.stdout.toString(), stderr: out.stderr.toString() };
}

test('each step of the edit bench does the pinned work, and --compare reads its report', () => {
  const first = bench(...SMALL_FLAGS, '--warmup', '1', '--json');
  // The bench writes each step that did work in proportion to the sheet to stderr.
  expect(first.exitCode, first.stderr).toBe(0);
  const report = JSON.parse(first.stdout) as Report;
  const got = Object.fromEntries(
    report.results.map((r) => [`${r.name} ${r.workbook}`, r.perOperation])
  );
  expect(Object.keys(got).sort()).toEqual(Object.keys(EXPECTED).sort());
  // Each pin by itself, so a failure names its step and workbook, for example `commit grid`.
  for (const [step, pinned] of Object.entries(EXPECTED))
    expect(got[step], `${step}; if the change is on purpose, update EXPECTED`).toEqual(pinned);
  expect(regressions(report.results)).toEqual([]);
  const dir = mkdtempSync(join(tmpdir(), 'xlsx-edit-bench-'));
  try {
    const file = join(dir, 'before.json');
    writeFileSync(file, first.stdout);
    const second = bench(...SMALL_FLAGS, '--warmup', '1', '--json', '--compare', file);
    expect(second.exitCode, second.stderr).toBe(0);
    const compared = (JSON.parse(second.stdout) as Report).comparison ?? [];
    expect(compared.length).toBe(report.results.length);
    for (const c of compared)
      expect([
        c.rowsMeasuredChange,
        c.rowsCachedChange,
        c.sheetIndexesChange,
        c.cellsIndexedChange,
      ]).toEqual([0, 0, 0, 0]);
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
  // A wrong flag stops the command before any work.
  expect(bench('--rows', '0').exitCode).toBe(2);
  const oneSheet = bench('--sheets', '1');
  expect(oneSheet.exitCode).toBe(2);
  expect(oneSheet.stderr).toContain('--sheets takes 2 or more');
}, 120_000);

test('four times the rows, and a keystroke, a commit, or a scroll measures as many', async () => {
  const grid = async (rows: number) =>
    run({ ...SMALL, workbooks: ['grid'], rows }).then((r) =>
      Object.fromEntries(r.map((s) => [s.name, s.perOperation]))
    );
  const small = await grid(1000);
  const large = await grid(4000);
  expect(large.open!.rowsMeasured).toBe(4 * small.open!.rowsMeasured);
  for (const step of ['keystroke', 'commit', 'commit-text', 'scroll'])
    expect(large[step]!.rowsMeasured, step).toBe(small[step]!.rowsMeasured);
  // A keystroke and a scroll step read no row at any size.
  for (const step of ['keystroke', 'scroll']) expect(large[step], step).toEqual(NONE);
}, 120_000);

test('the gate names a step that measures rows or builds an index', () => {
  const step = (name: string, w: Work): ScenarioResult => ({
    name,
    workbook: 'grid',
    cells: 1,
    workbookSha256: 'a',
    total: { medianMs: 0, p95Ms: 0 },
    perOperation: w,
  });
  expect(
    regressions([
      step('open', work(9, 9, 9, 9)),
      step('keystroke', work(2, 0, 0, 0)),
      step('scroll', work(0, 0, 1, 10)),
      step('sheet-switch', work(0, 5, 0, 0)),
      step('commit', work(2, 9, 1, 10)),
      step('commit-text', work(1, 9, 1, 10)),
      // A commit that measures no row kept a stale height for the row that it changed.
      step('commit', work(0, 10, 1, 10)),
    ])
  ).toEqual([
    'keystroke on grid measured 2 rows',
    'scroll on grid built 1 sheet index',
    'sheet-switch on grid read 5 cached rows',
    'commit on grid measured 2 rows, not 1',
    'commit on grid measured 0 rows, not 1',
  ]);
});

test('--compare gives the change of each step, and refuses a report of other sizes or bytes', () => {
  const step = (medianMs: number, w: Work, workbookSha256 = 'a'): ScenarioResult => ({
    name: 'commit',
    workbook: 'grid',
    cells: 1,
    workbookSha256,
    total: { medianMs, p95Ms: medianMs },
    perOperation: w,
  });
  const before = { schema: 1 as const, args: SMALL, results: [step(10, work(2, 9, 1, 10))] };
  expect(compareReports(before, SMALL, [step(5, work(1, 9, 1, 10))])).toEqual([
    {
      name: 'commit',
      workbook: 'grid',
      medianChangePct: -50,
      rowsMeasuredChange: -1,
      rowsCachedChange: 0,
      sheetIndexesChange: 0,
      cellsIndexedChange: 0,
    },
  ]);
  expect(() => compareReports(before, { ...SMALL, rows: SMALL.rows + 1 }, [])).toThrow(
    `--compare: the report has ${SMALL.rows} rows`
  );
  // A size of a workbook that the run does not use does not matter.
  expect(() =>
    compareReports(before, { ...SMALL, workbooks: ['grid'], sheets: 3 }, [])
  ).not.toThrow();
  // A change to the workbook builder gives other bytes at the same sizes.
  expect(() => compareReports(before, SMALL, [step(5, work(1, 9, 1, 10), 'b')])).toThrow(
    '--compare: the grid workbook has other bytes than in the report'
  );
  const noArgs = { schema: 1, results: [] } as unknown as typeof before;
  expect(() => compareReports(noArgs, SMALL, [])).toThrow('--compare takes a report of --json');
  // A report with no sizes, for example changed by hand.
  const noSizes = { ...before, args: {} } as unknown as typeof before;
  expect(() => compareReports(noSizes, SMALL, [])).toThrow('--compare takes a report of --json');
  // A result with no time, no counter, or no SHA-256, for example in a report changed by hand.
  const [result] = before.results;
  for (const changed of [
    { ...result, total: undefined },
    { ...result, perOperation: null },
    { ...result, perOperation: { rowsMeasured: 1 } },
    { ...result, workbookSha256: undefined },
  ]) {
    const report = { ...before, results: [changed] } as unknown as typeof before;
    expect(() => compareReports(report, SMALL, [])).toThrow('--compare takes a report of --json');
  }
});

test('the commits stay inside the rows of the data, and runs that need more rows stop', async () => {
  // 22 commits in 40 rows start above the middle, and each one does the same work.
  const results = await run({ ...SMALL, workbooks: ['grid'], rows: 40, runs: 10 });
  expect(regressions(results)).toEqual([]);
  expect(results.find((r) => r.name === 'commit')!.perOperation).toEqual(work(1, 39, 1, 800));
  const message =
    'the sheets workbook has 100 data rows, and 1 warmup and 50 timed runs need 102 rows for ' +
    'their commits: use fewer runs or warmup runs';
  await expect(benchWorkbook('sheets', { ...SMALL, sheets: 2, runs: 50 })).rejects.toThrow(message);
  // The run checks each workbook before the first one runs.
  expect(() => checkRuns({ ...SMALL, runs: 50 })).toThrow(message);
}, 120_000);

test('the commits of the strings workbook skip a last row that is part of a row', async () => {
  // 1001 strings give 50 full rows and a last row with one cell. 46 commits fit in the full rows,
  // so each commit replaces a cell, and each one does the same work.
  const strings = { ...SMALL, workbooks: ['strings'] as Args['workbooks'], strings: 1001 };
  const results = await run({ ...strings, runs: 20, warmup: 3 });
  expect(regressions(results)).toEqual([]);
  expect(results.find((r) => r.name === 'commit-text')!.perOperation).toEqual(work(1, 50, 1, 1001));
  expect(() => checkRuns({ ...strings, strings: 999, runs: 22, warmup: 3 })).toThrow(
    'the strings workbook has 49 data rows'
  );
}, 120_000);
