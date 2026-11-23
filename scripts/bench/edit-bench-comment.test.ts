// The sticky comment of the speed bench job, after DOCX Editor's
// `scripts/bench/edit-bench-comment.test.ts` at 66819cf1a (EigenPal).

import { describe, expect, test } from 'bun:test';
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

import { change, COMMENT_MARKER, foldRuns, renderComment } from './edit-bench-comment.ts';
import type { Report } from './edit-bench.ts';

function report(
  o: { medianMs?: number; rowsCached?: number; sha?: string; rows?: number } = {}
): Report {
  const medianMs = o.medianMs ?? 10;
  return {
    schema: 1,
    runtime: 'bun 1.4.4',
    args: {
      workbooks: ['grid'],
      rows: o.rows ?? 20_000,
      sheets: 200,
      strings: 200_000,
      runs: 20,
      warmup: 5,
      json: true,
    },
    results: [
      {
        name: 'commit',
        workbook: 'grid',
        cells: 400_000,
        workbookSha256: o.sha ?? 'abc',
        total: { medianMs, p95Ms: medianMs * 1.2 },
        perOperation: {
          rowsMeasured: 1,
          rowsCached: o.rowsCached ?? 19_999,
          sheetIndexes: 1,
          cellsIndexed: 400_000,
        },
      },
    ],
  };
}

const side = (...medians: number[]) => foldRuns(medians.map((medianMs) => report({ medianMs })));

describe('the speed bench comment', () => {
  test('a comparable base gives a table of changes, with the marker first', () => {
    const body = renderComment(side(8), side(10));
    expect(body.startsWith(COMMENT_MARKER)).toBe(true);
    expect(body).toContain('## Speed bench');
    expect(body).toContain('| commit | grid | 10.00 ms | 8.00 ms | -20.0%, faster |');
    expect(body).toContain('Work counters: no change.');
  });

  test('the comparison table gives the head p95, and the open has its note', () => {
    const head = report({ medianMs: 8 });
    head.results.push({ ...head.results[0]!, name: 'open' });
    const body = renderComment(foldRuns([head]), side(10));
    expect(body).toContain('| Step | Workbook | Base median | Head median | Change | Head p95 |');
    expect(body).toContain('| commit | grid | 10.00 ms | 8.00 ms | -20.0%, faster | 9.60 ms |');
    expect(body).toContain('The `open` step runs once in each round');
    expect(body).toContain(
      'The `edit-bench-report` artifact of the workflow run holds each report'
    );
    expect(renderComment(side(8), side(10))).not.toContain('runs once in each round');
  });

  test('the percent is rounded before the limits, so the words agree with the value', () => {
    expect(change(100, 105.04, 0)).toBe('+5.0%, within noise');
    expect(change(100, 94.96, 0)).toBe('-5.0%, within noise');
    expect(change(100, 105.06, 0)).toBe('+5.1%, slower');
    expect(change(100, 120.04, 0)).toBe('+20.0%, slower');
    expect(change(100, 120.06, 0)).toBe('+20.1%, slower (over 20%)');
    expect(change(0, 3, 0)).toBe('n/a');
  });

  test('a slower change over 20% gets a warning, and a small one is within noise', () => {
    expect(renderComment(side(15), side(10))).toContain('+50.0%, slower (over 20%)');
    expect(renderComment(side(11), side(10))).toContain('+10.0%, slower |');
    expect(renderComment(side(10.3), side(10))).toContain('+3.0%, within noise');
  });

  test('a changed work counter is listed with its values on both commits', () => {
    const body = renderComment(
      foldRuns([report({ rowsCached: 20_000 })]),
      foldRuns([report({ rowsCached: 19_999 })])
    );
    expect(body).toContain('Work counters that changed');
    expect(body).toContain('- `commit grid`: Rows cached 19999 to 20000');
  });

  test('a counter key that an object inherits shows by its key', () => {
    const head = report();
    Object.assign(head.results[0]!.perOperation, { toString: 4 });
    expect(renderComment(foldRuns([head]), side(10))).toContain(
      '- `commit grid`: `toString` none to 4'
    );
  });

  test('a counter on one commit only is listed, and the base still compares', () => {
    // Another commit can add or drop a counter, so its report has other keys than this Report.
    const head = report({ medianMs: 8 });
    (head.results[0]!.perOperation as unknown as Record<string, number>).rowsCopied = 5;
    const base = report({ medianMs: 10 });
    delete (base.results[0]!.perOperation as unknown as Partial<Record<string, number>>)
      .cellsIndexed;
    const body = renderComment(foldRuns([head]), foldRuns([base]));
    expect(body).toContain('| commit | grid | 10.00 ms | 8.00 ms |');
    expect(body).toContain('- `commit grid`: Cells indexed none to 400000');
    expect(body).toContain('- `commit grid`: `rowsCopied` none to 5');
  });

  test('no base, another setup, or other workbook bytes give the head only', () => {
    expect(renderComment(side(10))).toContain('The base commit has no edit bench');
    const sizes = renderComment(side(10), foldRuns([report({ rows: 2000 })]));
    expect(sizes).toContain('The base rounds used another setup');
    const bytes = renderComment(side(10), foldRuns([report({ sha: 'other' })]));
    expect(bytes).toContain('The grid workbook of the base commit has other bytes');
    for (const body of [sizes, bytes]) {
      expect(body).toContain('| Step | Workbook | Median | p95 |');
      expect(body).not.toContain('Base median');
    }
  });

  test('a step on one commit only still gets its row', () => {
    const head = report();
    head.results[0]!.name = 'commit-text';
    const body = renderComment(foldRuns([head]), side(10));
    expect(body).toContain('| commit-text | grid | — | 10.00 ms | n/a |');
    expect(body).toContain('| commit | grid | 10.00 ms | — | n/a |');
  });
});

describe('the interleaved rounds', () => {
  test('two rounds give the mean of their medians', () => {
    const body = renderComment(side(10, 12), side(10, 10));
    expect(body).toContain('| commit | grid | 10.00 ms | 11.00 ms |');
    expect(body).toContain('Rounds: head 2, base 2, interleaved.');
  });

  test('a change inside the spread of the rounds is within noise, and outside it is not', () => {
    // Head rounds 10 and 14: mean 12, spread 4 / 12 = 33%. The +20% change stays within noise.
    const noisy = renderComment(side(10, 14), side(10));
    expect(noisy).toContain('| +20.0%, within noise |');
    expect(noisy).toContain('(at most 33.3%)');
    expect(renderComment(side(12, 12), side(10))).toContain('| +20.0%, slower |');
  });

  test('the spread of the base rounds also widens the noise of a step', () => {
    // Base rounds 8 and 12: mean 10, spread 40%. A head of 13 is +30%, within that spread.
    expect(renderComment(side(13), side(8, 12))).toContain('| +30.0%, within noise |');
  });

  test('the floor is rounded as the change is, so the row and the footer agree', () => {
    // Head 10.7 and 10.0, base 9.3 and 10.0: the change and the base spread both round to 7.3%.
    const body = renderComment(side(10.7, 10), side(9.3, 10));
    expect(body).toContain('| +7.3%, within noise |');
    expect(body).toContain('(at most 7.3%)');
  });

  test('the footer says when a side has fewer rounds', () => {
    expect(renderComment(side(10, 10), side(10))).toContain(
      'interleaved (a round is missing, so the drift cancels less).'
    );
    expect(renderComment(side(10, 10), side(10, 10))).toContain('interleaved. Each round:');
  });

  test('a round with another setup is left out of its side', () => {
    const base = foldRuns([report({ medianMs: 10 }), report({ medianMs: 30, rows: 2000 })]);
    expect(base.rounds).toBe(1);
    expect(base.steps.get('commit grid')!.medianMs).toBe(10);
  });

  test('rounds with medians of 0 have no spread, not a spread of NaN', () => {
    const body = renderComment(side(0, 0), side(10));
    expect(body).toContain('-100.0%, faster');
    expect(body).not.toContain('NaN');
    expect(change(0, 0, 0)).toBe('0.0%, within noise');
  });

  test('a round with other workbook bytes is left out of its side', () => {
    const head = foldRuns([report({ medianMs: 10 }), report({ medianMs: 20, sha: 'other' })]);
    expect(head.rounds).toBe(1);
    expect(renderComment(head, side(10))).toContain('| commit | grid | 10.00 ms | 10.00 ms |');
  });
});

test('the workflow finds the comment by the marker that the script writes', () => {
  const workflow = readFileSync(join(import.meta.dir, '../../.github/workflows/bench.yml'), 'utf8');
  expect(workflow).toContain(`const marker = '${COMMENT_MARKER}';`);
});

test('the command prints its usage for --help', () => {
  const run = Bun.spawnSync(
    [process.execPath, join(import.meta.dir, 'edit-bench-comment.ts'), '--help'],
    { stdout: 'pipe', stderr: 'pipe' }
  );
  expect(run.exitCode).toBe(0);
  expect(run.stdout.toString()).toContain('Usage: bun scripts/bench/edit-bench-comment.ts');
});

test('the command refuses an unknown flag or a flag with no value', () => {
  for (const flags of [['--hed', 'x', '--out', 'y'], ['--head']]) {
    const run = Bun.spawnSync(
      [process.execPath, join(import.meta.dir, 'edit-bench-comment.ts'), ...flags],
      { stdout: 'pipe', stderr: 'pipe' }
    );
    expect(run.exitCode).toBe(2);
    expect(run.stderr.toString()).toContain('Usage: bun scripts/bench/edit-bench-comment.ts');
    expect(run.stderr.toString()).not.toContain(' at ');
  }
});

test('the command reads the reports and writes the comment; a bad report is left out', () => {
  const dir = mkdtempSync(join(tmpdir(), 'xlsx-bench-comment-'));
  try {
    const file = (name: string, body: unknown) => {
      const path = join(dir, name);
      writeFileSync(path, typeof body === 'string' ? body : JSON.stringify(body));
      return path;
    };
    const out = join(dir, 'comment.md');
    const run = Bun.spawnSync(
      [
        process.execPath,
        join(import.meta.dir, 'edit-bench-comment.ts'),
        '--head',
        file('head.json', report({ medianMs: 8 })),
        '--base',
        file('base.json', report({ medianMs: 10 })),
        '--base',
        file('cut.json', '{"schema":1,"res'),
        '--base',
        file('shape.json', { schema: 1, results: [{ name: 'commit' }] }),
        '--out',
        out,
      ],
      { stdout: 'pipe', stderr: 'pipe' }
    );
    expect(run.exitCode, run.stderr.toString()).toBe(0);
    expect(run.stderr.toString()).toContain('The base report is left out: ');
    expect(run.stderr.toString()).toContain('cut.json: ');
    expect(run.stderr.toString()).toContain('shape.json: not a report of bench:edit -- --json');
    expect(readFileSync(out, 'utf8')).toContain('| commit | grid | 10.00 ms | 8.00 ms |');
    const none = Bun.spawnSync(
      [process.execPath, join(import.meta.dir, 'edit-bench-comment.ts'), '--out', out],
      { stdout: 'pipe', stderr: 'pipe' }
    );
    expect(none.exitCode).toBe(2);
    expect(none.stderr.toString()).toContain('Usage: bun scripts/bench/edit-bench-comment.ts');
    const bad = Bun.spawnSync(
      [
        process.execPath,
        join(import.meta.dir, 'edit-bench-comment.ts'),
        '--head',
        join(dir, 'cut.json'),
        '--out',
        out,
      ],
      { stdout: 'pipe', stderr: 'pipe' }
    );
    expect(bad.exitCode).toBe(2);
    expect(bad.stderr.toString()).toContain('The head report is left out: ');
    expect(bad.stderr.toString()).toContain('No head report can be read.');
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
});
