// The sticky pull request comment of the speed bench job (`.github/workflows/bench.yml`). It reads
// the reports of `bun run bench:edit -- --json` on the head and on the base commit, and writes
// one table: the median time of each step on both commits, the change, and the work counters
// that changed.
//
// Each side gives several rounds, because the job interleaves them (head, base, base, head, two
// times). A whole round takes the momentary state of the machine, and that order cancels the
// drift to first order. The median of a step on one side is the mean of the medians of its
// rounds, and its p95 is the mean of the p95 values of its rounds. A change inside the spread of
// the round medians on either side, or inside 5%, shows as within noise.
//
// The times come from a shared runner, so the comment is information only: nothing here fails the
// job. The gate is `edit-bench-gates.test.ts`, which pins the work counters in `bun run test`.
//
// Usage: bun scripts/bench/edit-bench-comment.ts --head <report> [--head <report>]
//          [--base <report> ...] --out <comment.md>
//
// Model: DOCX Editor's `scripts/bench/edit-bench-comment.mjs` at 0eb629bcd (EigenPal), with the
// interleaved rounds and the spread of 66819cf1a. Differences, each for a reason:
// - TypeScript, run by Bun as the other bench scripts are, so it reads the report types of
//   `edit-bench.ts`.
// - Words, not colored marks: "slower", "faster", and "within noise", so no meaning is carried by
//   color alone (the accessibility rule of the documentation style in AGENTS.md).
// - The two sides are comparable only with the same sizes, runs, and workbooks, and the same
//   SHA-256 of each workbook, as `--compare` checks. DOCX Editor checks the SHA-256 of its one
//   fixture.
// - The work counters compare by key, as DOCX Editor's `numericLeaves` does; the report of this
//   bench holds one level of counters, so no deeper walk is needed. DOCX Editor lists only a
//   counter on both sides. This script also lists a counter on one side only, as "none" on the
//   other, because a counter that a change adds or drops is a change of the work that it counts.
// - The noise floor of a step is rounded to one decimal, as the change is, so the floor in the
//   footer and the words of each row agree.
// - The head-only table has no min and max columns: the report of this bench gives the median
//   and the p95 of each step only, as PPTX Editor's report does.
// - The `open` step runs once in each round, so the comment says that its value comes from one
//   sample in each round.
// - The footer says when the two sides have a different number of rounds: then the drift does
//   not cancel fully.
// - The comment holds no JSON of the reports. The workflow artifact holds each report, so the
//   comment stays far under the size limit of a comment, and the comment names the artifact.
// - `--help` prints the usage and exits with code 0, as the other bench commands do.
// - No browser section: the job runs only the edit bench. DOCX Editor added its browser times in
//   a later step.

import { readFileSync, writeFileSync } from 'node:fs';

import { stopOnUsageError, UsageError } from './args.ts';
import type { Report } from './edit-bench.ts';

/** The first line of each comment, which the workflow finds to update the comment. */
export const COMMENT_MARKER = '<!-- edit-bench-report -->';

/** A change within this many percent is noise. The spread of the rounds widens it for a step. */
const NOISE_PCT = 5;

/** A slower change above this many percent gets a warning. */
const WARN_PCT = 20;

/** The names of the counters that the bench writes. Another counter shows by its key. */
const COUNTER_NAMES: Readonly<Record<string, string>> = {
  rowsMeasured: 'Rows measured',
  rowsCached: 'Rows cached',
  sheetIndexes: 'Sheet indexes',
  cellsIndexed: 'Cells indexed',
};

/** Where the reader finds the reports and the cause of a failed round. */
const ARTIFACT_NOTE =
  'The `edit-bench-report` artifact of the workflow run holds each report, and the job log ' +
  'gives the cause of a failed round.';

const USAGE =
  'Usage: bun scripts/bench/edit-bench-comment.ts --head <report> [--head <report>] ' +
  '[--base <report> ...] --out <comment.md>';

/** One step of one side: the mean of the medians and p95 values of its rounds. */
interface Step {
  name: string;
  workbook: string;
  workbookSha256: string;
  medianMs: number;
  p95Ms: number;
  /** The spread of the round medians, in percent of their mean; 0 for one round. */
  spreadPct: number;
  /** The work counters of the step: each numeric value of `perOperation`. */
  counters: Map<string, number>;
}

/** The rounds of one side, folded into one value for each step. */
export interface Side {
  /** The rounds that the side holds. */
  rounds: number;
  /** The sizes, runs, and workbooks of each round: two sides compare only when they are equal. */
  setup: string;
  /** The steps, by `step workbook`, in the order of the first round. */
  steps: Map<string, Step>;
}

type Result = Report['results'][number];

const keyOf = (r: { name: string; workbook: string }) => `${r.name} ${r.workbook}`;
const mean = (xs: readonly number[]) => xs.reduce((a, b) => a + b, 0) / xs.length;

/** Each numeric value of `perOperation`, so a counter that a later bench adds still compares. */
function countersOf(r: Result): Map<string, number> {
  const out = new Map<string, number>();
  for (const [k, v] of Object.entries(r.perOperation ?? {})) if (Number.isFinite(v)) out.set(k, v);
  return out;
}

/** The report in `file`, or an error that names the file. */
export function readReport(file: string): Report {
  let report: Report;
  try {
    report = JSON.parse(readFileSync(file, 'utf8')) as Report;
  } catch (error) {
    throw new Error(`${file}: ${(error as Error).message}`);
  }
  const args = report?.args;
  const ok =
    report?.schema === 1 &&
    Array.isArray(report.results) &&
    report.results.length > 0 &&
    report.results.every(
      (r) =>
        typeof r?.name === 'string' &&
        typeof r.workbook === 'string' &&
        typeof r.workbookSha256 === 'string' &&
        Number.isFinite(r.total?.medianMs) &&
        Number.isFinite(r.total?.p95Ms) &&
        typeof r.perOperation === 'object' &&
        r.perOperation !== null
    ) &&
    typeof args === 'object' &&
    args !== null &&
    Array.isArray(args.workbooks) &&
    (['rows', 'sheets', 'strings', 'runs', 'warmup'] as const).every((k) =>
      Number.isInteger(args[k])
    );
  if (!ok) throw new Error(`${file}: not a report of bench:edit -- --json`);
  return report;
}

const setupOf = (r: Report): string => {
  const a = r.args;
  return (
    `${a.workbooks.join(', ')}; ${a.rows} rows, ${a.sheets} sheets, ${a.strings} strings; ` +
    `${a.runs} timed runs after ${a.warmup} untimed`
  );
};

const shasOf = (r: Report): string =>
  [...new Set(r.results.map((x) => `${x.workbook}=${x.workbookSha256}`))].sort().join(' ');

/**
 * The rounds of one side as one value for each step. A round with another setup or other workbook
 * bytes than the first round is left out, with a line on stderr.
 */
export function foldRuns(reports: readonly Report[]): Side {
  const [first] = reports;
  if (!first) throw new Error('foldRuns takes one or more reports');
  const rounds = reports.filter((r) => {
    if (setupOf(r) === setupOf(first) && shasOf(r) === shasOf(first)) return true;
    console.error('A round with another setup or other workbook bytes is left out.');
    return false;
  });
  const steps = new Map<string, Step>();
  for (const result of first.results) {
    const same = rounds
      .map((r) => r.results.find((x) => keyOf(x) === keyOf(result)))
      .filter((x): x is Result => x !== undefined);
    const medians = same.map((x) => x.total.medianMs);
    const medianMs = mean(medians);
    steps.set(keyOf(result), {
      name: result.name,
      workbook: result.workbook,
      workbookSha256: result.workbookSha256,
      medianMs,
      p95Ms: mean(same.map((x) => x.total.p95Ms)),
      // All medians 0 is no spread, not 0 / 0.
      spreadPct:
        medians.length > 1 && medianMs > 0
          ? ((Math.max(...medians) - Math.min(...medians)) / medianMs) * 100
          : 0,
      counters: countersOf(result),
    });
  }
  return { rounds: rounds.length, setup: setupOf(first), steps };
}

const ms = (x: number) => `${x.toFixed(2)} ms`;

/**
 * The change from `base` to `head`, in words. The percent is rounded to one decimal before the
 * limits apply, as in DOCX Editor, so the shown value and its words agree.
 */
export function change(baseMs: number, headMs: number, floorPct: number): string {
  if (baseMs === 0) return headMs === 0 ? '0.0%, within noise' : 'n/a';
  const pct = Number((((headMs - baseMs) / baseMs) * 100).toFixed(1));
  const text = `${pct >= 0 ? '+' : ''}${pct.toFixed(1)}%`;
  const floor = Math.max(NOISE_PCT, Number(floorPct.toFixed(1)));
  if (pct > floor)
    return pct > WARN_PCT ? `${text}, slower (over ${WARN_PCT}%)` : `${text}, slower`;
  if (pct < -floor) return `${text}, faster`;
  return `${text}, within noise`;
}

/** Why the base side does not compare with the head side, or undefined when it does. */
function notComparable(head: Side, base: Side | undefined): string | undefined {
  if (!base)
    return (
      'The base commit has no edit bench, or no base round gave a report, ' +
      'so the table gives the head only.'
    );
  if (base.setup !== head.setup)
    return `The base rounds used another setup (${base.setup}), so the table gives the head only.`;
  for (const [key, step] of head.steps) {
    const other = base.steps.get(key);
    if (other && other.workbookSha256 !== step.workbookSha256)
      return (
        `The ${step.workbook} workbook of the base commit has other bytes, ` +
        'so the table gives the head only.'
      );
  }
  return undefined;
}

/** The lines of the counters of `key` that differ between `base` and `head`. */
function counterChanges(key: string, base: Step, head: Step): string[] {
  const out: string[] = [];
  for (const k of new Set([...base.counters.keys(), ...head.counters.keys()])) {
    const b = base.counters.get(k);
    const h = head.counters.get(k);
    const name = Object.hasOwn(COUNTER_NAMES, k) ? COUNTER_NAMES[k] : `\`${k}\``;
    if (b !== h) out.push(`- \`${key}\`: ${name} ${b ?? 'none'} to ${h ?? 'none'}`);
  }
  return out;
}

/** The note on the open step, which runs once in each round. */
function openNote(side: Side): string[] {
  if (![...side.steps.values()].some((s) => s.name === 'open')) return [];
  return [
    'The `open` step runs once in each round, so its value is the mean of one sample from each ' +
      'round, and its p95 equals its median.',
  ];
}

/** The markdown of the comment. */
export function renderComment(head: Side, base?: Side): string {
  const lines = [
    COMMENT_MARKER,
    '## Speed bench',
    '',
    'The median time of each step of `bun run bench:edit` on a shared CI runner. ' +
      'The times are information only, and no time fails this job. The work counters are ' +
      'exact, and `scripts/bench/edit-bench-gates.test.ts` pins them in `bun run test`.',
    '',
  ];
  const note = notComparable(head, base);
  if (note !== undefined || !base) {
    lines.push(`> ${note}`, '', '| Step | Workbook | Median | p95 |', '| --- | --- | --: | --: |');
    for (const s of head.steps.values())
      lines.push(`| ${s.name} | ${s.workbook} | ${ms(s.medianMs)} | ${ms(s.p95Ms)} |`);
    lines.push(
      '',
      `Rounds: ${head.rounds}. Each round: ${head.setup}.`,
      ...openNote(head),
      ARTIFACT_NOTE
    );
    return `${lines.join('\n')}\n`;
  }
  lines.push(
    '| Step | Workbook | Base median | Head median | Change | Head p95 |',
    '| --- | --- | --: | --: | --- | --: |'
  );
  const counterLines: string[] = [];
  for (const [key, s] of head.steps) {
    const b = base.steps.get(key);
    if (!b) {
      lines.push(`| ${s.name} | ${s.workbook} | — | ${ms(s.medianMs)} | n/a | ${ms(s.p95Ms)} |`);
      continue;
    }
    const floor = Math.max(s.spreadPct, b.spreadPct);
    lines.push(
      `| ${s.name} | ${s.workbook} | ${ms(b.medianMs)} | ${ms(s.medianMs)} | ` +
        `${change(b.medianMs, s.medianMs, floor)} | ${ms(s.p95Ms)} |`
    );
    counterLines.push(...counterChanges(key, b, s));
  }
  for (const [key, b] of base.steps)
    if (!head.steps.has(key))
      lines.push(`| ${b.name} | ${b.workbook} | ${ms(b.medianMs)} | — | n/a | — |`);
  const spreads = [...head.steps.values(), ...base.steps.values()].map((s) => s.spreadPct);
  lines.push(
    '',
    `Rounds: head ${head.rounds}, base ${base.rounds}, interleaved` +
      (head.rounds === base.rounds ? '' : ' (a round is missing, so the drift cancels less)') +
      `. Each round: ${head.setup}. ` +
      `A change inside ${NOISE_PCT}%, or inside the spread of the round medians of its step ` +
      `(at most ${Math.max(0, ...spreads).toFixed(1)}%), is within noise.`,
    ...openNote(head),
    ARTIFACT_NOTE,
    ''
  );
  if (counterLines.length === 0) lines.push('Work counters: no change.');
  else
    lines.push(
      'Work counters that changed. The counters are exact, so check that each change is ' +
        'intended. A lower count is less work. If `edit-bench-gates.test.ts` pins the step, ' +
        'update its pin in the same pull request.',
      '',
      ...counterLines
    );
  return `${lines.join('\n')}\n`;
}

/** The reports of `files`; a file that cannot be read is left out, with its reason on stderr. */
function readAll(files: readonly string[], side: string): Report[] {
  const out: Report[] = [];
  for (const file of files) {
    try {
      out.push(readReport(file));
    } catch (error) {
      console.error(`The ${side} report is left out: ${(error as Error).message}`);
    }
  }
  return out;
}

function main(argv: readonly string[]): void {
  const head: string[] = [];
  const base: string[] = [];
  let out: string | undefined;
  if (argv.includes('--help')) {
    console.log(USAGE);
    return;
  }
  for (let i = 0; i < argv.length; i += 2) {
    const flag = argv[i]!;
    if (!['--head', '--base', '--out'].includes(flag))
      throw new UsageError(`unknown argument: ${flag}\n${USAGE}`);
    const value = argv[i + 1];
    if (value === undefined || value.startsWith('--'))
      throw new UsageError(`${flag} takes a value\n${USAGE}`);
    if (flag === '--head') head.push(value);
    else if (flag === '--base') base.push(value);
    else out = value;
  }
  if (head.length === 0 || out === undefined) throw new UsageError(USAGE);
  const heads = readAll(head, 'head');
  if (heads.length === 0) throw new UsageError('No head report can be read.');
  const bases = readAll(base, 'base');
  writeFileSync(
    out,
    renderComment(foldRuns(heads), bases.length > 0 ? foldRuns(bases) : undefined)
  );
}

if (import.meta.main) {
  try {
    main(process.argv.slice(2));
  } catch (error) {
    stopOnUsageError(error);
  }
}
