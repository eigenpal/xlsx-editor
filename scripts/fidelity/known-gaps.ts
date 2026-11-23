// The fixture gate of `bun run fidelity`: a known-gap ratchet (design D8).
//
// Same role as pptx-editor/scripts/fidelity/known-gaps.ts at 1943d2c9 (EigenPal). Difference:
// each gap records the measured value of each failing check, as design D8 states, and the gate
// compares the run with it inside `GAP_TOLERANCE`. A fixture with a gap still reports
// "fail (known gap)", never a pass. A missing or stale reference fails the gate.
//
// Compared with 1943d2c9; later commits 67ae7309 and 06534363 checked: 67ae7309 is already
// equal, and 06534363 does not change this file. 67ae7309 changes the remedy for a reference PDF
// with foreign fonts to "export its reference PDF again". Here the remedy is
// `FOREIGN_REFERENCE_REMEDY` (`font-names.ts`), which says the same, and the run reports it.

import { spawnSync } from 'node:child_process';
import { readFileSync, writeFileSync } from 'node:fs';
import {
  FAILURE_KINDS,
  GAP_TOLERANCE,
  KNOWN_GAPS,
  pixelLimit,
  THRESHOLDS,
  type FailureKind,
  type FontEnvironment,
  type KnownGap,
  type KnownGapList,
} from './thresholds.ts';

/** The result of one fixture. */
export interface FixtureResult {
  readonly name: string;
  /** `compared`: both PDFs were measured. `error`: the fixture could not be compared. */
  readonly status: 'compared' | 'error';
  readonly error?: string;
  /** The checks that the page fails. */
  readonly failures?: readonly FailureKind[];
  /** The measured value of every check. */
  readonly values?: Readonly<Record<FailureKind, number>>;
  /** Reference lines that the comparison checks. */
  readonly linesTotal?: number;
  readonly linesMatched?: number;
  /** Of the matched lines, those that a cell clip shortens on both pages (`LineScore.clipped`). */
  readonly clippedLines?: readonly string[];
  /**
   * Reference lines that a clip cuts on the reference page and that the painted page does not hold
   * cut at the same edge (`LineScore.uncut`).
   */
  readonly uncutLines?: readonly string[];
  /**
   * Of `uncutLines`, the lines that no painted line matches (`LineScore.lostUncut`): the painted
   * page lost the line or cut it elsewhere.
   */
  readonly lostUncutLines?: readonly string[];
  /**
   * Reference lines that the painted clip cuts at an end that the reference clip holds whole
   * (`LineScore.extraCut`).
   */
  readonly extraCutLines?: readonly string[];
  /**
   * The ends of matched lines where the reference clip is unknown, so no clip rule checks them
   * (`LineScore.unknownReference`).
   */
  readonly unknownReferenceEnds?: readonly string[];
  /** Missing lines that the reference clip cuts on both sides (`LineScore.bothCut`). */
  readonly bothCutLines?: readonly string[];
  readonly missingLines?: readonly string[];
  readonly extraLines?: readonly string[];
  /**
   * Reference text items that the line scores leave out as rotated, mirrored, or vertical
   * (`PdfItems.rotated`): no clip rule checks their cells.
   */
  readonly rotatedItems?: number;
  /**
   * Reference lines of one character other than a digit (`LineScore.oneCharacter`): no line score
   * and no clip rule checks them.
   */
  readonly oneCharacterLines?: number;
  readonly diffBox?: readonly [number, number, number, number];
  /** What the harness page reported: the sheet, the counts, layout problems, and fonts. */
  readonly page?: Readonly<Record<string, unknown>>;
}

/**
 * One problem of the gate. `record` problems go away with `--update-gaps`; `fix` problems need a
 * fix of the code or of the fixture. A `note` does not fail the gate: outside CI, other tools than
 * the ones that measured the gaps. Their scores still have to stay inside `GAP_TOLERANCE`.
 */
export interface GateProblem {
  readonly kind: 'fix' | 'record' | 'note';
  readonly text: string;
  /** For a `fix` problem of one fixture: the fixture, and what fails (`FixCause`). */
  readonly fixture?: string;
  readonly cause?: FixCause;
}

/**
 * What a `fix` problem of a fixture fails: the comparison (`error`), a check that is not a known
 * gap or is worse than it (`gap`), the count of checked lines (`lines`), or a clipped line that
 * the list does not name or that the painted page no longer holds cut (`clip`).
 */
export type FixCause = 'error' | 'gap' | 'lines' | 'clip';

/**
 * The causes of the `fix` problems of the fixture `name`, in the order of `FixCause`, without
 * repeats. The report and the console name them in the status of the fixture.
 */
export function fixCauses(name: string, problems: readonly GateProblem[]): FixCause[] {
  const found = new Set(
    problems.filter((p) => p.kind === 'fix' && p.fixture === name).map((p) => p.cause!)
  );
  return (['error', 'gap', 'lines', 'clip'] as const).filter((c) => found.has(c));
}

/** The tools of a run. */
export interface GapTools {
  readonly poppler: string;
  readonly chromium: string;
}

/** Half the last digit of a recorded value (`round`): the rounding noise of a record. */
const half = (kind: FailureKind): number => (kind === 'pixels' ? 5e-6 : 5e-4);

/** Whether `value` is worse than `recorded` beyond the tolerance of its check. */
export function worse(kind: FailureKind, value: number, recorded: number): boolean {
  const t = GAP_TOLERANCE[kind] + half(kind);
  return kind === 'line-breaks' ? value < recorded - t : value > recorded + t;
}

/**
 * Whether a recorded gap value `next` is looser than the recorded value `was` by any amount. The
 * ratchet against the merge base uses it: a record never grows, even inside `GAP_TOLERANCE`.
 */
export function looserRecord(kind: FailureKind, next: number, was: number): boolean {
  return kind === 'line-breaks' ? next < was : next > was;
}

/** The problem of a run in a font environment that has no measured known-gap list. */
export function unmeasuredProblem(env: FontEnvironment): GateProblem {
  return {
    kind: 'fix',
    text: `the ${env} font environment has no measured known-gap list in scripts/fidelity/thresholds.ts, so this run is a gap, never a pass; measure it with \`bun run fidelity --update-gaps\` in that environment and review the diff`,
  };
}

/** Whether `value` is better than `recorded` beyond the tolerance of its check. */
export function better(kind: FailureKind, value: number, recorded: number): boolean {
  const t = GAP_TOLERANCE[kind] + half(kind);
  return kind === 'line-breaks' ? value > recorded + t : value < recorded - t;
}

/**
 * How far `value` can still move toward worse before the gate fails it against `recorded`: the
 * rest of `GAP_TOLERANCE`, in the units of the check. Negative when the value fails.
 */
export function headroom(kind: FailureKind, value: number, recorded: number): number {
  const t = GAP_TOLERANCE[kind] + half(kind);
  return kind === 'line-breaks' ? value - (recorded - t) : recorded + t - value;
}

/**
 * The text of a headroom or of another amount of a check, as the report writes it: percentage
 * points for a share, whole lines for extra text. By default it rounds down, so it never shows
 * more room than the gate gives; `nearest` rounds to the nearest shown digit.
 */
export function headroomText(kind: FailureKind, v: number, nearest = false): string {
  const down = (x: number, digits: number) =>
    (nearest ? Math.round(x * 10 ** digits) : Math.floor(x * 10 ** digits + 1e-9)) / 10 ** digits;
  if (kind === 'pixels') return `${(down(v, 5) * 100).toFixed(3)} percentage points`;
  if (kind === 'line-breaks') return `${(down(v, 3) * 100).toFixed(1)} percentage points`;
  if (kind === 'extra-text') {
    const lines = Math.max(0, nearest ? Math.round(v) : Math.floor(v + 1e-9));
    return `${lines} line${lines === 1 ? '' : 's'}`;
  }
  return `${down(v, 2).toFixed(2)} pt`;
}

const fmt = (kind: FailureKind, v: number): string =>
  kind === 'pixels'
    ? `${(v * 100).toFixed(3)}%`
    : kind === 'line-breaks'
      ? `${(v * 100).toFixed(1)}% of lines`
      : kind === 'extra-text'
        ? `${v} lines`
        : `${v.toFixed(2)} pt`;

/** The text of a measured value, as the report writes it. */
export const valueText = fmt;

/** The threshold of one check, as the report writes it. */
export function thresholdText(kind: FailureKind, fixture = ''): string {
  if (kind === 'pixels') {
    const limit = pixelLimit(fixture) * 100;
    return `≤ ${limit < 0.1 ? Number(limit.toPrecision(2)) : limit.toFixed(1)}%`;
  }
  if (kind === 'line-breaks') return 'all lines';
  if (kind === 'extra-text') return `≤ ${THRESHOLDS.maxExtraLines} lines`;
  if (kind === 'line-position') return `≤ ${THRESHOLDS.maxLineShiftPt} pt`;
  return `≤ ${THRESHOLDS.maxBaselineDeltaPt} pt`;
}

/** The gap entry that a result records: the value of each failing check. */
export function gapOf(r: FixtureResult): KnownGap | undefined {
  if (r.status !== 'compared' || !r.failures?.length || !r.values) return undefined;
  const gap: Partial<Record<FailureKind, number>> = {};
  for (const k of r.failures) gap[k] = round(k, r.values[k]);
  return gap;
}

const round = (kind: FailureKind, v: number): number =>
  kind === 'pixels' ? Math.round(v * 1e5) / 1e5 : Math.round(v * 1000) / 1000;

/** Whether a tool text names a version. A text with no digit, such as '', matches no list. */
const named = (text: string): boolean => /\d/.test(text);

/**
 * Whether a known-gap list was measured with exactly the tools of a run. A tool text with no
 * version number is as far as possible from every list: it never matches.
 */
export function sameTools(list: GapTools, tools: GapTools): boolean {
  return (
    named(tools.poppler) &&
    named(tools.chromium) &&
    list.poppler === tools.poppler &&
    list.chromium === tools.chromium
  );
}

/** The toolchain of a list as one text: `<poppler>; <chromium>`. */
export const toolchainText = (tools: GapTools): string => `${tools.poppler}; ${tools.chromium}`;

/** The list of exactly the tools of a run, or `undefined`. */
export function exactList(
  lists: readonly KnownGapList[],
  tools: GapTools
): KnownGapList | undefined {
  return lists.find((l) => sameTools(l, tools));
}

/** The tighter of two recorded values of one check. */
const tighter = (kind: FailureKind, a: number, b: number): number =>
  kind === 'line-breaks' ? Math.max(a, b) : Math.min(a, b);

/**
 * The tightest recorded value of each fixture and check across `lists`. A list that measured a
 * fixture and records no gap for a check passes it, and that pass is the tightest value. Line
 * counts take the largest count. `undefined` when there is no list. One list is returned as is.
 */
export function tightestList(lists: readonly KnownGapList[]): KnownGapList | undefined {
  if (lists.length <= 1) return lists[0];
  const fixtureLines: Record<string, number> = {};
  const fixtures: Record<string, KnownGap> = {};
  const names = new Set(
    lists.flatMap((l) => [...Object.keys(l.fixtureLines), ...Object.keys(l.fixtures)])
  );
  for (const name of names) {
    const measured = lists.filter(
      (l) => Object.hasOwn(l.fixtureLines, name) || Object.hasOwn(l.fixtures, name)
    );
    for (const l of measured)
      if (Object.hasOwn(l.fixtureLines, name))
        fixtureLines[name] = Math.max(fixtureLines[name] ?? 0, l.fixtureLines[name]);
    const gap: Partial<Record<FailureKind, number>> = {};
    for (const kind of FAILURE_KINDS) {
      let value: number | undefined;
      for (const l of measured) {
        const recorded = Object.hasOwn(l.fixtures, name) ? l.fixtures[name][kind] : undefined;
        if (recorded === undefined) {
          value = undefined;
          break;
        }
        value = value === undefined ? recorded : tighter(kind, value, recorded);
      }
      if (value !== undefined) gap[kind] = value;
    }
    if (Object.keys(gap).length) fixtures[name] = gap;
  }
  // Clipped lines take, for each line, the fewest times that a list names it: a list that names
  // none allows none.
  const fixtureClipped: Record<string, string[]> = {};
  for (const name of new Set(lists.flatMap((l) => Object.keys(l.fixtureClipped ?? {})))) {
    const counts = lists.map((l) => lineCounts(clippedOf(l, name)));
    const fewest: string[] = [];
    for (const [text, n] of counts[0]!)
      for (let i = 0; i < Math.min(n, ...counts.map((c) => c.get(text) ?? 0)); i += 1)
        fewest.push(text);
    if (fewest.length) fixtureClipped[name] = fewest.sort();
  }
  const unique = (texts: readonly string[]): string => [...new Set(texts)].join(' or ');
  return {
    poppler: `the tightest values of ${unique(lists.map((l) => l.poppler))}`,
    chromium: unique(lists.map((l) => l.chromium)),
    fixtureLines,
    ...(lists.some((l) => l.fixtureClipped) ? { fixtureClipped } : {}),
    fixtures,
  };
}

/** The problem of a CI run whose toolchain has no exact list. */
export function noExactListProblem(env: FontEnvironment, tools: GapTools): GateProblem {
  return {
    kind: 'fix',
    text: `no measured list for this toolchain (${toolchainText(tools)}) in the ${env} font environment; CI compares only with the list of its own tools, so measure it with \`bun run fidelity --update-gaps\` and review the diff`,
  };
}

/** The input of the gate of one run. */
export interface GateInput {
  readonly results: readonly FixtureResult[];
  readonly env: FontEnvironment;
  /** The tools of the run itself. */
  readonly tools: GapTools;
  /** The lists of the run's font environment. */
  readonly lists: readonly KnownGapList[];
  /** Whether every fixture ran. */
  readonly complete: boolean;
  /** Whether the run is in CI (`process.env.CI` is set). */
  readonly ci: boolean;
  /** The list that `--update-gaps` wrote for this run (`planUpdate`): the gate checks against it. */
  readonly recorded?: KnownGapList;
}

/**
 * The gate of one run. With `recorded`, the run checks against the list that it wrote.
 * Otherwise it uses the list of exactly its tools. With no such list, it compares with the
 * tightest recorded values of its environment (`tightestList`); in CI that is also a `fix`.
 * An environment with no list has no measurement and fails.
 */
export function gateRun(input: GateInput): { list?: KnownGapList; problems: GateProblem[] } {
  const { results, env, tools, lists, complete, ci } = input;
  if (input.recorded)
    return { list: input.recorded, problems: checkGate(results, input.recorded, tools, complete) };
  const exact = exactList(lists, tools);
  if (exact) return { list: exact, problems: checkGate(results, exact, tools, complete) };
  const tightest = tightestList(lists);
  if (!tightest)
    return {
      problems: [
        unmeasuredProblem(env),
        ...checkGate(results, recordedList([], tools), tools, false),
      ],
    };
  const problems = checkGate(results, tightest, tools, complete);
  return {
    list: tightest,
    problems: ci ? [noExactListProblem(env, tools), ...problems] : problems,
  };
}

/**
 * The lists of a font environment after `--update-gaps`: the list of the same tools is replaced,
 * or the new list is added. The lists of other tools stay as they are.
 */
export function withList(lists: readonly KnownGapList[], list: KnownGapList): KnownGapList[] {
  const at = lists.findIndex((l) => sameTools(l, list));
  return at < 0 ? [...lists, list] : lists.map((l, i) => (i === at ? list : l));
}

/**
 * Check the results against the known gaps of one font environment. `complete` is false when
 * only some fixtures ran: then a listed fixture that did not run is no problem.
 */
export function checkGate(
  results: readonly FixtureResult[],
  list: KnownGapList,
  tools: GapTools,
  complete: boolean
): GateProblem[] {
  const problems: GateProblem[] = [];
  const fix = (text: string, fixture: string, cause: FixCause): void =>
    void problems.push({ kind: 'fix', text, fixture, cause });
  const record = (text: string): void => void problems.push({ kind: 'record', text });
  const otherTools = !sameTools(list, tools);
  // With other raster tools (outside CI only, `gateRun`), a better score is a note, not a
  // request to tighten: the list records the values of other tools. A worse score still fails.
  const improved = (text: string): void =>
    void problems.push({ kind: otherTools ? 'note' : 'record', text });
  if (otherTools)
    problems.push({
      kind: 'note',
      text: `no known-gap list matches the tools of this run (${tools.poppler || 'no Poppler'} and ${tools.chromium || 'no Chromium'}); the gate compares with ${list.poppler || 'no Poppler'} and ${list.chromium || 'no Chromium'}`,
    });
  for (const r of results) {
    if (r.status === 'error') {
      fix(`${r.name}: ${r.error ?? 'not compared'}`, r.name, 'error');
      continue;
    }
    const gap: KnownGap = Object.hasOwn(list.fixtures, r.name) ? list.fixtures[r.name] : {};
    const failures = new Set(r.failures ?? []);
    for (const k of FAILURE_KINDS) {
      const value = r.values![k];
      const recorded = gap[k];
      if (failures.has(k) && recorded === undefined)
        fix(
          `${r.name}: fails ${k} (${fmt(k, value)}, threshold ${thresholdText(k, r.name)}); not a known gap`,
          r.name,
          'gap'
        );
      else if (failures.has(k) && worse(k, value, recorded!))
        fix(
          `${r.name}: ${k} is ${fmt(k, value)}, worse than its known gap ${fmt(k, recorded!)}`,
          r.name,
          'gap'
        );
      else if (failures.has(k) && better(k, value, recorded!))
        improved(
          `${r.name}: ${k} improved to ${fmt(k, value)} from ${fmt(k, recorded!)}; tighten the gap`
        );
      else if (!failures.has(k) && recorded !== undefined)
        improved(`${r.name}: passes ${k} now; remove the gap`);
    }
    const lines = Object.hasOwn(list.fixtureLines, r.name) ? list.fixtureLines[r.name] : undefined;
    if (lines !== undefined && (r.linesTotal ?? 0) < lines)
      fix(
        `${r.name}: checks ${r.linesTotal ?? 0} reference lines, fewer than ${lines}; ${FEWER_LINES_REMEDY}`,
        r.name,
        'lines'
      );
    else if (lines !== (r.linesTotal ?? 0))
      improved(`${r.name}: checks ${r.linesTotal ?? 0} reference lines; record the count`);
    // Clipped lines compare as a multiset: a line that a clip cuts twice needs two entries.
    const clipped = lineCounts(r.clippedLines ?? []);
    const allowed = lineCounts(clippedOf(list, r.name));
    for (const [text, n] of clipped)
      if (n > (allowed.get(text) ?? 0))
        fix(
          `${r.name}: a cell clip cuts the line \`${text}\` ${timesText(n)}, and the list names it ${timesText(allowed.get(text) ?? 0)}; ${CLIPPED_RUN_REMEDY}`,
          r.name,
          'clip'
        );
    // A reference line that is still cut, and that the painted page no longer holds cut, is a
    // regression: the gate never asks to record it away.
    const uncut = lineCounts(r.uncutLines ?? []);
    for (const [text, n] of uncut)
      fix(
        `${r.name}: the reference PDF clips the line \`${text}\`, and ${uncutCause(r, text)} (${timesText(n)}); ${UNCUT_REMEDY}`,
        r.name,
        'clip'
      );
    // The converse: the painted clip cuts a line that the reference page shows whole. Never
    // recorded.
    for (const [text, n] of lineCounts(r.extraCutLines ?? []))
      fix(
        `${r.name}: ${extraCutText(text)} (${timesText(n)}); ${EXTRA_CUT_REMEDY}`,
        r.name,
        'clip'
      );
    // A line end with an unknown reference clip is a skipped check: a coverage gap, never a pass.
    for (const [text, n] of lineCounts(r.unknownReferenceEnds ?? []))
      fix(
        isPaintedSplitKey(text)
          ? `${r.name}: the painted page splits the text of one reference item at \`${text}\` (${timesText(n)}), so no clip rule checks that end; ${PAINTED_SPLIT_REMEDY}`
          : isUnknownGlyphsKey(text)
            ? `${r.name}: the reference PDF has glyphs that do not account for the text at \`${text}\` (${timesText(n)}), so no glyph rule checks that end; ${UNKNOWN_GLYPHS_REMEDY}`
            : `${r.name}: the reference PDF has no known clip at \`${text}\` (${timesText(n)}), so no clip rule checks that end; ${UNKNOWN_REFERENCE_REMEDY}`,
        r.name,
        'clip'
      );
    // A named clipped line whose reference clip or glyphs are unknown now: the unknown end is the
    // cause, and a new record would hide it.
    const unknownLines = new Map<string, boolean>();
    for (const key of r.unknownReferenceEnds ?? []) {
      if (isPaintedSplitKey(key)) continue;
      const text = key.slice(0, key.lastIndexOf(' | '));
      unknownLines.set(text, (unknownLines.get(text) ?? false) || isUnknownGlyphsKey(key));
    }
    for (const [text, n] of allowed)
      if (n > (clipped.get(text) ?? 0) + (uncut.get(text) ?? 0)) {
        const glyphs = unknownLines.get(text.slice(0, text.lastIndexOf(' | ')));
        if (glyphs === true)
          fix(
            `${r.name}: the list names the clipped line \`${text}\` ${timesText(n)}, and the reference glyphs of that line are unknown now, so no glyph rule checks it; a new record would hide it; ${UNKNOWN_GLYPHS_REMEDY}`,
            r.name,
            'clip'
          );
        else if (glyphs === false)
          fix(
            `${r.name}: the list names the clipped line \`${text}\` ${timesText(n)}, and the reference clip of that line is unknown now, so no rule matches it; a new record would hide it; ${UNKNOWN_REFERENCE_REMEDY}`,
            r.name,
            'clip'
          );
        else
          improved(
            `${r.name}: a cell clip cuts the line \`${text}\` ${timesText(clipped.get(text) ?? 0)} now, and the list names it ${timesText(n)}; record the list again`
          );
      }
  }
  if (complete) {
    const ran = new Set(results.map((r) => r.name));
    for (const name of Object.keys(list.fixtures))
      if (!ran.has(name)) record(`${name}: has a known gap but did not run; remove it`);
  }
  return problems;
}

/**
 * The remedy of a run for a clipped line that its list does not name: the line scores cannot
 * check the glyphs that a clip hides. The owner route holds only for a new fixture or a fixture
 * with a new reference PDF: the merge-base check refuses a new clipped line of any other fixture
 * of the merge base (`CLIPPED_BASE_REMEDY`).
 */
export const CLIPPED_RUN_REMEDY =
  'fix the layout or the fixture; for a new fixture or a new reference PDF only, an owner may name the line in `fixtureClipped` of each list and run `bun run fidelity:check-gaps` with `XLSX_FIDELITY_ACCEPT_CLIPPED=<fixture>`';

/**
 * The remedy of a run for a reference line that its clip cuts and that the painted page does not
 * hold cut at the same edge: the paint lost or moved the clip of the cell, or the layout lost the
 * line.
 */
export const UNCUT_REMEDY =
  'fix the clip of the cell in the layout or the paint; a lost cut is a regression, not a record';

/**
 * Why a reference line that its clip cuts is not cut on the painted page
 * (`FixtureResult.uncutLines`, by `key`). The cause comes from `compareLines` (`lostUncutLines`),
 * never from a text match: the painted line is gone, or it is not cut at that edge.
 */
export function uncutCause(r: FixtureResult, key: string): string {
  return r.lostUncutLines?.includes(key)
    ? 'the painted page lost the line or cut it elsewhere'
    : 'the painted page does not hold it cut at that edge';
}

/**
 * Whether an extra cut of `compareLines` names an end whose painted clip is unknown (`... |
 * <side> clip unknown`), not a painted clip edge that cuts the line.
 */
export function isUnknownClipKey(key: string): boolean {
  return key.endsWith(' clip unknown');
}

/**
 * Whether a key of an unknown end (`LineScore.unknownReference`) names an inner end of a painted
 * piece that splits the text of one reference item, not an unknown reference clip: the clip may be
 * known, and the painted split is the cause.
 */
export function isPaintedSplitKey(key: string): boolean {
  return key.endsWith(' painted split inside a reference item');
}

/**
 * Whether a key of an unknown reference end (`LineScore.unknownReference`) names reference glyphs
 * that do not account for the text of the line, such as a ligature, and not an unknown clip.
 */
export function isUnknownGlyphsKey(key: string): boolean {
  return key.endsWith(' reference glyphs unknown');
}

/** What an extra cut of `compareLines` (`LineScore.extraCut`) says, for its key. */
export function extraCutText(key: string): string {
  return isUnknownClipKey(key)
    ? `the reference clip holds the line \`${key}\` whole at an end, and the painted clip there is unknown: no painted run with a one-rectangle clip that holds the baseline lies at that end, or runs with different clips start an item of the next cell there`
    : `the painted page cuts the line \`${key}\` at an end that the reference clip holds whole`;
}

/**
 * The remedy of a run for a line end with an unknown reference clip: the reference PDF writes no
 * clip there as one rectangle that holds the baseline, or its items do not split at a cell.
 */
export const UNKNOWN_REFERENCE_REMEDY =
  'export the reference PDF again, or name the limit of the clip reader in docs/site/content/known-approximations.mdx and fix it; a skipped check is a gap, not a pass';

/**
 * The remedy of a run for an inner end of a painted piece that splits the text of one reference
 * item (`isPaintedSplitKey`). Mostly the cause is the layout or the paint. The reference PDF can
 * also draw two cells as one item (approximation "Reference items across cells").
 */
export const PAINTED_SPLIT_REMEDY =
  'check the reference runs of that item, which `textRuns` in `scripts/fidelity/pdf-clips.ts` reads from the reference PDF of the fixture: if they hold the text of two cells, the reference PDF draws two cells as one item (approximation "Reference items across cells"); else fix the layout or the paint of the cell, which splits text that the reference PDF draws as one item; a skipped check is a gap, not a pass';

/**
 * The remedy of a run for a cut line end whose reference glyphs do not account for the text of
 * the line (`isUnknownGlyphsKey`), such as a ligature or a glyph without Unicode.
 */
export const UNKNOWN_GLYPHS_REMEDY =
  'export the reference PDF again, or name the limit of the glyph reader (`textRuns` in `pdf-clips.ts`) in docs/site/content/known-approximations.mdx and fix it; a skipped check is a gap, not a pass';

/**
 * The remedy of a run for a line that the painted page cuts at an end that the reference clip
 * holds whole, or whose painted clip is unknown at such an end (`extraCutText`).
 */
export const EXTRA_CUT_REMEDY =
  'fix the clip of the cell in the layout or the paint; an extra cut is a regression, not a record';

/**
 * The remedy of the merge-base check for a fixture of the merge base that gains a clipped line:
 * no acceptance applies to it.
 */
export const CLIPPED_BASE_REMEDY = 'fix the layout or the fixture';

/**
 * The remedy of the merge-base check for a fixture of the merge base whose reference PDF changed
 * and that gains a clipped line, such as a clip edge that the new reference PDF moves: an owner
 * accepts them for that run.
 */
export const CLIPPED_REFERENCE_REMEDY =
  'fix the layout or the fixture, or, because the fixture has a new reference PDF, as an owner only, run `bun run fidelity:check-gaps` with `XLSX_FIDELITY_ACCEPT_CLIPPED=<fixture>`';

/**
 * The remedy for a clipped line of a fixture that the merge base does not have, in a change that
 * removes a fixture of the merge base: the fixture can be a renamed one, so it is not new.
 */
export const CLIPPED_RENAMED_REMEDY =
  'a fixture of the merge base is gone in this change, so no fixture is new; keep the old fixture name, or remove the old fixture and add the new one in separate changes; then, as an owner only, run `bun run fidelity:check-gaps` with `XLSX_FIDELITY_ACCEPT_CLIPPED=<fixture>` in the change that adds it';

/**
 * The remedy for a fixture of the merge base that the change removes or renames while it adds a
 * fixture: the checked lines of the old fixture would otherwise go with it.
 */
export const RENAMED_LINES_REMEDY =
  'keep the old fixture name, or remove the old fixture and add the new one in separate changes';

/**
 * The remedy for a fixture that checks fewer reference lines than its list or its merge base
 * records. No list takes fewer lines, also not with a new reference PDF.
 */
export const FEWER_LINES_REMEDY =
  'fix the layout so that the fixture checks its lines again, or, as an owner only, remove the fixture and add it as a new one in separate changes; a fixture with a known gap cannot be removed, so fix its gap first, or keep the fixture';

/**
 * The remedy of the merge-base check for a new fixture with clipped lines: an owner accepts them
 * for that run.
 */
export const CLIPPED_NEW_REMEDY =
  'fix the layout or the fixture, or, as an owner only, run `bun run fidelity:check-gaps` with `XLSX_FIDELITY_ACCEPT_CLIPPED=<fixture>`';

/** A count of times as text: `1 time`, `2 times`. */
export const timesText = (n: number): string => `${n} ${n === 1 ? 'time' : 'times'}`;

/** How many times each line occurs in `lines`. */
export function lineCounts(lines: readonly string[]): Map<string, number> {
  const counts = new Map<string, number>();
  for (const text of lines) counts.set(text, (counts.get(text) ?? 0) + 1);
  return counts;
}

/** The clipped lines that a list allows a fixture (`KnownGapList.fixtureClipped`). */
export function clippedOf(list: KnownGapList, name: string): readonly string[] {
  const record = list.fixtureClipped ?? {};
  return Object.hasOwn(record, name) ? record[name] : [];
}

/** The known-gap list that a complete run records. */
export function recordedList(results: readonly FixtureResult[], tools: GapTools): KnownGapList {
  const fixtures: Record<string, KnownGap> = {};
  const fixtureLines: Record<string, number> = {};
  const fixtureClipped: Record<string, readonly string[]> = {};
  for (const r of [...results].sort((a, b) => (a.name < b.name ? -1 : 1))) {
    if (r.status !== 'compared') continue;
    fixtureLines[r.name] = r.linesTotal ?? 0;
    if (r.clippedLines?.length) fixtureClipped[r.name] = [...r.clippedLines].sort();
    const gap = gapOf(r);
    if (gap) fixtures[r.name] = gap;
  }
  return { ...tools, fixtureLines, fixtureClipped, fixtures };
}

const START = 'export const KNOWN_GAPS';

/** Format `file` with Prettier: its exit code and the first line of its error. */
function prettierWrite(file: string): { readonly status: number | null; readonly error: string } {
  const run = spawnSync('bunx', ['prettier', '--write', file], { encoding: 'utf8' });
  return { status: run.status, error: run.error?.message ?? run.stderr.split('\n')[0] ?? '' };
}

/**
 * Rewrite the known gaps of one font environment and the list's toolchain in `thresholds.ts`,
 * then format the file (`format`, replaced in tests). The lists of other toolchains stay as
 * they are. A format that fails throws: the list is written, but `bun run format:check` would
 * fail on it.
 */
export function writeGaps(
  file: string,
  env: FontEnvironment,
  list: KnownGapList,
  format: (file: string) => {
    readonly status: number | null;
    readonly error: string;
  } = prettierWrite
): void {
  const all = { ...KNOWN_GAPS, [env]: withList(KNOWN_GAPS[env] ?? [], list) };
  const text = readFileSync(file, 'utf8');
  const at = text.indexOf(START);
  if (at < 0) throw new Error(`${file} has no ${START}`);
  const body = `${START}: Readonly<Partial<Record<FontEnvironment, readonly KnownGapList[]>>> = ${JSON.stringify(all, null, 2)};\n`;
  writeFileSync(file, text.slice(0, at) + body);
  const formatted = format(file);
  if (formatted.status !== 0)
    throw new Error(`prettier did not format ${file}: ${formatted.error}; run bun run format`);
}
