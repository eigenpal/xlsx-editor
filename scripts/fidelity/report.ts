// The report of `bun run fidelity`: `fidelity-report/report.md` and `results.json`, and for each
// fixture the rasters of both pages and the diff overlay. Same role as
// pptx-editor/scripts/fidelity/report.ts at 1943d2c9 (EigenPal); it writes Markdown instead of
// HTML, so no markup is built from file data.

import { writeFileSync } from 'node:fs';
import { isAbsolute, join, relative, resolve } from 'node:path';
import {
  FAILURE_KINDS,
  FIXTURE_PIXEL_MISMATCH,
  type FailureKind,
  type FontEnvironment,
  type KnownGapList,
} from './thresholds.ts';
import {
  headroom,
  headroomText,
  fixCauses,
  uncutCause,
  isUnknownClipKey,
  isPaintedSplitKey,
  isUnknownGlyphsKey,
  valueText,
  thresholdText,
  type FixtureResult,
  type GateProblem,
  type GapTools,
} from './known-gaps.ts';

export interface ReportContext {
  readonly env: FontEnvironment;
  /** The fonts of the run: the reference faces, or the substitutes. */
  readonly fonts: string;
  readonly tools: GapTools;
  /**
   * The known-gap list that the gate compared with: the list of the run's toolchain, or the
   * tightest recorded values of the environment (`tightestList`). Absent when it has none.
   */
  readonly gaps?: KnownGapList;
  readonly problems: readonly GateProblem[];
}

// Escape the backslash before the pipe, so a backslash in the text cannot cancel the escape of
// the pipe after it and end the table cell.
const cell = (s: string): string =>
  s
    .replace(/\\/g, '\\\\')
    .replace(/\|/g, '\\|')
    .replace(/\r?\n|\r/g, ' ');

/**
 * The verdict of a fixture in the report, after the gate: a fixture with a `fix` problem fails
 * with its causes (`fixCauses`), also when every check passes, such as a clipped line that the
 * list does not name.
 */
export function verdict(
  r: FixtureResult,
  gaps: KnownGapList | undefined,
  problems: readonly GateProblem[] = []
): string {
  if (r.status === 'error') return 'error';
  if (!gaps) return r.failures?.length ? 'fail (no measured record)' : 'unmeasured (no record)';
  const causes = fixCauses(r.name, problems);
  if (causes.length) return `fail (${causes.join(', ')})`;
  if (!r.failures?.length) return 'pass';
  return Object.hasOwn(gaps.fixtures, r.name) ? 'fail (known gap)' : 'fail';
}

/** The coverage gap of the harness: it renders only the print profile. */
export const SCREEN_PROFILE_GAP =
  'The harness renders the print profile (`PRINT_PROFILE`) only. The screen profile (`SCREEN_PROFILE`), the default of the editor, has no fidelity measurement.';

/** The coverage gap of the harness: no check compares the face or size that the canvas draws. */
export const DRAWN_FONT_GAP =
  'The harness does not check that each glyph draws in the face and size of its cell.';

/** The coverage gap of the line scores: they score the nearer end of a line, as PPTX Editor does. */
export const LINE_WIDTH_GAP =
  'The line scores take the nearer end of each line, so no check bounds the width of a line or the place of a word inside it. The pixel score counts such glyphs only as a share of the whole page.';

/**
 * The coverage gap of the font substitution: the fixtures name the faces that the harness
 * draws, so no fixture shows that a substitute keeps the advances of the family it replaces.
 */
export const SUBSTITUTE_FACE_GAP =
  'No fixture names a family that the editor draws with a substitute face, such as Calibri. The fixtures name the substitute faces themselves, so no fidelity check measures the substitution. A row that cites a fidelity fixture shows the substitute faces only.';

/** The coverage gap of the pixel score on a fixture whose rasterizer difference nearly fills the limit. */
export const RASTERIZER_MARGIN_GAP =
  'On a fixture with many glyphs, the rasterizer difference takes most of the page limit. So the pixel score of such a fixture cannot see a paint change that is smaller than the rest of its limit. Design D8 (`openspec/changes/typed-ooxml-workbook-editor/design.md`) gives the measured `number-formats` scores and margins on each recorded toolchain. `bun scripts/fidelity/rasterizer-control.ts <fixture>` compares the editor page with a pdf.js raster in the same Chromium, but no gate runs it.';

/**
 * Every coverage gap of the harness that holds on each run. `bun run fidelity` prints each one
 * after "Coverage gap: ". `docs/site/content/known-approximations.mdx` lists the same text under
 * "Coverage gaps", and `harness.test.ts` keeps the two equal. The gaps with a count of their own
 * (`oneCharacterGap`, `rotatedTextGap`) print only on a run that has such text.
 */
export const COVERAGE_GAPS = [
  SCREEN_PROFILE_GAP,
  DRAWN_FONT_GAP,
  LINE_WIDTH_GAP,
  RASTERIZER_MARGIN_GAP,
  SUBSTITUTE_FACE_GAP,
] as const;

/** A coverage gap as the run and the report print it. */
export const coverageGapLine = (gap: string): string => `Coverage gap: ${gap}`;

/**
 * The coverage gap of the line scores: a line of one character other than a digit, such as a line
 * of stacked text. `count` is the number of such reference lines in the run.
 */
export function oneCharacterGap(count: number): string {
  return coverageGapLine(
    `A line of one character other than a digit has no line score, so no clip rule checks its cell (approximation "Line scores of clipped text"); ${count} ${count === 1 ? 'line' : 'lines'} on this run.`
  );
}

/** The reference lines of the run of one character other than a digit. */
export function oneCharacterLines(results: readonly FixtureResult[]): number {
  return results.reduce((n, r) => n + (r.oneCharacterLines ?? 0), 0);
}

/**
 * The coverage gap of the line scores: `pdfItems` leaves out rotated, mirrored, and vertical text,
 * so its lines have no score. `count` is the number of such reference items in the run.
 */
export function rotatedTextGap(count: number): string {
  return coverageGapLine(
    `Rotated, mirrored, and vertical text has no line score, so no clip rule checks its cell; ${count} ${count === 1 ? 'item' : 'items'} on this run.`
  );
}

/** The reference text items of the run that the line scores leave out as rotated. */
export function rotatedItems(results: readonly FixtureResult[]): number {
  return results.reduce((n, r) => n + (r.rotatedItems ?? 0), 0);
}

/**
 * The environment variable that names another report folder. Tests and probes set it. The run
 * and the helper tools (`glyph-share.ts`, `rasterizer-control.ts`) read the folder through
 * `reportDir`, so they read the same folder.
 */
export const REPORT_DIR_ENV = 'XLSX_FIDELITY_REPORT_DIR';

const REPO_ROOT = join(import.meta.dir, '..', '..');

/**
 * The report folder of a run: `XLSX_FIDELITY_REPORT_DIR`, or `fidelity-report/` of the repository.
 * A relative value resolves against the repository root, because `bun run` starts the run there
 * while a helper tool starts in the shell folder.
 */
export function reportDir(env: Readonly<Record<string, string | undefined>> = process.env): string {
  const dir = env[REPORT_DIR_ENV];
  return dir ? resolve(REPO_ROOT, dir) : join(REPO_ROOT, 'fidelity-report');
}

/**
 * The remedy of a helper tool that finds no raster of a run (`first`) or a raster older than the
 * reference PDF (`again`): run the harness, with the same `XLSX_FIDELITY_REPORT_DIR` when it is
 * set, so that the run writes the folder that the tool reads.
 */
export function runFirstText(
  env: Readonly<Record<string, string | undefined>> = process.env,
  when: 'first' | 'again' = 'first'
): string {
  return env[REPORT_DIR_ENV]
    ? `run \`bun run fidelity\` with the same ${REPORT_DIR_ENV} ${when}`
    : `run \`bun run fidelity\` ${when}`;
}

/**
 * The path of `report.md` in `reportDir` that the run prints: relative to `cwd` (the current
 * folder, as in the source) when the folder is inside it, and absolute when it is outside (a
 * relative path would climb with `..`).
 */
export function reportPathText(cwd: string, reportDir: string): string {
  const file = resolve(reportDir, 'report.md');
  const rel = relative(cwd, file);
  return rel === '..' || rel.startsWith('../') || rel.startsWith('..\\') || isAbsolute(rel)
    ? file
    : rel;
}

/**
 * The report line of a known gap that still fails: the measured and recorded values, how much
 * worse than the record the value is, and how far it can move before the gate fails. A
 * `line-breaks` score gets worse as it drops; every other score gets worse as it grows.
 */
export function knownGapNote(kind: FailureKind, value: number, recorded: number): string {
  const over = kind === 'line-breaks' ? recorded - value : value - recorded;
  const worse = over > 0 ? `, ${headroomText(kind, over, true)} worse than the record` : '';
  const room = headroom(kind, value, recorded);
  const move = kind === 'line-breaks' ? 'drop' : 'grow';
  const margin =
    room >= 0
      ? `it can ${move} ${headroomText(kind, room)} more before the gate fails`
      : `it is ${headroomText(kind, -room, true)} past the tolerance, so the gate fails`;
  return `Known gap, ${kind}: measured ${valueText(kind, value)}, recorded ${valueText(kind, recorded)}${worse}; ${margin}`;
}

/** Write the report into `dir`. */
export function writeReport(
  dir: string,
  results: readonly FixtureResult[],
  ctx: ReportContext
): void {
  const out: string[] = [
    '# Fidelity report',
    '',
    `- Font environment: ${ctx.env}`,
    `- Fonts: ${cell(ctx.fonts)}`,
    `- Tools: ${ctx.tools.poppler}, ${ctx.tools.chromium}`,
    ...(ctx.gaps ? [`- Known gaps measured with: ${ctx.gaps.poppler}, ${ctx.gaps.chromium}`] : []),
    `- Thresholds: ${FAILURE_KINDS.map((k) => `${k} ${thresholdText(k)}`).join('; ')}`,
    ...Object.keys(FIXTURE_PIXEL_MISMATCH).map(
      (name) => `- Pixel limit of ${name}: ${thresholdText('pixels', name)}`
    ),
    '',
  ];
  for (const gap of COVERAGE_GAPS) out.push(coverageGapLine(gap), '');
  const single = oneCharacterLines(results);
  if (single) out.push(oneCharacterGap(single), '');
  const rotated = rotatedItems(results);
  if (rotated) out.push(rotatedTextGap(rotated), '');
  if (!ctx.gaps)
    out.push(
      `The ${ctx.env} font environment has no measured known-gap list in \`scripts/fidelity/thresholds.ts\`. The missing measurement is a gap, never a pass, and the gate fails.`,
      ''
    );
  if (ctx.env === 'substitute')
    out.push(
      'The run drew the open faces of `@docx-editor.dev/fonts`, because `XLSX_REFERENCE_FONTS` names no folder. The fixtures name these faces.',
      ''
    );
  out.push(
    '| Fixture | Verdict | Pixels | Lines | Baseline | Line position | Extra lines |',
    '| --- | --- | --- | --- | --- | --- | --- |'
  );
  for (const r of results) {
    if (r.status === 'error') {
      out.push(`| ${r.name} | error: ${cell(r.error ?? '')} | | | | | |`);
      continue;
    }
    const v = r.values!;
    out.push(
      `| ${r.name} | ${verdict(r, ctx.gaps, ctx.problems)} | ${valueText('pixels', v.pixels)} | ${r.linesMatched} of ${r.linesTotal} | ${valueText('baseline', v.baseline)} | ${valueText('line-position', v['line-position'])} | ${v['extra-text']} |`
    );
  }
  out.push('');
  for (const r of results) {
    if (r.status !== 'compared') continue;
    const notes: string[] = [];
    if (r.missingLines?.length)
      notes.push(
        `- Reference lines not found: ${r.missingLines.map((l) => `\`${cell(l)}\``).join(', ')}`
      );
    if (r.clippedLines?.length)
      notes.push(
        `- Matched lines that a cell clip shortens on both pages: ${r.clippedLines.map((l) => `\`${cell(l)}\``).join(', ')}. The printed PDF of the painted page leaves out the glyphs that the clip hides; the reference PDF keeps them as text.`
      );
    const lostCut = (r.uncutLines ?? []).filter((l) => uncutCause(r, l).includes('lost'));
    const keptUncut = (r.uncutLines ?? []).filter((l) => !lostCut.includes(l));
    if (keptUncut.length)
      notes.push(
        `- Reference lines that a cell clip cuts and that the painted page does not hold cut at that edge: ${keptUncut.map((l) => `\`${cell(l)}\``).join(', ')}. Fix the clip of the cell; a lost cut is a regression.`
      );
    if (lostCut.length)
      notes.push(
        `- Reference lines that a cell clip cuts and that the painted page lost or cut elsewhere: ${lostCut.map((l) => `\`${cell(l)}\``).join(', ')}. Fix the clip of the cell; a lost cut is a regression.`
      );
    const extraCut = (r.extraCutLines ?? []).filter((l) => !isUnknownClipKey(l));
    const unknownClip = (r.extraCutLines ?? []).filter(isUnknownClipKey);
    if (extraCut.length)
      notes.push(
        `- Lines that the painted page cuts at an end that the reference clip holds whole: ${extraCut.map((l) => `\`${cell(l)}\``).join(', ')}. Fix the clip of the cell; an extra cut is a regression.`
      );
    if (unknownClip.length)
      notes.push(
        `- Lines that the reference clip holds whole at an end where the painted clip is unknown: no painted run with a one-rectangle clip that holds the baseline lies there, or runs with different clips start an item of the next cell there: ${unknownClip.map((l) => `\`${cell(l)}\``).join(', ')}. Fix the clip of the cell; an unknown painted clip at that end is a regression.`
      );
    const unknownClipEnds = (r.unknownReferenceEnds ?? []).filter(
      (l) => !isUnknownGlyphsKey(l) && !isPaintedSplitKey(l)
    );
    const unknownGlyphEnds = (r.unknownReferenceEnds ?? []).filter(isUnknownGlyphsKey);
    const paintedSplitEnds = (r.unknownReferenceEnds ?? []).filter(isPaintedSplitKey);
    if (paintedSplitEnds.length)
      notes.push(
        `- Inner ends of painted pieces that split the text of one reference item, which no clip rule checks: ${paintedSplitEnds.map((l) => `\`${cell(l)}\``).join(', ')}. Check the reference runs of each item (\`textRuns\` in \`scripts/fidelity/pdf-clips.ts\`): two cells in one item is approximation "Reference items across cells"; else fix the layout or the paint of the cell. A skipped check is a gap, not a pass.`
      );
    if (unknownClipEnds.length)
      notes.push(
        `- Line ends with no known reference clip, which no clip rule checks: ${unknownClipEnds.map((l) => `\`${cell(l)}\``).join(', ')}. A skipped check is a gap, not a pass.`
      );
    if (unknownGlyphEnds.length)
      notes.push(
        `- Cut line ends whose reference glyphs do not account for the text, such as a ligature, which no glyph rule checks: ${unknownGlyphEnds.map((l) => `\`${cell(l)}\``).join(', ')}. A skipped check is a gap, not a pass.`
      );
    if (r.bothCutLines?.length)
      notes.push(
        `- Missing lines that the reference clip cuts on both sides: ${r.bothCutLines.map((l) => `\`${cell(l)}\``).join(', ')}. The line scores match no part of such a line (approximation "Line scores of clipped text").`
      );
    if (r.extraLines?.length)
      notes.push(`- Extra lines: ${r.extraLines.map((l) => `\`${cell(l)}\``).join(', ')}`);
    const known = ctx.gaps?.fixtures[r.name];
    // Each check that still fails against its record. A check that passes again has no gap to
    // describe: the gate asks to remove its record.
    for (const [kind, recorded] of Object.entries(known ?? {}) as [FailureKind, number][]) {
      const value = r.values?.[kind];
      if (typeof value !== 'number' || !r.failures?.includes(kind)) continue;
      notes.push(`- ${knownGapNote(kind, value, recorded)}`);
    }
    if (r.diffBox)
      notes.push(`- Differing pixels inside [${r.diffBox.join(', ')}] (px at 144 dpi)`);
    const problems = r.page?.layoutProblems;
    if (Array.isArray(problems) && problems.length)
      notes.push(`- Layout problems: ${problems.join(', ')}`);
    if (notes.length === 0) continue;
    out.push(`## ${r.name}`, '', ...notes, '');
  }
  if (ctx.problems.length) {
    out.push('## Gate problems', '');
    for (const p of ctx.problems) out.push(`- ${p.kind}: ${p.text}`);
    out.push('');
  }
  writeFileSync(join(dir, 'report.md'), out.join('\n'));
  writeFileSync(
    join(dir, 'results.json'),
    `${JSON.stringify({ env: ctx.env, fonts: ctx.fonts, tools: ctx.tools, results }, null, 2)}\n`
  );
}

/**
 * The console lines of the clip results of a fixture, which the CI log records, so that a CI list
 * can name them (`fixtureClipped`).
 */
export function clipLogLines(r: FixtureResult): string[] {
  const named = (label: string, lines: readonly string[] | undefined): string[] =>
    (lines ?? []).map((line) => `  ${label}: ${JSON.stringify(line)}`);
  const unknown = r.unknownReferenceEnds ?? [];
  return [
    ...named('clipped', r.clippedLines),
    ...named('uncut', r.uncutLines),
    ...named('extra cut', r.extraCutLines),
    ...named(
      'unknown reference clip',
      unknown.filter((l) => !isUnknownGlyphsKey(l) && !isPaintedSplitKey(l))
    ),
    ...named('unknown reference glyphs', unknown.filter(isUnknownGlyphsKey)),
    ...named('painted split inside a reference item', unknown.filter(isPaintedSplitKey)),
    ...named('missing, cut on both sides', r.bothCutLines),
  ];
}
