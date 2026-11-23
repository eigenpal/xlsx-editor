// The Markdown report of `bun run fidelity` (`report.ts`).

import { afterEach, describe, expect, test } from 'bun:test';
import { mkdtempSync, readFileSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import type { FixtureResult, GateProblem } from './known-gaps.ts';
import {
  clipLogLines,
  DRAWN_FONT_GAP,
  oneCharacterGap,
  REPORT_DIR_ENV,
  reportDir,
  reportPathText,
  runFirstText,
  rotatedItems,
  rotatedTextGap,
  writeReport,
} from './report.ts';

const dirs: string[] = [];
afterEach(() => {
  for (const d of dirs.splice(0)) rmSync(d, { recursive: true, force: true });
});

const values = {
  pixels: 0,
  'line-breaks': 1,
  baseline: 0,
  'line-position': 0,
  'extra-text': 0,
};

function report(result: Partial<FixtureResult>, problems: readonly GateProblem[] = []): string {
  const dir = mkdtempSync(join(tmpdir(), 'xlsx-report-'));
  dirs.push(dir);
  writeReport(
    dir,
    [
      {
        name: 'preserved-parts',
        status: 'compared',
        failures: [],
        values,
        linesTotal: 2,
        linesMatched: 2,
        ...result,
      },
    ],
    {
      env: 'substitute',
      fonts: 'substitutes',
      tools: { poppler: 'poppler 1', chromium: 'chromium 1' },
      gaps: {
        poppler: 'poppler 1',
        chromium: 'chromium 1',
        fixtureLines: { 'preserved-parts': 2 },
        fixtures: {},
      },
      problems,
    }
  );
  return readFileSync(join(dir, 'report.md'), 'utf8');
}

describe('the fidelity report', () => {
  test('a fixture with clipped lines names them', () => {
    expect(report({ clippedLines: ['Kept content'] })).toContain(
      '- Matched lines that a cell clip shortens on both pages: `Kept content`.'
    );
    expect(report({})).not.toContain('cell clip');
    expect(report({ bothCutLines: ['Wide title'] })).toContain(
      '- Missing lines that the reference clip cuts on both sides: `Wide title`.'
    );
  });

  test('the drawn font gap names its issue and no layout check', () => {
    expect(DRAWN_FONT_GAP).toContain('face and size of its cell');
    expect(DRAWN_FONT_GAP).not.toContain('layout');
    expect(report({})).toContain(DRAWN_FONT_GAP);
    // The one-character limit of the line scores is stated with its count when the run has such
    // lines.
    expect(oneCharacterGap(8)).toEndWith('; 8 lines on this run.');
    expect(oneCharacterGap(1)).toEndWith('; 1 line on this run.');
    expect(report({})).not.toContain('one character other than a digit');
    expect(report({ oneCharacterLines: 8 })).toContain(oneCharacterGap(8));
  });

  test('rotated, mirrored, and vertical text is a stated gap with its count', () => {
    expect(rotatedTextGap(3)).toBe(
      'Coverage gap: Rotated, mirrored, and vertical text has no line score, so no clip rule checks its cell; 3 items on this run.'
    );
    expect(rotatedTextGap(1)).toEndWith('; 1 item on this run.');
    // The report states it only when a reference PDF of the run has such text.
    expect(report({})).not.toContain('rotated, mirrored');
    expect(report({ rotatedItems: 2 })).toContain(rotatedTextGap(2));
    const error: FixtureResult = { name: 'formulas', status: 'error', error: 'no reference' };
    expect(rotatedItems([error, { ...error, status: 'compared', rotatedItems: 4 }])).toBe(4);
  });

  test('an end with unknown reference glyphs has its own note and console label', () => {
    const clip = 'Kept @ 47.0 | right reference clip unknown';
    const glyphs = 'Kept content @ 47.0 | right reference glyphs unknown';
    const text = report({ unknownReferenceEnds: [clip, glyphs] });
    expect(text).toContain(
      '- Line ends with no known reference clip, which no clip rule checks: `Kept @ 47.0 \\| right reference clip unknown`. A skipped check is a gap, not a pass.'
    );
    expect(text).toContain(
      '- Cut line ends whose reference glyphs do not account for the text, such as a ligature, which no glyph rule checks: `Kept content @ 47.0 \\| right reference glyphs unknown`. A skipped check is a gap, not a pass.'
    );
    expect(
      clipLogLines({ name: 'a', status: 'compared', unknownReferenceEnds: [clip, glyphs] })
    ).toEqual([
      `  unknown reference clip: ${JSON.stringify(clip)}`,
      `  unknown reference glyphs: ${JSON.stringify(glyphs)}`,
    ]);
    // An inner end of a painted split inside a reference item has its own label and note.
    const split = '1 @ 47.0 | right painted split inside a reference item';
    expect(clipLogLines({ name: 'a', status: 'compared', unknownReferenceEnds: [split] })).toEqual([
      `  painted split inside a reference item: ${JSON.stringify(split)}`,
    ]);
    expect(report({ unknownReferenceEnds: [split] })).toContain(
      '- Inner ends of painted pieces that split the text of one reference item'
    );
    expect(report({ unknownReferenceEnds: [split] })).not.toContain(
      '- Line ends with no known reference clip'
    );
    // A reference line that is missing and cut on both sides has its own label in the log.
    const both = 'Kept content @ 47.0';
    expect(clipLogLines({ name: 'a', status: 'compared', bothCutLines: [both] })).toEqual([
      `  missing, cut on both sides: ${JSON.stringify(both)}`,
    ]);
  });

  test('an uncut line that the painted page lost has its own note', () => {
    const key = 'Kept content @ 47.0 | right clip 89.0';
    expect(report({ uncutLines: [key], lostUncutLines: [key] })).toContain(
      '- Reference lines that a cell clip cuts and that the painted page lost or cut elsewhere: `Kept content @ 47.0 \\| right clip 89.0`.'
    );
  });

  test('a fixture with a fix problem of the gate fails with its cause, and names uncut lines', () => {
    expect(report({})).toContain('| preserved-parts | pass |');
    const key = 'Kept content @ 47.0 | right clip 89.0';
    const problems: GateProblem[] = [
      { kind: 'fix', text: 'preserved-parts: uncut', fixture: 'preserved-parts', cause: 'clip' },
      { kind: 'fix', text: 'other: uncut', fixture: 'other', cause: 'gap' },
      { kind: 'record', text: 'preserved-parts: record the count' },
    ];
    const text = report(
      {
        uncutLines: [key],
        extraCutLines: ['Next @ 47.0 | right clip 66.0', 'Kept @ 61.0 | left clip unknown'],
      },
      problems
    );
    expect(text).toContain(
      '- Lines that the painted page cuts at an end that the reference clip holds whole: `Next @ 47.0 \\| right clip 66.0`. Fix the clip of the cell; an extra cut is a regression.'
    );
    expect(text).toContain(
      '- Lines that the reference clip holds whole at an end where the painted clip is unknown: no painted run with a one-rectangle clip that holds the baseline lies there, or runs with different clips start an item of the next cell there: `Kept @ 61.0 \\| left clip unknown`. Fix the clip of the cell; an unknown painted clip at that end is a regression.'
    );
    expect(text).toContain('| preserved-parts | fail (clip) |');
    expect(text).toContain(
      '- Reference lines that a cell clip cuts and that the painted page does not hold cut at that edge: `Kept content @ 47.0 \\| right clip 89.0`.'
    );
  });
});

describe('the report folder of a run', () => {
  test('is the folder that the variable names, or fidelity-report/ of the repository', () => {
    expect(reportDir({ [REPORT_DIR_ENV]: '/tmp/elsewhere' })).toBe('/tmp/elsewhere');
    expect(reportDir({})).toBe(join(import.meta.dir, '..', '..', 'fidelity-report'));
    expect(reportDir({ [REPORT_DIR_ENV]: '' })).toBe(reportDir({}));
    expect(reportDir({ [REPORT_DIR_ENV]: 'out' })).toBe(join(import.meta.dir, '..', '..', 'out'));
  });

  test('a helper tool names the variable in its remedy when the variable is set', () => {
    expect(runFirstText({})).toBe('run `bun run fidelity` first');
    expect(runFirstText({ [REPORT_DIR_ENV]: '/tmp/elsewhere' })).toBe(
      `run \`bun run fidelity\` with the same ${REPORT_DIR_ENV} first`
    );
    expect(runFirstText({}, 'again')).toBe('run `bun run fidelity` again');
    expect(runFirstText({ [REPORT_DIR_ENV]: 'out' }, 'again')).toBe(
      `run \`bun run fidelity\` with the same ${REPORT_DIR_ENV} again`
    );
  });
});

describe('the report path that the run prints', () => {
  test('is relative inside the current folder and absolute outside it', () => {
    const cwd = join(tmpdir(), 'xlsx-cwd');
    expect(reportPathText(cwd, join(cwd, 'fidelity-report'))).toBe(
      join('fidelity-report', 'report.md')
    );
    const outside = join(tmpdir(), 'xlsx-report-elsewhere');
    expect(reportPathText(cwd, outside)).toBe(join(outside, 'report.md'));
    expect(reportPathText(cwd, join(cwd, '..', 'xlsx-cwd-sibling'))).toBe(
      join(tmpdir(), 'xlsx-cwd-sibling', 'report.md')
    );
  });
});
