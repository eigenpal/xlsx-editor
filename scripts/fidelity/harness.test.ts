// The pure parts of the fidelity harness: rasters, PDF lines, line matching, and the known-gap
// gate. The printed lines of the painted page: `printed-lines.test.ts`. The ratchet against the
// merge base: `check-known-gaps.test.ts`. `bun run fidelity` runs the browser part.

import { describe, expect, test } from 'bun:test';
import { copyFileSync, mkdtempSync, readFileSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import {
  compareLines,
  differingPixelsIn,
  failureKinds,
  isOneCharacter,
  lineKey,
} from './compare.ts';
import { FIXTURES } from '../fixtures/catalog.ts';
import { workbookFamilies } from './fonts.ts';
import {
  checkGate,
  CLIPPED_RUN_REMEDY,
  EXTRA_CUT_REMEDY,
  exactList,
  gateRun,
  headroom,
  headroomText,
  worse,
  recordedList,
  sameTools,
  thresholdText,
  tightestList,
  PAINTED_SPLIT_REMEDY,
  UNCUT_REMEDY,
  UNKNOWN_GLYPHS_REMEDY,
  UNKNOWN_REFERENCE_REMEDY,
  withList,
  writeGaps,
  type FixtureResult,
} from './known-gaps.ts';
import { groupLines } from './pdf-lines.ts';
import { COVERAGE_GAPS, knownGapNote, writeReport } from './report.ts';
import { downsample, type PdfLine } from './poppler.ts';
import { FIXTURE_DIR } from './references.ts';
import * as thresholds from './thresholds.ts';
import type { KnownGapList } from './thresholds.ts';

const line = (text: string, xMin: number, baseline: number, xMax = xMin + 50): PdfLine => ({
  text,
  xMin,
  xMax,
  baseline,
  size: 11,
});

const tools = { poppler: 'poppler 1', chromium: 'chromium 1' };
const values = {
  pixels: 0,
  'line-breaks': 1,
  baseline: 0,
  'line-position': 0,
  'extra-text': 0,
};
const result = (patch: Partial<FixtureResult>): FixtureResult => ({
  name: 'basic-values',
  status: 'compared',
  failures: [],
  values,
  linesTotal: 3,
  linesMatched: 3,
  ...patch,
});
const list = (fixtures: KnownGapList['fixtures']): KnownGapList => ({
  ...tools,
  fixtureLines: { 'basic-values': 3 },
  fixtures,
});

describe('rasters and PDF lines', () => {
  test('a 2 x 2 box filter averages each block', () => {
    const data = new Uint8Array([
      0, 0, 0, 255, 200, 200, 200, 255, 100, 100, 100, 255, 100, 100, 100, 255,
    ]);
    const out = downsample({ width: 2, height: 2, data }, 2);
    expect([...out.data]).toEqual([100, 100, 100, 255]);
  });

  test('text items on one baseline group into a line, split at a wide gap', () => {
    const item = (text: string, x: number, width: number) => ({ text, x, y: 50, width, size: 11 });
    const lines = groupLines(
      [item('Kept', 39, 20), item('content', 62, 35), item('5', 134, 5)],
      612,
      792
    );
    expect(lines.map((l) => l.text)).toEqual(['Kept content', '5']);
    expect(lines[0]!.baseline).toBe(50);
  });
});

describe('line matching', () => {
  test('a reference line matches the painted line with the same text', () => {
    const score = compareLines([line('Kept content', 39.2, 50)], [line('Kept content', 39, 50)]);
    expect(score.matched).toBe(1);
    expect(score.worstShift).toBeCloseTo(0.2);
  });

  test('a line elsewhere still matches, and reports its shift', () => {
    const score = compareLines([line('Kept content', 80, 50)], [line('Kept content', 39, 50)]);
    expect(score.matched).toBe(1);
    expect(score.worstShift).toBeGreaterThan(0.5);
  });

  test('number lines are checked, with their signs; single glyphs are not, except a digit', () => {
    expect(lineKey('-3,5')).toBe('-3,5');
    expect(lineKey('#####')).toBe('#####');
    expect(lineKey('1 234,50')).toBe('1234,50');
    expect(lineKey('7')).toBe('7');
    expect(lineKey('A')).toBe('');
    // A line of one character other than a digit has no key, and the run counts it as a gap.
    expect([isOneCharacter('A'), isOneCharacter(' x '), isOneCharacter('7')]).toEqual([
      true,
      true,
      false,
    ]);
    expect([isOneCharacter('AB'), isOneCharacter('')]).toEqual([false, false]);
    expect(lineKey('%')).toBe('');
    expect(lineKey('• Bullet text')).toBe('Bullettext');
    const score = compareLines([line('3,5', 165, 61)], [line('-3,5', 165, 61)]);
    expect([score.total, score.matched, score.extra]).toEqual([1, 0, ['3,5']]);
  });

  test('reference lines that one line of ours holds side by side match it', () => {
    const score = compareLines(
      [line('-3,5 1234568', 165, 61, 236)],
      [line('-3,5', 165, 61, 182.6), line('1234568', 194, 61, 235.6)]
    );
    expect([score.total, score.matched, score.extra]).toEqual([2, 2, []]);
    // Without items, the group's width (70.6 points) is scored against the painted line's (71).
    expect(score.worstShift).toBeCloseTo(0.4, 9);
  });

  test('an unchecked reference line can complete the group', () => {
    const score = compareLines(
      [line('x 1,2E+11', 490, 538, 543)],
      [line('x', 490, 538, 495.6), line('1,2E+11', 507, 538, 543.6)]
    );
    expect([score.total, score.matched, score.extra]).toEqual([1, 1, []]);
  });

  test('reference pieces more than about one em apart do not join', () => {
    // A 30 point gap at 11 points is far wider than the one em at which `groupLines` splits.
    const score = compareLines(
      [line('-3,5 1234568', 165, 61, 236)],
      [line('-3,5', 165, 61, 182.6), line('1234568', 212.6, 61, 254.2)]
    );
    expect([score.total, score.matched]).toEqual([2, 0]);
    expect(score.extra).toEqual(['-3,5 1234568']);
  });

  test('each joined piece is scored against the item of ours with the same characters', () => {
    const ours: PdfLine = {
      ...line('-3,5 1234568', 165, 61, 236),
      items: [
        { at: 0, length: 4, xMin: 165, xMax: 183 },
        { at: 4, length: 7, xMin: 194, xMax: 236 },
      ],
    };
    const near = compareLines(
      [ours],
      [line('-3,5', 165, 61, 182.6), line('1234568', 193, 61, 235)]
    );
    expect(near.matched).toBe(2);
    expect(near.worstShift).toBe(1);
    // A piece that does not start on an painted item does not join it.
    const split = compareLines([ours], [line('-3,51', 165, 61, 188), line('234568', 194, 61, 236)]);
    expect(split.matched).toBe(0);
  });

  test('each step takes its pairs nearest pair first, in any order of the reference lines', () => {
    const both = (ours: PdfLine[], reference: PdfLine[]) =>
      [reference, [...reference].reverse()].map((r) => {
        const s = compareLines(ours, r);
        return [s.matched, s.worstBaseline, s.worstShift];
      });
    // The same text on another row: 120 and 119 pair first, so 100 pairs with 140.
    const total = (b: number) => line('Total', 0, b, 30);
    expect(both([total(119), total(140)], [total(100), total(120)])).toEqual([
      [2, 40, 0],
      [2, 40, 0],
    ]);
    // A bullet that one PDF writes as a letter: the reference line at 20 and the painted line at
    // 19 pair first.
    const bullet = (x: number): PdfLine => ({ ...line('q Total', x, 100, x + 30), marker: true });
    const plain = (x: number) => line('Total', x, 100, x + 30);
    expect(both([plain(19), plain(40)], [bullet(0), bullet(20)])).toEqual([
      [2, 0, 40],
      [2, 0, 40],
    ]);
    // Painted pieces that join to hold a reference line: the pieces from 19 pair with the
    // reference line at 20 first.
    const pieces = [19, 40].flatMap((x) => [
      line('AB', x, 100, x + 10),
      line('CD', x + 11, 100, x + 21),
    ]);
    expect(both(pieces, [line('ABCD', 0, 100, 25), line('ABCD', 20, 100, 45)])).toEqual([
      [2, 0, 36],
      [2, 0, 36],
    ]);
  });

  test('a tie of the reference line falls to the painted baseline', () => {
    // `Total` at 100 lies 0.3 points from painted lines at 99.7 and 100.3, and `Total` at 100.6
    // only from the one at 100.3. The one at 99.7 pairs with 100 first, in any painted order.
    const total = (b: number) => line('Total', 10, b, 40);
    for (const ours of [
      [total(100.3), total(99.7)],
      [total(99.7), total(100.3)],
    ]) {
      const s = compareLines(ours, [total(100), total(100.6)]);
      expect(s.matched).toBe(2);
      expect(s.worstBaseline).toBeCloseTo(0.3, 6);
    }
  });

  test('a tie of reference lines falls to the painted lines, in any painted order', () => {
    const total = (x: number) => line('Total', x, 100, x + 20);
    for (const ours of [
      [total(5), total(15)],
      [total(15), total(5)],
    ])
      expect(compareLines(ours, [total(10), total(20)]).worstShift).toBe(5);
  });
});

describe('differing pixels in glyph boxes', () => {
  test('counts the red pixels of a diff inside and outside the boxes', () => {
    // A 4 x 1 diff: red, gray, red, red.
    const data = [255, 0, 0, 255, 128, 128, 128, 255, 255, 0, 0, 255, 255, 0, 0, 255];
    const diff = { width: 4, height: 1, data };
    expect(differingPixelsIn(diff, [[0, 0, 2, 1]])).toEqual({ inside: 1, outside: 2, covered: 2 });
  });
});

describe('the known-gap gate', () => {
  test('the headroom of a known gap is the rest of its tolerance, as the value of the check', () => {
    const t = thresholds.GAP_TOLERANCE.pixels;
    // A value at its record keeps the whole tolerance and half the last digit of the record.
    const noise = 5e-6;
    expect(headroom('pixels', 0.00861, 0.00861)).toBeCloseTo(t + noise, 9);
    expect(headroom('pixels', 0.009, 0.00861)).toBeCloseTo(t + noise - 0.00039, 9);
    // A line-break share fails when it falls, so its headroom is the fall that it can take.
    const tl = thresholds.GAP_TOLERANCE['line-breaks'];
    expect(headroom('line-breaks', 0.9, 0.9)).toBeCloseTo(tl + 5e-4, 9);
    expect(headroomText('pixels', 0.00061)).toBe('0.061 percentage points');
    expect(headroomText('line-breaks', 0.02)).toBe('2.0 percentage points');
    expect(headroomText('extra-text', 1.0005)).toBe('1 line');
    expect(headroomText('extra-text', 0.0005)).toBe('0 lines');
    expect(headroomText('baseline', 0.5)).toBe('0.50 pt');
    expect(headroomText('pixels', 0.00039, true)).toBe('0.039 percentage points');
  });

  test('a value has no headroom left exactly when the gate fails it, for every check', () => {
    const mismatches: string[] = [];
    for (const kind of thresholds.FAILURE_KINDS) {
      const recorded = kind === 'line-breaks' ? 0.9 : kind === 'pixels' ? 0.008 : 2;
      const step = kind === 'line-breaks' || kind === 'pixels' ? 1e-5 : 1e-3;
      for (let i = -2000; i <= 2000; i += 1) {
        const value = recorded + i * step;
        if (headroom(kind, value, recorded) < 0 !== worse(kind, value, recorded))
          mismatches.push(`${kind} ${value}`);
      }
    }
    expect(mismatches).toEqual([]);
  });

  test('a pass with no gap has no problem', () => {
    expect(checkGate([result({})], list({}), tools, true)).toEqual([]);
  });

  test('a new failure needs a fix', () => {
    const r = result({ failures: ['pixels'], values: { ...values, pixels: 0.01 } });
    expect(checkGate([r], list({}), tools, true).map((p) => p.kind)).toEqual(['fix']);
  });

  test('a score worse than its gap needs a fix; a better one is recorded', () => {
    const worse = result({ failures: ['pixels'], values: { ...values, pixels: 0.02 } });
    const better = result({ failures: ['pixels'], values: { ...values, pixels: 0.006 } });
    const gaps = list({ 'basic-values': { pixels: 0.01 } });
    expect(checkGate([worse], gaps, tools, true).map((p) => p.kind)).toEqual(['fix']);
    expect(checkGate([better], gaps, tools, true).map((p) => p.kind)).toEqual(['record']);
  });

  test('a listed check that passes now is recorded', () => {
    const gaps = list({ 'basic-values': { baseline: 1 } });
    expect(checkGate([result({})], gaps, tools, true)[0]!.text).toContain('remove the gap');
  });

  test('a missing reference fails, and fewer lines fail', () => {
    const missing = result({ status: 'error', error: 'no reference' });
    expect(checkGate([missing], list({}), tools, true).map((p) => p.kind)).toEqual(['fix']);
    const fewer = result({ linesTotal: 2, linesMatched: 2 });
    expect(checkGate([fewer], list({}), tools, true).map((p) => p.kind)).toEqual(['fix']);
  });

  test('other tools are a note, not a failure', () => {
    const problems = checkGate([result({})], list({}), { ...tools, poppler: 'poppler 2' }, true);
    expect(problems.map((p) => p.kind)).toEqual(['note']);
  });

  test('with other tools, a better score is a note and a worse one still fails', () => {
    const other = { ...tools, poppler: 'poppler 2' };
    const gaps = list({ 'basic-values': { pixels: 0.01 } });
    const better = result({ failures: ['pixels'], values: { ...values, pixels: 0.006 } });
    const passes = result({});
    const worse = result({ failures: ['pixels'], values: { ...values, pixels: 0.02 } });
    expect(checkGate([better], gaps, other, true).map((p) => p.kind)).toEqual(['note', 'note']);
    expect(checkGate([passes], gaps, other, true).map((p) => p.kind)).toEqual(['note', 'note']);
    expect(checkGate([worse], gaps, other, true).map((p) => p.kind)).toEqual(['note', 'fix']);
  });

  test('with other tools, more reference lines are a note; fewer still fail', () => {
    const other = { ...tools, poppler: 'poppler 2' };
    const more = result({ linesTotal: 4, linesMatched: 4 });
    const fewer = result({ linesTotal: 2, linesMatched: 2 });
    expect(checkGate([more], list({}), other, true).map((p) => p.kind)).toEqual(['note', 'note']);
    expect(checkGate([fewer], list({}), other, true).map((p) => p.kind)).toEqual(['note', 'fix']);
    expect(checkGate([more], list({}), tools, true).map((p) => p.kind)).toEqual(['record']);
  });

  test('a recorded list holds the failing values and the line counts', () => {
    const r = result({ failures: ['baseline'], values: { ...values, baseline: 0.98642 } });
    expect(recordedList([r], tools)).toEqual({
      ...list({ 'basic-values': { baseline: 0.986 } }),
      fixtureClipped: {},
    });
    expect(
      recordedList([result({ clippedLines: ['Kept content', 'A line'] })], tools).fixtureClipped
    ).toEqual({ 'basic-values': ['A line', 'Kept content'] });
  });

  test('a clipped line that the list does not name fails; a named one that is gone asks for a record', () => {
    const named = (...lines: string[]): KnownGapList => ({
      ...list({}),
      fixtureClipped: { 'basic-values': lines },
    });
    // A list without the record allows none.
    expect(checkGate([result({ clippedLines: ['Kept content'] })], list({}), tools, true)).toEqual([
      {
        kind: 'fix',
        text: `basic-values: a cell clip cuts the line \`Kept content\` 1 time, and the list names it 0 times; ${CLIPPED_RUN_REMEDY}`,
        fixture: 'basic-values',
        cause: 'clip',
      },
    ]);
    // A run offers the owner route only for a new fixture or a new reference PDF, and names
    // its acceptance.
    expect(CLIPPED_RUN_REMEDY).toBe(
      'fix the layout or the fixture; for a new fixture or a new reference PDF only, an owner may name the line in `fixtureClipped` of each list and run `bun run fidelity:check-gaps` with `XLSX_FIDELITY_ACCEPT_CLIPPED=<fixture>`'
    );
    expect(
      checkGate([result({ clippedLines: ['Kept content'] })], named('Kept content'), tools, true)
    ).toEqual([]);
    // Another line fails, even when the count stays the same.
    expect(
      checkGate([result({ clippedLines: ['Other'] })], named('Kept content'), tools, true)[0]!.kind
    ).toBe('fix');
    // Lines compare as a multiset: a line that a clip cuts twice needs two entries.
    const twice = result({ clippedLines: ['Kept content', 'Kept content'] });
    expect(checkGate([twice], named('Kept content'), tools, true)[0]!.kind).toBe('fix');
    expect(checkGate([twice], named('Kept content', 'Kept content'), tools, true)).toEqual([]);
    expect(checkGate([result({})], named('Kept content'), tools, true)).toEqual([
      {
        kind: 'record',
        text: 'basic-values: a cell clip cuts the line `Kept content` 0 times now, and the list names it 1 time; record the list again',
      },
    ]);
    // A named line that the reference PDF still clips, and that the painted page no longer holds
    // cut, is a regression to fix: the gate does not ask to record it away.
    const lost = result({ uncutLines: ['Kept content'] });
    expect(checkGate([lost], named('Kept content'), tools, true)).toEqual([
      {
        kind: 'fix',
        text: `basic-values: the reference PDF clips the line \`Kept content\`, and the painted page does not hold it cut at that edge (1 time); ${UNCUT_REMEDY}`,
        fixture: 'basic-values',
        cause: 'clip',
      },
    ]);
    expect(checkGate([lost], list({}), tools, true).map((p) => p.kind)).toEqual(['fix']);
    // When the painted line is gone, the line says so. The cause comes from `compareLines`
    // (`lostUncutLines`): a missing line of the same text does not change it.
    const gone = result({
      uncutLines: ['Kept content @ 47.0 | right clip 89.0'],
      lostUncutLines: ['Kept content @ 47.0 | right clip 89.0'],
    });
    expect(checkGate([gone], list({}), tools, true)[0]!.text).toBe(
      `basic-values: the reference PDF clips the line \`Kept content @ 47.0 | right clip 89.0\`, and the painted page lost the line or cut it elsewhere (1 time); ${UNCUT_REMEDY}`
    );
    const kept = result({
      uncutLines: ['Kept content @ 47.0 | right clip 89.0'],
      missingLines: ['Kept content'],
    });
    expect(checkGate([kept], list({}), tools, true)[0]!.text).toContain(
      'the painted page does not hold it cut at that edge'
    );
    // An error of a fixture names the fixture once.
    const error: FixtureResult = { name: 'basic-values', status: 'error', error: 'no reference' };
    expect(checkGate([error], list({}), tools, true)[0]!.text).toBe('basic-values: no reference');
    // The converse, an extra painted cut, is a regression too, also when the list names it.
    const extra = result({ extraCutLines: ['Kept content @ 47.0 | right clip 66.0'] });
    expect(
      checkGate([extra], named('Kept content @ 47.0 | right clip 66.0'), tools, true).filter(
        (p) => p.kind === 'fix'
      )
    ).toEqual([
      {
        kind: 'fix',
        text: `basic-values: the painted page cuts the line \`Kept content @ 47.0 | right clip 66.0\` at an end that the reference clip holds whole (1 time); ${EXTRA_CUT_REMEDY}`,
        fixture: 'basic-values',
        cause: 'clip',
      },
    ]);
    // An end whose painted clip is unknown says so: no painted edge cuts the line there.
    const unknown = result({ extraCutLines: ['Kept content @ 47.0 | left clip unknown'] });
    expect(checkGate([unknown], list({}), tools, true)).toEqual([
      {
        kind: 'fix',
        text: `basic-values: the reference clip holds the line \`Kept content @ 47.0 | left clip unknown\` whole at an end, and the painted clip there is unknown: no painted run with a one-rectangle clip that holds the baseline lies at that end, or runs with different clips start an item of the next cell there (1 time); ${EXTRA_CUT_REMEDY}`,
        fixture: 'basic-values',
        cause: 'clip',
      },
    ]);
    // A line end with an unknown reference clip is a skipped check: it fails as a coverage gap.
    const skipped = result({
      unknownReferenceEnds: ['Kept @ 47.0 | right reference clip unknown'],
    });
    expect(checkGate([skipped], list({}), tools, true)).toEqual([
      {
        kind: 'fix',
        text: `basic-values: the reference PDF has no known clip at \`Kept @ 47.0 | right reference clip unknown\` (1 time), so no clip rule checks that end; ${UNKNOWN_REFERENCE_REMEDY}`,
        fixture: 'basic-values',
        cause: 'clip',
      },
    ]);
    // An inner end of a painted piece that splits the text of one reference item names the paint
    // as the cause, not the reference clip.
    const split = result({
      unknownReferenceEnds: ['1 @ 47.0 | right painted split inside a reference item'],
    });
    expect(checkGate([split], list({}), tools, true)).toEqual([
      {
        kind: 'fix',
        text: `basic-values: the painted page splits the text of one reference item at \`1 @ 47.0 | right painted split inside a reference item\` (1 time), so no clip rule checks that end; ${PAINTED_SPLIT_REMEDY}`,
        fixture: 'basic-values',
        cause: 'clip',
      },
    ]);
    // A painted split of a named clipped line is no unknown reference clip of that line: the list
    // is out of date, not the reference clip.
    const splitNamed = result({
      unknownReferenceEnds: ['Kept content @ 47.0 | right painted split inside a reference item'],
    });
    expect(
      checkGate([splitNamed], named('Kept content @ 47.0 | right clip 89.0'), tools, true).map(
        (p) => p.text
      )
    ).toEqual([
      `basic-values: the painted page splits the text of one reference item at \`Kept content @ 47.0 | right painted split inside a reference item\` (1 time), so no clip rule checks that end; ${PAINTED_SPLIT_REMEDY}`,
      'basic-values: a cell clip cuts the line `Kept content @ 47.0 | right clip 89.0` 0 times now, and the list names it 1 time; record the list again',
    ]);
    // Reference glyphs that do not account for the text, such as a ligature, skip the glyph rule:
    // the end fails with the same remedy.
    const glyphless = result({
      unknownReferenceEnds: ['Kept @ 47.0 | right reference glyphs unknown'],
    });
    // A named clipped line with unknown reference glyphs gets only the fix advice: a new record
    // would hide the cause.
    const namedGlyphless = named('Kept content @ 47.0 | right clip 89.0');
    const glyphsNow = result({
      unknownReferenceEnds: ['Kept content @ 47.0 | right reference glyphs unknown'],
    });
    expect(
      checkGate([glyphsNow], namedGlyphless, tools, true).map((p) => [p.kind, p.text])
    ).toEqual([
      [
        'fix',
        `basic-values: the reference PDF has glyphs that do not account for the text at \`Kept content @ 47.0 | right reference glyphs unknown\` (1 time), so no glyph rule checks that end; ${UNKNOWN_GLYPHS_REMEDY}`,
      ],
      [
        'fix',
        `basic-values: the list names the clipped line \`Kept content @ 47.0 | right clip 89.0\` 1 time, and the reference glyphs of that line are unknown now, so no glyph rule checks it; a new record would hide it; ${UNKNOWN_GLYPHS_REMEDY}`,
      ],
    ]);
    expect(checkGate([glyphless], list({}), tools, true).map((p) => p.text)).toEqual([
      `basic-values: the reference PDF has glyphs that do not account for the text at \`Kept @ 47.0 | right reference glyphs unknown\` (1 time), so no glyph rule checks that end; ${UNKNOWN_GLYPHS_REMEDY}`,
    ]);
    // A named clipped line whose reference clip is unknown now names that cause, not a new record.
    const lostClip = result({
      unknownReferenceEnds: ['Kept content @ 47.0 | right reference clip unknown'],
    });
    const named1 = named('Kept content @ 47.0 | right clip 89.0');
    expect(checkGate([lostClip], named1, tools, true).map((p) => [p.kind, p.text])).toEqual([
      [
        'fix',
        `basic-values: the reference PDF has no known clip at \`Kept content @ 47.0 | right reference clip unknown\` (1 time), so no clip rule checks that end; ${UNKNOWN_REFERENCE_REMEDY}`,
      ],
      [
        'fix',
        `basic-values: the list names the clipped line \`Kept content @ 47.0 | right clip 89.0\` 1 time, and the reference clip of that line is unknown now, so no rule matches it; a new record would hide it; ${UNKNOWN_REFERENCE_REMEDY}`,
      ],
    ]);
    // Across toolchains, each line takes the fewest times that a list names it.
    const two = tightestList([
      named('A', 'B', 'B'),
      { ...named('B', 'B', 'C'), poppler: 'poppler 2' },
    ]);
    expect(two!.fixtureClipped).toEqual({ 'basic-values': ['B', 'B'] });
    expect(tightestList([list({}), { ...list({}), poppler: 'poppler 2' }])!.fixtureClipped).toBe(
      undefined
    );
  });
});

describe('the known-gap list of a toolchain', () => {
  const at = (poppler: string, chromium: string, pixels: number): KnownGapList => ({
    poppler,
    chromium,
    fixtureLines: { 'basic-values': 3 },
    fixtures: { 'basic-values': { pixels } },
  });
  const local = at('poppler 26.10.0, pdf.js 5.4.624', 'chromium 158.0.8127.18', 0.00861);
  const ci = at('poppler 24.02.0, pdf.js 5.4.624', 'chromium 158.0.8127.18', 0.00677);

  test('a run uses the list of exactly its tools', () => {
    expect(exactList([local, ci], ci)).toBe(ci);
    expect(exactList([ci, local], local)).toBe(local);
    expect(exactList([local, ci], { ...ci, poppler: 'poppler 24.03.0, pdf.js 5.4.624' })).toBe(
      undefined
    );
    expect(exactList([], ci)).toBeUndefined();
  });

  test('an empty or digit-free tool text matches no list', () => {
    const blank = { ...local, poppler: '' };
    expect(sameTools(blank, blank)).toBe(false);
    expect(sameTools(local, { ...local, chromium: 'chromium' })).toBe(false);
    expect(exactList([blank, local], { ...local, poppler: '' })).toBeUndefined();
  });

  test('the tightest list takes the tightest value of each check; a pass is the tightest', () => {
    const a: KnownGapList = {
      ...local,
      fixtureLines: { 'basic-values': 3, 'dates-1904': 6 },
      fixtures: {
        'basic-values': { pixels: 0.009, 'line-breaks': 0.9 },
        'dates-1904': { baseline: 1 },
      },
    };
    const b: KnownGapList = {
      ...ci,
      fixtureLines: { 'basic-values': 4, 'dates-1904': 6 },
      fixtures: { 'basic-values': { pixels: 0.007, 'line-breaks': 0.8, baseline: 1 } },
    };
    const t = tightestList([a, b])!;
    expect(t.fixtures).toEqual({ 'basic-values': { pixels: 0.007, 'line-breaks': 0.9 } });
    expect(t.fixtureLines).toEqual({ 'basic-values': 4, 'dates-1904': 6 });
    expect(t.poppler).toContain(local.poppler);
    expect(t.poppler).toContain(ci.poppler);
    expect(tightestList([a])).toBe(a);
    expect(tightestList([])).toBeUndefined();
  });

  const run = { poppler: 'poppler 24.08.0, pdf.js 5.4.624', chromium: 'chromium 158.0.8127.18' };
  const gate = (r: FixtureResult, tools = run, inCi = false) =>
    gateRun({
      results: [r],
      env: 'substitute',
      tools,
      lists: [local, ci],
      complete: true,
      ci: inCi,
    });

  test('outside CI, a run with other tools compares with the tightest values', () => {
    expect(gate(result({})).list?.fixtures['basic-values']).toEqual({ pixels: 0.00677 });
    const better = result({ failures: ['pixels'], values: { ...values, pixels: 0.005 } });
    // Worse than the CI list, although better than the local list: a fix.
    const worse = result({ failures: ['pixels'], values: { ...values, pixels: 0.0085 } });
    const unlisted = result({ failures: ['baseline'], values: { ...values, baseline: 1 } });
    const missing = result({ status: 'error', error: 'no reference' });
    const fewer = result({ linesTotal: 2, linesMatched: 2 });
    const kinds = (r: FixtureResult) => gate(r).problems.map((p) => p.kind);
    expect(gate(better).problems[0]!.text).toContain('poppler 24.02.0');
    expect(kinds(better)).toEqual(['note', 'note']);
    expect(kinds(result({}))).toEqual(['note', 'note']);
    expect(kinds(worse)).toEqual(['note', 'fix']);
    expect(kinds(unlisted)).toEqual(['note', 'note', 'fix']);
    expect(kinds(missing)).toEqual(['note', 'fix']);
    expect(kinds(fewer)).toEqual(['note', 'note', 'fix']);
  });

  test('in CI, a run with no list of exactly its tools needs a fix', () => {
    const problems = gate(result({}), run, true).problems;
    expect(problems[0]!.kind).toBe('fix');
    expect(problems[0]!.text).toContain('no measured list for this toolchain');
    // With its own list, a pass of a listed check asks for a new record, never a note.
    expect(gate(result({}), ci, true).problems.map((p) => p.kind)).toEqual(['record']);
    // An empty tool text is as far as possible from every list: never its list.
    const empty = gate(result({}), { poppler: '', chromium: '' }, true).problems;
    expect(empty[0]!.text).toContain('no measured list for this toolchain');
  });

  test('a run with no list in its environment fails as unmeasured', () => {
    const problems = gateRun({
      results: [result({})],
      env: 'reference',
      tools: run,
      lists: [],
      complete: true,
      ci: false,
    }).problems;
    expect(problems.map((p) => p.kind)).toEqual(['fix', 'record']);
    expect(problems[0]!.text).toContain('no measured known-gap list');
  });

  test('--update-gaps checks the run against the list that it wrote', () => {
    const failing = result({ failures: ['pixels'], values: { ...values, pixels: 0.02 } });
    const written = recordedList([failing], run);
    const out = gateRun({
      results: [failing],
      env: 'substitute',
      tools: run,
      lists: [local, ci],
      complete: true,
      ci: true,
      recorded: written,
    });
    expect(written.poppler).toBe(run.poppler);
    expect(written.chromium).toBe(run.chromium);
    expect(written.fixtures).toEqual({ 'basic-values': { pixels: 0.02 } });
    expect(out.list).toBe(written);
    expect(out.problems).toEqual([]);
  });

  test('--update-gaps replaces the list of its tools, or adds one', () => {
    const remeasured = at(ci.poppler, ci.chromium, 0.006);
    expect(withList([local, ci], remeasured)).toEqual([local, remeasured]);
    const other = at('poppler 25.01.0, pdf.js 5.4.624', ci.chromium, 0.007);
    expect(withList([local, ci], other)).toEqual([local, ci, other]);
    expect(withList([], other)).toEqual([other]);
  });

  test('--update-gaps rewrites only the list of the run in thresholds.ts', async () => {
    // A new directory for each run: the dynamic import below caches modules by path, so a
    // reused path would read the thresholds.ts of an earlier run.
    const dir = mkdtempSync(join(tmpdir(), 'xlsx-gaps-test-'));
    try {
      const file = join(dir, 'thresholds.ts');
      copyFileSync(join(import.meta.dir, 'thresholds.ts'), file);
      const before = thresholds.KNOWN_GAPS.substitute ?? [];
      const remeasured = { ...before[1]!, fixtures: {} };
      writeGaps(file, 'substitute', remeasured);
      const after = ((await import(file)) as typeof thresholds).KNOWN_GAPS;
      expect(after.substitute).toEqual([before[0]!, remeasured]);
      expect(after.reference).toEqual(thresholds.KNOWN_GAPS.reference);
      // A format that fails says so, in one line.
      expect(() =>
        writeGaps(file, 'substitute', remeasured, () => ({ status: 2, error: 'SyntaxError' }))
      ).toThrow(`prettier did not format ${file}: SyntaxError; run bun run format`);
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  });

  test('the substitute environment has a list for the CI tools and for the local tools', () => {
    // The Chromium of the pinned Playwright: a Playwright bump measures and records new lists.
    const lists = thresholds.KNOWN_GAPS.substitute ?? [];
    const run = { poppler: 'poppler 24.02.0, pdf.js 6.5.438', chromium: 'chromium 159.0.8154.6' };
    const local = { ...run, poppler: 'poppler 26.10.0, pdf.js 6.5.438' };
    for (const tools of [run, local]) {
      const list = exactList(lists, tools);
      expect(list).toBeDefined();
      expect(list?.fixtures['number-formats']).toBeUndefined();
    }
  });
});

test('the font families of a fixture come from its styles and theme', () => {
  const families = workbookFamilies(
    new Uint8Array(readFileSync(join(FIXTURE_DIR, 'basic-values.xlsx')))
  );
  expect(families.has('carlito')).toBe(true);
});

describe('coverage gaps', () => {
  test('the public list names every coverage gap that the harness prints', () => {
    const doc = readFileSync(
      join(import.meta.dir, '../../docs/site/content/known-approximations.mdx'),
      'utf8'
    );
    const start = doc.indexOf('\n## Coverage gaps\n');
    expect(start).toBeGreaterThanOrEqual(0);
    const end = doc.indexOf('\n## ', start + 1);
    const section = doc.slice(start, end < 0 ? undefined : end);
    const listed = [...section.matchAll(/^- (.+)$/gm)].map((m) => m[1]);
    expect(listed).toEqual([...COVERAGE_GAPS]);
  });
});

describe('the pixel limit of each fixture', () => {
  // A page with no line differences, so that only the pixel score decides.
  const clean = compareLines([], []);
  test('fills-patterns fails above its own limit, where another fixture passes', () => {
    expect(failureKinds({ mismatch: 0.0003 }, clean, 'fills-patterns')).toEqual(['pixels']);
    expect(failureKinds({ mismatch: 0.0003 }, clean, 'basic-values')).toEqual([]);
    expect(thresholdText('pixels', 'fills-patterns')).toBe('≤ 0.02%');
    expect(thresholdText('pixels', 'basic-values')).toBe('≤ 0.5%');
  });
  test('each limit names a catalog fixture and is tighter than the page limit', () => {
    const names = new Set(FIXTURES.map((f) => f.name));
    for (const [name, limit] of Object.entries(thresholds.FIXTURE_PIXEL_MISMATCH)) {
      expect(names.has(name)).toBe(true);
      expect(limit).toBeLessThan(thresholds.THRESHOLDS.maxPixelMismatch);
    }
  });
});

describe('the report line of a known gap', () => {
  test('a value inside the tolerance gives the room that is left', () => {
    expect(knownGapNote('pixels', 0.0072, 0.007)).toBe(
      'Known gap, pixels: measured 0.720%, recorded 0.700%, 0.020 percentage points worse than the record; it can grow 0.080 percentage points more before the gate fails'
    );
  });

  test('a value past the tolerance says that the gate fails, with no negative room', () => {
    const note = knownGapNote('pixels', 0.0085, 0.007);
    expect(note).toContain('0.150 percentage points worse than the record');
    expect(note).toContain('past the tolerance, so the gate fails');
    expect(note).not.toContain('-');
  });

  test('a line-breaks score drops as it gets worse', () => {
    const note = knownGapNote('line-breaks', 0.9, 0.9);
    expect(note).toContain('it can drop');
    expect(note).not.toContain('worse than the record');
  });

  test('the report writes the line for a failing known gap', () => {
    const dir = mkdtempSync(join(tmpdir(), 'fidelity-report-'));
    try {
      const tools = { poppler: 'poppler 1', chromium: 'chromium 1' };
      writeReport(
        dir,
        [
          {
            name: 'number-formats',
            status: 'compared',
            failures: ['pixels'],
            values: {
              pixels: 0.0072,
              'line-breaks': 1,
              baseline: 0,
              'line-position': 0,
              'extra-text': 0,
            },
            linesTotal: 1,
            linesMatched: 1,
            missingLines: [],
            extraLines: [],
          },
        ],
        {
          env: 'substitute',
          fonts: 'substitutes',
          tools,
          gaps: { ...tools, fixtureLines: {}, fixtures: { 'number-formats': { pixels: 0.007 } } },
          problems: [],
        }
      );
      const report = readFileSync(join(dir, 'report.md'), 'utf8');
      expect(report).toContain(`- ${knownGapNote('pixels', 0.0072, 0.007)}`);
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  });
});
