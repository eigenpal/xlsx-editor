// The painted side of the line scores of the fidelity harness: the printed PDF of the painted page,
// the clipped lines, the font names of both PDFs, and the print sequence of the harness page. The
// clip rectangles of a PDF: `pdf-clips.test.ts`. `bun run fidelity` runs the browser part.

import { describe, expect, test } from 'bun:test';
import { copyFileSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join } from 'node:path';
import { compareTracked } from './test-glyphs.ts';
import {
  baseFontName,
  familyName,
  FOREIGN_REFERENCE_REMEDY,
  foreignReferenceFonts,
  macStyle,
  postScriptName,
  referenceFamily,
} from './font-names.ts';
import { postScriptFamilies } from './fonts.ts';
import type { TextRun } from './pdf-clips.ts';
import { groupLines, pdfItems, printedPageProblem, type PdfItems } from './pdf-lines.ts';
import type { PdfLine } from './poppler.ts';
import { printPainted, type PrintablePage } from './print.ts';
import { FIXTURE_DIR } from './references.ts';

const line = (text: string, xMin: number, baseline: number, xMax = xMin + 50): PdfLine => ({
  text,
  xMin,
  xMax,
  baseline,
  size: 11,
});

/** A text run of a PDF at baseline 47, from `x0` to `x1`, held by `clip`. */
const run = (x0: number, x1: number, clip: TextRun['clip'], baseline = 47): TextRun => ({
  x0,
  x1,
  baseline,
  clip,
});

describe('lines that a cell clip cuts', () => {
  // The reference PDF keeps every glyph of `Kept content` and clips it at 89 points; the printed
  // PDF has no `t`, which the painted clip, also ending at 89 points, hides entirely. The `n`
  // crosses the edge, so the painted line ends past it, at 93.8 points. The two PDFs write the clip
  // of the cell one point apart at the top and on the left, as the fixture `preserved-parts` does.
  const reference = [line('Kept content', 39, 47, 97.7)];
  const ours = [line('Kept conten', 39, 47, 93.8)];
  const refRuns = [run(39, 94.6, [37, 37, 89, 50])];
  const ourRuns = (edge: number, end = 93.8) => [run(39, end, [36, 36, edge, 50])];
  const clips = (edge: number, end?: number) => ({
    ours: ourRuns(edge, end),
    reference: refRuns,
  });
  const key = 'Kept content @ 47.0 | right clip 89.0';

  test('the painted line that ends at the painted clip edge matches the reference line that it starts', () => {
    const score = compareTracked(ours, reference, clips(89));
    expect([score.total, score.matched, score.clipped, score.extra]).toEqual([1, 1, [key], []]);
    // The uncut edge is scored: the left edge here.
    expect(score.worstShift).toBe(0);
  });

  test('without the clips of both PDFs, the shorter line is missing and extra', () => {
    const none = compareTracked(ours, reference);
    expect([none.matched, none.clipped, none.missing, none.extra]).toEqual([
      0,
      [],
      ['Kept content'],
      ['Kept conten'],
    ]);
    // A missing clip on either side is no evidence.
    expect(compareTracked(ours, reference, { ours: ourRuns(89), reference: [] }).matched).toBe(0);
    expect(compareTracked(ours, reference, { ours: [], reference: refRuns }).matched).toBe(0);
    const unclipped = { ours: [run(39, 93.8, 'none')], reference: refRuns };
    expect(compareTracked(ours, reference, unclipped).matched).toBe(0);
    // A clip that is not one rectangle is no evidence.
    const path = { ours: [run(39, 93.8, 'non-rectangular')], reference: refRuns };
    expect(compareTracked(ours, reference, path).matched).toBe(0);
    // A reference clip that is not one rectangle leaves the line missing, and names its ends.
    const refPath = { ours: ourRuns(89), reference: [run(39, 94.6, 'non-rectangular')] };
    const unknown = compareTracked(ours, reference, refPath);
    expect([unknown.matched, unknown.unknownReference]).toEqual([
      0,
      [
        'Kept content @ 47.0 | right reference clip unknown',
        'Kept content @ 47.0 | left reference clip unknown',
      ],
    ]);
    // A clip that does not hold the baseline is no evidence.
    const above = { ours: [run(39, 93.8, [36, 10, 89, 30])], reference: refRuns };
    expect(compareTracked(ours, reference, above).matched).toBe(0);
    // A reference line that ends inside its clip is not cut by it.
    const short = [line('Kept content', 39, 47, 89.3)];
    expect(compareTracked(ours, short, clips(89)).matched).toBe(0);
  });

  test('the painted clip edge may move within the line position threshold, with the same name', () => {
    // A layout fix that moves the painted clip 0.1 points keeps the name of the line: it names the
    // reference clip edge.
    expect(compareTracked(ours, reference, clips(89.1)).clipped).toEqual([key]);
    expect(compareTracked(ours, reference, clips(88.6)).clipped).toEqual([key]);
  });

  test('a clip of ours 2 points wider or narrower than the reference clip fails', () => {
    const wide = compareTracked([line('Kept content', 39, 47, 97.7)], reference, clips(91, 97.7));
    expect(wide.clipped).toEqual([]);
    const early = compareTracked([line('Kept con', 39, 47, 87.5)], reference, clips(87, 87.5));
    expect([early.matched, early.missing]).toEqual([0, ['Kept content']]);
  });

  test('the painted line that drops a glyph that the clip shows fails', () => {
    // `Kept conte` ends at 88.3 points, but the painted clip shows the `n` up to 89 points.
    const dropped = compareTracked([line('Kept conte', 39, 47, 88.3)], reference, clips(89, 88.3));
    expect([dropped.matched, dropped.missing]).toEqual([0, ['Kept content']]);
  });

  test('each clipped line is named by its own baseline', () => {
    // The same cell text in two rows, at baselines 47 and 62.
    const below = (l: PdfLine): PdfLine => ({ ...l, baseline: 62 });
    const twice = compareTracked(
      [...ours, ...ours.map(below)],
      [...reference, ...reference.map(below)],
      {
        ours: [...ourRuns(89), run(39, 93.8, [36, 51, 89, 65], 62)],
        reference: [...refRuns, run(39, 94.6, [37, 52, 89, 65], 62)],
      }
    );
    expect(twice.clipped).toEqual([key, 'Kept content @ 62.0 | right clip 89.0']);
  });

  test('the painted line that a clip cuts on the left matches the reference line that it ends', () => {
    const score = compareTracked(
      [line('content', 59.6, 47, 97.7)],
      [line('Kept content', 39, 47, 97.6)],
      {
        ours: [run(59.6, 97.7, [60, 36, 120, 50])],
        reference: [run(39, 97.6, [60, 37, 120, 50])],
      }
    );
    expect([score.matched, score.clipped]).toEqual([1, ['Kept content @ 47.0 | left clip 60.0']]);
    expect(score.worstShift).toBeCloseTo(0.1, 9);
    // The painted line that starts 3 points inside the left edge dropped a glyph that the clip
    // shows.
    const late = compareTracked(
      [line('ontent', 63, 47, 97.7)],
      [line('Kept content', 39, 47, 97.6)],
      {
        ours: [run(63, 97.7, [60, 36, 120, 50])],
        reference: [run(39, 97.6, [60, 37, 120, 50])],
      }
    );
    expect(late.matched).toBe(0);
  });

  test('the painted text must be the start or the end of the reference text', () => {
    const typo = [line('Kept cantent', 39, 47, 93.8)];
    expect(compareTracked(typo, reference, clips(89)).matched).toBe(0);
  });

  test('the text of the next cell stays its own line, also where the overflow overlaps it', () => {
    // The reference text runs past its clip at 89 points under `Next`, the text of the next cell
    // with its own clip. `groupLines` starts a new line where the clip of an item differs.
    const cellA = [37, 37, 89, 50] as const;
    const cellB = [89, 37, 160, 50] as const;
    const item = (text: string, x: number, width: number, clip?: TextRun['clip']) => ({
      text,
      x,
      y: 47,
      width,
      size: 11,
      ...(clip ? { clip } : {}),
    });
    for (const nx of [91, 134]) {
      const ref = groupLines(
        [item('Kept content', 39, 58.7, cellA), item('Next', nx, 20, cellB)],
        612,
        792
      );
      const ours = groupLines(
        [item('Kept conten', 39, 54.8, [36, 36, 89, 50]), item('Next', nx, 20, [89, 36, 160, 50])],
        612,
        792
      );
      expect([ref.map((l) => l.text), ours.map((l) => l.text)]).toEqual([
        ['Kept content', 'Next'],
        ['Kept conten', 'Next'],
      ]);
      const score = compareTracked(ours, ref, {
        reference: [run(39, 94.6, cellA), run(nx, nx + 20, cellB)],
        ours: [run(39, 93.8, [36, 36, 89, 50]), run(nx, nx + 20, [89, 36, 160, 50])],
      });
      // The run of `Next` overlaps the line `Kept content`, but it is not inside its item, so it
      // does not give that line its clip.
      expect([score.matched, score.total, score.clipped, score.uncut]).toEqual([2, 2, [key], []]);
    }
    // Items without a clip keep the gap rule alone: the overlap joins them.
    const joined = groupLines([item('Kept content', 39, 58.7), item('Next', 91, 20)], 612, 792);
    expect(joined.map((l) => l.text)).toEqual(['Kept contentNext']);
  });

  test('two cells under one reference clip keep the clip rules at their inner ends', () => {
    // The reference PDF clips only text that overflows: `right` and `abababababab`, the texts of
    // two cells, share the clip of the print area and form one line. Each painted cell has its own
    // clip, so the painted page has two lines.
    const area = [36, 36, 576, 756] as const;
    const item = (text: string, x: number, width: number, clip: TextRun['clip']) => ({
      text,
      x,
      y: 47,
      width,
      size: 11,
      clip,
    });
    const ref = groupLines(
      [item('right', 66, 21.5, area), item('abababababab', 91, 50, area)],
      612,
      792
    );
    expect(ref.map((l) => l.text)).toEqual(['right abababababab']);
    const score = (edge: number) => {
      const ours = groupLines(
        [
          item('right', 66, 21.5, [36, 36, edge, 50]),
          item('abababababab', 91, 50, [89, 36, 143, 50]),
        ],
        612,
        792
      );
      expect(ours.map((l) => l.text)).toEqual(['right', 'abababababab']);
      return compareTracked(ours, ref, {
        reference: [run(66, 87.5, area), run(91, 141, area)],
        ours: [run(66, 87.5, [36, 36, edge, 50]), run(91, 141, [89, 36, 143, 50])],
      });
    };
    // The clip of the painted cell holds its text whole: the run passes.
    const whole = score(89);
    expect([whole.matched, whole.total, whole.extraCut, whole.uncut]).toEqual([1, 1, [], []]);
    // The painted clip ends 2 points inside the last glyph of `right`: an extra cut at the inner end.
    const cut = score(85.5);
    expect([cut.matched, cut.extraCut]).toEqual([1, ['right @ 47.0 | right clip 85.5']]);
    // A painted clip that is not one rectangle at the inner end is unknown: it fails too.
    const ours = groupLines(
      [item('right', 66, 21.5, [36, 36, 89, 50]), item('abababababab', 91, 50, [89, 36, 143, 50])],
      612,
      792
    );
    const unknown = compareTracked(ours, ref, {
      reference: [run(66, 87.5, area), run(91, 141, area)],
      ours: [run(66, 87.5, 'non-rectangular'), run(91, 141, [89, 36, 143, 50])],
    });
    // The outer left end of the line is unknown as well.
    expect(unknown.extraCut).toEqual([
      'right abababababab @ 47.0 | left clip unknown',
      'right @ 47.0 | right clip unknown',
    ]);

    // A reference clip that is not one rectangle turns no rule on: each end of the line is named.
    const refUnknown = compareTracked(ours, ref, {
      reference: [run(66, 87.5, 'non-rectangular'), run(91, 141, 'non-rectangular')],
      ours: [run(66, 87.5, [36, 36, 85.5, 50]), run(91, 141, [89, 36, 143, 50])],
    });
    expect([refUnknown.extraCut, refUnknown.unknownReference]).toEqual([
      [],
      [
        'right abababababab @ 47.0 | right reference clip unknown',
        'right abababababab @ 47.0 | left reference clip unknown',
        'right @ 47.0 | right reference clip unknown',
        'abababababab @ 47.0 | left reference clip unknown',
      ],
    ]);
    // Reference items that do not split at the painted lines name their inner ends as a painted
    // split, not as an unknown reference clip.
    const oneItem = groupLines([item('right abababababab', 66, 75, area)], 612, 792);
    const unsplit = compareTracked(ours, oneItem, {
      reference: [run(66, 141, area)],
      ours: [run(66, 87.5, [36, 36, 89, 50]), run(91, 141, [89, 36, 143, 50])],
    });
    expect(unsplit.unknownReference).toEqual([
      'right @ 47.0 | right painted split inside a reference item',
      'abababababab @ 47.0 | left painted split inside a reference item',
    ]);
  });

  test('a reference line that its clip cuts needs the painted line cut at the same edge', () => {
    // The painted paint loses the clip of the cell: every glyph shows, and the texts match in full.
    const full = [line('Kept content', 39, 47, 97.7)];
    for (const clip of [[0, 0, 612, 792] as const, 'none' as const]) {
      const lost = compareTracked(full, reference, {
        ours: [run(39, 97.7, clip)],
        reference: refRuns,
      });
      expect([lost.matched, lost.clipped, lost.uncut, lost.lostUncut]).toEqual([1, [], [key], []]);
    }
    // The painted clip 2 points to the right of the reference clip: not cut at the same edge.
    const moved = compareTracked(full, reference, {
      ours: [run(39, 97.7, [36, 36, 91, 50])],
      reference: refRuns,
    });
    expect(moved.uncut).toEqual([key]);
    // A clip at 92 points crosses the last glyph, which starts at 89.5 points: it keeps the whole
    // text, and cuts it at the same edge.
    const crossed = compareTracked(full, reference, {
      ours: [run(39, 97.7, [36, 36, 92, 50])],
      reference: [run(39, 94.6, [37, 37, 92, 50])],
    });
    expect([crossed.matched, crossed.uncut]).toEqual([1, []]);
    // The painted page that loses the line names it too, as lost (`lostUncut`).
    const gone = compareTracked([], reference, clips(89));
    expect([gone.missing, gone.uncut, gone.lostUncut]).toEqual([['Kept content'], [key], [key]]);
    // A line that no clip cuts on the reference page needs nothing.
    const plain = compareTracked(full, reference, {
      ours: [run(39, 97.7, 'none')],
      reference: [run(39, 97.7, [37, 37, 120, 50])],
    });
    expect([plain.matched, plain.uncut]).toEqual([1, []]);
  });

  test('a reference piece of a joined line of ours compares with the part that holds it', () => {
    // The reference PDF writes `Kept content`, cut at 92 points inside its last glyph, and `abc`
    // of the next cell as two lines; the painted line holds both, and its last run, of `abc`, has
    // no clip.
    const ours: PdfLine = {
      ...line('Kept contentabc', 39, 47, 108),
      items: [
        { at: 0, length: 11, xMin: 39, xMax: 97.7 },
        { at: 11, length: 3, xMin: 99, xMax: 108 },
      ],
    };
    const score = compareTracked(
      [ours],
      [line('Kept content', 39, 47, 97.7), line('abc', 99, 47, 108)],
      {
        ours: [run(39, 97.7, [36, 36, 92, 50]), run(99, 108, 'none')],
        reference: [run(39, 94.6, [37, 37, 92, 50]), run(99, 108, [92, 37, 140, 50])],
      }
    );
    expect([score.matched, score.uncut]).toEqual([2, []]);
  });

  test('the painted clip that cuts a line that the reference page shows whole is an extra cut', () => {
    // The reference line fits its clip; the painted clip ends at 66 points, inside the painted last
    // glyph, which the printed PDF keeps, so the texts match in full.
    const whole = [line('Overflow text', 10, 47, 70)];
    const extra = compareTracked(whole, whole, {
      reference: [run(10, 70, [0, 37, 200, 50])],
      ours: [run(10, 70, [8, 36, 66, 50])],
    });
    expect([extra.matched, extra.uncut, extra.extraCut]).toEqual([
      1,
      [],
      ['Overflow text @ 47.0 | right clip 66.0'],
    ]);
    // On the left side too.
    const left = compareTracked(whole, whole, {
      reference: [run(10, 70, [0, 37, 200, 50])],
      ours: [run(10, 70, [14, 36, 200, 50])],
    });
    expect(left.extraCut).toEqual(['Overflow text @ 47.0 | left clip 14.0']);
    // Controls: the painted line inside the painted clip, the painted clip within the threshold of
    // the painted line end, a line with no clip on the painted page, and no clip evidence.
    for (const ours of [
      run(10, 70, [8, 36, 120, 50]),
      run(10, 70, [8, 36, 69.6, 50]),
      run(10, 70, 'none'),
    ])
      expect(
        compareTracked(whole, whole, { reference: [run(10, 70, [0, 37, 200, 50])], ours: [ours] })
          .extraCut
      ).toEqual([]);
    expect(compareTracked(whole, whole).extraCut).toEqual([]);
  });

  test('only known reference clips apply the cut rules', () => {
    const kept = [line('Kept content', 39, 47, 97.7)];
    const ours = (clip: TextRun['clip']) => [run(39, 97.7, clip)];
    // A: the reference clips the line on both sides, and the painted page is the same: both edges
    // compare as cuts, and both hold.
    const both = compareTracked(kept, kept, {
      reference: [run(39, 97.7, [42, 37, 95, 50])],
      ours: ours([42, 36, 95, 50]),
    });
    expect([both.matched, both.uncut, both.extraCut]).toEqual([1, [], []]);
    // The painted page that loses the clip of such a line names both cuts.
    const lostBoth = compareTracked(kept, kept, {
      reference: [run(39, 97.7, [42, 37, 95, 50])],
      ours: ours('none'),
    });
    expect(lostBoth.uncut).toEqual([
      'Kept content @ 47.0 | right clip 95.0',
      'Kept content @ 47.0 | left clip 42.0',
    ]);
    // B: the reference clip is not one rectangle, so it gives no evidence: the painted cut is no
    // extra cut.
    const unknown = compareTracked(kept, kept, {
      reference: [run(39, 97.7, 'non-rectangular')],
      ours: ours([36, 36, 89, 50]),
    });
    expect([unknown.matched, unknown.uncut, unknown.extraCut]).toEqual([1, [], []]);
    // C: no reference run lies on the baseline: no evidence either.
    const none = compareTracked(kept, kept, { reference: [], ours: ours([36, 36, 89, 50]) });
    expect([none.matched, none.uncut, none.extraCut]).toEqual([1, [], []]);
    // A reference run with no clip is known: the painted cut there is an extra cut.
    const free = compareTracked(kept, kept, {
      reference: [run(39, 97.7, 'none')],
      ours: ours([36, 36, 89, 50]),
    });
    expect(free.extraCut).toEqual(['Kept content @ 47.0 | right clip 89.0']);
    // A painted clip that is unknown at an end that the reference holds whole fails that end.
    const unknownOurs = compareTracked(kept, kept, {
      reference: [run(39, 97.7, 'none')],
      ours: [run(39, 97.7, 'non-rectangular')],
    });
    expect(unknownOurs.extraCut).toEqual([
      'Kept content @ 47.0 | right clip unknown',
      'Kept content @ 47.0 | left clip unknown',
    ]);
  });

  test('each end that the reference clip holds whole keeps the painted line inside the painted clip', () => {
    const kept = [line('Kept content', 39, 47, 97.7)];
    // The reference clip cuts the line on the right, inside its last glyph, and holds its left end
    // whole. The painted clip cuts the right end at the same edge, and also cuts the left end.
    const leftToo = compareTracked(kept, kept, {
      reference: [run(39, 94.6, [37, 37, 92, 50])],
      ours: [run(39, 97.7, [42, 37, 92, 50])],
    });
    expect([leftToo.matched, leftToo.uncut, leftToo.extraCut]).toEqual([
      1,
      [],
      ['Kept content @ 47.0 | left clip 42.0'],
    ]);
    // The same end of a line that the painted page holds only in part, as a clipped line. Glyphs
    // of 7.5 points: the clip at 89 points leaves the seven of `Kept con`.
    const clipped = compareTracked(
      [line('Kept con', 39, 47, 89.2)],
      [line('Kept content', 39, 47, 121.5)],
      {
        reference: [run(39, 121.5, [37, 37, 89, 50])],
        ours: [run(39, 89.2, [42, 37, 89, 50])],
      }
    );
    expect([clipped.clipped, clipped.extraCut]).toEqual([
      ['Kept content @ 47.0 | right clip 89.0'],
      ['Kept content @ 47.0 | left clip 42.0'],
    ]);
    // An end with an unknown reference clip applies no rule; the known end still does. Each end
    // takes the clip of the run that draws the item at that end.
    const twoItems: PdfLine = {
      ...kept[0]!,
      items: [
        { at: 0, length: 4, xMin: 39, xMax: 60 },
        { at: 4, length: 7, xMin: 62, xMax: 97.7 },
      ],
    };
    const halfKnown = compareTracked([twoItems], [twoItems], {
      reference: [run(39, 60, 'non-rectangular'), run(62, 97.7, [37, 37, 120, 50])],
      ours: [run(39, 60, [37, 37, 120, 50]), run(62, 97.7, [37, 37, 92, 50])],
    });
    expect(halfKnown.extraCut).toEqual(['Kept content @ 47.0 | right clip 92.0']);
  });

  test('a clip edge inside a space of the reference text cuts the painted line at the glyph before it', () => {
    // "Kept" from 39 to 58, a space, and "content" from 61 to 97.7.
    const spaced: PdfLine = {
      ...line('Kept content', 39, 47, 97.7),
      items: [
        { at: 0, length: 4, xMin: 39, xMax: 58 },
        { at: 4, length: 7, xMin: 61, xMax: 97.7 },
      ],
    };
    const runs = (clip: TextRun['clip']) => [run(39, 58, clip), run(61, 97.7, clip)];
    // Right: the reference clip ends at 60, in the space. Paint draws no space, so the painted line
    // ends at 58, before the painted clip edge.
    const right = compareTracked([line('Kept', 39, 47, 58)], [spaced], {
      reference: runs([37, 37, 60, 50]),
      ours: [run(39, 58, [37, 37, 60, 50])],
    });
    expect([right.clipped, right.uncut, right.missing]).toEqual([
      ['Kept content @ 47.0 | right clip 60.0'],
      [],
      [],
    ]);
    // Left: the reference clip starts at 59.5, in the space.
    const left = compareTracked([line('content', 61, 47, 97.7)], [spaced], {
      reference: runs([59.5, 37, 200, 50]),
      ours: [run(61, 97.7, [59.5, 37, 200, 50])],
    });
    expect([left.clipped, left.uncut, left.missing]).toEqual([
      ['Kept content @ 47.0 | left clip 59.5'],
      [],
      [],
    ]);
    // An edge past the advance start of the next word is no space: the painted line lost a
    // glyph that the clip shows.
    const past = compareTracked([line('Kept', 39, 47, 58)], [spaced], {
      reference: runs([37, 37, 64, 50]),
      ours: [run(39, 58, [37, 37, 64, 50])],
    });
    expect([past.clipped, past.uncut, past.missing]).toEqual([
      [],
      ['Kept content @ 47.0 | right clip 64.0'],
      ['Kept content'],
    ]);
    // An extra painted cut with its edge in the space is an extra cut, not a missing line.
    const extra = compareTracked([line('Kept', 39, 47, 58)], [spaced], {
      reference: runs([37, 37, 200, 50]),
      ours: [run(39, 58, [37, 37, 60, 50])],
    });
    expect([extra.extraCut, extra.missing, extra.extra]).toEqual([
      ['Kept content @ 47.0 | right clip 60.0'],
      [],
      [],
    ]);
  });

  test('a clip edge in a space inside one reference item takes the word advances of its run', () => {
    // The reference PDFs keep a cell string as one item with its inner spaces, and one run
    // whose space glyph splits its words: "Kept" from 39 to 58 and "content" from 61 to 97.7.
    const oneItem: PdfLine = {
      ...line('Kept content', 39, 47, 97.7),
      items: [{ at: 0, length: 11, xMin: 39, xMax: 97.7 }],
    };
    const refRun = (clip: TextRun['clip']): TextRun => ({
      ...run(39, 97.7, clip),
      words: [
        [39, 58],
        [61, 97.7],
      ],
    });
    const right = compareTracked([line('Kept', 39, 47, 58)], [oneItem], {
      reference: [refRun([37, 37, 60, 50])],
      ours: [run(39, 58, [37, 37, 60, 50])],
    });
    expect([right.clipped, right.uncut]).toEqual([['Kept content @ 47.0 | right clip 60.0'], []]);
    const left = compareTracked([line('content', 61, 47, 97.7)], [oneItem], {
      reference: [refRun([59.5, 37, 200, 50])],
      ours: [run(61, 97.7, [59.5, 37, 200, 50])],
    });
    expect([left.clipped, left.uncut]).toEqual([['Kept content @ 47.0 | left clip 59.5'], []]);
    // The word before the space comes from the reference text, also when the painted line
    // joins two words into one.
    const three: PdfLine = {
      ...line('Ab cd ef', 39, 47, 70),
      items: [{ at: 0, length: 6, xMin: 39, xMax: 70 }],
    };
    const joined = compareTracked([line('Abcd', 39, 47, 55)], [three], {
      reference: [
        {
          ...run(39, 70, [37, 37, 56.5, 50]),
          words: [
            [39, 45],
            [47, 55],
            [58, 70],
          ],
        },
      ],
      ours: [run(39, 55, [37, 37, 56.5, 50])],
    });
    expect([joined.clipped, joined.uncut]).toEqual([['Ab cd ef @ 47.0 | right clip 56.5'], []]);
    // The cut end is scored: a painted word that ends 2 points before the advance end of the
    // reference word is not cut.
    const short = compareTracked([line('Kept', 39, 47, 56)], [oneItem], {
      reference: [refRun([37, 37, 60, 50])],
      ours: [run(39, 56, [37, 37, 60, 50])],
    });
    expect([short.clipped, short.uncut]).toEqual([[], ['Kept content @ 47.0 | right clip 60.0']]);
  });

  test('a clip edge in a space of the preserved-parts reference cuts the painted line there', async () => {
    const pdf = await pdfItems(join(FIXTURE_DIR, 'preserved-parts.pdf'));
    const reference = groupLines(pdf.items, 612, 792);
    const kept = reference.find((l) => l.text === 'Kept content')!;
    const keptRun = pdf.runs.find((r) => Math.abs(r.baseline - kept.baseline) < 0.25)!;
    // Its words end at 59.7 and start again at 62.0. A clip edge at 60.8 lies in the space.
    expect(keptRun.words?.map(([a, b]) => [+a.toFixed(1), +b.toFixed(1)])).toEqual([
      [39, 59.7],
      [62, 97.7],
    ]);
    const clip: TextRun['clip'] = [37, 37, 60.8, 50];
    const runs = pdf.runs.map((r) => (r === keptRun ? { ...r, clip } : r));
    const ours = [
      line('Kept', 39, 47, keptRun.words![0]![1]),
      ...reference.filter((l) => l !== kept),
    ];
    const score = compareTracked(ours, reference, {
      reference: runs,
      ours: [run(39, keptRun.words![0]![1], clip), ...pdf.runs.filter((r) => r !== keptRun)],
    });
    expect([score.clipped, score.uncut, score.missing]).toEqual([
      ['Kept content @ 47.0 | right clip 60.8'],
      [],
      [],
    ]);
  });

  test('an extra cut that hides whole glyphs is an extra cut, not a missing line', () => {
    const kept = [line('Kept content', 10, 100, 60)];
    // The painted clip at 40 hides the glyphs of "tent" entirely: the printed PDF keeps "Kept con".
    const lost = compareTracked([line('Kept con', 10, 100, 41)], kept, {
      reference: [run(10, 60, [0, 90, 200, 110], 100)],
      ours: [run(10, 41, [0, 90, 40, 110], 100)],
    });
    expect([lost.matched, lost.missing, lost.extra, lost.extraCut]).toEqual([
      1,
      [],
      [],
      ['Kept content @ 100.0 | right clip 40.0'],
    ]);
    // On the left, the painted line holds the end of the reference text.
    const left = compareTracked([line('content', 29, 100, 60)], kept, {
      reference: [run(10, 60, [0, 90, 200, 110], 100)],
      ours: [run(29, 60, [30, 90, 200, 110], 100)],
    });
    expect([left.missing, left.extraCut]).toEqual([[], ['Kept content @ 100.0 | left clip 30.0']]);
    // With an unknown reference clip, the line stays missing and the painted line stays extra.
    const unknown = compareTracked([line('Kept con', 10, 100, 41)], kept, {
      reference: [run(10, 60, 'non-rectangular', 100)],
      ours: [run(10, 41, [0, 90, 40, 110], 100)],
    });
    expect([unknown.missing, unknown.extra, unknown.extraCut]).toEqual([
      ['Kept content'],
      ['Kept con'],
      [],
    ]);
  });

  test('a line that a clip cuts on both sides stays missing', () => {
    // A centered text that runs past both edges of its cell: a named limit of the rule.
    const both = compareTracked([line('ept conten', 41, 47, 93.8)], reference, {
      ours: [run(41, 93.8, [42, 36, 89, 50])],
      reference: [run(39, 94.6, [42, 37, 89, 50])],
    });
    expect([both.matched, both.missing, both.uncut]).toEqual([0, ['Kept content'], []]);
    // The run names it, so that its note says why it is missing.
    expect(both.bothCut).toEqual(['Kept content']);
  });
});

describe('the font names of the PDFs', () => {
  test('a reference PostScript name alone gives its family', () => {
    expect(referenceFamily('AAAAAC+Calibri')).toBe('Calibri');
    expect(referenceFamily('Calibri-Bold')).toBe('Calibri');
    expect(referenceFamily('Calibri-BoldItalic')).toBe('Calibri');
    expect(referenceFamily('Calibri-Light')).toBe('Calibri Light');
    expect(baseFontName('BAAAAA+Carlito-Bold')).toBe('Carlito-Bold');
  });

  test('a vendor suffix and a comma style are not part of the family', () => {
    expect(referenceFamily('ArialMT')).toBe('Arial');
    expect(referenceFamily('Arial-BoldMT')).toBe('Arial');
    expect(referenceFamily('Arial-BoldItalicMT')).toBe('Arial');
    expect(referenceFamily('TimesNewRomanPSMT')).toBe('Times New Roman');
    expect(referenceFamily('TimesNewRomanPS-BoldMT')).toBe('Times New Roman');
    expect(referenceFamily('TimesNewRomanPS-ItalicMT')).toBe('Times New Roman');
    expect(referenceFamily('CourierNewPS-BoldMT')).toBe('Courier New');
    expect(referenceFamily('CourierNewPSMT')).toBe('Courier New');
    expect(referenceFamily('Calibri,Bold')).toBe('Calibri');
    expect(referenceFamily('Calibri,BoldItalic')).toBe('Calibri');
  });

  test('a reference font must be embedded and named by the workbook', () => {
    const named = new Set(['calibri', 'timesnewroman', 'couriernew']);
    expect(
      foreignReferenceFonts(
        [
          { name: 'AAAAAB+Calibri-Bold', embedded: true },
          { name: 'TimesNewRomanPS-BoldMT', embedded: true },
          { name: 'CourierNewPS-BoldMT', embedded: true },
        ],
        named
      )
    ).toEqual([]);
    expect(
      foreignReferenceFonts(
        [
          { name: 'ArialMT', embedded: true },
          { name: 'Calibri', embedded: false },
        ],
        named
      )
    ).toEqual(['ArialMT', 'Calibri (not embedded)']);
    // A font file of the harness resolves its PostScript name through its name table, before
    // the name itself: here a face whose name looks like Calibri belongs to another family.
    const known = new Map([
      ['Calibri-Bold', 'Carlito'],
      ['XYZ123', 'Calibri'],
    ]);
    expect(
      foreignReferenceFonts(
        [
          { name: 'AAAAAB+Calibri-Bold', embedded: true },
          { name: 'AAAAAC+XYZ123', embedded: true },
        ],
        named,
        known
      )
    ).toEqual(['AAAAAB+Calibri-Bold']);
    // The refusal names its next step.
    expect(FOREIGN_REFERENCE_REMEDY).toContain('export the reference PDF again');
    expect(FOREIGN_REFERENCE_REMEDY).toContain('fidelity:record-reference');
  });

  test('the family of a face is its name id 1, not its typographic family', () => {
    // A font file with one table, `name`: Windows US English records for the family (id 1),
    // the PostScript name (id 6), and the typographic family (id 16).
    const records: [number, string][] = [
      [1, 'Calibri Light'],
      [6, 'Calibri-Light'],
      [16, 'Calibri'],
    ];
    const text = records.map(([, t]) => [...t].flatMap((c) => [0, c.charCodeAt(0)]));
    const header = 6 + records.length * 12;
    const name: number[] = [0, 0, 0, records.length, header >> 8, header & 255];
    let at = 0;
    records.forEach(([id], i) => {
      const len = text[i]!.length;
      name.push(0, 3, 0, 1, 0x04, 0x09, 0, id, len >> 8, len & 255, at >> 8, at & 255);
      at += len;
    });
    name.push(...text.flat());
    const dir = [0, 1, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0];
    const offset = dir.length + 16;
    const bytes = new Uint8Array([
      ...dir,
      ...[...'name'].map((c) => c.charCodeAt(0)),
      0,
      0,
      0,
      0,
      0,
      0,
      0,
      offset,
      0,
      0,
      name.length >> 8,
      name.length & 255,
      ...name,
    ]);
    expect(familyName(bytes)).toBe('Calibri Light');
    expect(postScriptName(bytes)).toBe('Calibri-Light');
    const fontsDir = mkdtempSync(join(tmpdir(), 'xlsx-fonts-'));
    try {
      writeFileSync(join(fontsDir, 'light.ttf'), bytes);
      expect(postScriptFamilies([fontsDir])).toEqual(new Map([['Calibri-Light', 'Calibri Light']]));
    } finally {
      rmSync(fontsDir, { recursive: true, force: true });
    }
  });

  test('the name table of a substitute gives the PostScript name a PDF writes', () => {
    const fonts = dirname(
      Bun.resolveSync(
        '@docx-editor.dev/fonts/package.json',
        join(import.meta.dir, '../../packages/core')
      )
    );
    const file = join(fonts, 'assets', 'Carlito-Bold.ttf');
    const bytes = new Uint8Array(readFileSync(file));
    expect(postScriptName(bytes)).toBe('Carlito-Bold');
    expect(familyName(bytes)).toBe('Carlito');
    expect(macStyle(bytes)).toEqual({ bold: true, italic: false });
    // Truncated bytes have no table: no name, never a read past the end.
    expect(postScriptName(bytes.subarray(0, 40))).toBeUndefined();
    // A font folder resolves each PostScript name to the family of its name table.
    const dir = mkdtempSync(join(tmpdir(), 'xlsx-fonts-'));
    try {
      copyFileSync(file, join(dir, 'Carlito-Bold.ttf'));
      expect(postScriptFamilies([dir])).toEqual(new Map([['Carlito-Bold', 'Carlito']]));
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  });
});

describe('the printed page of ours', () => {
  const items = (n: number, patch: Partial<PdfItems> = {}): PdfItems => ({
    items: Array.from({ length: n }, () => ({ text: 'A1', x: 40, y: 50, width: 10, size: 11 })),
    rotated: 0,
    size: [612, 792],
    pages: 1,
    runs: [],
    fonts: [{ name: 'Carlito-Regular', embedded: true }],
    ...patch,
  });
  const loaded = new Set(['Carlito-Regular']);

  test('pdfItems reads a PDF from bytes as from its path, with its fonts and pages', async () => {
    const path = join(FIXTURE_DIR, 'preserved-parts.pdf');
    const bytes = new Uint8Array(readFileSync(path));
    const fromBytes = await pdfItems(bytes);
    // pdf.js names fonts by document (`g_d0_f1`), so the font names differ between two reads.
    const plain = (r: PdfItems) => ({ ...r, items: r.items.map(({ font: _, ...i }) => i) });
    expect(plain(fromBytes)).toEqual(plain(await pdfItems(path)));
    expect(fromBytes.items.map((i) => [i.text, i.postScript])).toEqual([
      ['Kept content', 'Carlito-Regular'],
      ['5', 'Carlito-Bold'],
    ]);
    expect([fromBytes.rotated, fromBytes.pages]).toEqual([0, 1]);
    expect(fromBytes.fonts).toEqual([
      { name: 'Carlito-Bold', embedded: true },
      { name: 'Carlito-Regular', embedded: true },
    ]);
    // pdf.js takes over the buffer that it reads; the caller's bytes stay usable.
    expect(bytes.byteLength).toBeGreaterThan(0);
  });

  test('a printed page must be one page of the reference size, in loaded fonts, with text', () => {
    expect(printedPageProblem(items(3), [612, 792], 3, loaded)).toBeUndefined();
    expect(printedPageProblem(items(3), [612.4, 792], 3, loaded)).toBeUndefined();
    expect(printedPageProblem(items(3), [595, 842], 3, loaded)).toBe(
      'the printed page is 612 x 792 pt, the reference page 595 x 842 pt; set the `@page` size in `page/index.html` to the reference page size'
    );
    expect(printedPageProblem(items(3, { pages: 2 }), [612, 792], 3, loaded)).toBe(
      'the printed PDF has 2 pages, not 1; check the `@page` size and the canvas size of the harness page (`page/index.html`, `page/main.ts`)'
    );
    // A system font, or a font that the PDF does not embed, is not a face of the page.
    const foreign = items(3, {
      fonts: [
        { name: 'Carlito-Regular', embedded: true },
        { name: 'Helvetica', embedded: true },
        { name: 'Carlito-Regular', embedded: false },
      ],
    });
    expect(printedPageProblem(foreign, [612, 792], 3, loaded)).toBe(
      'the printed PDF draws with fonts that the page did not load: Helvetica, Carlito-Regular (not embedded); make the paint use the faces that the measurer loaded, or load the missing face'
    );
    // A canvas that Chromium prints as an image holds no text.
    expect(printedPageProblem(items(0, { fonts: [] }), [612, 792], 4, loaded)).toBe(
      'the paint drew 4 texts, but the printed page holds no text; paint the canvas during `beforeprint` (`page/main.ts`), or Chromium prints it as an image'
    );
    // Rotated text counts as text, and a page without painted text needs none.
    expect(printedPageProblem(items(0, { rotated: 2 }), [612, 792], 2, loaded)).toBeUndefined();
    expect(printedPageProblem(items(0, { fonts: [] }), [612, 792], 0, loaded)).toBeUndefined();
  });
});

describe('the print sequence of the harness page', () => {
  /**
   * A page that records the calls in order. `paints` is how often its print paint runs while
   * it prints; `armed` is whether a fixture is rendered.
   */
  function fakePage(paints: number, armed = true, failPdf = false) {
    const calls: string[] = [];
    const page: PrintablePage = {
      async evaluate<R>(fn: () => R | Promise<R>): Promise<R> {
        const text = String(fn);
        if (text.includes('harnessPrint')) {
          calls.push('print');
          return armed as R;
        }
        calls.push('end');
        return paints as R;
      },
      async pdf(options) {
        calls.push(`pdf ${options.preferCSSPageSize} ${options.printBackground}`);
        if (failPdf) throw new Error('printing failed');
        return new Uint8Array([37, 80, 68, 70]);
      },
    };
    return { page, calls };
  }

  test('arms the print paint, prints with backgrounds, and ends the print', async () => {
    const { page, calls } = fakePage(1);
    expect([...(await printPainted(page))]).toEqual([37, 80, 68, 70]);
    expect(calls).toEqual(['print', 'pdf true true', 'end']);
  });

  test('refuses a PDF unless the print paint ran exactly once', async () => {
    await expect(printPainted(fakePage(0).page)).rejects.toThrow(
      'the print paint ran 0 times while Chromium printed, not once; arm the paint with `harnessPrint` and paint once in the `beforeprint` listener (`page/main.ts`)'
    );
    await expect(printPainted(fakePage(2).page)).rejects.toThrow('ran 2 times');
  });

  test('ends the print when printing fails, and refuses a page with no fixture', async () => {
    const failing = fakePage(1, true, true);
    await expect(printPainted(failing.page)).rejects.toThrow('printing failed');
    expect(failing.calls).toEqual(['print', 'pdf true true', 'end']);
    const empty = fakePage(0, false);
    await expect(printPainted(empty.page)).rejects.toThrow(
      'the page has no rendered fixture to print; call `harnessRender` and check that it returns ok first'
    );
    expect(empty.calls).toEqual(['print']);
  });
});
