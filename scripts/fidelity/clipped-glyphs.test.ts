// The text and glyph conditions of a line that a cell clip cuts: the painted line holds the
// reference glyphs that the clip leaves, a clip can leave one character, and leading symbols
// compare alike on both sides. Also the clip of each end, and the tolerance of each edge. The
// other clip rules of `compareLines`: `printed-lines.test.ts`.

import { describe, expect, test } from 'bun:test';
import { compareLines, spanOf } from './compare.ts';
import { compareTracked } from './test-glyphs.ts';
import { extraCutText } from './known-gaps.ts';
import { referenceEnd } from './clip-rules.ts';
import { THRESHOLDS } from './thresholds.ts';
import type { TextRun } from './pdf-clips.ts';
import type { PdfLine } from './poppler.ts';

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

describe('the glyphs of a line that a cell clip cuts', () => {
  // The reference PDF keeps every glyph of `Kept content` and clips it at 89 points.
  const reference = [line('Kept content', 39, 47, 97.7)];
  const refRuns = [run(39, 94.6, [37, 37, 89, 50])];
  const ourRuns = (edge: number, end = 93.8) => [run(39, end, [36, 36, edge, 50])];
  const clips = (edge: number, end?: number) => ({
    ours: ourRuns(edge, end),
    reference: refRuns,
  });
  const key = 'Kept content @ 47.0 | right clip 89.0';

  test('a clipped line holds the glyphs that the reference clip leaves, no more and no less', () => {
    // Glyph advances of `Kept content`: 11 glyphs from 39 to 94.6 points. The clip at 89 points
    // leaves the glyphs whose advance starts before it: `Keptconten`.
    const step = (94.6 - 39) / 11;
    const glyphs = Array.from(
      { length: 11 },
      (_, i) => [39 + i * step, 39 + (i + 1) * step] as const
    );
    const evidence = (end: number) => ({
      ours: [run(39, end, [36, 36, 89, 50])],
      reference: [{ ...run(39, 94.6, [37, 37, 89, 50]), glyphs }],
    });
    const right = compareTracked([line('Kept conten', 39, 47, 93.8)], reference, evidence(93.8));
    expect([right.matched, right.clipped]).toEqual([1, [key]]);
    // Wider painted glyphs reach the edge with fewer of them: the line does not match.
    for (const part of ['Kept cont', 'Kept c', 'K']) {
      const score = compareTracked([line(part, 39, 47, 89)], reference, evidence(89));
      expect([score.matched, score.clipped, score.uncut]).toEqual([0, [], [key]]);
    }
    // A painted line with every glyph, drawn narrower so that it ends at the clip edge, shows the
    // `t` that the reference clip hides: it matches by its text, and fails as uncut.
    const narrow = compareTracked([line('Kept content', 39, 47, 88.9)], reference, evidence(88.9));
    expect([narrow.matched, narrow.uncut]).toEqual([1, [key]]);
    // A run of the next cell under another clip, inside the reach of the overflow, is not a glyph
    // of the line.
    // Here the next cell holds the digit `7` at 91 to 93 points.
    const neighbour = compareTracked(
      [line('Kept conten', 39, 47, 93.8), line('7', 91, 47, 93)],
      [...reference, line('7', 91, 47, 93)],
      {
        ours: [run(39, 93.8, [36, 36, 89, 50]), run(91, 93, [89, 36, 140, 50])],
        reference: [
          { ...run(39, 94.6, [37, 37, 89, 50]), glyphs },
          { ...run(91, 93, [89, 37, 140, 50]), glyphs: [[91, 93]] },
        ],
      }
    );
    // The glyphs of the line come from the runs under its own clip only: the `7` would make the
    // reference glyphs 12 for 11 characters, an unknown end.
    expect([
      neighbour.matched,
      neighbour.clipped,
      neighbour.uncut,
      neighbour.unknownReference,
    ]).toEqual([2, [key], [], []]);
    // Two digits `77` from 90 to 97 points, inside the overflow on both pages.
    const two = compareTracked(
      [line('Kept conten', 39, 47, 93.8), line('77', 90, 47, 97)],
      [...reference, line('77', 90, 47, 97)],
      {
        ours: [run(39, 93.8, [36, 36, 89, 50]), run(90, 97, [89, 36, 140, 50])],
        reference: [
          { ...run(39, 94.6, [37, 37, 89, 50]), glyphs },
          {
            ...run(90, 97, [89, 37, 140, 50]),
            glyphs: [
              [90, 93.5],
              [93.5, 97],
            ],
          },
        ],
      }
    );
    expect([two.matched, two.clipped, two.uncut, two.extraCut, two.unknownReference]).toEqual([
      2,
      [key],
      [],
      [],
      [],
    ]);
    // A run of the next cell that ends past the line, under another clip, does not end the line.
    // Here the next cell holds the digit `5` at 92.5 to 98.08 points on both pages.
    const past = compareTracked(
      [line('Kept conten', 39, 47, 93.8), line('5', 92.5, 47, 98.08)],
      [...reference, line('5', 92.5, 47, 98.08)],
      {
        ours: [run(39, 93.8, [36, 36, 89, 50]), run(92.5, 98.08, [89, 36, 140, 50])],
        reference: [
          { ...run(39, 94.6, [37, 37, 89, 50]), glyphs },
          { ...run(92.5, 98.08, [89, 37, 140, 50]), glyphs: [[92.5, 98.08]] },
        ],
      }
    );
    expect([past.matched, past.clipped, past.uncut, past.extraCut, past.unknownReference]).toEqual([
      2,
      [key],
      [],
      [],
      [],
    ]);
    // A reference ligature draws `ep` as one glyph, so the reference glyphs do not account for
    // the text of the line. The text rule matches the line, and its cut end is named as unknown,
    // which fails the page with the remedy of an unknown reference clip. It is not uncut.
    const ligature = [glyphs[0]!, [glyphs[1]![0], glyphs[2]![1]] as const, ...glyphs.slice(3)];
    const unknown = compareTracked([line('Kept conten', 39, 47, 93.8)], reference, {
      ...evidence(93.8),
      reference: [{ ...run(39, 94.6, [37, 37, 89, 50]), glyphs: ligature }],
    });
    expect([unknown.matched, unknown.clipped, unknown.uncut, unknown.unknownReference]).toEqual([
      1,
      [key],
      [],
      ['Kept content @ 47.0 | right reference glyphs unknown'],
    ]);
    // A painted line that lost the clip stays uncut, also with unknown reference glyphs.
    const lost = compareTracked([line('Kept content', 39, 47, 97.7)], reference, {
      ours: [run(39, 97.7, 'none')],
      reference: [{ ...run(39, 94.6, [37, 37, 89, 50]), glyphs: ligature }],
    });
    expect([lost.matched, lost.uncut]).toEqual([1, [key]]);
  });

  test('a run that starts a line of the next cell does not end a painted line', () => {
    // The painted `77` of the next cell starts inside the reach of the overflow and ends past the
    // painted line. Its run starts the item of its own line, so it does not end `Kept conten`.
    const score = compareTracked(
      [line('Kept conten', 39, 47, 93.8), line('77', 90, 47, 94.2)],
      [...reference, line('77', 90, 47, 94.2)],
      {
        ours: [run(39, 93.8, [36, 36, 89, 50]), run(90, 94.2, [89, 36, 140, 50])],
        reference: [run(39, 94.6, [37, 37, 89, 50]), run(90, 94.2, [89, 37, 140, 50])],
      }
    );
    expect([score.clipped, score.uncut, score.extraCut, score.unknownReference]).toEqual([
      [key],
      [],
      [],
      [],
    ]);
    // On the left, the first run decides: `K` under the clip at 46 points, `e` under another one.
    const left = compareTracked([line('Ke', 46, 47, 60)], [line('XKe', 39, 47, 60)], {
      ours: [run(46, 50, [46, 36, 100, 50]), run(50, 60, [0, 36, 100, 50])],
      reference: [run(39, 60, [46, 37, 100, 50])],
    });
    expect([left.clipped, left.uncut, left.extraCut]).toEqual([
      ['XKe @ 47.0 | left clip 46.0'],
      [],
      [],
    ]);
    // The same in another run order, and with a run of `e` drawn again inside the run of `Ke`:
    // the run that starts first decides, not the first one listed or the one that ends first.
    for (const ours of [
      [run(50, 60, [0, 36, 100, 50]), run(46, 50, [46, 36, 100, 50])],
      [run(46, 60, [46, 36, 100, 50]), run(52, 56, [0, 36, 100, 50])],
    ]) {
      const score = compareTracked([line('Ke', 46, 47, 60)], [line('XKe', 39, 47, 60)], {
        ours,
        reference: [run(39, 60, [46, 37, 100, 50])],
      });
      expect([score.clipped, score.uncut]).toEqual([['XKe @ 47.0 | left clip 46.0'], []]);
    }
  });

  test('the runs of the next cell inside the end item of a line are not its end', () => {
    // `AW` of cell A runs past its clip at 89 points under `is` of cell B, which starts at 92
    // points. Each page draws `i` and `s` with their own runs, under the clip of cell B.
    const items = (text: string, xMin: number, xMax: number): PdfLine => ({
      ...line(text, xMin, 47, xMax),
      items: [{ at: 0, length: text.length, xMin, xMax }],
    });
    const cellA = [36, 36, 89, 50] as const;
    const cellB = [89, 36, 142, 50] as const;
    const pages = (o: {
      sEnd?: number;
      iStart?: number;
      a?: TextRun;
      w?: TextRun;
      stray?: TextRun[];
    }) => {
      const sEnd = o.sEnd ?? 96.1;
      const lines = [items('AW', 79, 95.8), items('is', 92, sEnd)];
      return compareTracked(lines, lines, {
        ours: [
          o.a ?? run(79, 86, cellA),
          o.w ?? run(86, 95.8, cellA),
          run(o.iStart ?? 92, 94.44, cellB),
          run(94.44, sEnd, cellB),
          ...(o.stray ?? []),
        ],
        reference: [run(79, 95.8, [37, 37, 89, 50]), run(92, sEnd, [89, 37, 143, 51])],
      });
    };
    const passes = (score: ReturnType<typeof pages>): void =>
      expect([score.matched, score.uncut, score.extraCut, score.unknownReference]).toEqual([
        2,
        [],
        [],
        [],
      ]);
    // The `s` ends inside or past the end item of `AW`, also when its run starts the item 0.4
    // points late (`ON_ITEM_PT`).
    for (const sEnd of [95.5, 96.1]) passes(pages({ sEnd }));
    passes(pages({ iStart: 92.4 }));
    // A run that ends 0.6 points past the end item, or starts 0.6 points before it, is not in it.
    passes(pages({ stray: [run(95, 96.4, 'none')] }));
    passes(pages({ stray: [run(78.4, 80, [80, 36, 200, 50])] }));
    // A `W` without its clip is the end of `AW`, also where it starts 0.6 points before `is`.
    for (const w of [run(86, 95.8, 'none'), run(91.4, 95.8, 'none')]) {
      const score = pages({ a: run(79, w.x0, cellA), w });
      expect(score.uncut).toEqual(['AW @ 47.0 | right clip 89.0']);
    }
  });

  test('runs with different clips at the start of an item of the next cell make the end unknown', () => {
    // On the reference page, `W` of `AW` starts at 92 points, where `is` of the next cell starts
    // too: no run tells which line it draws. The painted `AW` has no clip, so a guess for the next
    // cell would pass it. The ends of `AW` are unknown instead, and fail as a skipped check.
    const items = (text: string, xMin: number, xMax: number): PdfLine => ({
      ...line(text, xMin, 47, xMax),
      items: [{ at: 0, length: text.length, xMin, xMax }],
    });
    const lines = [items('AW', 79, 95.8), items('is', 92, 96.1)];
    const cellB = [89, 37, 143, 51] as const;
    const score = compareTracked(lines, lines, {
      ours: [
        run(79, 95.8, 'none'),
        run(92, 94.44, [89, 36, 142, 50]),
        run(94.44, 96.1, [89, 36, 142, 50]),
      ],
      reference: [
        run(79, 92, [37, 37, 89, 50]),
        run(92, 95.8, [37, 37, 89, 50]),
        run(92, 94.44, cellB),
        run(94.44, 96.1, cellB),
      ],
    });
    expect(score.unknownReference).toEqual([
      'AW @ 47.0 | right reference clip unknown',
      'AW @ 47.0 | left reference clip unknown',
    ]);
  });

  test('runs with different clips at the start of an item of the next cell make a painted end unknown', () => {
    // The painted `W` of `AW` starts at 92 points, where the painted `is` of the next cell starts
    // too. The reference `AW` has no clip. The ends of the painted `AW` are unknown, and the
    // message names the tie as a cause.
    const items = (text: string, xMin: number, xMax: number): PdfLine => ({
      ...line(text, xMin, 47, xMax),
      items: [{ at: 0, length: text.length, xMin, xMax }],
    });
    const lines = [items('AW', 79, 95.8), items('is', 92, 96.1)];
    const cellB = [89, 36, 142, 50] as const;
    const score = compareTracked(lines, lines, {
      ours: [
        run(79, 92, [36, 36, 140, 50]),
        run(92, 95.8, [36, 36, 140, 50]),
        run(92, 94.44, cellB),
        run(94.44, 96.1, cellB),
      ],
      reference: [
        run(79, 95.8, 'none'),
        run(92, 94.44, [89, 37, 143, 51]),
        run(94.44, 96.1, [89, 37, 143, 51]),
      ],
    });
    expect(score.extraCut).toEqual([
      'AW @ 47.0 | right clip unknown',
      'AW @ 47.0 | left clip unknown',
    ]);
    for (const key of score.extraCut)
      expect(extraCutText(key)).toContain(
        'runs with different clips start an item of the next cell there'
      );
  });

  test('a line that a clip cuts does not take a whole line of the same text on another row', () => {
    // Two rows cut `This text is longer than its cell` at 108 points on both pages; a third row
    // shows it whole. The rows that the clip cuts match their own cut lines, and the whole row
    // matches its whole line.
    const text = 'This text is longer than its cell';
    const cut = (baseline: number) => [37, baseline - 10, 108, baseline + 3] as const;
    const reference = [47, 75, 103].map((b) => line(text, 39, b, 170.5));
    const ours = [
      line('This text is longe', 39, 47, 110),
      line('This text is longe', 39, 75, 110),
      line(text, 39, 103, 170.5),
    ];
    const runs = (x1: number) => [
      run(39, x1, cut(47), 47),
      run(39, x1, cut(75), 75),
      run(39, 170.5, 'none', 103),
    ];
    const score = compareTracked(ours, reference, { ours: runs(110), reference: runs(170.5) });
    expect([score.matched, score.total, score.worstBaseline]).toEqual([3, 3, 0]);
    expect(score.clipped).toEqual([
      `${text} @ 47.0 | right clip 108.0`,
      `${text} @ 75.0 | right clip 108.0`,
    ]);
    expect([score.uncut, score.missing, score.extra]).toEqual([[], [], []]);
  });

  test('lines of the same text on one row pair nearest pair first, in any reference order', () => {
    // `Revenue` in A1 is cut at 60 points on both pages; B1 shows it whole. The whole painted line
    // pairs with B1, its nearest twin, and A1 matches as clipped.
    const ours = [line('Reve', 39, 47, 62), line('Revenue', 62, 47, 100)];
    const evidence = {
      ours: [run(39, 62, [36, 36, 60, 50]), run(62, 100, [60, 36, 140, 50])],
      reference: [run(39, 77, [37, 37, 60, 50]), run(62, 100, 'none')],
    };
    const cells = [line('Revenue', 39, 47, 77), line('Revenue', 62, 47, 100)];
    for (const reference of [cells, [...cells].reverse()]) {
      const score = compareTracked(ours, reference, evidence);
      expect([score.matched, score.clipped]).toEqual([2, ['Revenue @ 47.0 | right clip 60.0']]);
      expect([score.uncut, score.extraCut, score.missing, score.extra]).toEqual([[], [], [], []]);
    }
    // Two copies of a long text on one row: A1 is cut at 108 points, C1 shows it whole at 200.
    const text = 'This text is longer than its cell';
    const runs = (x1: number) => [run(39, x1, [37, 37, 108, 50]), run(200, 331.5, 'none')];
    const score = compareTracked(
      [line('This text is longe', 39, 47, 110), line(text, 200, 47, 331.5)],
      [line(text, 39, 47, 170.5), line(text, 200, 47, 331.5)],
      { ours: runs(110), reference: runs(170.5) }
    );
    expect([score.matched, score.clipped, score.worstShift]).toEqual([
      2,
      [`${text} @ 47.0 | right clip 108.0`],
      0,
    ]);
    expect([score.uncut, score.extraCut]).toEqual([[], []]);
  });

  test('the clip passes take a painted line within the baseline threshold, and score it', () => {
    // The painted line lies 0.3 points low, inside `maxBaselineDeltaPt`.
    const reference = [line('Kept content', 39, 47, 97.7)];
    const ours = [line('Kept conten', 39, 47.3, 93.8)];
    const painted = [run(39, 93.8, [36, 36, 89, 50], 47.3)];
    const cut = compareTracked(ours, reference, {
      ours: painted,
      reference: [run(39, 94.6, [37, 37, 89, 50])],
    });
    expect([cut.matched, cut.clipped, cut.uncut]).toEqual([
      1,
      ['Kept content @ 47.0 | right clip 89.0'],
      [],
    ]);
    expect(cut.worstBaseline).toBeCloseTo(0.3, 6);
    // The reference page shows the line whole: the painted clip cuts it, an extra cut.
    const extra = compareTracked(ours, reference, {
      ours: painted,
      reference: [run(39, 97.7, 'none')],
    });
    expect([extra.matched, extra.extraCut, extra.missing]).toEqual([
      1,
      ['Kept content @ 47.0 | right clip 89.0'],
      [],
    ]);
    expect(extra.worstBaseline).toBeCloseTo(0.3, 6);
  });

  test('a tighter line position threshold does not change which ends count as cut or whole', () => {
    // The reference line ends 0.4 points past its clip edge at 89 points: within `clipEdgePt`, so
    // the clip holds the end whole. Tightening `maxLineShiftPt` to 0.3 keeps that verdict.
    const reference = line('Kept content', 39, 47, 89.4);
    const refRuns = [run(39, 89.4, [37, 37, 89, 50])];
    const verdicts = () =>
      [[37, 37, 89, 50] as const, 'none' as const].map((clip) => {
        const evidence = { reference: refRuns, ours: [run(39, 89.4, clip)] };
        const score = compareTracked([line('Kept content', 39, 47, 89.4)], [reference], evidence);
        return [
          referenceEnd(reference, evidence, 'right'),
          score.matched,
          score.uncut,
          score.extraCut,
        ];
      });
    const before = verdicts();
    expect(before[0]![0]).toEqual({ kind: 'whole' });
    const limits = THRESHOLDS as { maxLineShiftPt: number };
    const was = limits.maxLineShiftPt;
    try {
      limits.maxLineShiftPt = 0.3;
      expect(verdicts()).toEqual(before);
    } finally {
      limits.maxLineShiftPt = was;
    }
  });

  test('a line that no rule matches on its row matches the same text on another row', () => {
    // The painted line lies 1 point below the reference line, past `maxBaselineDeltaPt`. It still
    // matches, and the baseline check scores the shift.
    const score = compareTracked([line('Kept', 39, 48, 60)], [line('Kept', 39, 47, 60)]);
    expect([score.matched, score.total, score.worstBaseline]).toEqual([1, 1, 1]);
    expect([score.missing, score.extra]).toEqual([[], []]);
  });

  test('a joined line counts the items of each piece from the start of the joined text', () => {
    const first: PdfLine = {
      ...line('ab c', 10, 47, 30),
      items: [
        { at: 0, length: 2, xMin: 10, xMax: 20 },
        { at: 2, length: 1, xMin: 25, xMax: 30 },
      ],
    };
    const joined = spanOf([first, line('de', 40, 47, 50)]);
    expect([joined.text, joined.xMin, joined.xMax, joined.items]).toEqual([
      'ab c de',
      10,
      50,
      [
        { at: 0, length: 2, xMin: 10, xMax: 20 },
        { at: 2, length: 1, xMin: 25, xMax: 30 },
        { at: 3, length: 2, xMin: 40, xMax: 50 },
      ],
    ]);
  });

  test('the clip of the run that draws the last glyph decides an end, on both pages', () => {
    // Eleven glyphs of `Kept content`; the clip at 89 points leaves ten. Each page draws the last
    // glyph of the line with its own run.
    const step = (94.6 - 39) / 11;
    const g = Array.from({ length: 11 }, (_, i) => [39 + i * step, 39 + (i + 1) * step] as const);
    const whole = [line('Kept content', 39, 47, 94.6)];
    const refRuns = [{ ...run(39, 94.6, [37, 37, 89, 50]), glyphs: g }];
    // The painted `n` under its own clip: no clip, a wider one, or one at 80 points. Only the clip
    // at 89 points cuts the line at the reference edge.
    const painted = (last: TextRun['clip']) =>
      compareTracked([line('Kept conten', 39, 47, 93.8)], whole, {
        ours: [run(39, g[9]![0], [36, 36, 89, 50]), run(g[9]![0], 93.8, last)],
        reference: refRuns,
      });
    const same = painted([36, 36, 89, 50]);
    expect([same.matched, same.clipped, same.uncut]).toEqual([1, [key], []]);
    for (const last of ['none', [36, 36, 200, 50], [36, 36, 80, 50]] as const) {
      const score = painted(last);
      expect([score.matched, score.clipped, score.uncut, score.missing]).toEqual([
        0,
        [],
        [key],
        ['Kept content'],
      ]);
    }
    // The reference `t` under another clip than the rest of the line. Without a clip, or under a
    // wider one, the reference shows the `t` whole: the painted cut at 89 points is an extra cut.
    // Under a clip at 91 points, the reference cuts the line there: the painted line is uncut.
    const refLast = (last: TextRun['clip']) =>
      compareTracked([line('Kept conten', 39, 47, 93.8)], whole, {
        ours: [run(39, 93.8, [36, 36, 89, 50])],
        reference: [
          { ...run(39, g[9]![1], [37, 37, 89, 50]), glyphs: g.slice(0, 10) },
          { ...run(g[10]![0], 94.6, last), glyphs: g.slice(10) },
        ],
      });
    for (const last of ['none', [37, 37, 120, 50]] as const) {
      const score = refLast(last);
      expect([score.matched, score.clipped, score.extraCut]).toEqual([1, [], [key]]);
    }
    const at91 = refLast([37, 37, 91, 50]);
    expect([at91.matched, at91.uncut, at91.missing]).toEqual([
      0,
      ['Kept content @ 47.0 | right clip 91.0'],
      ['Kept content'],
    ]);
    // The last glyph `2` of `id 12` under a clip at 59 points, while the reference holds the line
    // whole: the painted end is an extra cut, as when the clip of the whole cell moves.
    const id = [line('id 12', 39, 47, 61.6)];
    const cutTwo = compareTracked(id, id, {
      ours: [run(39, 56, [36, 36, 100, 50]), run(56, 61.6, [36, 36, 59, 50])],
      reference: [run(39, 61.6, [37, 37, 100, 50])],
    });
    expect(cutTwo.extraCut).toEqual(['id 12 @ 47.0 | right clip 59.0']);
  });

  test('the painted clip within the threshold of the reference clip is part of a cut', () => {
    // The painted line holds the ten glyphs that the reference clip leaves, but its own clip lies
    // 2 points to the right or to the left of the reference clip.
    for (const edge of [91, 87]) {
      const score = compareTracked([line('Kept conten', 39, 47, 93.8)], reference, clips(edge));
      expect([score.matched, score.clipped, score.uncut, score.missing]).toEqual([
        0,
        [],
        [key],
        ['Kept content'],
      ]);
    }
    // The same on a cut on the left: `XKe` cut at 46 points leaves `Ke`.
    const left = (edge: number) =>
      compareTracked([line('Ke', 46, 47, 60)], [line('XKe', 39, 47, 60)], {
        ours: [run(46, 60, [edge, 36, 100, 50])],
        reference: [run(39, 60, [46, 37, 100, 50])],
      });
    expect(left(46).clipped).toEqual(['XKe @ 47.0 | left clip 46.0']);
    for (const edge of [44, 48]) {
      const score = left(edge);
      expect([score.matched, score.uncut, score.missing]).toEqual([
        0,
        ['XKe @ 47.0 | left clip 46.0'],
        ['XKe'],
      ]);
    }
  });

  test('a glyph that ends more than the threshold before a left clip edge is hidden', () => {
    // `X` ends at 45.2 points, 0.8 points before the clip edge at 46: the clip hides it, so a
    // painted line that shows it fails as uncut.
    const xk = [line('XK', 39, 47, 53)];
    const score = compareLines(xk, xk, {
      ours: [run(39, 53, [46, 36, 100, 50])],
      reference: [
        {
          ...run(39, 53, [46, 37, 100, 50]),
          glyphs: [
            [39, 45.2],
            [46.2, 53],
          ],
        },
      ],
    });
    expect([score.matched, score.uncut]).toEqual([1, ['XK @ 47.0 | left clip 46.0']]);
  });

  test('a run without glyph advances skips the glyph rule, as only a run that a test builds does', () => {
    // `Kept cont` holds fewer glyphs than the clip leaves, so the tracked runs refuse it.
    const narrow = [line('Kept cont', 39, 47, 89)];
    expect(compareTracked(narrow, reference, clips(89, 89)).uncut).toEqual([key]);
    const untracked = compareLines(narrow, reference, clips(89, 89));
    expect([untracked.matched, untracked.clipped]).toEqual([1, [key]]);
  });

  test('a painted piece that splits a reference item names both of its inner ends as a painted split', () => {
    // The reference line `ab cdef` has the items `ab` and `cdef`; the painted pieces `abc` and
    // `def` split inside `cdef`, so no item of the reference holds `abc` alone.
    const area = [0, 0, 600, 800] as const;
    const ref: PdfLine = {
      ...line('ab cdef', 10, 47, 40),
      items: [
        { at: 0, length: 2, xMin: 10, xMax: 20 },
        { at: 2, length: 4, xMin: 22, xMax: 40 },
      ],
    };
    const first: PdfLine = {
      ...line('abc', 10, 47, 25),
      items: [
        { at: 0, length: 2, xMin: 10, xMax: 20 },
        { at: 2, length: 1, xMin: 22, xMax: 25 },
      ],
    };
    const second: PdfLine = {
      ...line('def', 27, 47, 40),
      items: [{ at: 0, length: 3, xMin: 27, xMax: 40 }],
    };
    const score = compareTracked([first, second], [ref], {
      reference: [run(10, 20, area), run(22, 40, area)],
      ours: [
        run(10, 20, [0, 0, 26, 800]),
        run(22, 25, [0, 0, 26, 800]),
        run(27, 40, [26, 0, 600, 800]),
      ],
    });
    expect([score.matched, score.uncut, score.extraCut, score.unknownReference]).toEqual([
      1,
      [],
      [],
      [
        'abc @ 47.0 | right painted split inside a reference item',
        'def @ 47.0 | left painted split inside a reference item',
      ],
    ]);
  });

  test('a joined line holds the glyphs that the clip leaves, on both join paths', () => {
    // Two reference pieces on one baseline hold one painted line: `Kept content`, cut at 94
    // points, and `abc` of the next cell. The clip crosses the last glyph, so it leaves all 11.
    const joined: PdfLine = {
      ...line('Kept contentabc', 39, 47, 108),
      items: [
        { at: 0, length: 11, xMin: 39, xMax: 94.6 },
        { at: 11, length: 3, xMin: 99, xMax: 108 },
      ],
    };
    const pieces = compareTracked(
      [joined],
      [line('Kept content', 39, 47, 94.6), line('abc', 99, 47, 108)],
      {
        ours: [run(39, 94.6, [36, 36, 94, 50]), run(99, 108, 'none')],
        reference: [run(39, 94.6, [37, 37, 94, 50]), run(99, 108, [94, 37, 140, 50])],
      }
    );
    expect([pieces.matched, pieces.uncut, pieces.extraCut]).toEqual([2, [], []]);
    // The converse: one reference line, cut at 94 points, that two painted lines hold together.
    // The clip crosses the last glyph, so it leaves all 11.
    const whole = [line('Kept content', 39, 47, 94.6)];
    const parts = compareTracked(
      [line('Kept', 39, 47, 59.2), line('content', 64.3, 47, 94.6)],
      whole,
      {
        ours: [run(39, 59.2, [36, 36, 94, 50]), run(64.3, 94.6, [36, 36, 94, 50])],
        reference: [run(39, 94.6, [37, 37, 94, 50])],
      }
    );
    expect([parts.matched, parts.uncut]).toEqual([1, []]);
  });

  test('a cut that leaves one character, on the left or as an extra cut', () => {
    // `XK` cut on the left at 46 points, between its glyphs, leaves `K`.
    const left = compareTracked([line('K', 46, 47, 53)], [line('XK', 39, 47, 53)], {
      ours: [run(46, 53, [46, 36, 100, 50])],
      reference: [
        {
          ...run(39, 53, [46, 37, 100, 50]),
          glyphs: [
            [39, 46],
            [46, 53],
          ],
        },
      ],
    });
    expect([left.matched, left.clipped, left.uncut]).toEqual([
      1,
      ['XK @ 47.0 | left clip 46.0'],
      [],
    ]);
    // The reference holds `KX` whole; the painted clip at 46 points leaves `K`: an extra cut.
    const whole = [{ ...run(39, 53, [36, 36, 200, 50]) }];
    const extra = compareTracked([line('K', 39, 47, 46)], [line('KX', 39, 47, 53)], {
      ours: [run(39, 46, [36, 36, 46, 50])],
      reference: whole,
    });
    expect([extra.matched, extra.extraCut]).toEqual([1, ['KX @ 47.0 | right clip 46.0']]);
    // A painted line that ends 14 points before the painted clip edge dropped glyphs that the
    // clip shows: it is no cut, and the reference line stays missing.
    const dropped = compareTracked(
      [line('Kept', 39, 47, 60)],
      [line('Kept content', 39, 47, 97.7)],
      {
        ours: [run(39, 60, [36, 36, 74, 50])],
        reference: [run(39, 94.6, [36, 36, 200, 50])],
      }
    );
    expect([dropped.matched, dropped.extraCut, dropped.missing]).toEqual([0, [], ['Kept content']]);
  });

  test('a clip that leaves one character matches only by the clip rules', () => {
    // The reference clip at 46 points cuts the second character of each text.
    const refRuns1 = [run(39, 53, [37, 37, 46, 50])];
    const cutTo = (whole: string, part: string, clip: TextRun['clip'] = [36, 36, 46, 50]) =>
      compareTracked([line(part, 39, 47, 46.5)], [line(whole, 39, 47, 53)], {
        ours: [run(39, 46.5, clip)],
        reference: refRuns1,
      });
    for (const [whole, part] of [
      ['KX', 'K'],
      ['ID', 'I'],
      ['K1', 'K'],
    ] as const) {
      const score = cutTo(whole, part);
      expect([score.matched, score.clipped, score.uncut]).toEqual([
        1,
        [`${whole} @ 47.0 | right clip 46.0`],
        [],
      ]);
    }
    // A lost clip, a wider clip, or another character still fails as uncut.
    for (const score of [
      cutTo('KX', 'K', 'none'),
      cutTo('KX', 'K', [36, 36, 60, 50]),
      cutTo('KX', 'J'),
    ])
      expect([score.matched, score.clipped, score.uncut]).toEqual([
        0,
        [],
        ['KX @ 47.0 | right clip 46.0'],
      ]);
  });

  test('a cut line whose key drops its leading symbols on one side only still matches', () => {
    // `lineKey` keeps leading symbols only for text with two letters in a row: `[h]:mm` keys as
    // `[h]:mm` but `[h]:` as `h:`. The prefix test compares the visible text of both sides alike.
    // Glyphs of 13 points from 39 points: the clip at 89 points leaves four.
    for (const [whole, part] of [
      ['[h]:mm', '[h]:'],
      ['(a) Note', '(a) N'],
    ] as const) {
      const end = 39 + 13 * whole.replace(/\s+/g, '').length;
      const score = compareTracked([line(part, 39, 47, 93.8)], [line(whole, 39, 47, end)], {
        ours: ourRuns(89),
        reference: [run(39, end, [37, 37, 89, 50])],
      });
      expect([score.matched, score.clipped, score.uncut]).toEqual([
        1,
        [`${whole} @ 47.0 | right clip 89.0`],
        [],
      ]);
    }
  });
});

describe('a clipped line beside a short cell, or on a moved row', () => {
  // `Kept content` is clipped at 89 points; the reference glyphs split its run evenly.
  const step = (94.6 - 39) / 11;
  const glyphs = Array.from(
    { length: 11 },
    (_, i) => [39 + i * step, 39 + (i + 1) * step] as const
  );

  test('a one-character cell beside a clipped line joins nothing', () => {
    // The next cell holds `t`, `7`, or `nt`. `t` has no key, so it does not join `Kept conten`
    // to hold `Kept content`; the clipped pass matches the line.
    for (const next of ['t', '7', 'nt']) {
      const w = next.length * 2.5;
      const s = compareTracked(
        [line('Kept conten', 39, 47, 93.8), line(next, 91, 47, 91 + w)],
        [line('Kept content', 39, 47, 97.7), line(next, 91, 47, 91 + w)],
        {
          ours: [run(39, 93.8, [36, 36, 89, 50]), run(91, 91 + w, [89, 36, 140, 50])],
          reference: [
            { ...run(39, 94.6, [37, 37, 89, 50]), glyphs },
            {
              ...run(91, 91 + w, [89, 37, 140, 50]),
              glyphs:
                next.length === 1
                  ? [[91, 91 + w]]
                  : [
                      [91, 93.5],
                      [93.5, 91 + w],
                    ],
            },
          ],
        }
      );
      expect([s.clipped, s.uncut, s.unknownReference, s.missing, s.extra]).toEqual([
        ['Kept content @ 47.0 | right clip 89.0'],
        [],
        [],
        [],
        [],
      ]);
      expect([s.matched, s.total]).toEqual(next === 't' ? [1, 1] : [2, 2]);
    }
  });

  test('a clipped line on a row moved 1 point is out of place, not lost', () => {
    // Three rows of `Kept content` (clipped) and `Plain text`; the painted page moves each row
    // 1 point down. Both lines report the baseline, and the clip rules pass.
    const page = (shift: number) => {
      const lines: PdfLine[] = [];
      const runs: TextRun[] = [];
      for (let r = 0; r < 3; r += 1) {
        const b = 47 + 15 * r + shift;
        lines.push(line(shift ? 'Kept conten' : 'Kept content', 39, b, shift ? 93.8 : 97.7));
        lines.push(line('Plain text', 150, b, 190));
        runs.push(run(39, shift ? 93.8 : 94.6, [shift ? 36 : 37, b - 10, 89, b + 3], b));
        runs.push(run(150, 190, [148, b - 10, 200, b + 3], b));
      }
      return { lines, runs };
    };
    const ref = page(0);
    const refRuns = ref.runs.map((r, i) => (i % 2 ? r : { ...r, glyphs }));
    const ours = page(1);
    const s = compareTracked(ours.lines, ref.lines, { ours: ours.runs, reference: refRuns });
    expect([s.matched, s.uncut, s.lostUncut, s.missing, s.extra]).toEqual([6, [], [], [], []]);
    expect(s.worstBaseline).toBeCloseTo(1, 6);
    expect(s.clipped).toHaveLength(3);
  });
});
