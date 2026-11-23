// The text runs of a PDF with the clip rectangle of each (`pdf-clips.ts`). The line scores of
// clipped lines: `printed-lines.test.ts`.

import { describe, expect, test } from 'bun:test';
import { join } from 'node:path';
import { OPS } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { compareLines } from './compare.ts';
import {
  clipOfItem,
  intersect,
  rectangleOf,
  splitItemByRuns,
  textRuns,
  type TextRun,
} from './pdf-clips.ts';
import { pdfItems } from './pdf-lines.ts';
import type { PdfLine } from './poppler.ts';
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

describe('the clip rectangles of a PDF', () => {
  const page = (x: number, y: number): [number, number] => [x, 100 - y];
  /** A rectangle path of pdf.js `constructPath` data. */
  const rect = (x0: number, y0: number, x1: number, y1: number) => [
    OPS.endPath,
    [new Float32Array([0, x0, y0, 1, x1, y0, 1, x1, y1, 1, x0, y1, 4])],
    new Float32Array([x0, y0, x1, y1]),
  ];
  const glyphs = (text: string) => [[...text].map((c) => ({ unicode: c, width: 500 }))];
  const show = (x: number, text: string): [number, unknown][] => [
    [OPS.beginText, []],
    [OPS.setFont, ['F1', 10]],
    [OPS.setTextMatrix, [new Float32Array([1, 0, 0, 1, x, 50])]],
    [OPS.showText, glyphs(text)],
    [OPS.endText, []],
  ];
  const allOf = (ops: [number, unknown][]) =>
    textRuns(
      ops.map(([f]) => f),
      ops.map(([, a]) => a),
      page
    );
  /** The runs without their glyph advances, which only the first test checks. */
  const runsOf = (ops: [number, unknown][]) =>
    allOf(ops).map(({ glyphs: _, text: _text, state: _state, ...r }) => r);

  test('a text run takes the clip that holds it, and restore ends the clip', () => {
    const runs = runsOf([
      [OPS.save, []],
      [OPS.clip, null],
      [OPS.constructPath, rect(10, 40, 30, 60)],
      ...show(12, 'ab'),
      [OPS.restore, []],
      ...show(40, 'cd'),
    ]);
    expect(runs).toEqual([
      { x0: 12, x1: 22, baseline: 50, clip: [10, 40, 30, 60] },
      { x0: 40, x1: 50, baseline: 50, clip: 'none' },
    ]);
    // Each glyph that is not white space records its advance extent.
    expect(allOf([...show(12, 'a b')]).map((r) => r.glyphs)).toEqual([
      [
        [12, 17],
        [22, 27],
      ],
    ]);
  });

  test('nested clips intersect, and the current matrix maps the path', () => {
    const runs = runsOf([
      [OPS.eoClip, null],
      [OPS.constructPath, rect(0, 0, 50, 100)],
      [OPS.transform, [2, 0, 0, 1, 0, 0]],
      [OPS.clip, null],
      [OPS.constructPath, rect(10, 40, 30, 60)],
      ...show(12, 'ab'),
    ]);
    // The second rectangle spans 20 to 60 points under the matrix; the first ends at 50.
    expect(runs[0]!.clip).toEqual([20, 40, 50, 60]);
  });

  test('a clip path that is not one axis-aligned rectangle makes the clip unknown', () => {
    const triangle = [
      OPS.endPath,
      [new Float32Array([0, 10, 40, 1, 30, 40, 1, 20, 60, 4])],
      new Float32Array([10, 40, 30, 60]),
    ];
    expect(
      runsOf([[OPS.clip, null], [OPS.constructPath, triangle], ...show(12, 'ab')])[0]!.clip
    ).toBe('non-rectangular');
    const rotated = runsOf([
      [OPS.transform, [0.8, 0.6, -0.6, 0.8, 0, 0]],
      [OPS.clip, null],
      [OPS.constructPath, rect(10, 40, 30, 60)],
      ...show(12, 'ab'),
    ]);
    expect(rotated[0]!.clip).toBe('non-rectangular');
    // A filled or stroked path, such as a border or a cell fill, is not a clip.
    const border = runsOf([[OPS.constructPath, rect(10, 40, 30, 60)], ...show(12, 'ab')]);
    expect(border[0]!.clip).toBe('none');
  });

  test('the text operators place each run', () => {
    const at = (ops: [number, unknown][]) =>
      runsOf([[OPS.beginText, []], [OPS.setFont, ['F1', 10]], ...ops, [OPS.endText, []]]).map(
        ({ x0, x1, baseline }) => [x0, x1, baseline]
      );
    const matrix: [number, unknown] = [OPS.setTextMatrix, [new Float32Array([1, 0, 0, 1, 10, 50])]];
    // `Td` moves the start of the line; a glyph is 5 points wide at 10 points.
    expect(
      at([
        matrix,
        [OPS.showText, glyphs('ab')],
        [OPS.moveText, [5, -10]],
        [OPS.showText, glyphs('cd')],
      ])
    ).toEqual([
      [10, 20, 50],
      [15, 25, 60],
    ]);
    // `TD` sets the leading that `T*` then moves by.
    expect(
      at([
        matrix,
        [OPS.setLeadingMoveText, [0, -12]],
        [OPS.showText, glyphs('ab')],
        [OPS.nextLine, []],
        [OPS.showText, glyphs('cd')],
      ])
    ).toEqual([
      [10, 20, 62],
      [10, 20, 74],
    ]);
    // `TL` then `T*`.
    expect(
      at([matrix, [OPS.setLeading, [8]], [OPS.nextLine, []], [OPS.showText, glyphs('ab')]])
    ).toEqual([[10, 20, 58]]);
    // `Tz` scales the advance, and `Tc` adds to each glyph but the last: a run ends at its last
    // glyph.
    expect(at([matrix, [OPS.setHScale, [50]], [OPS.showText, glyphs('ab')]])).toEqual([
      [10, 15, 50],
    ]);
    expect(at([matrix, [OPS.setCharSpacing, [1]], [OPS.showText, glyphs('ab')]])).toEqual([
      [10, 21, 50],
    ]);
    // A number of `TJ` moves the next glyph by its thousandths of the font size; one after the
    // last glyph does not move the end of the run.
    const spaced = [[{ unicode: 'a', width: 500 }, -1000, { unicode: 'b', width: 500 }, 500]];
    expect(at([matrix, [OPS.showSpacedText, spaced]])).toEqual([[10, 30, 50]]);
  });

  test('a form paints as a save and restore, under its matrix and clipped to its box', () => {
    const runs = runsOf([
      [
        OPS.paintFormXObjectBegin,
        [
          [2, 0, 0, 2, 0, 0],
          [0, 0, 20, 30],
        ],
      ],
      // A clip and a font size inside the form end with it.
      [OPS.clip, null],
      [OPS.constructPath, rect(0, 0, 15, 30)],
      ...show(5, 'ab'),
      [OPS.paintFormXObjectEnd, []],
      ...show(40, 'cd'),
    ]);
    expect(runs).toEqual([
      { x0: 10, x1: 30, baseline: 0, clip: [0, 40, 30, 100] },
      { x0: 40, x1: 50, baseline: 50, clip: 'none' },
    ]);
  });

  test('a group form takes its bounding box from its group', () => {
    const runs = runsOf([
      [OPS.beginGroup, [{ matrix: [2, 0, 0, 2, 0, 0], bbox: [0, 0, 20, 30] }]],
      [OPS.paintFormXObjectBegin, [[2, 0, 0, 2, 0, 0], null]],
      ...show(5, 'ab'),
      [OPS.paintFormXObjectEnd, []],
      [OPS.endGroup, [{}]],
      ...show(40, 'cd'),
    ]);
    expect(runs.map((r) => r.clip)).toEqual([[0, 40, 40, 100], 'none']);
  });

  test('a run that a text rise or a text clip places has an unknown clip', () => {
    const clipOf = (ops: [number, unknown][]) => runsOf(ops).map((r) => r.clip);
    expect(clipOf([[OPS.setTextRise, [2]], ...show(12, 'ab')])).toEqual(['non-rectangular']);
    expect(clipOf([[OPS.setTextRise, [0]], ...show(12, 'ab')])).toEqual(['none']);
    // Render mode 7 adds the text to the clip path: the text after it clips with an unknown path.
    const textClip = clipOf([
      [OPS.setTextRenderingMode, [7]],
      ...show(12, 'ab'),
      ...show(40, 'cd'),
    ]);
    expect(textClip).toEqual(['non-rectangular', 'non-rectangular']);
    expect(clipOf([[OPS.setTextRenderingMode, [3]], ...show(12, 'ab')])).toEqual(['none']);
  });

  test('restore brings back the text state of its save', () => {
    // A character spacing that a save and restore hold does not reach the run after them.
    const spaced = runsOf([
      [OPS.save, []],
      [OPS.setCharSpacing, [5]],
      [OPS.restore, []],
      ...show(10, 'ab'),
    ]);
    expect([spaced[0]!.x0, spaced[0]!.x1]).toEqual([10, 20]);
    // Nor a text rise: the run keeps its rectangular clip.
    const risen = runsOf([
      [OPS.clip, null],
      [OPS.constructPath, rect(5, 40, 30, 60)],
      [OPS.save, []],
      [OPS.setTextRise, [3]],
      [OPS.restore, []],
      ...show(10, 'ab'),
    ]);
    expect(risen[0]!.clip).toEqual([5, 40, 30, 60]);
    // Nor a render mode that adds text to the clip path, nor a font size.
    const plain = runsOf([
      [OPS.save, []],
      [OPS.setTextRenderingMode, [7]],
      [OPS.restore, []],
      [OPS.beginText, []],
      [OPS.setTextMatrix, [new Float32Array([1, 0, 0, 1, 10, 50])]],
      [OPS.save, []],
      [OPS.setFont, ['F1', 20]],
      [OPS.restore, []],
      [OPS.showText, glyphs('ab')],
      [OPS.endText, []],
    ]);
    expect(plain).toEqual([{ x0: 10, x1: 10, baseline: 50, clip: 'none' }]);
    // A state set inside the save holds until the restore.
    const inside = runsOf([[OPS.save, []], [OPS.setCharSpacing, [5]], ...show(10, 'ab')]);
    expect(inside[0]!.x1).toBe(25);
  });

  test('a run ends at its last glyph that is not a space, in any font matrix', () => {
    // `Tf 1` with a text matrix that scales by 11, `Tc`, and a trailing space glyph: the run
    // ends at the end of `t`, not after the space or the spacing.
    const ops: [number, unknown][] = [
      [OPS.beginText, []],
      [OPS.setFont, ['F1', 1]],
      [OPS.setCharSpacing, [0.1]],
      [OPS.setWordSpacing, [0.2]],
      [OPS.setTextMatrix, [new Float32Array([11, 0, 0, 11, 39, 53])]],
      [
        OPS.showText,
        [
          [
            { unicode: 'K', width: 500 },
            { unicode: 't', width: 500 },
            { unicode: ' ', width: 250, isSpace: true },
          ],
        ],
      ],
      [OPS.endText, []],
    ];
    const [end] = runsOf(ops);
    // 11 x (0.5 + 0.1 + 0.5) points after 39.
    expect(end!.x1).toBeCloseTo(51.1, 9);
    // The reference line runs past its clip at 45 points; the painted line, cut nowhere, runs to
    // 51.1 points too, so the lost cut is reported.
    const runs = runsOf([[OPS.clip, null], [OPS.constructPath, rect(37, 40, 45, 60)], ...ops]);
    const kt = line('Kt', 39, 47, 51.1);
    const score = compareLines([kt], [kt], {
      ours: [{ ...runs[0]!, baseline: 47, clip: 'none' }],
      reference: [{ ...runs[0]!, baseline: 47 }],
    });
    expect([score.matched, score.uncut]).toEqual([1, ['Kt @ 47.0 | right clip 45.0']]);
  });

  test('a run starts at its first glyph that is not a space', () => {
    const lead = [
      [
        { unicode: ' ', width: 250, isSpace: true },
        { unicode: ' ', width: 250, isSpace: true },
        { unicode: 'a', width: 500 },
      ],
    ];
    const runs = runsOf([
      [OPS.beginText, []],
      [OPS.setFont, ['F1', 10]],
      [OPS.setCharSpacing, [1]],
      [OPS.setTextMatrix, [new Float32Array([1, 0, 0, 1, 10, 50])]],
      [OPS.showText, lead],
      [OPS.endText, []],
    ]);
    // Two spaces of 2.5 points, each with 1 point of character spacing, come before the `a`.
    expect([runs[0]!.x0, runs[0]!.x1]).toEqual([17, 22]);
  });

  test('word spacing applies to each space glyph', () => {
    const words = [
      [
        { unicode: 'a', width: 500 },
        { unicode: ' ', width: 250, isSpace: true },
        { unicode: 'b', width: 500 },
      ],
    ];
    const runs = runsOf([
      [OPS.setWordSpacing, [4]],
      [OPS.beginText, []],
      [OPS.setFont, ['F1', 10]],
      [OPS.setTextMatrix, [new Float32Array([1, 0, 0, 1, 10, 50])]],
      [OPS.showText, words],
      [OPS.endText, []],
    ]);
    // 5 + 2.5 + 4 + 5 points.
    expect(runs[0]!.x1).toBe(26.5);
  });

  test('an item takes the clip of the run that starts last at its start', () => {
    const cellA = [37, 37, 89, 50] as const;
    const cellB = [89, 37, 143, 50] as const;
    // The run of the overflow from 39 to 97.7 points and the run of the next cell from 91 points.
    const runs = [run(39, 97.7, cellA), run(91, 111, cellB)];
    expect(clipOfItem(runs, 39, 47, 0.25)).toEqual(cellA);
    expect(clipOfItem(runs, 91, 47, 0.25)).toEqual(cellB);
    expect(clipOfItem(runs, 95, 47, 0.25)).toEqual(cellB);
    // Another baseline, or a start that no run holds, has no clip.
    expect(clipOfItem(runs, 39, 60, 0.25)).toBeUndefined();
    expect(clipOfItem(runs, 120, 47, 0.25)).toBeUndefined();
  });

  test('a run that a space glyph splits records the ink of each word', () => {
    expect(runsOf(show(10, 'ab cd'))).toEqual([
      {
        x0: 10,
        x1: 35,
        baseline: 50,
        clip: 'none',
        words: [
          [10, 20],
          [25, 35],
        ],
      },
    ]);
    // A run of one word records none.
    expect(runsOf(show(10, 'ab'))[0]!.words).toBeUndefined();
    // A glyph with an empty Unicode value is a glyph, not a space: it neither ends a word nor
    // the run.
    const empty = runsOf([
      [OPS.beginText, []],
      [OPS.setFont, ['F1', 10]],
      [OPS.setTextMatrix, [new Float32Array([1, 0, 0, 1, 10, 50])]],
      [
        OPS.showText,
        [
          [
            { unicode: 'a', width: 500 },
            { unicode: '', width: 500 },
            { unicode: ' ', width: 500 },
            { unicode: '', width: 500 },
          ],
        ],
      ],
    ]);
    expect(empty).toEqual([
      {
        x0: 10,
        x1: 30,
        baseline: 50,
        clip: 'none',
        words: [
          [10, 20],
          [25, 30],
        ],
      },
    ]);
  });

  test('a clip with no path makes the clip unknown', () => {
    for (const empty of [
      [OPS.endPath, [null], null],
      [OPS.endPath, [], null],
      [OPS.endPath, null, null],
    ])
      expect(
        runsOf([[OPS.clip, null], [OPS.constructPath, empty], ...show(12, 'ab')])[0]!.clip
      ).toBe('non-rectangular');
  });

  test('the reference PDFs carry the cell clips of their clipped lines', async () => {
    const preserved = await pdfItems(join(FIXTURE_DIR, 'preserved-parts.pdf'));
    const kept = preserved.runs.find((r) => Math.abs(r.baseline - 47) < 0.25 && r.x0 < 40);
    expect(kept?.clip).toEqual([37, 37, 89, 50]);
    const formats = await pdfItems(join(FIXTURE_DIR, 'number-formats.pdf'));
    const code = formats.runs.find((r) => Math.abs(r.baseline - 327) < 0.25 && r.x0 > 300);
    expect(code?.clip).toEqual([307, 317, 414, 330]);
  });
});

describe('the rectangle of a clip path', () => {
  // pdf.js `constructPath` data: 0 is a move, 1 a line, 4 closes the path.
  const path = (...points: [number, number][]): number[] =>
    points.flatMap(([x, y], i) => [i === 0 ? 0 : 1, x, y]).concat([4]);
  const identity = [1, 0, 0, 1, 0, 0] as const;
  const page = (x: number, y: number): [number, number] => [x, y];
  const read = (data: number[]) => rectangleOf(data, identity, page);

  test('an axis-aligned rectangle reads as its box', () => {
    expect(read(path([10, 20], [50, 20], [50, 40], [10, 40]))).toEqual([10, 20, 50, 40]);
    // Closed back to its first point, or drawn the other way round.
    expect(read(path([10, 20], [10, 40], [50, 40], [50, 20], [10, 20]))).toEqual([10, 20, 50, 40]);
  });

  test('a degenerate path reads as unknown', () => {
    // A line drawn back and forth: every edge is horizontal, and the box has no height.
    expect(read(path([10, 20], [50, 20], [10, 20], [50, 20]))).toBeUndefined();
    // A point, and an edge of no length.
    expect(read(path([10, 20], [10, 20], [10, 20], [10, 20]))).toBeUndefined();
    expect(read(path([10, 20], [50, 20], [50, 20], [10, 20]))).toBeUndefined();
    // Two horizontal edges in a row.
    expect(read(path([10, 20], [30, 20], [50, 20], [50, 40]))).toBeUndefined();
    // A slanted edge.
    expect(read(path([10, 20], [50, 25], [50, 40], [10, 40]))).toBeUndefined();
  });
});

describe('the clip of two clips in a row', () => {
  test('overlapping rectangles give their overlap, and disjoint ones read as unknown', () => {
    expect(intersect([10, 10, 50, 50], [30, 20, 80, 40])).toEqual([30, 20, 50, 40]);
    expect(intersect('none', [10, 10, 50, 50])).toEqual([10, 10, 50, 50]);
    // Side by side, and one above the other: no point is inside both.
    expect(intersect([10, 10, 50, 50], [60, 10, 90, 50])).toBe('non-rectangular');
    expect(intersect([10, 10, 50, 50], [10, 60, 50, 90])).toBe('non-rectangular');
    // Rectangles that share only an edge, or an overlap thinner than 0.01 points, leave no area.
    expect(intersect([10, 10, 50, 50], [50, 10, 90, 50])).toBe('non-rectangular');
    expect(intersect([10, 10, 50, 50], [10, 49.995, 50, 90])).toBe('non-rectangular');
  });
});

describe('the parts of a text item that covers runs of two clips', () => {
  const left: [number, number, number, number] = [37, 37, 108, 50];
  const right: [number, number, number, number] = [37, 37, 325, 107];
  const run = (x0: number, x1: number, text: string | undefined, clip: TextRun['clip']) =>
    ({ x0, x1, baseline: 47, clip, ...(text === undefined ? {} : { text }) }) as TextRun;

  test('splits where the clip changes, and keeps the text of each part', () => {
    const runs = [
      run(39, 170.5, 'This text is longer than its cell', left),
      run(172, 178, '2', right),
    ];
    const item = { text: 'This text is longer than its cell 2', x: 39, width: 138.6 };
    expect(splitItemByRuns(runs, item, 47, 1)).toEqual([
      { text: 'This text is longer than its cell', x: 39, width: 131.5 },
      { text: '2', x: 172, width: 6 },
    ]);
  });

  test('splits where a restore ends the graphics state of a run, with one clip', () => {
    const runs = [
      { ...run(39, 60, '84', 'none'), state: 1 },
      { ...run(62, 110, 'Shared text', 'none'), state: 2 },
    ];
    const item = { text: '84 Shared text', x: 39, width: 71 };
    expect(splitItemByRuns(runs, item, 47, 1).map((p) => p.text)).toEqual(['84', 'Shared text']);
  });

  test('keeps an item whole when its runs share one clip', () => {
    const runs = [run(39, 60, 'Hello', left), run(62, 90, 'world', left)];
    const item = { text: 'Hello world', x: 39, width: 51 };
    expect(splitItemByRuns(runs, item, 47, 1)).toEqual([item]);
  });

  test('keeps an item whole when its runs do not spell its text, or have no text', () => {
    const item = { text: 'Hello world', x: 39, width: 51 };
    expect(
      splitItemByRuns([run(39, 60, 'Hello', left), run(62, 90, 'there', right)], item, 47, 1)
    ).toEqual([item]);
    expect(
      splitItemByRuns([run(39, 60, undefined, left), run(62, 90, 'world', right)], item, 47, 1)
    ).toEqual([item]);
  });
});
