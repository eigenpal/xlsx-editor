// The joins of the line matching of the fidelity harness: several lines of one page that join to
// hold one line of the other page (`compareLines` in `compare.ts`). The other line matching tests
// are in `harness.test.ts`.

import { describe, expect, test } from 'bun:test';
import { compareLines } from './compare.ts';
import type { PdfLine } from './poppler.ts';
import * as thresholds from './thresholds.ts';

const line = (text: string, xMin: number, baseline: number, xMax = xMin + 50): PdfLine => ({
  text,
  xMin,
  xMax,
  baseline,
  size: 11,
});

describe('line joins', () => {
  test('reference pieces join by their own slack, never by the width of a painted piece', () => {
    const limits = thresholds.THRESHOLDS as { maxLineShiftPt: number };
    const was = limits.maxLineShiftPt;
    // Without painted items, the gap of the reference pieces is at most 1 em (11 points) plus
    // `joinGapPt`. A tighter line position threshold does not change that.
    const plain = (gap: number) =>
      compareLines(
        [line('ABCD', 0, 100, 20 + gap)],
        [line('AB', 0, 100, 10), line('CD', 10 + gap, 100, 20 + gap)]
      ).matched;
    try {
      limits.maxLineShiftPt = 0.2;
      expect([plain(11.4), plain(11.6)]).toEqual([2, 0]);
    } finally {
      limits.maxLineShiftPt = was;
    }
    // With painted items, the painted gap is at most 1 em, and the reference gap is at most that
    // plus `joinGapPt` at each end. Each piece 0.3 points off, with a painted gap of exactly 1 em,
    // joins and scores 0.3.
    const items = (a: [number, number], b: [number, number]): PdfLine => ({
      ...line('AB CD', a[0], 100, b[1]),
      items: [
        { at: 0, length: 2, xMin: a[0], xMax: a[1] },
        { at: 2, length: 2, xMin: b[0], xMax: b[1] },
      ],
    });
    const off = compareLines(
      [items([0.3, 10.3], [21.3, 31.3])],
      [line('AB', 0, 100, 10), line('CD', 21.6, 100, 31.6)]
    );
    expect([off.matched, off.missing, off.extra]).toEqual([2, [], []]);
    expect(off.worstShift).toBeCloseTo(0.3, 6);
    // A painted piece that is wider than its reference piece does not close a wide reference gap.
    for (const gap of [30, 100, 300]) {
      const wide = compareLines(
        [items([0, 29 + gap], [30 + gap, 60 + gap])],
        [line('AB', 0, 100, 30), line('CD', 30 + gap, 100, 60 + gap)]
      );
      expect([wide.matched, wide.missing, wide.extra]).toEqual([0, ['AB', 'CD'], ['AB CD']]);
    }
  });

  test('painted pieces that join hold the reference characters at their own places', () => {
    const reference: PdfLine = {
      ...line('Hello world', 0, 100, 52),
      size: 10,
      items: [
        { at: 0, length: 5, xMin: 0, xMax: 25 },
        { at: 5, length: 5, xMin: 27, xMax: 52 },
      ],
    };
    const ours = (dx: number) => [
      { ...line('Hello', 0, 100, 25), size: 10 },
      { ...line('world', 27 + dx, 100, 52 + dx), size: 10 },
    ];
    const score = (dx: number, ref: PdfLine) => {
      const s = compareLines(ours(dx), [ref]);
      return [s.matched, s.worstShift];
    };
    expect(score(0, reference)).toEqual([1, 0]);
    // A painted `world` 30 or 190 points to the right is no part of the line.
    for (const dx of [30, 190]) {
      expect(score(dx, reference)).toEqual([0, 0]);
      expect(score(dx, { ...reference, items: undefined })).toEqual([0, 0]);
    }
    // A painted `world` 0.4 points to the right scores its own distance.
    const near = compareLines(ours(0.4), [reference]);
    expect(near.matched).toBe(1);
    expect(near.worstShift).toBeCloseTo(0.4, 6);
  });

  test('a join decides by the grouping rule of each page, and scores each piece as one line', () => {
    // `Name 100` is one line of one page, with items at 0..30 and 38..53. The other page draws
    // `Name` 1.5 points narrower or wider toward the gap; its start is exact.
    const split = (end: number) => [
      { at: 0, length: 4, xMin: 0, xMax: end },
      { at: 4, length: 3, xMin: 38, xMax: 53 },
    ];
    const whole = (end: number): PdfLine => ({
      ...line('Name 100', 0, 100, 53),
      items: split(end),
    });
    for (const end of [28.5, 31.5]) {
      const forward = compareLines(
        [line('Name', 0, 100, end), line('100', 38, 100, 53)],
        [whole(30)]
      );
      const reverse = compareLines(
        [whole(end)],
        [line('Name', 0, 100, 30), line('100', 38, 100, 53)]
      );
      const single = compareLines([line('Name', 0, 100, end)], [line('Name', 0, 100, 30)]);
      expect([forward.matched, forward.worstShift]).toEqual([1, single.worstShift]);
      expect([reverse.matched, reverse.worstShift]).toEqual([2, single.worstShift]);
    }
    // A gap of the whole line of exactly 1 em (11 points), as PDF coordinates sum it, joins.
    const exact: PdfLine = {
      ...line('AB CD', 0, 100, 32.6),
      items: [
        { at: 0, length: 2, xMin: 0, xMax: 10.6 },
        { at: 2, length: 2, xMin: 21.6, xMax: 32.6 },
      ],
    };
    expect(21.6 - 10.6).toBeGreaterThan(11);
    expect(
      compareLines([exact], [line('AB', 0, 100, 10.6), line('CD', 21.6, 100, 32.6)]).matched
    ).toBe(2);
    // A wider gap of the whole line is two lines there: no join.
    const wide: PdfLine = {
      ...line('AB CD', 0, 100, 31.5),
      items: [
        { at: 0, length: 2, xMin: 0, xMax: 10 },
        { at: 2, length: 2, xMin: 21.5, xMax: 31.5 },
      ],
    };
    expect(
      compareLines([wide], [line('AB', 0, 100, 10), line('CD', 21.5, 100, 31.5)]).matched
    ).toBe(0);
  });

  test('the join gap limit of approximation "Join gap limit" is 1 em plus 1 point', () => {
    // `AB CD` is one line of one page, size 10, with a gap of exactly 1 em. The other page draws
    // `AB` narrower toward the gap, so the gap of the pieces is 10 points plus `narrower`.
    const whole: PdfLine = {
      ...line('AB CD', 0, 100, 30),
      size: 10,
      items: [
        { at: 0, length: 2, xMin: 0, xMax: 10 },
        { at: 2, length: 2, xMin: 20, xMax: 30 },
      ],
    };
    const join = (narrower: number) => {
      const pieces = [
        { ...line('AB', 0, 100, 10 - narrower), size: 10 },
        { ...line('CD', 20, 100, 30), size: 10 },
      ];
      const s = compareLines(pieces, [whole]);
      return [s.matched, s.missing, s.extra];
    };
    expect(join(0.99)).toEqual([1, [], []]);
    expect(join(1.01)).toEqual([0, ['AB CD'], ['AB', 'CD']]);
    // The same geometry as one line passes.
    expect(compareLines([line('AB', 0, 100, 8.99)], [line('AB', 0, 100, 10)]).worstShift).toBe(0);
  });

  test('a group of pieces pairs within the baseline threshold, as one line does', () => {
    const items = [
      { at: 0, length: 4, xMin: 0, xMax: 30 },
      { at: 4, length: 3, xMin: 38, xMax: 53 },
    ];
    const whole = (b: number): PdfLine => ({ ...line('Name 100', 0, b, 53), items });
    const pieces = (b: number) => [line('Name', 0, b, 30), line('100', 38, b, 53)];
    for (const d of [0.2, 0.3, 0.45]) {
      const forward = compareLines(pieces(100 + d), [whole(100)]);
      const reverse = compareLines([whole(100 + d)], pieces(100));
      expect([forward.matched, reverse.matched]).toEqual([1, 2]);
      expect(forward.worstBaseline).toBeCloseTo(d, 6);
      expect(reverse.worstBaseline).toBeCloseTo(d, 6);
    }
    // Past the threshold, the pieces do not join: the line is missing, and the pieces are extra.
    const far = compareLines(pieces(100.6), [whole(100)]);
    expect([far.matched, far.missing, far.extra]).toEqual([0, ['Name 100'], ['Name', '100']]);
  });

  test('each piece of a join lies within the baseline threshold of the line, with no rule between', () => {
    const items = [
      { at: 0, length: 4, xMin: 0, xMax: 30 },
      { at: 4, length: 3, xMin: 38, xMax: 53 },
    ];
    const whole = (b: number): PdfLine => ({ ...line('Name 100', 0, b, 53), items });
    const pieces = (a: number, b: number) => [line('Name', 0, a, 30), line('100', 38, b, 53)];
    // Pieces at 99.8 and 100.2 hold a line at 100, and the reverse; the baseline scores 0.2.
    const forward = compareLines(pieces(99.8, 100.2), [whole(100)]);
    const reverse = compareLines([whole(100)], pieces(99.8, 100.2));
    expect([forward.matched, reverse.matched]).toEqual([1, 2]);
    expect(forward.worstBaseline).toBeCloseTo(0.2, 6);
    expect(reverse.worstBaseline).toBeCloseTo(0.2, 6);
    // Pieces at 100 and 100.3 hold a line at 100.15.
    expect(compareLines(pieces(100, 100.3), [whole(100.15)]).matched).toBe(1);
    expect(compareLines([whole(100.15)], pieces(100, 100.3)).matched).toBe(2);
    // A piece 0.6 points off the line does not join: the line is missing, and the pieces are extra.
    const far = compareLines(pieces(100, 100.6), [whole(100)]);
    expect([far.matched, far.missing, far.extra]).toEqual([0, ['Name 100'], ['Name', '100']]);
    const farBack = compareLines([whole(100)], pieces(100, 100.6));
    expect([farBack.matched, farBack.missing, farBack.extra]).toEqual([
      0,
      ['Name', '100'],
      ['Name 100'],
    ]);
    // Of three pieces, the last one lies furthest from the line: its baseline is scored.
    const three: PdfLine = {
      ...line('AB CD EF', 0, 100, 40),
      size: 10,
      items: [
        { at: 0, length: 2, xMin: 0, xMax: 10 },
        { at: 2, length: 2, xMin: 15, xMax: 25 },
        { at: 4, length: 2, xMin: 30, xMax: 40 },
      ],
    };
    const threePieces = (b: number[]) =>
      [
        ['AB', 0, 10],
        ['CD', 15, 25],
        ['EF', 30, 40],
      ].map(([t, x0, x1], i) => ({
        ...line(t as string, x0 as number, b[i]!, x1 as number),
        size: 10,
      }));
    const worst = compareLines(threePieces([100.1, 99.9, 100.4]), [three]);
    expect(worst.matched).toBe(1);
    expect(worst.worstBaseline).toBeCloseTo(0.4, 6);
    // CD 3 points off its reference item: each piece is scored, not the width of the group.
    const moved = [
      { ...line('AB', 0, 100, 10), size: 10 },
      { ...line('CD', 18, 100, 28), size: 10 },
      { ...line('EF', 30, 100, 40), size: 10 },
    ];
    const shifted = compareLines(moved, [three]);
    expect(shifted.matched).toBe(1);
    expect(shifted.worstShift).toBeCloseTo(3, 6);
  });

  test('the em of a join is the size of the whole line, not of its pieces', () => {
    // A line of size 10 without items, and pieces of size 20 that are 15 points apart: past
    // 1 em of the line plus `joinGapPt`, though within 1 em of the pieces.
    const whole = { ...line('ABCD', 0, 100, 35), size: 10 };
    const pieces = [
      { ...line('AB', 0, 100, 10), size: 20 },
      { ...line('CD', 25, 100, 35), size: 20 },
    ];
    const s = compareLines(pieces, [whole]);
    expect([s.matched, s.missing, s.extra]).toEqual([0, ['ABCD'], ['AB', 'CD']]);
  });

  test('a painted piece scores each of its edges that lies on a reference item edge', () => {
    // The reference item `CD EF` spans two painted pieces. `CD` starts on its start, 7.5 points
    // left of it; `EF` ends on its end. The width of the group matches.
    const reference: PdfLine = {
      ...line('AB CD EF', 0, 100, 60.5),
      size: 10,
      items: [
        { at: 0, length: 2, xMin: 0, xMax: 20 },
        { at: 2, length: 4, xMin: 30, xMax: 60.5 },
      ],
    };
    const pieces = [
      { ...line('AB', 0, 100, 12.3), size: 10 },
      { ...line('CD', 22.5, 100, 36.5), size: 10 },
      { ...line('EF', 46.7, 100, 60.5), size: 10 },
    ];
    const s = compareLines(pieces, [reference]);
    expect(s.matched).toBe(1);
    expect(s.worstShift).toBeCloseTo(7.5, 6);
  });

  test('the characters of each painted piece start where the pieces before it end', () => {
    // The reference items `AB` and `CDEF`; the painted pieces `AB`, `CD`, and `EF`. The gap after
    // `CD` lies inside the reference item `CDEF`, so it is no gap of the reference line: 10.8
    // points are past 1 em plus `joinGapPt`.
    const reference: PdfLine = {
      ...line('AB CDEF', 0, 100, 60.8),
      size: 10,
      items: [
        { at: 0, length: 2, xMin: 0, xMax: 20 },
        { at: 2, length: 4, xMin: 30, xMax: 60.8 },
      ],
    };
    const pieces = [
      { ...line('AB', 0, 100, 20), size: 10 },
      { ...line('CD', 30, 100, 40), size: 10 },
      { ...line('EF', 50.8, 100, 60.8), size: 10 },
    ];
    const s = compareLines(pieces, [reference]);
    expect([s.matched, s.missing]).toEqual([0, ['AB CDEF']]);
  });

  test('the baseline of the worst reference piece of a join is scored', () => {
    const painted: PdfLine = {
      ...line('ABCD', 0, 100, 40),
      items: [
        { at: 0, length: 2, xMin: 0, xMax: 20 },
        { at: 2, length: 2, xMin: 25, xMax: 40 },
      ],
    };
    const s = compareLines([painted], [line('AB', 0, 99.9, 20), line('CD', 25, 100.4, 40)]);
    expect(s.matched).toBe(2);
    expect(s.worstBaseline).toBeCloseTo(0.4, 6);
  });

  test('a join of pieces on their own item edges scores those edges, not the group edges', () => {
    // Reference items `AB` at 0..10 and `CD` at 20..30; painted `AB` at 3..10 and `CD` at 20..33.
    // As two lines, each scores its nearer edge: 0. Joined, in either direction, it is the same.
    const reference: PdfLine = {
      ...line('AB CD', 0, 100, 30),
      items: [
        { at: 0, length: 2, xMin: 0, xMax: 10 },
        { at: 2, length: 2, xMin: 20, xMax: 30 },
      ],
    };
    const painted: PdfLine = {
      ...line('AB CD', 3, 100, 33),
      items: [
        { at: 0, length: 2, xMin: 3, xMax: 10 },
        { at: 2, length: 2, xMin: 20, xMax: 33 },
      ],
    };
    const ab = (x: number) => line('AB', x, 100, 10);
    const cd = (x: number) => line('CD', 20, 100, x);
    const score = (ours: PdfLine[], ref: PdfLine[]) => {
      const s = compareLines(ours, ref);
      return [s.matched, s.worstShift];
    };
    expect(score([ab(3), cd(33)], [ab(0), cd(30)])).toEqual([2, 0]);
    expect(score([ab(3), cd(33)], [reference])).toEqual([1, 0]);
    expect(score([painted], [ab(0), cd(30)])).toEqual([2, 0]);
  });

  test('a piece with no item edge scores the width of the group', () => {
    // Reference items `123` at 0..30 and `456` at 34..64. Painted `1` starts 2 points right,
    // `6` ends 2 points left, and `2345` has no edge on an item edge: the group is 4 points
    // narrower than the line. One digit is a line with a key, so it joins.
    const reference: PdfLine = {
      ...line('123456', 0, 100, 64),
      size: 10,
      items: [
        { at: 0, length: 3, xMin: 0, xMax: 30 },
        { at: 3, length: 3, xMin: 34, xMax: 64 },
      ],
    };
    const pieces = [
      { ...line('1', 2, 100, 10), size: 10 },
      { ...line('2345', 11, 100, 50), size: 10 },
      { ...line('6', 52, 100, 62), size: 10 },
    ];
    const s = compareLines(pieces, [reference]);
    expect(s.matched).toBe(1);
    expect(s.worstShift).toBeCloseTo(4, 6);
  });

  test('an inner end on a boundary of reference items keeps the clip rule of that item', () => {
    // Reference items `123` at 10..30 and `456` at 32..60; painted `12`, `3`, and `456`. The end
    // of `12` lies inside `123`: a painted split. The right end of `3` is the end of `123`, so
    // the clip rule checks it against `123`, and no painted split names it.
    const reference: PdfLine = {
      ...line('123 456', 10, 100, 60),
      items: [
        { at: 0, length: 3, xMin: 10, xMax: 30 },
        { at: 3, length: 3, xMin: 32, xMax: 60 },
      ],
    };
    const s = compareLines(
      [line('12', 10, 100, 22), line('3', 24, 100, 30), line('456', 32, 100, 60)],
      [reference],
      { ours: [], reference: [] }
    );
    expect(s.matched).toBe(1);
    expect(
      s.unknownReference.filter((k) => k.endsWith('painted split inside a reference item'))
    ).toEqual([
      '12 @ 100.0 | right painted split inside a reference item',
      '3 @ 100.0 | left painted split inside a reference item',
    ]);
    expect(s.unknownReference).toContain('123 @ 100.0 | right reference clip unknown');
  });

  test('a group shifted with no width change scores the edges of a line without items only', () => {
    // Reference items `AB`, `3DE6`, and `GH`. The painted `AB` starts 3 points right and `GH` ends
    // 3 points right; every inner edge on an item edge is exact, and `DE` has no item edge. Each
    // piece scores 0 as one line, and so does the join. Without items, the same pieces shift the
    // group 3 points, and the edges of the group score it.
    const reference: PdfLine = {
      ...line('AB 3DE6 GH', 0, 100, 100),
      size: 10,
      items: [
        { at: 0, length: 2, xMin: 0, xMax: 20 },
        { at: 2, length: 4, xMin: 30, xMax: 70 },
        { at: 6, length: 2, xMin: 80, xMax: 100 },
      ],
    };
    const pieces = [
      { ...line('AB', 3, 100, 20), size: 10 },
      { ...line('3', 30, 100, 39), size: 10 },
      { ...line('DE', 40, 100, 59), size: 10 },
      { ...line('6', 60, 100, 70), size: 10 },
      { ...line('GH', 80, 100, 103), size: 10 },
    ];
    const s = compareLines(pieces, [reference]);
    expect([s.matched, s.worstShift]).toEqual([1, 0]);
    const plain = compareLines(pieces, [{ ...reference, items: undefined }]);
    expect([plain.matched, plain.worstShift]).toEqual([1, 3]);
    // The converse: a painted line without items, 5 points right of reference pieces of the same
    // width, scores the edges of the group.
    const converse = compareLines(
      [line('ABCD', 5, 100, 36)],
      [line('AB', 0, 100, 10), line('CD', 21, 100, 31)]
    );
    expect([converse.matched, converse.worstShift]).toEqual([2, 5]);
  });

  test('a join piece moved past the gap limit or off the row is missing and extra', () => {
    // `84 | Shared text` of basic-values: reference items at 118..129.6 and 135..185.7, size 11.
    // Approximation "Join gap limit": the same move on one line fails on line position or on the
    // baseline.
    const whole = (dx: number): PdfLine => ({
      ...line('84 Shared text', 118, 100, 185.7 + dx),
      items: [
        { at: 0, length: 2, xMin: 118, xMax: 129.6 },
        { at: 2, length: 10, xMin: 135 + dx, xMax: 185.7 + dx },
      ],
    });
    const pieces = (dx: number, b = 100) => [
      line('84', 118, 100, 129.6),
      line('Shared text', 135 + dx, b, 185.7 + dx),
    ];
    const score = (ours: PdfLine[], ref: PdfLine[]) => {
      const s = compareLines(ours, ref);
      return [s.matched, s.missing, s.extra];
    };
    // `Shared text` 7 points right: the gap of the pieces is 12.4 points.
    expect(score(pieces(7), [whole(0)])).toEqual([0, ['84 Shared text'], ['84', 'Shared text']]);
    expect(score([whole(7)], pieces(0))).toEqual([0, ['84', 'Shared text'], ['84 Shared text']]);
    // `Shared text` 0.8 points low.
    expect(score(pieces(0, 100.8), [whole(0)])).toEqual([
      0,
      ['84 Shared text'],
      ['84', 'Shared text'],
    ]);
    expect(score([whole(0)], pieces(0, 100.8))).toEqual([
      0,
      ['84', 'Shared text'],
      ['84 Shared text'],
    ]);
    // The same moves on one line.
    const one = (dx: number, b = 100) =>
      compareLines(
        [line('Shared text', 135 + dx, b, 185.7 + dx)],
        [line('Shared text', 135, 100, 185.7)]
      );
    expect(one(7).worstShift).toBeCloseTo(7, 6);
    expect(one(0, 100.8).worstBaseline).toBeCloseTo(0.8, 6);
  });

  test('without a size on the whole line, the em of a join is the larger size of its pieces', () => {
    const whole = { ...line('ABCD', 0, 100, 35), size: undefined };
    const pieces = [
      { ...line('AB', 0, 100, 10), size: 10 },
      { ...line('CD', 25, 100, 35), size: 20 },
    ];
    expect(compareLines(pieces, [whole]).matched).toBe(1);
  });

  test('a gap of the whole line a few hundredths past 1 em is two lines there', () => {
    const whole: PdfLine = {
      ...line('AB CD', 0, 100, 31.03),
      items: [
        { at: 0, length: 2, xMin: 0, xMax: 10 },
        { at: 2, length: 2, xMin: 21.03, xMax: 31.03 },
      ],
    };
    const s = compareLines([whole], [line('AB', 0, 100, 10), line('CD', 21.03, 100, 31.03)]);
    expect(s.matched).toBe(0);
  });

  test('painted pieces that reference items do not split score the width of the group', () => {
    // One reference item holds `Hello world`; the painted `world` lies 3 points to the right.
    const reference: PdfLine = {
      ...line('Hello world', 0, 100, 52),
      size: 10,
      items: [{ at: 0, length: 10, xMin: 0, xMax: 52 }],
    };
    const ours = [
      { ...line('Hello', 0, 100, 25), size: 10 },
      { ...line('world', 30, 100, 55), size: 10 },
    ];
    for (const ref of [reference, { ...reference, items: undefined }]) {
      const s = compareLines(ours, [ref]);
      expect([s.matched, s.worstShift]).toEqual([1, 3]);
    }
  });

  test('each reference piece after the first joins by the split at its own characters', () => {
    // The painted `A BCD EF` has three items. The gap of the reference pieces `BCD` and `EF` is
    // 1 em plus 0.8 points: within the limit of a split of the painted items, and past the limit
    // where they do not split.
    const ours: PdfLine = {
      ...line('A BCD EF', 0, 100, 45),
      size: 10,
      items: [
        { at: 0, length: 1, xMin: 0, xMax: 5 },
        { at: 1, length: 3, xMin: 10, xMax: 25 },
        { at: 4, length: 2, xMin: 35, xMax: 45 },
      ],
    };
    const s = compareLines(
      [ours],
      [
        { ...line('A', 0, 100, 5), size: 10 },
        { ...line('BCD', 10, 100, 25), size: 10 },
        { ...line('EF', 35.8, 100, 45.8), size: 10 },
      ]
    );
    expect([s.matched, s.total, s.missing, s.extra]).toEqual([2, 2, [], []]);
  });

  test('a join stops at the first piece whose text does not continue the line', () => {
    // Thirty thousand `AB` pieces on the row of a line of `AB` and 60000 `C`s, on either page:
    // each start stops at its second piece, so the joins cost one step for each piece.
    const pieces = Array.from({ length: 30000 }, (_, i) => line('AB', i * 12, 100, i * 12 + 10));
    const long = line(`AB${'C'.repeat(60000)}`, 0, 100, 60);
    for (const [ours, reference] of [
      [pieces, [long]],
      [[long], pieces],
    ]) {
      const started = performance.now();
      const s = compareLines(ours, reference);
      expect(s.missing.length + s.extra.length).toBeGreaterThan(0);
      expect(performance.now() - started).toBeLessThan(1000);
    }
  });

  test('two cells that the reference PDF draws as one item read as a painted split', () => {
    // Approximation "Reference items across cells": the reference PDF holds `84` and `Shared text`
    // of two cells in one item, and the painted page draws two lines, each under its cell clip.
    const reference: PdfLine = {
      ...line('84 Shared text', 60, 92, 132.854),
      items: [{ at: 0, length: 12, xMin: 60, xMax: 132.854 }],
    };
    const page = [0, 0, 612, 792] as const;
    const s = compareLines(
      [line('84', 60, 92, 72.232), line('Shared text', 76.6, 92, 132.854)],
      [reference],
      {
        ours: [
          { x0: 60, x1: 72.232, baseline: 92, clip: [50, 82, 74, 102] },
          { x0: 76.6, x1: 132.854, baseline: 92, clip: [74, 82, 174, 102] },
        ],
        reference: [
          { x0: 60, x1: 72.232, baseline: 92, clip: page },
          { x0: 76.6, x1: 132.854, baseline: 92, clip: page },
        ],
      }
    );
    expect([s.matched, s.worstShift]).toEqual([1, 0]);
    expect(s.unknownReference).toEqual([
      '84 @ 92.0 | right painted split inside a reference item',
      'Shared text @ 92.0 | left painted split inside a reference item',
    ]);
  });
  test('each outer end of a join takes the baseline of the piece that holds it', () => {
    // Reference items `aaaa` at 50..70 and `bbbb` at 78..100, unclipped. The painted `bbbb` sits
    // 0.3 points below `aaaa`, within the baseline threshold of the line, so its right end lies on
    // its own baseline, as the same geometry as one painted line does.
    const reference: PdfLine = {
      ...line('aaaa bbbb', 50, 100, 100),
      items: [
        { at: 0, length: 4, xMin: 50, xMax: 70 },
        { at: 4, length: 4, xMin: 78, xMax: 100 },
      ],
    };
    const page = [0, 0, 612, 792] as const;
    for (const drop of [0.3, 0.45]) {
      const s = compareLines(
        [line('aaaa', 50, 100, 70), line('bbbb', 78, 100 + drop, 100)],
        [reference],
        {
          ours: [
            { x0: 50, x1: 70, baseline: 100, clip: [49, 90, 75, 104] },
            { x0: 78, x1: 100, baseline: 100 + drop, clip: [76, 90, 110, 104] },
          ],
          reference: [
            { x0: 50, x1: 70, baseline: 100, clip: page },
            { x0: 78, x1: 100, baseline: 100, clip: page },
          ],
        }
      );
      expect([s.matched, s.worstShift]).toEqual([1, 0]);
      expect(s.worstBaseline).toBeCloseTo(drop, 6);
      expect([s.extraCut, s.uncut, s.unknownReference]).toEqual([[], [], []]);
    }
  });
});
