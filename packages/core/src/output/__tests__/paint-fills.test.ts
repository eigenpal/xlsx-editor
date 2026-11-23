// The paint of pattern and gradient fills: the tile bitmap, the fallback to the mix, the pattern
// origin, and the geometry of gradients (scripts/fidelity/pdf-fills.test.ts compares the tiles,
// the linear axes, and the path positions with the reference PDF itself).

import { describe, expect, test } from 'bun:test';
import {
  linearEnds,
  MAX_TILE_SIDE,
  TILE_CACHE_AREA,
  paintFill,
  pathShapes,
  pixelEdges,
  tileBitmap,
  type FillContext,
} from '../paint-fills.ts';
import { fillPaint } from '../paint-style.ts';
import { paintedPosition } from '../../../../../scripts/fidelity/path-position.ts';

const cell = { x: 10, y: 20, width: 100, height: 40 };

/** A stub 2D context and document that record what the painter asks for. */
function stub(options: { pattern?: boolean; transform?: boolean } = {}) {
  const made: { width: number; height: number; rects: number[][] }[] = [];
  const doc = {
    createElement: () => {
      const canvas = {
        width: 0,
        height: 0,
        rects: [] as number[][],
        getContext: () => ({
          fillStyle: '',
          fillRect: (...r: number[]) => canvas.rects.push(r),
        }),
      };
      made.push(canvas);
      return canvas;
    },
  };
  const fills: unknown[] = [];
  const patterns: { source: unknown; matrix?: unknown }[] = [];
  const rects: number[][] = [];
  const axes: number[][] = [];
  const moves: (string | number)[][] = [];
  const points: number[][] = [];
  const g = {
    canvas: { ownerDocument: doc },
    fillStyle: '' as unknown,
    fillRect(...r: number[]) {
      fills.push(this.fillStyle);
      rects.push(r);
    },
    save() {},
    restore() {},
    translate: (...a: number[]) => moves.push(['translate', ...a]),
    scale: (...a: number[]) => moves.push(['scale', ...a]),
    createLinearGradient: (...axis: number[]) => {
      axes.push(axis);
      return { addColorStop() {} };
    },
    beginPath() {},
    closePath() {},
    fill() {},
    moveTo: (x: number, y: number) => points.push([x, y]),
    lineTo: (x: number, y: number) => points.push([x, y]),
    ...(options.pattern === false
      ? {}
      : {
          createPattern: (source: unknown) => {
            const p: { source: unknown; matrix?: unknown; setTransform?: (m: unknown) => void } = {
              source,
            };
            if (options.transform !== false) p.setTransform = (m) => (p.matrix = m);
            patterns.push(p);
            return p;
          },
        }),
  };
  return {
    g: g as unknown as CanvasRenderingContext2D,
    made,
    fills,
    patterns,
    rects,
    axes,
    points,
    moves,
  };
}

const grayPaint = fillPaint({
  kind: 'pattern',
  pattern: 'mediumGray',
  foreground: '000000',
  background: 'FFFFFF',
});

const context = (g: CanvasRenderingContext2D, ratio: number, origin = { x: 0, y: 0 }) =>
  ({ g, ratio, origin, patterns: new Map() }) satisfies FillContext;

describe('pattern paint', () => {
  test('a tile smaller than 8 device pixels paints the mix', () => {
    const s = stub();
    paintFill(context(s.g, 0.5), grayPaint, cell, 0);
    expect(s.fills).toEqual(['#808080']);
    expect(s.made).toEqual([]);
  });
  test('a tile of 16 device pixels at ratio 2, from the document of the target canvas', () => {
    const s = stub();
    paintFill(context(s.g, 2, { x: 3.3, y: -36 }), grayPaint, cell, 0);
    expect(s.made.length).toBe(1);
    expect([s.made[0]!.width, s.made[0]!.height]).toEqual([16, 16]);
    // The background, then each set pixel of mediumGray as a 2 by 2 device pixel square.
    expect(s.made[0]!.rects[0]).toEqual([0, 0, 16, 16]);
    expect(s.made[0]!.rects.slice(1, 3)).toEqual([
      [0, 0, 2, 2],
      [4, 0, 2, 2],
    ]);
    expect(s.patterns[0]!.matrix).toEqual({ a: 0.5, b: 0, c: 0, d: 0.5, e: 3.5, f: 4 });
    expect(s.fills).toEqual([s.patterns[0]]);
  });
  test('an origin far down a sheet gives the same phase in one bitmap period', () => {
    const s = stub();
    paintFill(context(s.g, 2, { x: 3.5, y: -1.5e7 - 1 }), grayPaint, cell, 0);
    // 16 device pixels for each bitmap: 7 and -3e7 - 2 device pixels give 7 and 14.
    expect(s.patterns[0]!.matrix).toEqual({ a: 0.5, b: 0, c: 0, d: 0.5, e: 3.5, f: 7 });
  });
  test('one paint makes each pattern once, and the document cache keeps the bitmap', () => {
    const s = stub();
    const fc = context(s.g, 2);
    paintFill(fc, grayPaint, cell, 0);
    paintFill(fc, grayPaint, { ...cell, x: 200 }, 0);
    expect(s.patterns.length).toBe(1);
    paintFill(context(s.g, 2), grayPaint, cell, 0);
    expect(s.made.length).toBe(1);
    expect(s.patterns.length).toBe(2);
  });
  test('a context with no pattern, or a pattern with no transform, paints the mix', () => {
    const a = stub({ pattern: false });
    paintFill(context(a.g, 2), grayPaint, cell, 0);
    expect(a.fills).toEqual(['#808080']);
    const b = stub({ transform: false });
    paintFill(context(b.g, 2), grayPaint, cell, 0);
    expect(b.fills).toEqual(['#808080']);
  });
  test('the bitmap holds the tiles whose period is nearest whole device pixels', () => {
    // 100 % zoom on a standard screen: 4/3 device pixels for each point, 3 tiles in 32 pixels.
    expect(tileBitmap(4 / 3)).toEqual({ tiles: 3, side: 32 });
    expect(tileBitmap(2)).toEqual({ tiles: 1, side: 16 });
    expect(tileBitmap(0.9)).toBeUndefined();
    expect(tileBitmap(0.99)).toBeUndefined();
    expect(tileBitmap(1)).toEqual({ tiles: 1, side: 8 });
    expect(tileBitmap(MAX_TILE_SIDE)).toBeUndefined();
    const edges = pixelEdges(11, 8);
    expect(edges[0]).toBe(0);
    expect(edges[8]).toBe(11);
    for (let i = 0; i < 8; i++) expect(edges[i + 1]! - edges[i]!).toBeGreaterThanOrEqual(1);
  });
});

describe('tile bitmap limits', () => {
  test('each pattern pixel covers one or more device pixels at every ratio from 1 to 8', () => {
    for (let ratio = 1; ratio <= 8; ratio += 0.001) {
      const bitmap = tileBitmap(ratio)!;
      const edges = pixelEdges(bitmap.side, 8 * bitmap.tiles);
      for (let j = 0; j < edges.length - 1; j++)
        expect(edges[j + 1]! - edges[j]!).toBeGreaterThanOrEqual(1);
    }
  });
  test('the document cache keeps no more device pixels than its area limit', () => {
    const s = stub();
    const side = MAX_TILE_SIDE;
    const count = TILE_CACHE_AREA / (side * side) + 4;
    const paint = (i: number) =>
      paintFill(
        context(s.g, side / 8),
        fillPaint({
          kind: 'pattern',
          pattern: 'mediumGray',
          foreground: i.toString(16).padStart(6, '0'),
          background: 'FFFFFF',
        }),
        cell,
        0
      );
    for (let i = 0; i < count; i++) paint(i);
    expect(s.made.length).toBe(count);
    // The newest bitmap stays; the oldest went first.
    paint(count - 1);
    expect(s.made.length).toBe(count);
    paint(0);
    expect(s.made.length).toBe(count + 1);
  });
});

describe('linear gradient fill', () => {
  test('the unit square is the box: the cell and its gridline pixel', () => {
    const s = stub();
    paintFill(
      context(s.g, 2),
      {
        kind: 'gradient',
        type: 'linear',
        degree: 0,
        stops: [],
        rect: { left: 0, right: 0, top: 0, bottom: 0 },
      },
      cell,
      1
    );
    expect(s.rects).toEqual([[0, 0, 1, 1]]);
    // The unit square is placed on the cell and scaled to the cell and its 1 point bleed.
    expect(s.moves).toEqual([
      ['translate', 10, 20],
      ['scale', 101, 41],
    ]);
  });
});

describe('path gradient fill', () => {
  test('the positions are those of the box: the cell and its gridline pixel', () => {
    const s = stub();
    const stops = [
      { offset: 0, color: '#FF0000' },
      { offset: 1, color: '#0000FF' },
    ];
    const inner = { left: 0.5, right: 0.5, top: 0.5, bottom: 0.5 };
    paintFill(
      context(s.g, 2),
      { kind: 'gradient', type: 'path', degree: 0, stops, rect: inner },
      cell,
      1
    );
    // The outlines reach the gridline pixel at x 111 and y 61.
    expect(Math.max(...s.points.map((p) => p[0]!))).toBe(111);
    expect(Math.max(...s.points.map((p) => p[1]!))).toBe(61);
    // Each gradient runs from the box center to a box edge.
    expect(s.axes).toEqual([
      [60.5, 0, 10, 0],
      [60.5, 0, 111, 0],
      [0, 40.5, 0, 20],
      [0, 40.5, 0, 61],
    ]);
  });
});

describe('linear gradient end points', () => {
  test('0 degrees runs from the left edge to the right edge', () => {
    expect(linearEnds(cell, 0)).toEqual([10, 40, 110, 40]);
  });
  test('45 degrees runs from corner to corner in a cell that is not square', () => {
    expect(linearEnds(cell, 45).map((v) => Math.round(v * 1e6) / 1e6)).toEqual([10, 20, 110, 60]);
  });
  test('270 degrees runs from the bottom edge to the top edge', () => {
    expect(linearEnds(cell, 270).map((v) => Math.round(v * 1e6) / 1e6)).toEqual([60, 60, 60, 20]);
  });
});

/** The position of the reference: the largest distance to an edge of the inner rectangle. */
function expected(
  inner: { left: number; right: number; top: number; bottom: number },
  u: number,
  v: number
): number {
  const parts = [
    inner.left > 0 ? (inner.left - u) / inner.left : 0,
    inner.right < 1 ? (u - inner.right) / (1 - inner.right) : 0,
    inner.top > 0 ? (inner.top - v) / inner.top : 0,
    inner.bottom < 1 ? (v - inner.bottom) / (1 - inner.bottom) : 0,
  ];
  return Math.max(0, ...parts);
}

describe('path gradient shapes', () => {
  const cases = [
    { left: 0.5, right: 0.5, top: 0.5, bottom: 0.5 },
    { left: 0, right: 0, top: 0, bottom: 0 },
    { left: 0.2, right: 0.4, top: 0.3, bottom: 0.6 },
    { left: 1, right: 1, top: 1, bottom: 1 },
  ];
  test.each(cases.map((c) => [JSON.stringify(c), c] as const))(
    'every point paints the position of the inner rectangle distance: %s',
    (_, inner) => {
      const shapes = pathShapes(cell, inner, 0.5);
      for (let i = 1; i < 40; i++)
        for (let j = 1; j < 20; j++) {
          const u = (i + 0.37) / 40;
          const v = (j + 0.29) / 20;
          const got = paintedPosition(shapes, cell.x + u * cell.width, cell.y + v * cell.height);
          expect(Math.abs(got - expected(inner, u, v))).toBeLessThan(1e-9);
        }
    }
  );
  test('an inner rectangle narrower than the overlap keeps its left part', () => {
    const inner = { left: 0.01, right: 0.01, top: 0.5, bottom: 0.5 };
    // The overlap of one device pixel at 4/3 device pixels for each point is wider than the inner
    // rectangle: samples inside it must still paint the left part.
    const shapes = pathShapes(cell, inner, 0.75);
    for (let i = 0; i < 100; i++) {
      const u = (i + 0.5) / 100;
      for (const v of [0.25, 0.5, 0.75]) {
        const got = paintedPosition(shapes, cell.x + u * cell.width, cell.y + v * cell.height);
        expect(Math.abs(got - expected(inner, u, v))).toBeLessThan(1e-9);
      }
    }
  });
  test('the first shape is the whole box in the color of position 0', () => {
    const [under] = pathShapes(cell, cases[0]!, 0.5);
    expect(under!.from).toEqual(under!.to);
    expect(under!.points).toEqual([[10, 20], [110, 20], [110, 60], [10, 60]]); // prettier-ignore
  });
});
