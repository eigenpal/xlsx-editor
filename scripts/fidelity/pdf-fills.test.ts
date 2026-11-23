// The paint tables of pattern and gradient fills against the reference PDF of `fills-patterns`.

import { beforeAll, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { linearEnds, pathShapes } from '../../packages/core/src/output/paint-fills.ts';
import { paintedPosition } from './path-position.ts';
import {
  PATTERN_PIXEL_PT,
  PATTERN_TILES,
  PATTERN_TILE_SIZE,
} from '../../packages/core/src/output/pattern-tiles.ts';
import { FILL_CASES, PATTERN_TYPES } from '../fixtures/fill-style-cases.ts';
import { readPdfFills } from './pdf-fills.ts';
import { DEVICE_SCALE, PX_PER_PT } from './page/scale.ts';
import { readXlsx } from '../../packages/core/src/store/index.ts';
import {
  layoutSheet,
  PRINT_PROFILE,
  type SheetLayout,
} from '../../packages/core/src/layout/index.ts';
import { substituteMeasurer } from '../../packages/core/src/layout/__tests__/test-fonts.ts';
import { fillPaint, type GradientStop } from '../../packages/core/src/output/paint-style.ts';

const pdf = readPdfFills(
  new Uint8Array(readFileSync(join(import.meta.dir, '..', '..', 'fixtures', 'fills-patterns.pdf')))
);

/** The pattern fills of one column, top first: B (red on yellow), D (blue), F (yellow). */
const column = (x: number) =>
  pdf.patterns.filter((p) => Math.abs(p.rect.x - x) < 1).sort((a, b) => a.rect.y - b.rect.y);

describe('pattern tiles', () => {
  const left = Math.min(...pdf.patterns.map((p) => p.rect.x));
  const columns = [...new Set(pdf.patterns.map((p) => p.rect.x))].sort((a, b) => a - b);
  test('each pattern is an 8 x 8 tile of 1 point pixels on the page grid', () => {
    for (const p of pdf.patterns) {
      expect(p.step).toBe(PATTERN_TILE_SIZE * PATTERN_PIXEL_PT);
      expect(p.pixels.length).toBe(PATTERN_TILE_SIZE);
      expect(p.pageAnchored).toBe(true);
    }
  });
  test('the tile table equals the reference tiles, and absent colors are black and white', () => {
    const [b, d, f] = [column(left), column(columns[1]!), column(columns[2]!)];
    expect(b.length).toBe(PATTERN_TYPES.length);
    PATTERN_TYPES.forEach((type, i) => {
      const bits = (p: (typeof b)[number], fg: string) =>
        p.pixels.map((row) => row.map((v) => (v === fg ? '1' : '0')).join(''));
      expect(bits(b[i]!, 'FF0000')).toEqual([...PATTERN_TILES[type]!]);
      expect(bits(d[i]!, '0000FF')).toEqual([...PATTERN_TILES[type]!]);
      expect(bits(f[i]!, '000000')).toEqual([...PATTERN_TILES[type]!]);
      expect(new Set(d[i]!.pixels.flat())).toEqual(
        new Set(PATTERN_TILES[type]!.join('').includes('0') ? ['0000FF', 'FFFFFF'] : ['0000FF'])
      );
      expect(new Set(f[i]!.pixels.flat()).has('000000')).toBe(true);
    });
  });
});

describe('linear gradients', () => {
  const linear = FILL_CASES.filter(
    (k) => k.label.startsWith('linear ') && k.refs[0]!.startsWith('K')
  );
  test('the end points equal the reference axis in the unit square of the cell', () => {
    const shadings = pdf.shadings
      .filter((s) => s.rect.width < 100)
      .sort((a, b) => a.rect.y - b.rect.y);
    const degrees = linear
      .filter((k) => !k.label.includes('three'))
      .map((k) => Number(/degree="(\d+)"/.exec(k.fill)![1]));
    // Each angle, then the three-stop case: one shading in the unit square for each cell.
    expect(shadings.length).toBe(degrees.length + 1);
    // The unit square of each shading is the paint box: the cell and its gridline pixel.
    const box = paintBoxAt('K2');
    for (const s of shadings)
      expect([s.rect.width, s.rect.height]).toEqual([box.width, box.height]);
    degrees.forEach((degree, i) => {
      const [x0, y0, x1, y1] = linearEnds({ x: 0, y: 0, width: 1, height: 1 }, degree);
      const s = shadings[i]!;
      // The reference may draw the axis reversed, with the colors reversed.
      const forward = s.from === 'FF0000';
      // Red at one end and blue at the other, in either order.
      expect([s.from, s.to].sort()).toEqual(['0000FF', 'FF0000']);
      const [a0, b0, a1, b1] = forward
        ? s.coords
        : [s.coords[2], s.coords[3], s.coords[0], s.coords[1]];
      expect([x0, y0, x1, y1].map((v) => Number(v.toFixed(6)) + 0)).toEqual(
        [a0, b0, a1, b1].map((v) => Number(v.toFixed(6)) + 0)
      );
    });
  });
});

/** The path gradient images of column K, top first: the three red-to-blue cases, then K20. */
const images = () => {
  const list = pdf.images.slice().sort((a, b) => a.rect.y - b.rect.y);
  expect(list.length).toBe(4);
  return list;
};
/**
 * The largest channel difference of a three-stop path pixel. The reference image has 1 point
 * pixels, so a position differs by up to 0.012; near the stop at 0.3 the color changes fastest.
 * Measured: 9.
 */
const PATH_COLOR_TOLERANCE = 12;

/** The overlap that paint gives a path gradient in the harness: one device pixel. */
const HARNESS_OVERLAP = 1 / (PX_PER_PT * DEVICE_SCALE);

/** The fixture sheet, laid out as the harness prints it. */
let sheet: SheetLayout;
beforeAll(async () => {
  const read = readXlsx(
    new Uint8Array(
      readFileSync(join(import.meta.dir, '..', '..', 'fixtures', 'fills-patterns.xlsx'))
    )
  );
  if (!read.ok) throw new Error('read');
  const view = { width: 600, height: 800, scrollX: 0, scrollY: 0 };
  const measurer = await substituteMeasurer();
  sheet = layoutSheet(read.pkg, 'Fills', view, { measurer, profile: PRINT_PROFILE })!;
});
/** The gradient paint of the cell at `ref`, resolved through the cascade. */
const gradientAt = (ref: string) => {
  const col = ref.charCodeAt(0) - 65;
  const row = Number(ref.slice(1)) - 1;
  const box = sheet.panes.flatMap((p) => p.cells).find((b) => b.row === row && b.col === col)!;
  const paint = fillPaint(box.fill);
  if (paint.kind !== 'gradient') throw new Error(`${ref} has no gradient`);
  return paint;
};
/**
 * The box that paint gives a gradient at `ref`, at the origin: the cell and its gridline pixel,
 * which is `72 / pixelsPerInch` points (`paint-sheet.ts`). The reference image has 1 point pixels.
 */
const paintBoxAt = (ref: string) => {
  const col = ref.charCodeAt(0) - 65;
  const row = Number(ref.slice(1)) - 1;
  const box = sheet.panes.flatMap((p) => p.cells).find((b) => b.row === row && b.col === col)!;
  const px = 72 / sheet.metrics.pixelsPerInch;
  return { x: 0, y: 0, width: box.rect.width + px, height: box.rect.height + px };
};
/** The gradient stops that paint gives the cell at `ref`, as `RRGGBB`. */
const stopsAt = (ref: string): { offset: number; color: string }[] =>
  gradientAt(ref).stops.map((s: GradientStop) => ({ offset: s.offset, color: s.color.slice(1) }));

describe('path gradients', () => {
  // The cases of column K, top first: centered, on a corner, and off center, red to blue. Each
  // inner rectangle comes from the cascade.
  test.each(['K12', 'K14', 'K16'].map((ref, k) => [ref, k] as const))(
    'the path geometry gives the position of each reference pixel: %s',
    (ref, k) => {
      const inner = gradientAt(ref).rect;
      const image = images()[k]!;
      const h = image.pixels.length;
      const w = image.pixels[0]!.length;
      // The paint box is the size of the reference image, so each image pixel is one point.
      const box = paintBoxAt(ref);
      expect([box.width, box.height]).toEqual([w, h]);
      const shapes = pathShapes(box, inner, HARNESS_OVERLAP);
      for (let y = 0; y < h; y++)
        for (let x = 0; x < w; x++) {
          // Red to blue in sRGB: the blue channel is the position.
          const reference = parseInt(image.pixels[y]![x]!.slice(4, 6), 16) / 255;
          expect(Math.abs(reference - paintedPosition(shapes, x + 0.5, y + 0.5))).toBeLessThan(
            0.02
          );
        }
    }
  );
});

describe('three stops', () => {
  const channels = (c: string) => [0, 2, 4].map((i) => parseInt(c.slice(i, i + 2), 16));
  /** The color of the stops at `t`, mixed in sRGB. */
  const mixAt = (stops: { offset: number; color: string }[], t: number): number[] => {
    const i = Math.max(0, stops.findIndex((s) => s.offset >= t) - 1);
    const [a, b] = [stops[i]!, stops[Math.min(i + 1, stops.length - 1)]!];
    const u = b.offset === a.offset ? 0 : (t - a.offset) / (b.offset - a.offset);
    const [ca, cb] = [channels(a.color), channels(b.color)];
    return ca.map((v, k) => v + (cb[k]! - v) * u);
  };

  test('the linear paint stops equal the reference shading, mixed in sRGB', () => {
    const stops = stopsAt('K18');
    expect(stops.map((s) => s.offset)).toEqual([0, 0.3, 1]);
    // K18 is the last shading of column K.
    const shading = pdf.shadings
      .filter((s) => s.rect.width < 100)
      .sort((a, b) => a.rect.y - b.rect.y)
      .at(-1)!;
    for (const s of stops) expect(shading.colorAt(s.offset)).toBe(s.color);
    for (const t of [0.15, 0.5, 0.75]) {
      const reference = channels(shading.colorAt(t));
      mixAt(stops, t).forEach((v, k) => expect(Math.abs(v - reference[k]!)).toBeLessThanOrEqual(2));
    }
  });

  test('the path paint stops give the color of each reference pixel', () => {
    const stops = stopsAt('K20');
    const image = images()[3]!;
    const h = image.pixels.length;
    const w = image.pixels[0]!.length;
    const inner = gradientAt('K20').rect;
    const box = paintBoxAt('K20');
    expect([box.width, box.height]).toEqual([w, h]);
    const shapes = pathShapes(box, inner, HARNESS_OVERLAP);
    let worst = 0;
    for (let y = 0; y < h; y++)
      for (let x = 0; x < w; x++) {
        const painted = mixAt(stops, paintedPosition(shapes, x + 0.5, y + 0.5));
        const reference = channels(image.pixels[y]![x]!);
        for (let k = 0; k < 3; k++) worst = Math.max(worst, Math.abs(painted[k]! - reference[k]!));
      }
    expect(worst).toBeLessThanOrEqual(PATH_COLOR_TOLERANCE);
  });
});
