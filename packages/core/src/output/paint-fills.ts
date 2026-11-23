// Cell fills: one color, a pattern tile, or a gradient, as the reference PDF of fixture
// `fills-patterns` draws them. `scripts/fidelity/pdf-fills.test.ts`
// compares the tiles, the linear axes, and the path gradient positions with that PDF.
//
// - A pattern repeats its 8 x 8 tile of 1 point pixels (`pattern-tiles.ts`) from one origin: the
//   sheet origin of the pane, or `PaintOptions.patternOrigin`. The tile is drawn in device pixels.
//   A square bitmap holds `k` by `k` tiles, with `k` from 1 to 8 chosen for the smallest period
//   drift for each tile. Pattern pixel `j` of the bitmap covers the device pixels from
//   `round(j * side / (8 * k))` up to `round((j + 1) * side / (8 * k))`, so each pattern pixel
//   covers one or more whole device pixels, and the period drifts by less than half a device
//   pixel for each bitmap. When one tile is smaller than 8 device pixels, or larger than
//   `MAX_TILE_SIDE`, the fill paints the mix of its colors (approximation "Fill pixels").
// - A gradient fills the box of its cell: the cell and the gridline pixel at its right and bottom,
//   as a plain fill does. The reference PDF clips each gradient to that box (67 by 23 points for
//   a cell of 66 by 22 points and its 1 point gridline), so the positions are those of the box.
//   A merged range is one box, so one gradient covers the whole range.
// The fill types are in `sml.xsd`: `x:patternFill` (`CT_PatternFill`, `@patternType` of
// `ST_PatternType`) and `x:gradientFill` (`CT_GradientFill`, `@type` of `ST_GradientType`). The
// schema gives the attributes; the geometry that this file paints is measured.
// - A linear gradient runs along `@degree` in the unit square of the box, which is then scaled to
//   the box: its end points are the center minus and plus `(cos, sin) * (|cos| + |sin|) / 2`, so
//   45 degrees runs corner to corner, and its lines of equal color join the other two corners.
// - A path gradient paints position 0 on its inner rectangle (`@left`, `@right`, `@top`,
//   `@bottom`, each default 0) and position 1 on the box edges,
//   with lines of equal color that are rectangles (`pathShapes`).
//
// The tile canvas is created from the document of the painter's own canvas, never from a script
// global. When there is none, or the context cannot make or place a pattern, the fill paints the
// mix of its colors.

import type { Rect } from '../layout/index.ts';
import type { FillPaint, GradientStop } from './paint-style.ts';
import { PATTERN_PIXEL_PT, PATTERN_TILES, PATTERN_TILE_SIZE } from './pattern-tiles.ts';

type G = CanvasRenderingContext2D;
type PatternPaint = Extract<FillPaint, { kind: 'pattern' }>;

/** What a fill paint needs from the sheet paint, for one pane. */
export interface FillContext {
  readonly g: G;
  /** Device pixels for each point. */
  readonly ratio: number;
  /** Where the tile grid starts, in viewport points. */
  readonly origin: { readonly x: number; readonly y: number };
  /** The patterns of this paint and pane, by tile key: a pattern once made is placed once. */
  readonly patterns: Map<string, CanvasPattern | null>;
}

/** The largest side of one tile, in device pixels. */
export const MAX_TILE_SIDE = 512;
/**
 * The most tile bitmaps kept for one document; the entry used least recently goes first. A view
 * with more pattern and color pairs builds bitmaps again on each paint.
 */
export const TILE_CACHE_CAP = 256;
/**
 * The most device pixels of all tile bitmaps kept for one document (64 MiB of RGBA): 64 bitmaps
 * of the largest side, so that a zoomed view with many patterns keeps its bitmaps. A view at
 * 100 % zoom uses bitmaps of 32 to 64 pixels on a side, about 4 MiB for 256 entries.
 */
export const TILE_CACHE_AREA = 16 * 1024 * 1024;
/** Built tile bitmaps, by document, so that a bitmap never crosses into another window. */
const TILE_CACHES = new WeakMap<Document, Map<string, HTMLCanvasElement>>();

/** Paint one fill over `rect`. `bleed` widens the fill over the right and bottom gridline. */
export function paintFill(fc: FillContext, paint: FillPaint, rect: Rect, bleed: number): void {
  const { g } = fc;
  if (paint.kind === 'none') return;
  const box = { x: rect.x, y: rect.y, width: rect.width + bleed, height: rect.height + bleed };
  if (paint.kind === 'color' || paint.kind === 'pattern') {
    g.fillStyle = paint.kind === 'color' ? paint.color : (patternStyle(fc, paint) ?? paint.mix);
    g.fillRect(box.x, box.y, box.width, box.height);
    return;
  }
  if (paint.type === 'linear') linear(g, box, paint.degree, paint.stops);
  else path(g, box, paint.rect, paint.stops, 1 / fc.ratio);
}

/** How many tiles one bitmap holds at `ratio`, and its side in device pixels. */
export function tileBitmap(ratio: number): { tiles: number; side: number } | undefined {
  const one = PATTERN_TILE_SIZE * PATTERN_PIXEL_PT * ratio;
  // A tile of 8 or more device pixels gives each pattern pixel one or more device pixels.
  if (one < PATTERN_TILE_SIZE || one > MAX_TILE_SIDE) return undefined;
  let best = { tiles: 1, side: Math.round(one), error: Math.abs(one - Math.round(one)) };
  for (let k = 2; k <= 8 && Math.round(one * k) <= MAX_TILE_SIDE; k++) {
    const error = Math.abs(one * k - Math.round(one * k)) / k;
    if (error < best.error - 1e-9) best = { tiles: k, side: Math.round(one * k), error };
  }
  return { tiles: best.tiles, side: best.side };
}

/** The device pixel edges of the pattern pixels of a bitmap: `count + 1` values. */
export function pixelEdges(side: number, count: number): number[] {
  return Array.from({ length: count + 1 }, (_, j) => Math.round((j * side) / count));
}

function patternStyle(fc: FillContext, paint: PatternPaint): CanvasPattern | undefined {
  const { g, ratio } = fc;
  const bitmap = tileBitmap(ratio);
  if (!bitmap || typeof g.createPattern !== 'function') return undefined;
  const key = `${paint.pattern}|${paint.foreground}|${paint.background}|${bitmap.side}|${bitmap.tiles}`;
  if (fc.patterns.has(key)) return fc.patterns.get(key) ?? undefined;
  const canvas = tileCanvas(g, paint, bitmap.side, bitmap.tiles, key);
  let pattern = canvas ? g.createPattern(canvas, 'repeat') : null;
  if (pattern && typeof pattern.setTransform === 'function') {
    // The bitmap is in device pixels: undo the paint scale, and start the grid on a device pixel.
    // The origin is taken within one bitmap period, so the matrix keeps small numbers far down a
    // sheet, where a canvas would round a large offset by more than one device pixel.
    const phase = (v: number) => {
      const d = Math.round(v * ratio) % bitmap.side;
      return (d < 0 ? d + bitmap.side : d) / ratio;
    };
    const e = phase(fc.origin.x);
    const f = phase(fc.origin.y);
    pattern.setTransform({ a: 1 / ratio, b: 0, c: 0, d: 1 / ratio, e, f });
  } else pattern = null;
  fc.patterns.set(key, pattern);
  return pattern ?? undefined;
}

function tileCanvas(
  g: G,
  paint: PatternPaint,
  side: number,
  tiles: number,
  key: string
): HTMLCanvasElement | undefined {
  const doc = (g.canvas as HTMLCanvasElement | undefined)?.ownerDocument;
  const rows = Object.hasOwn(PATTERN_TILES, paint.pattern)
    ? PATTERN_TILES[paint.pattern]
    : undefined;
  if (!doc || !rows) return undefined;
  let cache = TILE_CACHES.get(doc);
  if (!cache) {
    cache = new Map();
    TILE_CACHES.set(doc, cache);
  }
  const hit = cache.get(key);
  if (hit) {
    cache.delete(key);
    cache.set(key, hit);
    return hit;
  }
  const canvas = doc.createElement('canvas');
  canvas.width = side;
  canvas.height = side;
  const t = canvas.getContext('2d');
  if (!t) return undefined;
  const count = PATTERN_TILE_SIZE * tiles;
  const edge = pixelEdges(side, count);
  t.fillStyle = paint.background;
  t.fillRect(0, 0, side, side);
  t.fillStyle = paint.foreground;
  for (let r = 0; r < count; r++) {
    const bits = rows[r % PATTERN_TILE_SIZE]!;
    for (let c = 0; c < count; c++)
      if (bits[c % PATTERN_TILE_SIZE] === '1')
        t.fillRect(edge[c]!, edge[r]!, edge[c + 1]! - edge[c]!, edge[r + 1]! - edge[r]!);
  }
  let area = side * side;
  for (const c of cache.values()) area += c.width * c.height;
  for (const [old, c] of cache) {
    if (cache.size < TILE_CACHE_CAP && area <= TILE_CACHE_AREA) break;
    area -= c.width * c.height;
    cache.delete(old);
  }
  cache.set(key, canvas);
  return canvas;
}

/** The end points of a linear gradient along `degree`, clockwise from left to right. */
export function linearEnds(rect: Rect, degree: number): [number, number, number, number] {
  const rad = (degree * Math.PI) / 180;
  const cos = Math.cos(rad);
  const sin = Math.sin(rad);
  const reach = (Math.abs(cos) + Math.abs(sin)) / 2;
  const cx = rect.x + rect.width / 2;
  const cy = rect.y + rect.height / 2;
  const dx = cos * reach * rect.width;
  const dy = sin * reach * rect.height;
  return [cx - dx, cy - dy, cx + dx, cy + dy];
}

const UNIT: Rect = Object.freeze({ x: 0, y: 0, width: 1, height: 1 });

/**
 * A linear gradient in the unit square of the box, scaled to the box: its lines of equal color
 * join the same fractions of the box edges, as the reference PDF draws them.
 */
function linear(g: G, box: Rect, degree: number, stops: readonly GradientStop[]): void {
  if (box.width <= 0 || box.height <= 0) return;
  g.save();
  g.translate(box.x, box.y);
  g.scale(box.width, box.height);
  const grad = g.createLinearGradient(...linearEnds(UNIT, degree));
  for (const s of stops) grad.addColorStop(s.offset, s.color);
  g.fillStyle = grad;
  g.fillRect(0, 0, 1, 1);
  g.restore();
}

/** One shape of a path gradient: its outline, and the line from position 0 to position 1. */
export interface PathShape {
  readonly points: readonly (readonly [number, number])[];
  readonly from: readonly [number, number];
  readonly to: readonly [number, number];
}

/**
 * The shapes of a path gradient over `rect`, the box of the cell, in paint order, with the inner
 * rectangle `inner` as fractions of that box. Each part takes a linear gradient across its outer
 * edge. The shapes overlap, so that no anti-aliased edge of one shape lies over an empty pixel:
 * 0. the whole box in the color of position 0, under the others;
 * 1. the left and right parts, as bands of the full height that reach `overlap` points past the
 *    inner rectangle;
 * 2. the top part, which reaches down over the whole inner rectangle, where its gradient gives
 *    position 0;
 * 3. the bottom part, which reaches up over the whole inner rectangle.
 * Over each shared diagonal, both parts give the same position. A part of zero size is left out.
 */
export function pathShapes(
  rect: Rect,
  inner: { left: number; right: number; top: number; bottom: number },
  overlap: number
): PathShape[] {
  const { x, y, width: w, height: h } = rect;
  const L = x + inner.left * w;
  const R = x + inner.right * w;
  const T = y + inner.top * h;
  const B = y + inner.bottom * h;
  const X = x + w;
  const Y = y + h;
  const shapes: PathShape[] = [
    { points: [[x, y], [X, y], [X, Y], [x, Y]], from: [x, y], to: [x, y] }, // prettier-ignore
  ];
  const band = (x0: number, x1: number, from: number, to: number) =>
    shapes.push({ points: [[x0, y], [x1, y], [x1, Y], [x0, Y]], from: [from, 0], to: [to, 0] }); // prettier-ignore
  if (L > x) band(x, Math.min(X, L + overlap), L, x);
  // The right band never starts left of the inner rectangle, so it never covers the left part.
  if (X > R) band(Math.max(x, L, R - overlap), X, R, X);
  if (T > y || B > T)
    shapes.push({ points: [[x, y], [X, y], [R, T], [R, B], [L, B], [L, T]], from: [0, T], to: [0, y] }); // prettier-ignore
  if (Y > B || B > T)
    shapes.push({ points: [[x, Y], [X, Y], [R, B], [R, T], [L, T], [L, B]], from: [0, B], to: [0, Y] }); // prettier-ignore
  return shapes;
}

/** A path gradient over `box`, the cell and its gridline pixel, as the reference PDF clips it. */
function path(
  g: G,
  box: Rect,
  inner: { left: number; right: number; top: number; bottom: number },
  stops: readonly GradientStop[],
  overlap: number
): void {
  if (stops.length === 0) return;
  for (const shape of pathShapes(box, inner, overlap)) {
    const [fx, fy] = shape.from;
    const [tx, ty] = shape.to;
    if (fx === tx && fy === ty) g.fillStyle = stops[0]!.color;
    else {
      const grad = g.createLinearGradient(fx, fy, tx, ty);
      for (const s of stops) grad.addColorStop(s.offset, s.color);
      g.fillStyle = grad;
    }
    g.beginPath();
    shape.points.forEach(([px, py], i) => (i === 0 ? g.moveTo(px, py) : g.lineTo(px, py)));
    g.closePath();
    g.fill();
  }
}
