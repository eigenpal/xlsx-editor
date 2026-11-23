// The paint of a sheet window into a `<canvas>` (design D7): the target that the viewer and the
// fidelity harness use. The paint-target spike of D7 measured it against an SVG painter of the
// same records, and kept the canvas. The canvas works in points: its transform maps one point to
// `scale * devicePixelRatio` device pixels. Each pane clips to its own rectangle. In each pane,
// in paint order: the gridlines, the cell fills, the borders, the text, and the drawings.
//
// A canvas holds no text that assistive technology reads, so the viewer pairs it with the
// accessible grid of the same layout (`accessible-grid.ts`). Text draws glyph by glyph at the
// positions that layout measured, with the measured face (registered under its id by
// `registerFontFaces`), so no kerning or ligature of the browser moves a glyph. Underlines and
// strikes are rectangles from the face metrics, as in PPTX Editor's text paint. Pictures draw
// from package bytes (`pictures.ts`); nothing loads from a URL.

import type {
  CellBox,
  CellTextBox,
  DrawingLayout,
  PaneLayout,
  SheetLayout,
  TextSpan,
} from '../layout/index.ts';
import {
  borderLine,
  cssColor,
  cssRgba,
  DEFAULT_GRIDLINE_COLOR,
  fillPaint,
  lineBand,
  type BorderLine,
} from './paint-style.ts';
import { paintFill } from './paint-fills.ts';
import type { PictureStore } from './pictures.ts';

/** How to paint a sheet window. */
export interface PaintOptions {
  /** CSS pixels for each point. */
  readonly scale: number;
  /** The paint area, in points. */
  readonly width: number;
  readonly height: number;
  /**
   * Draw gridlines: the sheet view flag on screen (`x:sheetView/@showGridLines`), the print flag
   * on a printed page (`x:printOptions/@gridLines`).
   */
  readonly gridlines: boolean;
  /**
   * `RRGGBB`; `DEFAULT_GRIDLINE_COLOR` when absent. Layout does not read
   * `x:sheetView/@colorId` or `@defaultGridColor` (`CT_SheetView`).
   */
  readonly gridlineColor?: string;
  /** Draw the frozen pane edges (screen only). */
  readonly frozenLines?: boolean;
  /** Decoded pictures; pictures paint nothing without it. */
  readonly pictures?: PictureStore;
  /** Device pixels for each CSS pixel. Default 1. */
  readonly devicePixelRatio?: number;
  /**
   * Where the tile grid of pattern fills starts, in viewport points, for every pane. The
   * reference PDFs start it at the bottom-left corner of the page, so the fidelity harness passes
   * that corner. When absent, each pane starts it at its sheet origin, so a pattern moves with its
   * cells when the view scrolls.
   */
  readonly patternOrigin?: { readonly x: number; readonly y: number };
}

/** What a paint drew, for tests, the harness, and the viewer. */
export interface PaintStats {
  readonly cells: number;
  readonly texts: number;
  readonly drawings: number;
  /** Pictures that are still decoding: paint again when the store reports them ready. */
  readonly pendingPictures: number;
}

/** The largest side of the canvas backing store, in device pixels. */
export const MAX_CANVAS_SIDE = 16_384;

type G = CanvasRenderingContext2D;

interface Ctx {
  readonly g: G;
  readonly options: PaintOptions;
  /** One grid pixel, in points. */
  readonly px: number;
  /** Device pixels for each point. */
  readonly ratio: number;
  pending: number;
}

/**
 * A canvas font for a measured face id. The id keeps only `[A-Za-z0-9_-]`; a result with a
 * leading digit or a CSS keyword is not a valid family, so the canvas refuses the string and
 * keeps its earlier font. Exported for the paint tests.
 *
 * @internal
 */
export function fontOf(faceId: string, sizePt: number, bold: boolean, italic: boolean): string {
  const safe = faceId.replace(/[^A-Za-z0-9_-]/g, '');
  const family = safe.length > 0 ? `${safe}, sans-serif` : 'sans-serif';
  return `${italic ? 'italic ' : ''}${bold ? 'bold ' : ''}${sizePt}px ${family}`;
}

function gridlines(ctx: Ctx, pane: PaneLayout): void {
  if (pane.columns.length === 0 || pane.rows.length === 0) return;
  const { g, px } = ctx;
  const top = pane.rows[0]!.start;
  const last = pane.rows[pane.rows.length - 1]!;
  const bottom = last.start + last.size;
  const left = pane.columns[0]!.start;
  const end = pane.columns[pane.columns.length - 1]!;
  const right = end.start + end.size;
  g.fillStyle = cssColor(ctx.options.gridlineColor ?? DEFAULT_GRIDLINE_COLOR);
  for (const c of pane.columns) g.fillRect(c.start, top, px, bottom - top);
  g.fillRect(right, top, px, bottom - top);
  for (const r of pane.rows) g.fillRect(left, r.start, right - left, px);
  g.fillRect(left, bottom, right - left, px);
}

function fills(ctx: Ctx, pane: PaneLayout): void {
  const fc = {
    g: ctx.g,
    ratio: ctx.ratio,
    origin: ctx.options.patternOrigin ?? pane.sheetOrigin,
    patterns: new Map<string, CanvasPattern | null>(),
  };
  for (const box of pane.cells) paintFill(fc, fillPaint(box.fill), box.rect, ctx.px);
}

/** One border edge: a band of grid pixels along a gridline. */
function edge(
  ctx: Ctx,
  line: BorderLine,
  color: string,
  vertical: boolean,
  at: number,
  from: number,
  length: number
): void {
  const { g, px } = ctx;
  const [start, size] = lineBand(at, line.width, px);
  const bands: [number, number][] = line.double
    ? [
        [start, px],
        [start + 2 * px, px],
      ]
    : [[start, size]];
  for (const [s, w] of bands) {
    if (line.dash.length === 0) {
      g.fillStyle = color;
      if (vertical) g.fillRect(s, from, w, length + px);
      else g.fillRect(from, s, length + px, w);
      continue;
    }
    const mid = s + w / 2;
    g.strokeStyle = color;
    g.lineWidth = w;
    g.setLineDash(line.dash.map((d) => d * px));
    g.beginPath();
    g.moveTo(vertical ? mid : from, vertical ? from : mid);
    g.lineTo(vertical ? mid : from + length + px, vertical ? from + length + px : mid);
    g.stroke();
    g.setLineDash([]);
  }
}

function borders(ctx: Ctx, pane: PaneLayout): void {
  const { g, px } = ctx;
  for (const box of pane.cells) {
    const { x, y, width, height } = box.rect;
    const b = box.border;
    const sides = [
      [b.left, true, x, y, height],
      [b.right, true, x + width, y, height],
      [b.top, false, y, x, width],
      [b.bottom, false, y + height, x, width],
    ] as const;
    for (const [e, vertical, at, from, length] of sides) {
      const line = borderLine(e.style);
      if (line) edge(ctx, line, cssColor(e.color), vertical, at, from, length);
    }
    diagonals(g, box, px);
  }
}

function diagonals(g: G, box: CellBox, px: number): void {
  const b = box.border;
  const diagonal = borderLine(b.diagonal.style);
  if (!diagonal || !(b.diagonalUp || b.diagonalDown)) return;
  const { x, y, width, height } = box.rect;
  g.strokeStyle = cssColor(b.diagonal.color);
  g.lineWidth = diagonal.width * px;
  g.setLineDash(diagonal.dash.map((d) => d * px));
  g.beginPath();
  if (b.diagonalDown) {
    g.moveTo(x, y);
    g.lineTo(x + width, y + height);
  }
  if (b.diagonalUp) {
    g.moveTo(x, y + height);
    g.lineTo(x + width, y);
  }
  g.stroke();
  g.setLineDash([]);
}

/**
 * Paints one span. `last` is the canvas font that this text box set last, so a run of spans in
 * one font sets it once; the function returns the font that it leaves set. With a `clip`, it
 * draws only the glyphs that can reach it, so the cost of a very long span is the glyphs in view.
 */
function spanText(
  ctx: Ctx,
  span: TextSpan,
  baseline: number,
  last: string,
  clip: CellTextBox['clip'] | null
): string {
  const { g } = ctx;
  const s = span.style;
  const font = fontOf(s.faceId, s.sizePt, s.synthBold, s.synthItalic);
  if (font !== last) g.font = font;
  const color = cssColor(s.color);
  g.fillStyle = color;
  const text = span.text;
  const [from, to] = glyphsInClip(span, clip);
  for (let i = from; i < to;) {
    const cp = text.codePointAt(i)!;
    const n = cp > 0xffff ? 2 : 1;
    if (cp !== 32) g.fillText(text.slice(i, i + n), span.positions[i] ?? span.x, baseline);
    i += n;
  }
  let end = text.length;
  while (end > 0 && text[end - 1] === ' ') end -= 1;
  const x0 = span.positions[0] ?? span.x;
  const x1 = end >= text.length ? span.x + span.width : (span.positions[end] ?? x0);
  if (x1 <= x0) return font;
  const m = s.metrics;
  const thickness = Math.max(m.underlineThickness * s.sizePt, s.sizePt / 20);
  if (s.underline !== 'none') {
    const y = baseline - m.underlinePosition * s.sizePt;
    g.fillRect(x0, y, x1 - x0, thickness);
    if (s.underline === 'double' || s.underline === 'doubleAccounting')
      g.fillRect(x0, y + thickness * 2, x1 - x0, thickness);
  }
  if (s.strike) g.fillRect(x0, baseline - s.sizePt * 0.3 - thickness / 2, x1 - x0, thickness);
  return font;
}

/**
 * The range of the code units of `span` whose glyphs can reach `clip`, as `[from, to)`. A glyph
 * can overhang its advance, so a glyph is in range when its advance, widened by a full font
 * size on each side, meets the clip. The positions of a span grow from left to right, so a
 * binary search finds each end.
 */
function glyphsInClip(span: TextSpan, clip: CellTextBox['clip'] | null): [number, number] {
  const { text, positions } = span;
  if (!clip || positions.length !== text.length) return [0, text.length];
  const margin = span.style.sizePt;
  /** The first index whose position is past `x`, or at `x` when `orAt`. */
  const first = (x: number, orAt: boolean): number => {
    let lo = 0;
    let hi = positions.length;
    while (lo < hi) {
      const mid = (lo + hi) >>> 1;
      if (orAt ? positions[mid]! < x : positions[mid]! <= x) lo = mid + 1;
      else hi = mid;
    }
    return lo;
  };
  // The glyph before the first position at the left edge ends at that position.
  let from = Math.max(0, first(clip.x - margin, true) - 1);
  // A range never starts on the second half of a surrogate pair.
  const unit = text.charCodeAt(from);
  if (from > 0 && unit >= 0xdc00 && unit <= 0xdfff) from -= 1;
  return [from, first(clip.x + clip.width + margin, false)];
}

function texts(ctx: Ctx, pane: PaneLayout): number {
  const { g } = ctx;
  for (const box of pane.texts) {
    g.save();
    g.beginPath();
    g.rect(box.clip.x, box.clip.y, box.clip.width, box.clip.height);
    g.clip();
    textBox(ctx, box);
    g.restore();
  }
  return pane.texts.length;
}

/**
 * Whether no glyph of `span` can reach the clip of its text box. A glyph can overhang its advance,
 * so the span must lie a full font size away. A cell of many runs paints only the runs in view.
 */
const outsideClip = (span: TextSpan, clip: CellTextBox['clip']): boolean =>
  span.x + span.width + span.style.sizePt < clip.x ||
  span.x - span.style.sizePt > clip.x + clip.width;

function textBox(ctx: Ctx, box: CellTextBox): void {
  const t = box.text;
  if (t.rotation) {
    ctx.g.translate(t.rotation.originX, t.rotation.originY);
    ctx.g.rotate((-t.rotation.angle * Math.PI) / 180);
  }
  // Each text box paints inside its own save and restore, and a restore puts back an earlier
  // font, so the last font set is known only inside this box. It starts unknown.
  let font = '';
  for (const line of t.lines)
    for (const span of line.spans) {
      // The clip is in the coordinates of the sheet, and a rotated span is not.
      if (t.rotation) font = spanText(ctx, span, line.baseline, font, null);
      else if (!outsideClip(span, box.clip))
        font = spanText(ctx, span, line.baseline, font, box.clip);
    }
}

function tracePath(
  g: G,
  p: Extract<DrawingLayout, { kind: 'shape' }>['geometry']['paths'][number]
) {
  g.beginPath();
  for (const c of p.commands) {
    if (c.op === 'M') g.moveTo(c.x, c.y);
    else if (c.op === 'L') g.lineTo(c.x, c.y);
    else if (c.op === 'C') g.bezierCurveTo(c.x1, c.y1, c.x2, c.y2, c.x, c.y);
    else if (c.op === 'Q') g.quadraticCurveTo(c.x1, c.y1, c.x, c.y);
    else g.closePath();
  }
}

/** The overlay of a path fill mode of `ST_PathFillMode`: a color and its opacity. */
function shade(mode: string): [string, number] | undefined {
  if (mode === 'lighten') return ['#FFFFFF', 0.4];
  if (mode === 'lightenLess') return ['#FFFFFF', 0.2];
  if (mode === 'darken') return ['#000000', 0.4];
  if (mode === 'darkenLess') return ['#000000', 0.2];
  return undefined;
}

function drawing(ctx: Ctx, d: DrawingLayout): void {
  const { g, px } = ctx;
  const { x, y, width: w, height: h } = d.rect;
  g.save();
  g.translate(x, y);
  if (d.rotation) {
    g.translate(w / 2, h / 2);
    g.rotate((d.rotation * Math.PI) / 180);
    g.translate(-w / 2, -h / 2);
  }
  if (d.flipH || d.flipV) {
    g.translate(d.flipH ? w : 0, d.flipV ? h : 0);
    g.scale(d.flipH ? -1 : 1, d.flipV ? -1 : 1);
  }
  if (d.kind === 'shape') {
    for (const p of d.geometry.paths) {
      tracePath(g, p);
      if (d.fill && p.fill !== 'none') {
        g.fillStyle = cssRgba(d.fill);
        g.fill();
        const overlay = shade(p.fill);
        if (overlay) {
          g.globalAlpha = overlay[1];
          g.fillStyle = overlay[0];
          g.fill();
          g.globalAlpha = 1;
        }
      }
      if (p.stroke && d.line) {
        g.strokeStyle = cssRgba(d.line.color);
        g.lineWidth = d.line.width;
        g.stroke();
      }
    }
  } else if (d.kind === 'picture') {
    const state = ctx.options.pictures?.get(d.partName, d.contentType);
    if (state?.kind === 'loading') ctx.pending += 1;
    else if (state?.kind === 'ready') g.drawImage(state.bitmap, 0, 0, w, h);
    else if (state?.kind === 'unsupported') {
      g.fillStyle = '#F2F2F2';
      g.fillRect(0, 0, w, h);
      g.strokeStyle = '#BFBFBF';
      g.lineWidth = px;
      g.strokeRect(px / 2, px / 2, w - px, h - px);
    }
  }
  g.restore();
}

function paintPane(ctx: Ctx, pane: PaneLayout): number {
  const { g } = ctx;
  const c = pane.clip;
  g.save();
  g.beginPath();
  g.rect(c.x, c.y, c.width, c.height);
  g.clip();
  if (ctx.options.gridlines) gridlines(ctx, pane);
  fills(ctx, pane);
  borders(ctx, pane);
  const count = texts(ctx, pane);
  for (const d of pane.drawings) drawing(ctx, d);
  g.restore();
  return count;
}

/**
 * Paint a sheet window into `canvas`: the backing store is sized to the paint area at the device
 * pixel ratio, and the CSS size to the paint area at `scale`. Paint again with a new layout to
 * scroll or zoom.
 */
export function paintSheet(
  canvas: HTMLCanvasElement,
  layout: SheetLayout,
  options: PaintOptions
): PaintStats {
  const w = Math.max(0, options.width);
  const h = Math.max(0, options.height);
  const wanted = options.scale * (options.devicePixelRatio ?? 1);
  const ratio = Math.min(
    wanted,
    MAX_CANVAS_SIDE / Math.max(1, w),
    MAX_CANVAS_SIDE / Math.max(1, h)
  );
  const bw = Math.max(1, Math.round(w * ratio));
  const bh = Math.max(1, Math.round(h * ratio));
  if (canvas.width !== bw) canvas.width = bw;
  if (canvas.height !== bh) canvas.height = bh;
  canvas.style.width = `${w * options.scale}px`;
  canvas.style.height = `${h * options.scale}px`;
  let cells = 0;
  let drawings = 0;
  for (const pane of layout.panes) {
    cells += pane.cells.length;
    drawings += pane.drawings.length;
  }
  const g = canvas.getContext('2d');
  if (!g) return { cells, texts: 0, drawings, pendingPictures: 0 };
  g.setTransform(ratio, 0, 0, ratio, 0, 0);
  g.fillStyle = '#FFFFFF';
  g.fillRect(0, 0, w, h);
  const ctx: Ctx = {
    g,
    options,
    px: 72 / layout.metrics.pixelsPerInch,
    ratio,
    pending: 0,
  };
  // The browser's own kerning would move glyphs that layout placed one by one.
  if ('fontKerning' in g) g.fontKerning = 'none';
  g.textBaseline = 'alphabetic';
  let textCount = 0;
  for (const pane of layout.panes) textCount += paintPane(ctx, pane);
  if (options.frozenLines) {
    g.fillStyle = '#808080';
    if (layout.frozenEdges.x !== undefined) g.fillRect(layout.frozenEdges.x, 0, ctx.px, h);
    if (layout.frozenEdges.y !== undefined) g.fillRect(0, layout.frozenEdges.y, w, ctx.px);
  }
  return { cells, texts: textCount, drawings, pendingPictures: ctx.pending };
}
