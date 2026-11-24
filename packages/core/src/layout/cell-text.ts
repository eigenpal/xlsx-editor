// The text of one cell, laid out in its box (design D6): horizontal and vertical alignment,
// wrapping, indent, shrinking, rotation, and stacked text, in points on the pixel grid of a
// profile (metrics.ts). The alignment is the resolved `x:alignment` (`CT_CellAlignment`, sml.xsd):
// `@horizontal` (`ST_HorizontalAlignment`), `@vertical` (`ST_VerticalAlignment`), `@wrapText`,
// `@indent`, `@shrinkToFit`, and `@textRotation` (`ST_TextRotation`). The schema gives the value
// sets; the offsets and sizes in this file are measured.
//
// Measured on the reference PDFs (basic-values, number-formats, grid-structure, text-layout;
// quick reference, "Text in cells"):
// - Each cell owns the gridline pixel at its left and top edges (the grid of a printed page
//   starts at the margin). The text keeps a margin of `floor(MDW / 3)` pixels of its first
//   run's font on each side: Carlito 11 starts 3 points inside the cell (the gridline and
//   2 pixels) and ends 2 points inside it; Carlito 8 and 6 keep 1 pixel, Carlito 20 keeps 3.
// - Centered and distributed text starts at `left + round((free - width) / 2)`, half up.
// - `general` aligns text left, numbers right, and booleans and errors center.
// - An indent level is one maximum digit width of the default font, from the left for `left`,
//   `justify`, and `distributed`, and from the right for `right` and `distributed`.
// - `fill` repeats the text as often as it fits whole. `justify` spreads the spaces of each line
//   but the last; `distributed` spreads every line and centers a line without a space.
//   `centerContinuous` centers across the empty neighbor cells that share the alignment.
// - Lines of wrapped text are one row height apart: 14 points for Carlito 11 in print. A line's
//   baseline is its ascent below the top of the line. Top-aligned text starts below the top
//   gridline pixel; centered text centers the lines with their row heights; bottom-aligned text
//   has the descent of its last line above the bottom of the cell, and at least the descent of
//   the default font (Carlito 6 and 8 end 3 points above it, as Carlito 11 does).
// - `shrinkToFit` scales the font down to whole points: 11 points shrink to 6.
// - Text widths are whole pixels: each glyph advance rounds on its own (`pixelWidth`).
//
// - Rotated text aligns the bounding box of its line (see `rotated`); bottom-aligned text that
//   runs up or down, and stacked text, keep the text margin above the bottom of its cell.
//
// The approximation "Measured text offsets" names what differs from the PDFs by a measured amount:
// the horizontal place of stacked text and the vertical centering in a merged range. Tests pin each
// difference (sheet-layout.test.ts).

import type { ResolvedAlignment, ResolvedFont, Rgb } from '../cascade/index.ts';
import { breakOpportunities } from './text/line-break.ts';
import { fitsWidth } from './number-fit.ts';
import type { FaceMetrics, TextMeasurer } from './text/measurer.ts';
import {
  autoRowHeight,
  fontPixels,
  rowSize,
  type FontPixels,
  type GridProfile,
} from './metrics.ts';

/** A text of at least this length keeps its break opportunities between layouts. */
const LONG_TEXT_CHARS = 64 * 1024;
/** The long texts that keep their break opportunities. */
const LONG_TEXTS_KEPT = 2;
/** The break opportunities of the last long texts, oldest first. */
const longTextBreaks = new Map<string, Uint8Array>();

/** The break opportunities of `text`. Each paint of a very long cell reads the same ones. */
export function breaksOf(text: string): Uint8Array {
  if (text.length < LONG_TEXT_CHARS) return breakOpportunities(text);
  const known = longTextBreaks.get(text) ?? breakOpportunities(text);
  longTextBreaks.delete(text);
  longTextBreaks.set(text, known);
  if (longTextBreaks.size > LONG_TEXTS_KEPT)
    longTextBreaks.delete(longTextBreaks.keys().next().value as string);
  return known;
}

/**
 * Forget the kept break opportunities, so no text of a closed workbook stays in memory. The
 * editor calls it when it releases a workbook. The editors of a page share the cache, so the
 * next paint of a long text in another editor only finds its breaks again.
 */
export function forgetLongTexts(): void {
  longTextBreaks.clear();
}

/** The number of long texts that keep their break opportunities, for the tests. */
export function keptLongTexts(): number {
  return longTextBreaks.size;
}

/** A rectangle in points. */
export interface Rect {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
}

/** One run of cell text with its font. */
export interface TextRunInput {
  readonly text: string;
  readonly font: ResolvedFont;
  /** Spaces of `text` that are as wide as another character (`FormattedValue.pads`). */
  readonly pads?: readonly TextPad[];
}

/** A space at `index` of a run's text that is as wide as `char` in the run's font. */
export interface TextPad {
  readonly index: number;
  readonly char: string;
}

/** What kind of value the text shows: `general` alignment depends on it. */
export type ValueKind = 'text' | 'number' | 'boolean' | 'error';

/** The kind of a cell value for its alignment: a date aligns as a number. */
export function valueKind(v: { readonly kind: string }): ValueKind {
  if (v.kind === 'number' || v.kind === 'date') return 'number';
  if (v.kind === 'boolean' || v.kind === 'error') return v.kind;
  return 'text';
}

export interface CellTextInput {
  readonly runs: readonly TextRunInput[];
  readonly kind: ValueKind;
  /** The kind that `general` alignment follows, when it is not `kind` (`FormattedValue.alignAsText`). */
  readonly alignKind?: ValueKind;
  readonly alignment: ResolvedAlignment;
  /** The cell, merged range, or centered-across span that holds the text. */
  readonly box: Rect;
  /** The color of the number format section, over the font color. */
  readonly color?: Rgb;
}

/** What a painter needs to draw a span: the measured face and the font properties. */
export interface TextStyle {
  /** The id of the measured face (`ResolvedFace.id`): paint draws exactly these bytes. */
  readonly faceId: string;
  readonly family: string;
  readonly sizePt: number;
  readonly bold: boolean;
  readonly italic: boolean;
  readonly synthBold: boolean;
  readonly synthItalic: boolean;
  readonly underline: ResolvedFont['underline'];
  readonly strike: boolean;
  readonly color: Rgb;
  /** The metrics of the measured face, in em: paint draws underlines and strikes from them. */
  readonly metrics: FaceMetrics;
}

export interface TextSpan {
  readonly text: string;
  /** Left edge of the span, in points. */
  readonly x: number;
  readonly width: number;
  /**
   * The left edge of each UTF-16 unit of `text`, in points: whole-pixel advances, as layout
   * measured them, so paint draws each glyph where layout put it.
   */
  readonly positions: readonly number[];
  readonly style: TextStyle;
}

export interface TextLine {
  readonly x: number;
  readonly baseline: number;
  readonly width: number;
  readonly spans: readonly TextSpan[];
  /** Extra space after each space character, for `justify` and `distributed`. */
  readonly wordSpacing?: number;
}

export interface CellTextLayout {
  /** In points. With `rotation`, relative to the rotation origin. */
  readonly lines: readonly TextLine[];
  /** Counterclockwise degrees and the point that line coordinates are relative to. */
  readonly rotation?: {
    readonly angle: number;
    readonly originX: number;
    readonly originY: number;
  };
  /** The leftmost and rightmost point that the text covers. */
  readonly left: number;
  readonly right: number;
  /** The text area of the box, in points: where text shows without overflow. */
  readonly areaLeft: number;
  readonly areaRight: number;
  /** One unwrapped, unrotated line of text that may show across empty neighbor cells. */
  readonly mayOverflow: boolean;
}

export interface TextContext {
  readonly measurer: TextMeasurer;
  readonly profile: GridProfile;
  /** The maximum digit width of the default font, in pixels: the indent unit. */
  readonly maxDigitWidth: number;
  /**
   * The descent of the default font, in pixels. Bottom-aligned text keeps at least this much
   * below its last baseline (text-layout.pdf: Carlito 6 and 8 end 3 points above the bottom of
   * their cells, as Carlito 11 does). Zero when absent.
   */
  readonly minDescent?: number;
}

/** At most this many lines are laid out for one cell; the rest cannot show. */
export const MAX_CELL_LINES = 500;
/** An indent level is this many maximum digit widths of the default font. */
const INDENT_DIGITS = 1;
/** The left gridline pixel that each cell owns. */
const LEFT_GRIDLINE = 1;

/** The margin on each side of the text, in pixels, for a font's maximum digit width. */
export const textMargin = (maxDigitWidth: number): number => Math.floor(maxDigitWidth / 3);
const TOP_GRIDLINE = 1;

interface Measured {
  readonly text: string;
  /** Whole-pixel width of each UTF-16 unit. */
  readonly widths: Float64Array;
  /** The run of each unit. */
  readonly runOf: Uint16Array;
  readonly fonts: readonly FontPixels[];
  /** The line pitch of each run's font: the row height of that font, in pixels. */
  readonly pitches: readonly number[];
  readonly styles: readonly TextStyle[];
}

/**
 * Measure the runs of a cell. A line break (LF or CR) takes no width. With `breakAsSpace`, a
 * run of line breaks between two other characters takes the width of one space of the font of
 * its first break, so the words around it stay apart; the painter draws no glyph for it.
 */
function measureRuns(
  input: CellTextInput,
  ctx: TextContext,
  scale: number,
  breakAsSpace = false
): Measured {
  let text = '';
  for (const run of input.runs) text += run.text;
  const widths = new Float64Array(text.length);
  const runOf = new Uint16Array(text.length);
  const fonts: FontPixels[] = [];
  const pitches: number[] = [];
  const styles: TextStyle[] = [];
  let at = 0;
  input.runs.forEach((run, r) => {
    const font = scale === 1 ? run.font : { ...run.font, sizePt: run.font.sizePt * scale };
    const fp = fontPixels(ctx.measurer, font, ctx.profile);
    fonts.push(fp);
    pitches.push(
      (rowSize(autoRowHeight(fp, ctx.profile), ctx.profile) * ctx.profile.pixelsPerInch) / 72
    );
    styles.push({
      faceId: fp.face.id,
      family: fp.face.family,
      sizePt: font.sizePt,
      bold: font.bold,
      italic: font.italic,
      synthBold: fp.face.synthBold,
      synthItalic: fp.face.synthItalic,
      underline: font.underline,
      strike: font.strike,
      color: input.color ?? font.color,
      metrics: fp.face.metrics,
    });
    const advances = ctx.measurer.advances(fp.face, run.text);
    for (let i = 0; i < run.text.length; i += 1) {
      const c = run.text.charCodeAt(i);
      const a = advances[i]!;
      widths[at + i] = isBreak(c) || a === 0 ? 0 : Math.round(a * fp.sizePx);
      runOf[at + i] = Math.min(r, 0xffff);
    }
    for (const pad of run.pads ?? [])
      if (pad.index >= 0 && pad.index < run.text.length)
        widths[at + pad.index] = padWidth(ctx.measurer, fp, pad.char);
    at += run.text.length;
  });
  if (breakAsSpace) spaceBreaks(text, widths, runOf, fonts, ctx.measurer);
  return { text, widths, runOf, fonts, pitches, styles };
}

/**
 * Give each run of line breaks between two other characters the width of one space of the font
 * of its first break. Breaks at the start or the end of the text keep no width.
 */
function spaceBreaks(
  text: string,
  widths: Float64Array,
  runOf: Uint16Array,
  fonts: readonly FontPixels[],
  measurer: TextMeasurer
): void {
  for (let i = 0; i < text.length; i += 1) {
    if (!isBreak(text.charCodeAt(i))) continue;
    let j = i;
    while (j < text.length && isBreak(text.charCodeAt(j))) j += 1;
    const font = fonts[runOf[i]!];
    if (i > 0 && j < text.length && font) widths[i] = padWidth(measurer, font, ' ');
    i = j;
  }
}

const isSpace = (c: number) => c === 32 || c === 9 || c === 0x3000;
const isBreak = (c: number) => c === 10 || c === 13;

/** Line ranges `[start, end)` of a measured text. */
function breakLines(
  m: Measured,
  limit: number,
  wrap: boolean,
  stacked: boolean
): [number, number][] {
  const { text, widths } = m;
  const n = text.length;
  if (!wrap && !stacked) return [[0, n]];
  const ops = breaksOf(text);
  const lines: [number, number][] = [];
  let pos = 0;
  while (pos < n && lines.length < MAX_CELL_LINES) {
    if (stacked) {
      let end = pos + 1;
      while (end < n && widths[end] === 0 && !isBreak(text.charCodeAt(end))) end += 1;
      if (!isBreak(text.charCodeAt(pos))) lines.push([pos, end]);
      pos = end;
      continue;
    }
    let width = 0;
    let lastBreak = -1;
    let end = n;
    let next = n;
    for (let i = pos; i < n; i += 1) {
      const c = text.charCodeAt(i);
      if (c === 10 || (c === 13 && text.charCodeAt(i + 1) !== 10)) {
        end = i;
        next = i + 1;
        break;
      }
      width += widths[i]!;
      if (!isSpace(c) && width > limit && i > pos) {
        end = lastBreak > pos ? lastBreak : i;
        next = end;
        break;
      }
      if (ops[i + 1]) lastBreak = i + 1;
    }
    lines.push([pos, end]);
    pos = Math.max(next, pos + 1);
  }
  return lines;
}

interface Placed {
  start: number;
  end: number;
  /** Width without trailing spaces, pixels. */
  width: number;
  spaces: number;
  ascent: number;
  /** The tallest line height of the line's fonts. */
  height: number;
  /** The distance to the next line. */
  pitch: number;
}

function placeLine(m: Measured, start: number, end: number, trim: boolean): Placed {
  let last = end;
  if (trim) while (last > start && isSpace(m.text.charCodeAt(last - 1))) last -= 1;
  let width = 0;
  let spaces = 0;
  let ascent = 0;
  let height = 0;
  let pitch = 0;
  for (let i = start; i < last; i += 1) {
    width += m.widths[i]!;
    if (isSpace(m.text.charCodeAt(i))) spaces += 1;
  }
  const runs = new Set<number>();
  for (let i = start; i < end; i += 1) runs.add(m.runOf[i]!);
  if (runs.size === 0) runs.add(m.runOf[Math.min(start, m.runOf.length - 1)] ?? 0);
  for (const r of runs) {
    const f = m.fonts[r] ?? m.fonts[0]!;
    ascent = Math.max(ascent, f.ascent);
    height = Math.max(height, f.lineHeight);
    pitch = Math.max(pitch, m.pitches[r] ?? m.pitches[0]!);
  }
  return { start, end: last, width, spaces, ascent, height, pitch };
}

/** The spans of a line, from pixel `x`, in points (`k` points per pixel). */
function spansOf(m: Measured, line: Placed, x: number, k: number, gap: number): TextSpan[] {
  const spans: TextSpan[] = [];
  let i = line.start;
  let at = x;
  while (i < line.end) {
    const r = m.runOf[i]!;
    let j = i;
    let w = 0;
    let text = '';
    const positions: number[] = [];
    while (j < line.end && m.runOf[j] === r) {
      const c = m.text.charCodeAt(j);
      if (!isBreak(c)) {
        text += m.text[j];
        positions.push((at + w) * k);
      }
      w += m.widths[j]! + (isSpace(c) ? gap : 0);
      j += 1;
    }
    if (text.length) spans.push({ text, x: at * k, width: w * k, positions, style: m.styles[r]! });
    at += w;
    i = j;
  }
  return spans;
}

/** The horizontal rule that `general` stands for. */
function horizontalOf(alignment: ResolvedAlignment, kind: ValueKind): string {
  const h = alignment.horizontal;
  if (h !== 'general') return h;
  return kind === 'number' ? 'right' : kind === 'text' ? 'left' : 'center';
}

/**
 * The indent of each side of the text area, in pixels: one maximum digit width of the default
 * font for each level, from the left for `left`, `justify`, and `distributed`, and from the
 * right for `right` and `distributed`. Other alignments take no indent.
 */
function indentSides(
  a: ResolvedAlignment,
  horizontal: string,
  maxDigitWidth: number
): { left: number; right: number } {
  const indent = a.indent * INDENT_DIGITS * maxDigitWidth;
  return {
    left:
      horizontal === 'left' || horizontal === 'distributed' || horizontal === 'justify'
        ? indent
        : 0,
    right: horizontal === 'right' || horizontal === 'distributed' ? indent : 0,
  };
}

/** The width in points that the indent of a cell takes from its text area. */
export function indentWidth(
  alignment: ResolvedAlignment,
  kind: ValueKind,
  ctx: TextContext
): number {
  const sides = indentSides(alignment, horizontalOf(alignment, kind), ctx.maxDigitWidth);
  return ((sides.left + sides.right) * 72) / ctx.profile.pixelsPerInch;
}

/** Lay out the text of one cell in its box. */
export function layoutCellText(input: CellTextInput, ctx: TextContext): CellTextLayout {
  const k = 72 / ctx.profile.pixelsPerInch;
  const a = input.alignment;
  const h = horizontalOf(a, input.alignKind ?? input.kind);
  const boxLeft = input.box.x / k;
  const boxRight = (input.box.x + input.box.width) / k;
  const top = input.box.y / k + TOP_GRIDLINE;
  const bottom = (input.box.y + input.box.height) / k;
  // `ST_TextRotation`: 1 to 90 runs up, 91 to 180 runs down (90 - value degrees), and 255 is
  // stacked text.
  const stacked = a.textRotation === 255;
  const angle =
    a.textRotation > 0 && a.textRotation <= 90
      ? a.textRotation
      : a.textRotation > 90 && a.textRotation <= 180
        ? -(a.textRotation - 90)
        : 0;
  // Only text breaks into lines: a number shows on one line, and a number that does not fit
  // shows its fitted form or the `#` fill (`fitNumber`). The `@wrapText` of `CT_CellAlignment`
  // wraps "the text in a cell" (ECMA-376 Part 1, 18.8.1).
  const wrap =
    !angle &&
    !stacked &&
    input.kind !== 'number' &&
    (a.wrapText || h === 'justify' || h === 'distributed');
  // A line break in text that does not wrap shows as a space (approximation "Text layout
  // without a measurement"). Stacked text breaks its lines at it.
  const breakAsSpace = !wrap && !stacked;

  const { left: leftIndent, right: rightIndent } = indentSides(a, h, ctx.maxDigitWidth);
  // The margin follows the font of the first run: 2 pixels for Carlito 11, 1 for Carlito 8 and
  // a 6 point shrunk text, 3 for Carlito 20 (text-layout.pdf).
  const areaOf = (fonts: readonly FontPixels[]) => {
    const pad = textMargin(fonts[0]?.maxDigitWidth ?? ctx.maxDigitWidth);
    const l = boxLeft + LEFT_GRIDLINE + pad;
    const r = boxRight - pad;
    return { l, r, left: l + leftIndent, right: r - rightIndent };
  };
  let m = measureRuns(input, ctx, 1, breakAsSpace);
  let area = areaOf(m.fonts);
  if (a.shrinkToFit && !wrap && !angle && !stacked) {
    let total = 0;
    for (const w of m.widths) total += w;
    const room = Math.max(0, area.right - area.left);
    // The same test as a fit (`fitsWidth`), so the noise of a column offset shrinks nothing.
    if (!fitsWidth(total, room, false) && total > 0) {
      // The size rounds down to whole points: 11 points shrink to 6, not 6.46 (text-layout.pdf).
      const size = input.runs[0]?.font.sizePt ?? 11;
      const scale = Math.max(1, Math.floor(size * (room / total))) / size;
      m = measureRuns(input, ctx, scale, breakAsSpace);
      area = areaOf(m.fonts);
    }
  }
  const { left, right } = area;
  const avail = Math.max(0, right - left);
  if (h === 'fill' && !wrap && !angle && !stacked && m.text.length > 0) {
    let w = 0;
    for (const x of m.widths) w += x;
    const copies = w > 0 ? Math.max(1, Math.min(MAX_CELL_LINES, Math.floor(avail / w))) : 1;
    if (copies > 1)
      m = measureRuns(
        { ...input, runs: Array.from({ length: copies }, () => input.runs).flat() },
        ctx,
        1,
        breakAsSpace
      );
  }

  const ranges = breakLines(m, avail, wrap, stacked);
  const placed = ranges.map(([s, e]) => placeLine(m, s, e, wrap));
  const total = placed.reduce((sum, p) => sum + p.pitch, 0);
  const last = placed[placed.length - 1];
  const inked = last ? total - last.pitch + last.height : 0;

  // Bottom-aligned text that runs up or down, and stacked text, keep the text margin above the
  // bottom of the cell (text-layout.pdf: "Up 90" starts 2 points above it, and stacked text ends
  // 2 points higher). Only 90 degrees at the bottom is measured: -90 degrees (`textRotation`
  // 180) follows it without a measurement (approximation "Rotated text without a measurement").
  const pad = textMargin(m.fonts[0]?.maxDigitWidth ?? ctx.maxDigitWidth);
  if (angle) {
    // Only text is measured with rotation: numbers, booleans, and errors keep their `general`
    // rule of unrotated cells, without a measurement.
    const rotatedText = a.horizontal === 'general' && (input.alignKind ?? input.kind) === 'text';
    const marginAbove = Math.abs(angle) === 90 && a.vertical === 'bottom';
    return rotated(m, placed[0]!, angle, rotatedText ? rotatedGeneral(angle) : h, a.vertical, {
      left,
      right,
      top,
      bottom: marginAbove ? bottom - pad : bottom,
      k,
    });
  }

  const vertical = a.vertical;
  const y0 =
    vertical === 'top' || vertical === 'justify'
      ? top
      : vertical === 'center' || vertical === 'distributed'
        ? top + Math.round((bottom - top - total) / 2)
        : bottom - inked - (stacked ? pad : 0) - bottomRaise(last, ctx);
  const lines: TextLine[] = [];
  let y = y0;
  let minX = Infinity;
  let maxX = -Infinity;
  placed.forEach((p, i) => {
    const lastLine = i === placed.length - 1;
    let x = left;
    let gap = 0;
    const lineH = stacked && a.horizontal === 'general' ? 'center' : h;
    if (lineH === 'right') x = right - p.width;
    else if (lineH === 'center' || lineH === 'centerContinuous')
      x = left + Math.round((avail - p.width) / 2);
    else if ((lineH === 'justify' && !lastLine) || lineH === 'distributed') {
      if (p.spaces > 0) gap = (avail - p.width) / p.spaces;
      else if (lineH === 'distributed') x = left + Math.round((avail - p.width) / 2);
    }
    const width = p.width + gap * p.spaces;
    lines.push({
      x: x * k,
      baseline: (y + p.ascent) * k,
      width: width * k,
      spans: spansOf(m, p, x, k, gap),
      ...(gap ? { wordSpacing: gap * k } : {}),
    });
    minX = Math.min(minX, x);
    maxX = Math.max(maxX, x + width);
    y += p.pitch;
  });
  const single = !wrap && !stacked && placed.length === 1;
  return {
    lines,
    left: (Number.isFinite(minX) ? minX : left) * k,
    right: (Number.isFinite(maxX) ? maxX : left) * k,
    areaLeft: area.l * k,
    areaRight: area.r * k,
    mayOverflow: single && input.kind === 'text' && h !== 'fill' && !a.shrinkToFit,
  };
}

/**
 * The horizontal rule that `general` stands for in rotated text (text-layout.pdf): text turned
 * counterclockwise by less than 90 degrees aligns left, text turned up by 90 degrees or turned
 * clockwise aligns right. Measured for text values at 45, 90, and -45 degrees only.
 */
const rotatedGeneral = (angle: number): string => (angle > 0 && angle < 90 ? 'left' : 'right');

/**
 * One rotated line: its bounding box, from the ascent to the descent of the line, aligned in the
 * text area. Measured on text-layout.pdf at 45, 90, and -45 degrees:
 * - Left-aligned text starts its box at the left of the text area. Right-aligned text ends its
 *   box 1 pixel inside the right of the text area.
 * - Bottom-aligned text puts the top-left corner of its line (the start of the line, at the
 *   ascent) on a whole pixel, at or above the place where the box meets the bottom.
 */
function rotated(
  m: Measured,
  p: Placed,
  angle: number,
  h: string,
  vertical: string,
  area: { left: number; right: number; top: number; bottom: number; k: number }
): CellTextLayout {
  const { k } = area;
  const descent = p.height - p.ascent;
  const rad = (angle * Math.PI) / 180;
  const cos = Math.cos(rad);
  const sin = Math.sin(rad);
  const xs: number[] = [];
  const ys: number[] = [];
  for (const [x, y] of [
    [0, -p.ascent],
    [p.width, -p.ascent],
    [0, descent],
    [p.width, descent],
  ] as const) {
    // Counterclockwise on a page whose y axis points down.
    xs.push(x * cos + y * sin);
    ys.push(-x * sin + y * cos);
  }
  const minX = Math.min(xs[0]!, xs[1]!, xs[2]!, xs[3]!);
  const maxX = Math.max(xs[0]!, xs[1]!, xs[2]!, xs[3]!);
  const minY = Math.min(ys[0]!, ys[1]!, ys[2]!, ys[3]!);
  const maxY = Math.max(ys[0]!, ys[1]!, ys[2]!, ys[3]!);
  const bw = maxX - minX;
  const bh = maxY - minY;
  const bx =
    h === 'right'
      ? area.right - RIGHT_ROTATED - bw
      : h === 'center' || h === 'centerContinuous' || h === 'distributed'
        ? area.left + (area.right - area.left - bw) / 2
        : area.left;
  const by =
    vertical === 'top' || vertical === 'justify'
      ? area.top
      : vertical === 'center' || vertical === 'distributed'
        ? area.top + (area.bottom - area.top - bh) / 2
        : area.bottom - bh;
  const originX = bx - minX;
  let originY = by - minY;
  if (vertical === 'bottom') {
    // The corner at the start of the line and its ascent: ys[0], relative to the origin.
    const corner = Math.floor(originY + ys[0]! + SNAP_EPSILON);
    originY = corner - ys[0]!;
  }
  return {
    lines: [{ x: 0, baseline: 0, width: p.width * k, spans: spansOf(m, p, 0, k, 0) }],
    rotation: { angle, originX: originX * k, originY: originY * k },
    left: bx * k,
    right: (bx + bw) * k,
    areaLeft: area.left * k,
    areaRight: area.right * k,
    mayOverflow: false,
  };
}

/**
 * Right-aligned rotated text ends this many pixels inside the text area. Measured at 90 and -45
 * degrees ("Up 90" and "Down 45" in text-layout.pdf, both `general` text at the bottom).
 */
const RIGHT_ROTATED = 1;
/** Floating point noise that the whole-pixel snap of rotated text ignores, in pixels. */
const SNAP_EPSILON = 1e-9;

/** The whole-pixel width of one character, for a space that stands for it. */
function padWidth(measurer: TextMeasurer, fp: FontPixels, char: string): number {
  let w = 0;
  for (const a of measurer.advances(fp.face, char)) w += Math.round(a * fp.sizePx);
  return w;
}

/** How far a last line whose descent is under the default font's rises to keep that descent. */
function bottomRaise(last: Placed | undefined, ctx: TextContext): number {
  if (!last) return 0;
  return Math.max(0, (ctx.minDescent ?? 0) - (last.height - last.ascent));
}

/** The width of plain text in one font, in points: for fitting numbers. */
export function textWidth(
  ctx: TextContext,
  text: string,
  font: ResolvedFont,
  pads: readonly TextPad[] = []
): number {
  const fp = fontPixels(ctx.measurer, font, ctx.profile);
  const advances = ctx.measurer.advances(fp.face, text);
  const widths = new Float64Array(advances.length);
  for (let i = 0; i < advances.length; i += 1) widths[i] = Math.round(advances[i]! * fp.sizePx);
  for (const pad of pads)
    if (pad.index >= 0 && pad.index < widths.length)
      widths[pad.index] = padWidth(ctx.measurer, fp, pad.char);
  let w = 0;
  for (const x of widths) w += x;
  return (w * 72) / ctx.profile.pixelsPerInch;
}

/** The width in points that a cell box leaves for text in a font, without indent. */
export function textArea(box: Rect, ctx: TextContext, font: ResolvedFont): number {
  const k = 72 / ctx.profile.pixelsPerInch;
  const pad = textMargin(fontPixels(ctx.measurer, font, ctx.profile).maxDigitWidth);
  return Math.max(0, box.width - (LEFT_GRIDLINE + 2 * pad) * k);
}
