// The clip rectangle of each text run of a PDF page, read from its operator list (pdf.js
// `getOperatorList`). The fidelity harness matches a line that a cell clip cuts by the clips of
// both PDFs (`compareLines`), so the clip edges compare directly, with no raster test. The
// reference PDF writes a clip path for text that overflows its cell; its other text shares the
// clip of the print area. The printed PDF writes a clip path for the text of each cell.
//
// The reader tracks `save` and `restore`, the current transformation matrix, `clip` and
// `eoClip` with the path that follows them, the text matrix, and the text state (font size,
// character and word spacing, horizontal scale, leading, rise, and render mode), which `save`
// and `restore` keep with the graphics state (ISO 32000-1, 9.3.1). A clip path must be one
// axis-aligned rectangle in page space; any other clip path makes the clip of the text that it
// holds unknown (`'non-rectangular'`), and such text never matches as clipped. Nested clips
// intersect. PPTX Editor reads operator lists for `setFont` only (`pdf.ts`); it has no clip
// reader.
//
// A form XObject paints as a save and restore, under its form matrix and clipped to its bounding
// box, which pdf.js gives with `beginGroup` for a group form. Limits: the clip of a run is unknown
// (`'non-rectangular'`), so the run never matches as clipped, when its text rise is not zero, and
// after a text render mode that adds text to the clip path (modes 4 to 7). A clip path with no
// subpath is unknown too. A painted line that holds part of a line that a clip cuts on both sides
// stays missing (`compareLines`).

import { OPS } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { THRESHOLDS } from './thresholds.ts';

/** A rectangle in page points from the page's top-left: `[x0, y0, x1, y1]`. */
export type PageBox = readonly [number, number, number, number];

/** The clip of a text run: a rectangle, none, or a path that is not one rectangle. */
export type TextClipBox = PageBox | 'none' | 'non-rectangular';

/** One text-showing operator of a page: where it draws, and the clip that holds it. */
export interface TextRun {
  /** The left and right ends of the run's advance, in page points. */
  readonly x0: number;
  readonly x1: number;
  /** The baseline, from the page's top. */
  readonly baseline: number;
  readonly clip: TextClipBox;
  /**
   * The advance extents of the words of the run, left to right, in page points, when a white
   * space glyph splits it into two or more: the start of the advance of the first glyph of each
   * word, and the end of the advance of its last one. A glyph's side bearings are not read, so
   * the ink of a glyph can lie inside or past its advance. A run of one word has none; its word
   * is `x0` to `x1`.
   */
  readonly words?: readonly (readonly [number, number])[];
  /** The Unicode text of the run, with no white space at its ends. */
  readonly text?: string;
  /** The number of `restore` operators before the run: two runs with a `restore` between them
   * belong to two graphics states. */
  readonly state?: number;
  /**
   * The advance extent of each glyph of the run that is not white space, left to right, in page
   * points. `textRuns` always records it; a run that a test builds can leave it out.
   */
  readonly glyphs?: readonly (readonly [number, number])[];
}

type Matrix = readonly [number, number, number, number, number, number];

const IDENTITY: Matrix = [1, 0, 0, 1, 0, 0];

/** A matrix from operator arguments: six numbers, as an array or a typed array. */
const matrixOf = (v: unknown): Matrix => {
  const m = Array.from(v as ArrayLike<number>, Number);
  return m.length === 6 && m.every(Number.isFinite) ? (m as unknown as Matrix) : IDENTITY;
};

/** `a` then `b`, in the row-vector convention of PDF matrices. */
const mul = (a: Matrix, b: Matrix): Matrix => [
  a[0] * b[0] + a[1] * b[2],
  a[0] * b[1] + a[1] * b[3],
  a[2] * b[0] + a[3] * b[2],
  a[2] * b[1] + a[3] * b[3],
  a[4] * b[0] + a[5] * b[2] + b[4],
  a[4] * b[1] + a[5] * b[3] + b[5],
];

const apply = (m: Matrix, x: number, y: number): [number, number] => [
  m[0] * x + m[2] * y + m[4],
  m[1] * x + m[3] * y + m[5],
];

/** The text render modes from this one on add the text to the clip path (ISO 32000-1, 9.3.6). */
const ADD_TO_CLIP_MODE = 4;

/** The path commands of pdf.js `constructPath` data. */
const MOVE_TO = 0;
const LINE_TO = 1;
const CLOSE_PATH = 4;

/** Points of one closed path within this distance (points) share an edge. */
const SAME_EDGE_PT = THRESHOLDS.sameEdgePt;

/**
 * The rectangle in page points that a path of pdf.js `constructPath` data draws under `ctm`, or
 * undefined when the path is not one axis-aligned rectangle in page space. The four edges
 * alternate between horizontal and vertical, each with a length, so the path has two distinct x
 * values, two distinct y values, and an area above zero. A degenerate path, such as a line or a
 * point, reads as unknown.
 */
export function rectangleOf(
  data: ArrayLike<number> | null | undefined,
  ctm: Matrix,
  toPage: (x: number, y: number) => [number, number]
): PageBox | undefined {
  if (!data) return undefined;
  const points: [number, number][] = [];
  let i = 0;
  while (i < data.length) {
    const op = data[i]!;
    if (op === MOVE_TO || op === LINE_TO) {
      if (op === MOVE_TO && points.length) return undefined;
      points.push(toPage(...apply(ctm, data[i + 1]!, data[i + 2]!)));
      i += 3;
    } else if (op === CLOSE_PATH) i += 1;
    else return undefined;
  }
  if (points.length === 5) {
    const [first, last] = [points[0]!, points[4]!];
    if (Math.abs(first[0] - last[0]) > SAME_EDGE_PT || Math.abs(first[1] - last[1]) > SAME_EDGE_PT)
      return undefined;
    points.pop();
  }
  if (points.length !== 4) return undefined;
  const kinds: ('horizontal' | 'vertical')[] = [];
  for (let k = 0; k < 4; k += 1) {
    const [a, b] = [points[k]!, points[(k + 1) % 4]!];
    const dx = Math.abs(a[0] - b[0]) > SAME_EDGE_PT;
    const dy = Math.abs(a[1] - b[1]) > SAME_EDGE_PT;
    // Each edge moves along exactly one axis.
    if (dx === dy) return undefined;
    kinds.push(dx ? 'horizontal' : 'vertical');
  }
  if (kinds[0] === kinds[1] || kinds[1] === kinds[2] || kinds[2] === kinds[3]) return undefined;
  const xs = points.map((p) => p[0]);
  const ys = points.map((p) => p[1]);
  const box: PageBox = [Math.min(...xs), Math.min(...ys), Math.max(...xs), Math.max(...ys)];
  return box[2] - box[0] > SAME_EDGE_PT && box[3] - box[1] > SAME_EDGE_PT ? box : undefined;
}

/**
 * The clip of two clips in a row. Two rectangles that do not overlap, or whose overlap is not
 * wider and taller than `SAME_EDGE_PT`, leave no area: the clip reads as unknown
 * (`non-rectangular`), never as a box, as in `rectangleOf`.
 */
export const intersect = (a: TextClipBox, b: TextClipBox): TextClipBox => {
  if (a === 'non-rectangular' || b === 'non-rectangular') return 'non-rectangular';
  if (a === 'none') return b;
  if (b === 'none') return a;
  const box: PageBox = [
    Math.max(a[0], b[0]),
    Math.max(a[1], b[1]),
    Math.min(a[2], b[2]),
    Math.min(a[3], b[3]),
  ];
  return box[2] - box[0] > SAME_EDGE_PT && box[3] - box[1] > SAME_EDGE_PT ? box : 'non-rectangular';
};

/** A glyph of a `showText` argument, as pdf.js gives it. */
interface Glyph {
  readonly unicode?: string;
  readonly width?: number;
  /** A single-byte code 32, which word spacing applies to (ISO 32000-1, 9.3.3). */
  readonly isSpace?: boolean;
}

/** The text state that `save` and `restore` keep (ISO 32000-1, 9.3.1). */
interface TextState {
  readonly size: number;
  readonly leading: number;
  readonly charSpacing: number;
  readonly wordSpacing: number;
  readonly hScale: number;
  readonly rise: number;
  /** A render mode that adds text to the clip path is in effect. */
  readonly textClip: boolean;
}

const START_TEXT_STATE: TextState = {
  size: 0,
  leading: 0,
  charSpacing: 0,
  wordSpacing: 0,
  hScale: 1,
  rise: 0,
  textClip: false,
};

/**
 * The text runs of an operator list, with their clips. `toPage` maps a point of PDF user space to
 * page points from the page's top-left. The advance of a run counts its glyph widths, the spacing
 * of a `showSpacedText`, the character spacing, the word spacing of each space glyph (pdf.js
 * `isSpace`), and the horizontal scale. A run starts at its first glyph that is not a space and
 * ends at the end of its last one, without the character or word spacing after it. Glyph widths are
 * in thousandths of the font size, as for every font but a Type 3 font, whose font matrix the
 * reader does not read.
 */
export function textRuns(
  fnArray: readonly number[],
  argsArray: readonly unknown[],
  toPage: (x: number, y: number) => [number, number]
): TextRun[] {
  const runs: TextRun[] = [];
  let ctm: Matrix = IDENTITY;
  let clip: TextClipBox = 'none';
  let pendingClip = false;
  const stack: { ctm: Matrix; clip: TextClipBox; text: TextState }[] = [];
  let tm: Matrix = IDENTITY;
  let lm: Matrix = IDENTITY;
  let text: TextState = START_TEXT_STATE;
  const set = (change: Partial<TextState>): void => void (text = { ...text, ...change });
  let groupBox: ArrayLike<number> | undefined;
  let restores = 0;
  for (let i = 0; i < fnArray.length; i += 1) {
    const a = argsArray[i] as unknown[];
    switch (fnArray[i]) {
      case OPS.save:
        stack.push({ ctm, clip, text });
        break;
      case OPS.restore: {
        restores += 1;
        const s = stack.pop();
        if (s) ({ ctm, clip, text } = s);
        break;
      }
      case OPS.transform:
        ctm = mul(matrixOf(typeof a[0] === 'number' ? a : a[0]), ctm);
        break;
      case OPS.clip:
      case OPS.eoClip:
        pendingClip = true;
        break;
      case OPS.constructPath: {
        if (!pendingClip) break;
        pendingClip = false;
        const paths = a[1] as (ArrayLike<number> | null)[] | null;
        const box = paths && paths.length === 1 ? rectangleOf(paths[0]!, ctm, toPage) : undefined;
        clip = intersect(clip, box ?? 'non-rectangular');
        break;
      }
      case OPS.beginText:
        tm = IDENTITY;
        lm = IDENTITY;
        break;
      case OPS.setFont:
        set({ size: Number(a[1]) });
        break;
      case OPS.setCharSpacing:
        set({ charSpacing: Number(a[0]) });
        break;
      case OPS.setWordSpacing:
        set({ wordSpacing: Number(a[0]) });
        break;
      case OPS.setHScale:
        set({ hScale: Number(a[0]) / 100 });
        break;
      case OPS.setLeading:
        set({ leading: Number(a[0]) });
        break;
      case OPS.setTextRise:
        set({ rise: Number(a[0]) });
        break;
      case OPS.setTextRenderingMode:
        set({ textClip: Number(a[0]) >= ADD_TO_CLIP_MODE });
        break;
      case OPS.endText:
        // Text that a render mode adds to the clip path clips from here on.
        if (text.textClip) clip = intersect(clip, 'non-rectangular');
        break;
      case OPS.beginGroup:
        // pdf.js gives the bounding box of a group form with its group, not with the form.
        groupBox = (a?.[0] as { bbox?: ArrayLike<number> | null } | undefined)?.bbox ?? undefined;
        break;
      case OPS.paintFormXObjectBegin: {
        // A form XObject paints inside a save and restore of the graphics state, under its form
        // matrix, clipped to its bounding box (ISO 32000-1, 8.10.1).
        stack.push({ ctm, clip, text });
        const [matrix, own] = (a ?? []) as [unknown, ArrayLike<number> | null | undefined];
        const bbox = own ?? groupBox;
        groupBox = undefined;
        if (matrix) ctm = mul(matrixOf(matrix as ArrayLike<number>), ctm);
        if (bbox && bbox.length === 4) {
          const [x0, y0, x1, y1] = [bbox[0]!, bbox[1]!, bbox[2]!, bbox[3]!];
          const path = [
            MOVE_TO,
            x0,
            y0,
            LINE_TO,
            x1,
            y0,
            LINE_TO,
            x1,
            y1,
            LINE_TO,
            x0,
            y1,
            CLOSE_PATH,
          ];
          clip = intersect(clip, rectangleOf(path, ctm, toPage) ?? 'non-rectangular');
        }
        break;
      }
      case OPS.paintFormXObjectEnd: {
        const s = stack.pop();
        if (s) ({ ctm, clip, text } = s);
        break;
      }
      case OPS.setTextMatrix:
        tm = matrixOf(typeof a[0] === 'number' ? a : a[0]);
        lm = tm;
        break;
      case OPS.moveText:
        lm = mul([1, 0, 0, 1, Number(a[0]), Number(a[1])], lm);
        tm = lm;
        break;
      case OPS.setLeadingMoveText:
        set({ leading: -Number(a[1]) });
        lm = mul([1, 0, 0, 1, Number(a[0]), Number(a[1])], lm);
        tm = lm;
        break;
      case OPS.nextLine:
        lm = mul([1, 0, 0, 1, 0, -text.leading], lm);
        tm = lm;
        break;
      case OPS.showText:
      case OPS.showSpacedText: {
        let advance = 0;
        // The start of the first glyph that is not white space, and the end of the last one,
        // without the character spacing after it. A glyph with no Unicode value counts as a
        // glyph: only a white space glyph is a space.
        let start: number | undefined;
        let lastEnd = 0;
        let glyphs = 0;
        // The advance extents of the words, in text space: a white space glyph ends a word.
        const words: [number, number][] = [];
        let word: [number, number] | undefined;
        const marks: [number, number][] = [];
        let chars = '';
        for (const g of a[0] as (Glyph | number | null)[]) {
          if (typeof g === 'number') advance -= (g / 1000) * text.size;
          else if (g) {
            const width = ((g.width ?? 0) / 1000) * text.size;
            const blank = Boolean(g.unicode) && !g.unicode!.trim();
            chars += g.unicode ?? '';
            if (!blank) {
              glyphs += 1;
              start ??= advance;
              lastEnd = advance + width;
              marks.push([advance, lastEnd]);
              if (word) word[1] = lastEnd;
              else words.push((word = [advance, lastEnd]));
            } else word = undefined;
            advance += width + text.charSpacing + (g.isSpace ? text.wordSpacing : 0);
          }
        }
        advance *= text.hScale;
        const m = mul(tm, ctm);
        const [x0, y] = toPage(...apply(m, (start ?? 0) * text.hScale, 0));
        const [x1] = toPage(...apply(m, lastEnd * text.hScale, 0));
        // A run that a text rise or a text clip places has no known clip.
        const known = text.rise === 0 && !text.textClip;
        const pageWords = words.map(([w0, w1]): [number, number] => {
          const [p0] = toPage(...apply(m, w0 * text.hScale, 0));
          const [p1] = toPage(...apply(m, w1 * text.hScale, 0));
          return [Math.min(p0, p1), Math.max(p0, p1)];
        });
        pageWords.sort((p, q) => p[0] - q[0]);
        const pageGlyphs = marks.map(([g0, g1]): [number, number] => {
          const [p0] = toPage(...apply(m, g0 * text.hScale, 0));
          const [p1] = toPage(...apply(m, g1 * text.hScale, 0));
          return [Math.min(p0, p1), Math.max(p0, p1)];
        });
        pageGlyphs.sort((p, q) => p[0] - q[0]);
        if (glyphs)
          runs.push({
            x0: Math.min(x0, x1),
            x1: Math.max(x0, x1),
            baseline: y,
            clip: known ? clip : 'non-rectangular',
            ...(pageWords.length > 1 ? { words: pageWords } : {}),
            glyphs: pageGlyphs,
            text: chars.trim(),
            state: restores,
          });
        tm = mul([1, 0, 0, 1, advance, 0], tm);
        break;
      }
    }
  }
  return runs;
}

/** A point within this distance (points) of a run's start is on that run. */
const ON_RUN_PT = THRESHOLDS.onRunPt;

/**
 * The clip of the run that draws a text item that starts at `x` on the baseline `y`: of the runs
 * on that baseline whose advance holds `x`, the one that starts last. Undefined when no run
 * holds it.
 */
export function clipOfItem(
  runs: readonly TextRun[],
  x: number,
  y: number,
  sameRowPt: number
): TextClipBox | undefined {
  let found: TextRun | undefined;
  for (const r of runs) {
    if (Math.abs(r.baseline - y) > sameRowPt) continue;
    if (x < r.x0 - ON_RUN_PT || x > r.x1) continue;
    if (!found || r.x0 > found.x0) found = r;
  }
  return found?.clip;
}

/** A part of a text item: its text, the start of its first glyph, and its advance width. */
export interface ItemPart {
  readonly text: string;
  readonly x: number;
  readonly width: number;
}

/** The text with no white space, which item text and run text compare by. */
const compact = (text: string): string => text.replace(/\s+/gu, '');

/**
 * The parts of a text item that starts at `x` on the baseline `y`, split where two neighboring
 * runs under it have different clips, or a `restore` between them. pdf.js 6 joins the runs of
 * one font on one line into one item when they are close, also when each run is drawn in a
 * graphics state of its own, as a cell is; pdf.js 5 ended the item at the `setFont` of the next
 * state. The split gives each cell its own item again, so `clipOfItem` finds the clip of each
 * part. The item stays whole when its runs do not spell its text.
 */
export function splitItemByRuns(
  runs: readonly TextRun[],
  item: ItemPart,
  y: number,
  sameRowPt: number
): ItemPart[] {
  const under = runs
    .filter(
      (r) =>
        Math.abs(r.baseline - y) <= sameRowPt &&
        r.x0 >= item.x - ON_RUN_PT &&
        r.x1 <= item.x + item.width + ON_RUN_PT
    )
    .sort((p, q) => p.x0 - q.x0);
  if (under.length < 2 || under.some((r) => r.text === undefined)) return [item];
  if (compact(under.map((r) => r.text).join('')) !== compact(item.text)) return [item];
  const groups: TextRun[][] = [];
  for (const r of under) {
    const last = groups.at(-1);
    const prev = last?.at(-1);
    if (last && prev && prev.state === r.state && !otherClip(prev.clip, r.clip)) last.push(r);
    else groups.push([r]);
  }
  if (groups.length < 2) return [item];
  const parts: ItemPart[] = [];
  let at = 0;
  for (const group of groups) {
    // Take as many characters that are not white space as the runs of the group draw.
    let need = compact(group.map((r) => r.text).join('')).length;
    let end = at;
    while (end < item.text.length && need > 0) {
      if (/\S/u.test(item.text[end]!)) need -= 1;
      end += 1;
    }
    const x0 = group[0]!.x0;
    parts.push({ text: item.text.slice(at, end).trim(), x: x0, width: group.at(-1)!.x1 - x0 });
    at = end;
  }
  return parts;
}

/** Whether two clips of text items are known and differ: their texts belong to two cells. */
export function otherClip(a: TextClipBox | undefined, b: TextClipBox | undefined): boolean {
  if (a === undefined || b === undefined || a === b) return false;
  if (typeof a === 'string' || typeof b === 'string') return true;
  return a.some((v, i) => Math.abs(v - b[i]!) > SAME_EDGE_PT);
}
