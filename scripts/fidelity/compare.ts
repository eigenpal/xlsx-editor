// Copied from pptx-editor/scripts/fidelity/compare.ts at 1943d2c9 (EigenPal), and compared with
// its later commits: 67ae7309 does not change this file, and the line pairing of 06534363 is
// ported. Differences from 06534363:
// - The raster and line types are in `poppler.ts`, a page is a printed sheet page instead of a
//   slide, and `failureKinds` replaces `slidePasses`.
// - `lineKey` checks every line of two or more visible characters and every one-digit line,
//   numbers included, where the source needs two letters in a row: numbers are most of the text
//   of a sheet. A line without two letters in a row keeps its leading symbols, because a sign is
//   part of a number.
// - Each step that pairs by text takes its candidate pairs with `assign`, nearest pair first, as
//   the source does. The same text and the bullet form pair within `maxBaselineDeltaPt` (0.5 pt)
//   of a row, not within `LINE_PAIR_LIMIT_PT` (4 pt), so a line that a clip cuts does not take a
//   whole line of the same text on the next row. No step limits the nearer edge. A tie of the
//   reference lines falls to the painted lines, where the source leaves it to the input order.
// - The joins pair in both directions, as in the source, with no baseline rule between the pieces,
//   as in the source. A painted piece needs a key (`lineKey`), as the painted line of the converse
//   join does. Each joined text stays a start of the text of the line, so a start ends at the first
//   piece that does not continue it. Each piece lies within `maxBaselineDeltaPt` (0.5 pt) of the
//   baseline of the line of the other page, as one line does, where the source takes 0.25 pt
//   (`SAME_BASELINE_PT`). The painted page clips every cell, so a join is a common path here. The
//   baseline of the worst piece is scored, where the source scores the first piece. A join decides
//   only by the grouping rule of each page (`joinable`): where the items of the whole line split at
//   the pieces, its gap there is at most 1 em (`LINE_GAP_EM`), which restates the rule of
//   `groupLines` for the whole line, and the gap of the pieces is at most 1 em (`JOIN_GAP_EM`) plus
//   `joinGapPt` at each end; elsewhere the gap of the pieces is at most 1 em plus `joinGapPt`. The
//   source has no gap limit.
// - Each piece of a join is scored by the nearer of its start and end distances from the characters
//   that it holds in the line of the other page, in both directions, as one line is. The source
//   scores only the edges of the group. Here the edges of the group count only for a line of the
//   other page without items. A piece of the reference page must start and end on item edges of
//   the painted line. A painted piece scores the nearer of its edges that lie on item edges of the
//   reference line. When neither does, the width of the group is scored.
// - With clipping, each inner end of a painted piece of a join is checked by itself. An end inside
//   a reference item is named as a painted split inside a reference item, and no clip rule checks
//   it. An end on a boundary of two reference items gets the clip rules of a line end, against the
//   reference item on its side. The source has no clip rules.
// - A pass of `compareLines` (`clippedShift`) matches a line that a cell clip cuts: the printed
//   PDF leaves out a glyph that a clip hides entirely, and the reference PDF keeps it as text.
//   It compares the clip rectangles that both PDFs carry (`pdf-clips.ts`). A reference line that
//   its clip cuts also requires the painted line to be cut at the same edge (`LineScore.uncut`),
//   also when the texts match in full. At an end of a line that the reference clip holds whole,
//   the painted line must stay inside the painted clip (`LineScore.extraCut`), also when the
//   painted clip hides whole glyphs there. These clip rules are in `clip-rules.ts`. The clipped
//   pass and the extra-cut pass run after the joins and before the same-text step, and take the
//   reference lines in page order. The source has no clipped-line rule.
// - The same-text step takes the same text only, as the last step of the source does, at any
//   distance. So a join piece off the row leaves a missing line and extra lines, as in the
//   source. So does a join piece past the gap limit, which the source does not have.
// - Last, a clipped line that is left pairs with a cut of it on any row, nearest pair first, by
//   the rules of the clipped pass: a cut line that moved is out of place, as the last step of the
//   source holds for a whole line. The source has no clipped-line rule.
// - `differingPixelsIn` is not in the source.
//
// The per-page measurements of the fidelity harness: the pixels of the screenshot of the painted
// page against the raster of the reference page, and the text lines of the printed PDF against
// those of the reference PDF. Both sides' lines come from the same `groupLines` (`pdf-lines.ts`),
// so every rule applies to both alike. Pure functions; `run.ts` drives them.

import pixelmatch from 'pixelmatch';
import { PNG } from 'pngjs';
import {
  SIDES,
  cutsOurs,
  endClip,
  holdsEnd,
  holdsCut,
  keepsGlyphs,
  oursClipCuts,
  oursCut,
  reachesEdge,
  referenceCut,
  referenceEnd,
  type ClipCut,
  type ClipEvidence,
} from './clip-rules.ts';
import { LINE_GAP_EM } from './pdf-lines.ts';
import { normalize, type PdfLine, type Raster } from './poppler.ts';
import { pixelLimit, THRESHOLDS, type FailureKind } from './thresholds.ts';

export type { ClipCut, ClipEvidence } from './clip-rules.ts';

export interface PixelScore {
  /** Fraction of pixels that differ (pixelmatch, anti-aliasing tolerant). */
  readonly mismatch: number;
  /** The diff overlay as PNG, when asked for and some pixel differs. */
  readonly diff?: Buffer;
  /** The box around every differing pixel, `[x0, y0, x1, y1]` in pixels, or undefined. */
  readonly diffBox?: readonly [number, number, number, number];
}

/** Compare two rasters of the same size. */
export function comparePixels(ours: Raster, reference: Raster, withDiff = false): PixelScore {
  if (ours.width !== reference.width || ours.height !== reference.height)
    throw new Error('the two rasters have different sizes');
  const { width: w, height: h } = ours;
  const diff = new PNG({ width: w, height: h });
  const n = pixelmatch(ours.data, reference.data, diff.data, w, h, {
    threshold: THRESHOLDS.pixelThreshold,
  });
  return {
    mismatch: n / (w * h),
    diff: withDiff && n > 0 ? PNG.sync.write(diff, { filterType: 2 }) : undefined,
    diffBox: n ? differingBox(diff) : undefined,
  };
}

/** The box around the pixels that pixelmatch marks as different (pure red in its output). */
function differingBox(diff: PNG): [number, number, number, number] | undefined {
  let x0 = Infinity;
  let y0 = Infinity;
  let x1 = -1;
  let y1 = -1;
  for (let y = 0; y < diff.height; y += 1)
    for (let x = 0; x < diff.width; x += 1) {
      const i = (y * diff.width + x) * 4;
      if (diff.data[i] === 255 && diff.data[i + 1] === 0 && diff.data[i + 2] === 0) {
        if (x < x0) x0 = x;
        if (x > x1) x1 = x;
        if (y < y0) y0 = y;
        if (y > y1) y1 = y;
      }
    }
  return x1 < 0 ? undefined : [x0, y0, x1 + 1, y1 + 1];
}

/**
 * A bullet at the start of a line: symbols, punctuation, and private-use characters. A symbol
 * font can map one bullet glyph to other characters in each PDF, so line matching leaves bullets
 * out on both sides; the pixels check the glyph.
 */
const LEADING_BULLET = /^[\p{S}\p{P}\p{Co}]+\s*/u;

/**
 * Right-to-left scripts. A PDF text layer of shaped Arabic or Hebrew text holds glyph forms or
 * guessed characters, not the text, so these lines are checked by pixels only.
 */
const RTL = /[֐-ࣿיִ-﷿ﹰ-﻿]/u;

/** Control characters: a glyph that the PDF text layer cannot map. */
const UNMAPPED = /[\u0000-\u0008\u000E-\u001F�]/u;

/**
 * The text that line matching compares: normalized, without a leading bullet or whitespace. Empty
 * for a line that is not checked: right-to-left text, unmapped glyphs, and lines of fewer than two
 * visible characters (the one glyph per line of stacked text), except one digit: a one-digit number
 * is checked. Only a line with two letters in a row can start with a bullet: a line of numbers and
 * symbols keeps its leading signs and symbols (`-3,5`, `#####`).
 */
export function lineKey(text: string): string {
  const n = normalize(text);
  if (RTL.test(n) || UNMAPPED.test(n)) return '';
  const visible = n.replace(/\s+/g, '');
  if (Array.from(visible).length < 2 && !/^\p{Nd}$/u.test(visible)) return '';
  return /\p{L}{2}/u.test(n) ? n.replace(LEADING_BULLET, '').replace(/\s+/g, '') : visible;
}

/**
 * Whether a line holds one visible character other than a digit: `lineKey` gives it no key, so it
 * has no line score. Stacked text is such lines.
 */
export function isOneCharacter(text: string): boolean {
  const n = normalize(text);
  if (RTL.test(n) || UNMAPPED.test(n)) return false;
  const visible = n.replace(/\s+/g, '');
  return Array.from(visible).length === 1 && !/^\p{Nd}$/u.test(visible);
}

/**
 * The key without a first one-character word, for a line that starts with a list marker
 * (`PdfLine.marker`): a bullet that a symbol font writes as a letter (`q Hollow square`). Empty
 * for any other line, so a one-letter word of the text must match exactly. Used only when the
 * full keys do not match.
 */
function bulletless(line: PdfLine): string {
  if (!line.marker) return '';
  const n = normalize(line.text).replace(LEADING_BULLET, '');
  return /^\S\s/u.test(n) ? lineKey(n.slice(2)) : '';
}

export interface LineScore {
  /** Reference lines that the painted page holds. */
  readonly matched: number;
  /**
   * Of the matched lines, those that the painted page holds only in part, because a clip of its
   * cell hides the rest on both pages (`clippedShift`), each by `clippedKey`.
   */
  readonly clipped: readonly string[];
  /**
   * Reference lines that a clip cuts on the reference page and that the painted page does not hold
   * cut at the same edge (`cutsOurs`), each by `clippedKey`: the painted line runs past the edge,
   * the painted clip is elsewhere, or the painted page lost the line. Each one is a regression to
   * fix, not to record.
   */
  readonly uncut: readonly string[];
  /**
   * Of `uncut`, the reference lines that no painted line matches: the painted page lost the line or
   * cut it elsewhere.
   */
  readonly lostUncut: readonly string[];
  /**
   * Reference lines that the painted clip cuts at an end that the reference clip holds whole, each
   * by `clippedKey` with the edge of the painted clip: the painted line runs past the painted clip
   * edge by more than `clipEdgePt`, so the painted clip hides part of a glyph that the
   * reference page shows, or the line holds only the start or the end of the reference text and the
   * painted clip cuts it there. Each one is a regression to fix, not to record.
   */
  readonly extraCut: readonly string[];
  /**
   * The ends of matched and missing lines where the reference clip is unknown (`referenceEnd`),
   * each as `<text> @ <baseline> | <side> reference clip unknown`: no rule checks the clip there.
   * An inner end of a painted line whose characters the reference items do not split counts too.
   * Each one is a coverage gap, never a pass.
   */
  readonly unknownReference: readonly string[];
  /**
   * Of the missing lines, those that the reference clip cuts on both sides, such as centered text
   * wider than its cell: a named limit of the rule (approximation "Line scores of clipped text").
   */
  readonly bothCut: readonly string[];
  /** Reference lines that are checked (`lineKey`). */
  readonly total: number;
  /** Reference lines of one character other than a digit, which no rule checks (`isOneCharacter`). */
  readonly oneCharacter: number;
  /** Largest baseline distance of a matched line, in points. */
  readonly worstBaseline: number;
  /**
   * Largest horizontal distance of a matched line, in points: the nearer of the left-edge and
   * the right-edge distances.
   */
  readonly worstShift: number;
  /** Reference lines that the painted page does not hold (first 60 characters each). */
  readonly missing: readonly string[];
  /**
   * Checked painted lines (`lineKey`) that the reference page does not hold (first 60 characters
   * each). Each one fails the page (`THRESHOLDS.maxExtraLines`).
   */
  readonly extra: readonly string[];
}

/**
 * Pieces of one page join to hold one line of the other page only when each gap between two of
 * them is about one em: `groupLines` splits a line at a gap of more than `LINE_GAP_EM`, so a join
 * covers a gap of about one em, never a wider one (`joinable`). The size is that of the whole
 * line, or the larger size of the two pieces when the whole line has none; without any size, a
 * gap joins only within `joinGapPt`. The gate does not guard this constant.
 */
const JOIN_GAP_EM = LINE_GAP_EM;

/** The rounding error of a sum of PDF coordinates, in points: far below any measured distance. */
const ROUNDING_PT = 1e-6;

const bareOf = (line: PdfLine): string => normalize(line.text).replace(/\s+/g, '');

/**
 * Whether two pieces of one page, side by side, are close enough to join into the line `whole` of
 * the other page. The decision takes the grouping rule of each page only; `pieceShift` scores the
 * place of each piece by the nearer of its start and end distances, as for one line. The font size
 * of the em rules is the size of `whole`, or the larger size of the two pieces when `whole` has
 * none. The characters of `a` end at `at + bareOf(a).length` of the text of `whole` without white
 * space.
 * - When the items of `whole` end the characters of `a` and start those of `b`, the gap of `whole`
 *   there is at most `LINE_GAP_EM` of the font size, and the gap of the pieces is at most
 *   `JOIN_GAP_EM` of the font size plus `joinGapPt` at each of its two ends: each end of a piece
 *   can lie up to that far from its place in `whole`.
 * - Otherwise the gap of the pieces is at most `JOIN_GAP_EM` of the font size plus `joinGapPt`.
 * No piece width counts toward a gap. The clause on the gap of `whole` restates the rule of
 * `groupLines`, so it never decides on a line that `groupLines` made; it keeps other input, such as
 * the lines of a test, to that rule.
 */
function joinable(a: PdfLine, b: PdfLine, whole: PdfLine, at: number): boolean {
  const size = whole.size ?? Math.max(a.size ?? 0, b.size ?? 0);
  const gap = b.xMin - a.xMax;
  const split = at + bareOf(a).length;
  const end = whole.items?.find((s) => s.at + s.length === split);
  const next = whole.items?.find((s) => s.at === split);
  if (!end || !next) return gap <= JOIN_GAP_EM * size + THRESHOLDS.joinGapPt;
  const wholeGap = next.xMin - end.xMax;
  // `ROUNDING_PT` keeps a gap of exactly one em, summed from other edges, inside the rule.
  return (
    wholeGap <= LINE_GAP_EM * size + ROUNDING_PT &&
    gap <= JOIN_GAP_EM * size + 2 * THRESHOLDS.joinGapPt
  );
}

/**
 * A candidate pair of `compareLines`: reference lines and painted lines that hold the same text.
 * One side is one line; the other is one line, or several lines on one baseline that join to it.
 */
interface Candidate<E> {
  readonly refs: readonly PdfLine[];
  readonly ours: readonly E[];
  /** The distance that orders the pairs: of the left edges and the baselines. */
  readonly distance: number;
  readonly take: () => void;
}

/**
 * The largest distance of a joined piece from the items of the line `ours` of the other page that
 * hold the same characters, in points: the nearer of the start and the end distances, as for a
 * whole line. Undefined when a piece does not start and end on item edges of `ours`. A line
 * without items is scored by the width of the line against the width of the group instead.
 */
function pieceShift(ours: PdfLine, group: readonly PdfLine[]): number | undefined {
  const first = group[0];
  const last = group[group.length - 1];
  if (!ours.items) return Math.abs(ours.xMax - ours.xMin - (last.xMax - first.xMin));
  let worst = 0;
  let at = 0;
  for (const piece of group) {
    const length = bareOf(piece).length;
    const start = ours.items.find((s) => s.at === at);
    const end = ours.items.find((s) => s.at + s.length === at + length);
    if (!start || !end) return undefined;
    worst = Math.max(
      worst,
      Math.min(Math.abs(start.xMin - piece.xMin), Math.abs(end.xMax - piece.xMax))
    );
    at += length;
  }
  return worst;
}

/**
 * The largest distance of a painted piece of a join from the reference line `line`, in points.
 * A piece scores the nearer of its start and end distances from the items of `line` that start
 * and end its characters, or the one of them that an item edge gives. When an item edge gives
 * neither for some piece, the width of the group against the line counts too. A line without
 * items scores the width of the group only.
 */
function joinShift(line: PdfLine, group: readonly PdfLine[]): number {
  const first = group[0];
  const last = group[group.length - 1];
  const width = Math.abs(line.xMax - line.xMin - (last.xMax - first.xMin));
  if (!line.items) return width;
  let worst = 0;
  let at = 0;
  for (const piece of group) {
    const length = bareOf(piece).length;
    const start = line.items.find((s) => s.at === at);
    const end = line.items.find((s) => s.at + s.length === at + length);
    const d = Math.min(
      start ? Math.abs(start.xMin - piece.xMin) : Infinity,
      end ? Math.abs(end.xMax - piece.xMax) : Infinity
    );
    worst = Math.max(worst, d === Infinity ? width : d);
    at += length;
  }
  return worst;
}

/**
 * The part of the line `ours` that holds each piece of `group`, by the items of `ours` that hold
 * the same characters, as `pieceShift` pairs them. Each part has the text of its own items only,
 * without white space. The whole line `ours` for each piece when it has no items, or when its
 * items do not start and end with the characters of the piece. `ours` is a reference line split
 * by painted pieces, and then `compareLines` decides each inner end of such a piece by itself. Or
 * `ours` is a painted line split by reference pieces, which `pieceShift` takes only when each
 * piece starts and ends on a painted item edge, or when the painted line has no items: then each
 * reference piece compares with the whole painted line.
 */
function piecesOf(ours: PdfLine, group: readonly PdfLine[]): PdfLine[] {
  const items = ours.items;
  const chars = ours.text.replace(/\s+/g, '');
  let at = 0;
  return group.map((piece) => {
    const length = bareOf(piece).length;
    const own = items?.filter((s) => s.at >= at && s.at + s.length <= at + length) ?? [];
    const first = own[0];
    const last = own[own.length - 1];
    // The items hold the characters of the piece exactly, or the line does not split there.
    const exact = first?.at === at && last !== undefined && last.at + last.length === at + length;
    at += length;
    if (!exact) return ours;
    return {
      ...ours,
      text: chars.slice(first.at, last.at + last.length),
      xMin: first.xMin,
      xMax: last.xMax,
      items: own,
    };
  });
}

/**
 * The name of a clipped reference line in a known-gap list (`fixtureClipped`): its whole text,
 * its baseline, and the clip edge of the reference PDF that cuts it, each to a tenth of a point.
 * Two lines with the same text are two entries. Only a new reference PDF changes the edge, so a
 * layout fix that moves the painted clip within the threshold keeps the name.
 */
export const clippedKey = (line: PdfLine, clip: ClipCut): string =>
  `${normalize(line.text)} @ ${line.baseline.toFixed(1)} | ${clip.side} clip ${clip.edge.toFixed(1)}`;

/**
 * The position distance of a painted line that a clip cuts, from the reference line that holds
 * all of its text, with the reference clip edge that cuts it, or undefined when the clips do not
 * explain the difference. All of these hold, with `clipEdgePt` as the tolerance:
 * - The reference clip cuts the reference line on one side (`referenceCut`).
 * - The visible text of the painted line starts the reference text, for a cut on the right, or
 *   ends it, for a cut on the left, and holds the glyphs that the reference clip leaves
 *   (`holdsCut`).
 * - The painted line is cut at the same edge (`cutsOurs`).
 * Only the cut side is compared, and the uncut edge is scored.
 */
function clippedShift(
  ours: PdfLine,
  reference: PdfLine,
  evidence: ClipEvidence
): { readonly shift: number; readonly cut: ClipCut } | undefined {
  const cut = referenceCut(reference, evidence);
  if (!cut || !holdsCut(reference, ours, cut, evidence)) return undefined;
  if (!cutsOurs(ours, cut, evidence, reference)) return undefined;
  const shift =
    cut.side === 'right'
      ? Math.abs(ours.xMin - reference.xMin)
      : Math.abs(ours.xMax - reference.xMax);
  return { shift, cut };
}

/**
 * One line of the lines `pieces` of one row, with the text and the items of all of them, at the
 * baseline of the first piece. The `at` of each item counts from the start of the joined text,
 * and a piece without items is one item.
 */
export function spanOf(pieces: readonly PdfLine[]): PdfLine {
  let at = 0;
  const items = pieces.flatMap((p) => {
    const length = p.text.replace(/\s+/g, '').length;
    const own = p.items?.length ? p.items : [{ at: 0, length, xMin: p.xMin, xMax: p.xMax }];
    const shifted = own.map((s) => ({ ...s, at: s.at + at }));
    at += length;
    return shifted;
  });
  return {
    ...pieces[0]!,
    text: pieces.map((p) => p.text).join(' '),
    xMax: pieces[pieces.length - 1]!.xMax,
    items,
  };
}

/**
 * Compare the lines of the painted page with the lines of the reference page. Each step that
 * pairs by text takes its candidate pairs nearest pair first (by baseline and left edge), each
 * line at most once, so the pairs do not depend on the order of the lines. First, lines with the
 * same text (`lineKey`) on one row (within `maxBaselineDeltaPt`) pair, so repeated text does not
 * pair across cells. Then, on its row, the same text after a one-letter bullet. Then several
 * lines of either page on one baseline that join to hold one line of the other (a PDF can place
 * each word of a justified line apart): several reference pieces, with the unchecked lines there,
 * count each checked piece as matched, and each piece is scored as one line. Without `clipping`,
 * no clip rule runs, and an inner end of a painted split inside a reference item gets no key.
 * Then, with `clipping`, a reference line that is still missing matches an unused painted line on
 * its row that holds the start or the end of its text, when a clip cuts both pages there
 * (`clippedShift`), or when the painted clip cuts it (an extra cut). Then lines of the same text
 * that are left pair at any distance. Last, with `clipping`, a clipped line that is left pairs with
 * a cut of it at any distance, by the rules of `clippedShift`.
 */
export function compareLines(
  ours: readonly PdfLine[],
  reference: readonly PdfLine[],
  clipping?: ClipEvidence
): LineScore {
  // The lines of both pages go with the runs: a run that starts a line of the next cell does
  // not end another line (`endClip`).
  const clips: ClipEvidence | undefined = clipping && {
    ...clipping,
    oursLines: ours,
    referenceLines: reference,
  };
  const pool = ours.map((l) => ({
    ...l,
    key: lineKey(l.text),
    alt: bulletless(l),
    bare: normalize(l.text).replace(/\s+/g, ''),
    used: false,
  }));
  type Entry = (typeof pool)[number];
  let matched = 0;
  let total = 0;
  let worst = 0;
  let worstShift = 0;
  const unchecked: PdfLine[] = [];
  // A reference line that its clip cuts needs the painted line cut at the same edge, also when the
  // texts match in full: the painted page that loses the clip shows every glyph.
  const uncut: string[] = [];
  // The converse: at an end that the reference clip holds whole, the painted line stays inside the
  // painted clip.
  const extraCut: string[] = [];
  // The ends where no reference clip is known, so no clip rule applies.
  const unknownReference: string[] = [];
  const unknownEnd = (line: PdfLine, side: 'right' | 'left'): void =>
    void unknownReference.push(
      `${normalize(line.text)} @ ${line.baseline.toFixed(1)} | ${side} reference clip unknown`
    );
  // Each end of a matched line on its own: only a known reference clip gives evidence, so an end
  // with an unknown reference clip applies no rule, and `unknownReference` names it. A line that a
  // clip cuts on both sides compares both edges as cuts.
  const checkCut = (
    line: PdfLine,
    ours: PdfLine,
    evidence: ClipEvidence,
    sides: readonly ('right' | 'left')[] = SIDES
  ): void => {
    for (const side of sides) {
      const end = referenceEnd(line, evidence, side);
      if (end.kind === 'unknown') unknownEnd(line, side);
      // The painted line must be cut at the same edge, and hold only the glyphs that the edge
      // leaves: a painted line with every glyph that runs to the edge shows a hidden glyph. When
      // the reference glyphs are unknown, such as a ligature, no rule checks the glyphs: the end is
      // named as unknown, not as uncut.
      if (end.kind === 'cut') {
        const kept = keepsGlyphs(line, ours, end.cut, evidence);
        if (!cutsOurs(ours, end.cut, evidence, line) || !kept)
          uncut.push(clippedKey(line, end.cut));
        else if (kept === 'unknown')
          unknownReference.push(
            `${normalize(line.text)} @ ${line.baseline.toFixed(1)} | ${end.cut.side} reference glyphs unknown`
          );
      }
      if (end.kind === 'whole') {
        // A painted clip that is unknown at that end gives no evidence that the line stays inside
        // it: the end fails as an extra cut of an unknown edge.
        if (endClip(evidence.ours, ours, side, evidence.oursLines) === undefined) {
          extraCut.push(
            `${normalize(line.text)} @ ${line.baseline.toFixed(1)} | ${side} clip unknown`
          );
          continue;
        }
        const extra = oursCut(ours, evidence, side);
        if (extra) extraCut.push(clippedKey(line, extra));
      }
    }
  };
  // A painted line on the row of the reference line `line`.
  const near = (line: PdfLine, p: Entry): boolean =>
    Math.abs(p.baseline - line.baseline) <= THRESHOLDS.maxBaselineDeltaPt;
  // The reference lines that a pass paired, checked or not.
  const found = new Set<PdfLine>();
  // Pair the reference line `line` with the painted lines `hit`, and apply the clip rules.
  const take = (line: PdfLine, hit: readonly Entry[], shift = 0, groupEdges = true): void => {
    matched += 1;
    if (clips) {
      // Each outer end takes the baseline of the piece that holds it: the pieces of a join lie
      // within `maxBaselineDeltaPt` of the line, and the runs of an end lie on its own piece.
      const whole = spanOf(hit);
      for (const side of SIDES) {
        const end = side === 'left' ? hit[0]! : hit[hit.length - 1]!;
        checkCut(line, { ...whole, baseline: end.baseline }, clips, [side]);
      }
      // Several painted lines that together hold one reference line: the reference PDF clips only
      // text that overflows, so the cells of one row can share one clip there, while each painted
      // cell has its own. Each inner end of a painted piece keeps the rules of a line end, against
      // the reference clip of the characters that the piece holds.
      if (hit.length > 1) {
        const parts = piecesOf(line, hit);
        const chars = line.text.replace(/\s+/g, '');
        let at = 0;
        hit.forEach((piece, i) => {
          const part = parts[i]!;
          const inner = SIDES.filter((side) => (side === 'left' ? i > 0 : i < hit.length - 1));
          if (part !== line) checkCut({ ...part, text: piece.text }, piece, clips, inner);
          else
            for (const side of inner) {
              // An inner end on a boundary of two reference items keeps the rules of a line end,
              // against the item on its side. An end inside a reference item is a painted split:
              // the painted page splits text that the reference draws as one item, and no clip
              // rule checks that end.
              const offset = side === 'right' ? at + piece.bare.length : at;
              const own = line.items?.find((s) =>
                side === 'right' ? s.at + s.length === offset : s.at === offset
              );
              const other = line.items?.find((s) =>
                side === 'right' ? s.at === offset : s.at + s.length === offset
              );
              if (own && other)
                checkCut(
                  {
                    ...line,
                    text: chars.slice(own.at, own.at + own.length),
                    xMin: own.xMin,
                    xMax: own.xMax,
                    items: [{ ...own, at: 0 }],
                  },
                  piece,
                  clips,
                  [side]
                );
              else
                unknownReference.push(
                  `${normalize(piece.text)} @ ${line.baseline.toFixed(1)} | ${side} painted split inside a reference item`
                );
            }
          at += piece.bare.length;
        });
      }
    }
    // The baseline of a join is the baseline of its worst piece.
    for (const p of hit) worst = Math.max(worst, Math.abs(p.baseline - line.baseline));
    // The edges of the group score a single line, and a join only to a line without items, where
    // a shift of the whole group changes no width. Each piece holds the rest by its item edges.
    const last = hit[hit.length - 1];
    if (groupEdges)
      worstShift = Math.max(
        worstShift,
        Math.min(Math.abs(hit[0].xMin - line.xMin), Math.abs(last.xMax - line.xMax))
      );
    worstShift = Math.max(worstShift, shift);
  };
  // Several reference lines that together hold one painted line, the converse of `take`: each
  // checked one counts as matched. Each piece is scored by its own item edges (`pieceShift`); a
  // painted line without items scores the width and the edges of the group instead, as `take`
  // does for a reference line without items: a shift of the whole group changes no width.
  const takeGroup = (p: Entry, group: readonly PdfLine[], shift: number): void => {
    matched += group.filter((l) => lineKey(l.text)).length;
    if (clips) {
      // Each reference piece compares with the part of the painted line that holds its characters.
      const parts = piecesOf(p, group);
      group.forEach((l, i) => checkCut(l, parts[i]!, clips));
    }
    const first = group[0];
    const last = group[group.length - 1];
    for (const l of group) worst = Math.max(worst, Math.abs(l.baseline - p.baseline));
    if (!p.items)
      worstShift = Math.max(
        worstShift,
        Math.min(Math.abs(p.xMin - first.xMin), Math.abs(p.xMax - last.xMax))
      );
    worstShift = Math.max(worstShift, shift);
  };
  // Take the candidate pairs nearest first, each line at most once, as the source's `assign`: the
  // order depends only on where the lines are, not on where a line comes in its page. A tie of
  // the reference lines falls to the painted lines, which the source leaves to the input order.
  const assign = (candidates: Candidate<Entry>[]): void => {
    candidates.sort(
      (a, b) =>
        a.distance - b.distance ||
        a.refs[0]!.baseline - b.refs[0]!.baseline ||
        a.refs[0]!.xMin - b.refs[0]!.xMin ||
        a.ours[0]!.baseline - b.ours[0]!.baseline ||
        a.ours[0]!.xMin - b.ours[0]!.xMin
    );
    for (const c of candidates) {
      if (c.refs.some((l) => found.has(l)) || c.ours.some((p) => p.used)) continue;
      for (const l of c.refs) found.add(l);
      for (const p of c.ours) p.used = true;
      c.take();
    }
  };
  const distance = (line: PdfLine, p: PdfLine): number =>
    Math.hypot(p.xMin - line.xMin, p.baseline - line.baseline);
  const checked = reference.filter((l) => lineKey(l.text));
  total = checked.length;
  unchecked.push(...reference.filter((l) => !lineKey(l.text)));
  // The key and the bullet form of each checked line, computed once: a long line costs its length
  // once, not once for each painted line.
  const keys = new Map(checked.map((l) => [l, { key: lineKey(l.text), alt: bulletless(l) }]));
  // The single-line pairs of unused lines that `same` accepts.
  const pairs = (
    same: (line: PdfLine, p: Entry, key: string, alt: string) => boolean
  ): Candidate<Entry>[] =>
    checked.flatMap((line) =>
      found.has(line)
        ? []
        : pool
            .filter((p) => !p.used && same(line, p, keys.get(line)!.key, keys.get(line)!.alt))
            .map((p) => ({
              refs: [line],
              ours: [p],
              distance: distance(line, p),
              take: () => take(line, [p]),
            }))
    );
  // First the lines of the same text on one row: a line whose painted twin a clip cuts must not
  // take a whole painted line of the same text, on its row or on another row, before the clip
  // rules run.
  assign(pairs((line, p, key) => p.key === key && near(line, p)));
  // A bullet that one PDF writes as a letter and the other as a symbol, on its row.
  assign(
    pairs(
      (line, p, key, alt) =>
        near(line, p) && !!((alt && (p.alt === alt || p.key === alt)) || (p.alt && p.alt === key))
    )
  );
  // Several lines of either page that join to hold one line of the other. Each piece lies within
  // `maxBaselineDeltaPt` of the baseline of that line, as one line does, with no baseline rule
  // between the pieces, as in the source; the baseline of the worst piece is scored.
  const joins: Candidate<Entry>[] = [];
  for (const line of checked) {
    if (found.has(line)) continue;
    const bare = bareOf(line);
    // A painted line without a key, such as one character other than a digit, joins nothing, as
    // in the converse: a clip can leave one character of a cell, and the clipped pass checks it.
    const onRow = pool
      .filter(
        (p) =>
          !p.used &&
          !!p.key &&
          Math.abs(p.baseline - line.baseline) <= THRESHOLDS.maxBaselineDeltaPt
      )
      .sort((p, q) => p.xMin - q.xMin);
    for (let i = 0; i < onRow.length; i += 1) {
      // The joined text stays a start of the text of the line, so a start ends at the first piece
      // that does not continue it.
      let joined = onRow[i].bare;
      for (let j = i + 1; j < onRow.length && joined.length < bare.length; j += 1) {
        if (!joinable(onRow[j - 1], onRow[j], line, joined.length - onRow[j - 1].bare.length))
          break;
        joined += onRow[j].bare;
        if (!bare.startsWith(joined)) break;
        if (joined !== bare) continue;
        const hit = onRow.slice(i, j + 1);
        // Each painted piece is scored against the reference characters that it holds, and the
        // width of the group where no item edge gives a piece (`joinShift`). `take` names the inner
        // ends of pieces inside reference items as a painted split, each end by itself.
        const shift = joinShift(line, hit);
        joins.push({
          refs: [line],
          ours: hit,
          distance: distance(line, hit[0]),
          take: () => take(line, hit, shift, !line.items),
        });
      }
    }
  }
  // The converse: reference pieces, checked or not, that join to hold one painted line.
  for (const p of pool) {
    if (p.used || !p.key) continue;
    const onRow = reference
      .filter(
        (l) => !found.has(l) && Math.abs(l.baseline - p.baseline) <= THRESHOLDS.maxBaselineDeltaPt
      )
      .sort((a, b) => a.xMin - b.xMin);
    for (let i = 0; i < onRow.length; i += 1) {
      let joined = '';
      for (let j = i; j < onRow.length && joined.length < p.bare.length; j += 1) {
        const at = joined.length - (j > i ? bareOf(onRow[j - 1]).length : 0);
        if (j > i && !joinable(onRow[j - 1], onRow[j], p, at)) break;
        joined += bareOf(onRow[j]);
        if (!p.bare.startsWith(joined)) break;
        if (j === i || joined !== p.bare) continue;
        const group = onRow.slice(i, j + 1);
        const shift = pieceShift(p, group);
        if (shift !== undefined)
          joins.push({
            refs: group,
            ours: [p],
            distance: distance(group[0], p),
            take: () => takeGroup(p, group, shift),
          });
      }
    }
  }
  assign(joins);
  const lost = checked.filter((l) => !found.has(l));
  // A reference line whose text runs past the clip of its cell: the reference PDF keeps every
  // glyph of it, and the printed PDF keeps only the glyphs that the clip does not hide entirely.
  const clipped: string[] = [];
  // A painted line that can hold a cut of a reference line: one character other than a digit has
  // no key, and a clip can leave it. No unused painted line has the key of a lost line here: the
  // same-text pairs of its row, and later of any row, took them first.
  const cuttable = (p: Entry): boolean => !p.used && (!!p.key || isOneCharacter(p.text));
  const takeClipped = (l: PdfLine, hit: Entry, shift: number, cut: ClipCut): void => {
    hit.used = true;
    found.add(l);
    matched += 1;
    clipped.push(clippedKey(l, cut));
    // The other end of the line keeps the rules of a matched line.
    checkCut(l, hit, clips!);
    worst = Math.max(worst, Math.abs(hit.baseline - l.baseline));
    worstShift = Math.max(worstShift, shift);
  };
  for (const l of lost) {
    if (!clips) break;
    if (found.has(l)) continue;
    let hit: Entry | undefined;
    let shift = Infinity;
    let cut: ClipCut | undefined;
    for (const p of pool) {
      // A painted line of one character has no key, but a clip can leave one character of the
      // reference text: it matches only by the clip rules of `clippedShift`.
      if (!cuttable(p) || !near(l, p)) continue;
      const s = clippedShift(p, l, clips);
      if (s !== undefined && s.shift < shift) {
        hit = p;
        shift = s.shift;
        cut = s.cut;
      }
    }
    if (hit && cut) takeClipped(l, hit, shift, cut);
  }
  // A reference line that the painted clip cuts at an end that the reference clip holds whole: when
  // the clip hides whole glyphs, the painted line holds only the start (a cut on the right) or the
  // end (a cut on the left) of the reference text. It matches, and its key names the regression.
  for (const l of lost) {
    if (!clips) break;
    if (found.has(l)) continue;
    let hit: { entry: Entry; cut: ClipCut; shift: number } | undefined;
    for (const side of SIDES) {
      if (referenceEnd(l, clips, side).kind !== 'whole') continue;
      for (const p of pool) {
        if (!cuttable(p) || !near(l, p)) continue;
        if (!holdsEnd(l.text, p.text, side)) continue;
        const cut = oursClipCuts(p, l, clips, side);
        // A painted line that ends before the painted clip edge dropped glyphs that the clip shows:
        // it is no cut there, and the reference line stays missing.
        if (!cut || !reachesEdge(p, side, cut.edge, clips, l)) continue;
        const shift = side === 'right' ? Math.abs(p.xMin - l.xMin) : Math.abs(p.xMax - l.xMax);
        if (!hit || shift < hit.shift) hit = { entry: p, cut, shift };
      }
    }
    if (!hit) continue;
    hit.entry.used = true;
    found.add(l);
    matched += 1;
    extraCut.push(clippedKey(l, hit.cut));
    worst = Math.max(worst, Math.abs(hit.entry.baseline - l.baseline));
    worstShift = Math.max(worstShift, hit.shift);
  }
  // Last, the lines of the same text that are left pair at any distance, nearest pair first: a
  // line that moved is out of place, not missing and extra.
  assign(pairs((_, p, key) => p.key === key));
  // Then a clipped line that is left pairs with a cut of it on any row, nearest pair first, as the
  // same text does: a cut line that moved is out of place, not lost. Each pair is one line of each
  // page, so the cost is the number of lost lines times the number of painted lines.
  if (clips)
    assign(
      lost.flatMap((l) =>
        found.has(l)
          ? []
          : pool.flatMap((p) => {
              if (!cuttable(p)) return [];
              const s = clippedShift(p, l, clips);
              return s
                ? [
                    {
                      refs: [l],
                      ours: [p],
                      distance: distance(l, p),
                      take: () => takeClipped(l, p, s.shift, s.cut),
                    },
                  ]
                : [];
            })
      )
    );
  const gone = lost.filter((l) => !found.has(l));
  // A reference line that its clip cuts and that the painted page lost: its key names the
  // regression. A line that the clip cuts on both sides is named as such, so that its note says why
  // it is lost.
  const bothCut: string[] = [];
  const lostUncut: string[] = [];
  if (clips)
    for (const l of gone) {
      const cut = referenceCut(l, clips);
      // A missing line with an unknown reference clip: no rule could match it as clipped.
      for (const side of SIDES)
        if (referenceEnd(l, clips, side).kind === 'unknown') unknownEnd(l, side);
      if (cut) {
        uncut.push(clippedKey(l, cut));
        lostUncut.push(clippedKey(l, cut));
      } else if (SIDES.every((side) => referenceEnd(l, clips, side).kind === 'cut'))
        bothCut.push(normalize(l.text).slice(0, 60));
    }
  const missing = gone.map((l) => normalize(l.text).slice(0, 60));
  return {
    matched,
    clipped,
    uncut,
    lostUncut,
    extraCut,
    unknownReference,
    bothCut,
    total,
    oneCharacter: reference.filter((l) => isOneCharacter(l.text)).length,
    worstBaseline: worst,
    worstShift,
    missing,
    extra: pool.filter((p) => p.key && !p.used).map((p) => normalize(p.text).slice(0, 60)),
  };
}

/** The thresholds that the page of `fixture` misses, in the order of `FAILURE_KINDS`. */
export function failureKinds(pixels: PixelScore, lines: LineScore, fixture: string): FailureKind[] {
  const kinds: FailureKind[] = [];
  const lineMatch = lines.total ? lines.matched / lines.total : 1;
  if (pixels.mismatch > pixelLimit(fixture)) kinds.push('pixels');
  if (lineMatch < THRESHOLDS.minLineBreakMatch) kinds.push('line-breaks');
  if (lines.worstBaseline > THRESHOLDS.maxBaselineDeltaPt) kinds.push('baseline');
  if (lines.worstShift > THRESHOLDS.maxLineShiftPt) kinds.push('line-position');
  if (lines.extra.length > THRESHOLDS.maxExtraLines) kinds.push('extra-text');
  return kinds;
}

/** The measured value of each check, in the units of `GAP_TOLERANCE`. */
export function checkValues(pixels: PixelScore, lines: LineScore): Record<FailureKind, number> {
  return {
    pixels: pixels.mismatch,
    'line-breaks': lines.total ? lines.matched / lines.total : 1,
    baseline: lines.worstBaseline,
    'line-position': lines.worstShift,
    'extra-text': lines.extra.length,
  };
}

/** A box in raster pixels, `[x0, y0, x1, y1]`, the end edges excluded. */
export type PixelBox = readonly [number, number, number, number];

/**
 * How many differing pixels of a pixelmatch diff image (pure red) lie inside any of `boxes`, how
 * many lie outside all of them, and how many pixels the boxes cover.
 */
export function differingPixelsIn(
  diff: { readonly width: number; readonly height: number; readonly data: ArrayLike<number> },
  boxes: readonly PixelBox[]
): { inside: number; outside: number; covered: number } {
  const { width: w, height: h } = diff;
  const mask = new Uint8Array(w * h);
  for (const [x0, y0, x1, y1] of boxes)
    for (let y = Math.max(0, Math.floor(y0)); y < Math.min(h, Math.ceil(y1)); y += 1)
      mask.fill(1, y * w + Math.max(0, Math.floor(x0)), y * w + Math.min(w, Math.ceil(x1)));
  let inside = 0;
  let outside = 0;
  let covered = 0;
  for (let i = 0; i < w * h; i += 1) {
    covered += mask[i]!;
    const p = i * 4;
    if (diff.data[p] !== 255 || diff.data[p + 1] !== 0 || diff.data[p + 2] !== 0) continue;
    if (mask[i]) inside += 1;
    else outside += 1;
  }
  return { inside, outside, covered };
}
