// The clip rules of the line check (`compareLines` in `compare.ts`): which clip holds each end of a
// line, where the reference clip cuts a reference line, whether the painted line is cut at the
// same edge, and which glyphs of the reference text the edge leaves. The text runs of both PDFs and
// their clips come from `pdf-clips.ts`. PPTX Editor's `compare.ts` has no clip rules.

import type { PageBox, TextClipBox, TextRun } from './pdf-clips.ts';
import { normalize, type PdfLine } from './poppler.ts';
import { THRESHOLDS } from './thresholds.ts';

/** A text run lies on the baseline of a line when the two are within this distance (points). */
export const SAME_BASELINE_PT = THRESHOLDS.sameBaselinePt;

/** Two clip rectangles whose edges lie this close (points) are one clip. */
export const SAME_CLIP_PT = THRESHOLDS.sameClipPt;

/** A run that starts within this distance (points) of the start of an item draws that item. */
const ON_ITEM_PT = THRESHOLDS.onItemPt;

/**
 * The evidence that a clip hides the end of a line on both pages: the text runs of the printed
 * PDF and of the reference PDF, each with the clip rectangle that holds it (`textRuns` in
 * `pdf-clips.ts`). `compareLines` adds the lines of each page, so that a run that starts an item
 * of another line is not the end of a line (`endClip`).
 */
export interface ClipEvidence {
  readonly ours: readonly TextRun[];
  readonly reference: readonly TextRun[];
  readonly oursLines?: readonly PdfLine[];
  readonly referenceLines?: readonly PdfLine[];
}

/** The clip edge of the reference PDF that cuts a line: its side and its position in points. */
export interface ClipCut {
  readonly side: 'right' | 'left';
  readonly edge: number;
}

/** Whether two clips of text runs are one clip: the same kind, or rectangles within `SAME_CLIP_PT`. */
export function sameClip(a: TextClipBox, b: TextClipBox): boolean {
  if (typeof a === 'string' || typeof b === 'string') return a === b;
  return a.every((v, i) => Math.abs(v - b[i]!) <= SAME_CLIP_PT);
}

/** The items of a line, or the line itself, for a line without items. */
const spansOf = (line: PdfLine): readonly { readonly xMin: number; readonly xMax: number }[] =>
  line.items?.length ? line.items : [line];

/**
 * The runs on the baseline of `line` that lie inside its item at the `side` end (the line itself,
 * for a line without items), with `clipEdgePt` as the tolerance. A run that draws another line
 * of `lines` on that baseline, such as the next cell, is left out, also where it overlaps the
 * line: a run that starts an item of that line, or a run inside that item under the clip of the
 * run that starts it. A printed PDF can draw each glyph of the next cell with its own run.
 * Undefined, unknown, when runs with different clips start within `ON_ITEM_PT` of an item of
 * another line that reaches the end item: then no run tells which line it draws.
 */
function endRuns(
  runs: readonly TextRun[],
  line: PdfLine,
  side: 'right' | 'left',
  lines: readonly PdfLine[]
): TextRun[] | undefined {
  const tolerance = THRESHOLDS.clipEdgePt;
  const spans = spansOf(line);
  const item = spans.reduce((a, b) =>
    side === 'right' ? (b.xMax > a.xMax ? b : a) : b.xMin < a.xMin ? b : a
  );
  const onBaseline = (y: number): boolean => Math.abs(y - line.baseline) <= SAME_BASELINE_PT;
  const starts = (x: number, of: readonly { readonly xMin: number }[]): boolean =>
    of.some((s) => Math.abs(s.xMin - x) <= ON_ITEM_PT);
  const onLine = runs.filter((r) => onBaseline(r.baseline));
  // The clip of the runs that start at `x`, within `ON_ITEM_PT`: undefined when none does, and
  // `ambiguous` when such runs have different clips.
  const startClip = (x: number): TextRun['clip'] | 'ambiguous' | undefined => {
    const at = onLine.filter((r) => Math.abs(r.x0 - x) <= ON_ITEM_PT);
    if (!at.length) return undefined;
    return at.every((r) => sameClip(r.clip, at[0]!.clip)) ? at[0]!.clip : 'ambiguous';
  };
  // The items of the other lines on the baseline, each with the clip of the run that starts it.
  const others = lines
    .filter((l) => onBaseline(l.baseline))
    .flatMap((l) => spansOf(l))
    .filter((s) => !starts(s.xMin, spans))
    .map((s) => ({ ...s, clip: startClip(s.xMin) }));
  const reaches = (s: { readonly xMin: number; readonly xMax: number }): boolean =>
    s.xMin <= item.xMax + tolerance && s.xMax >= item.xMin - tolerance;
  if (others.some((s) => s.clip === 'ambiguous' && reaches(s))) return undefined;
  const otherLine = (r: TextRun): boolean =>
    others.some(
      (s) =>
        Math.abs(r.x0 - s.xMin) <= ON_ITEM_PT ||
        (s.clip !== undefined &&
          s.clip !== 'ambiguous' &&
          sameClip(r.clip, s.clip) &&
          r.x0 >= s.xMin - tolerance &&
          r.x1 <= s.xMax + tolerance)
    );
  return onLine.filter(
    (r) => r.x0 >= item.xMin - tolerance && r.x1 <= item.xMax + tolerance && !otherLine(r)
  );
}

/**
 * The clip of the run at the `side` end of a line: of the runs inside the item at that end
 * (`endRuns`), the one that draws the last glyph on that side, the one that ends last (right) or
 * starts first (left). A printed PDF can draw each glyph with its own run, so the clip of that
 * glyph decides, also when the other glyphs of the item have another clip. `'none'` when the end
 * run has no clip. Undefined, unknown, when no run lies there, or the clip is not one rectangle
 * or does not hold the baseline. `lines` are the lines of the page of `runs`.
 */
export function endClip(
  runs: readonly TextRun[],
  line: PdfLine,
  side: 'right' | 'left',
  lines: readonly PdfLine[] = []
): PageBox | 'none' | undefined {
  const inside = endRuns(runs, line, side, lines);
  if (!inside?.length) return undefined;
  const end = inside.reduce((a, b) =>
    side === 'right' ? (b.x1 > a.x1 ? b : a) : b.x0 < a.x0 ? b : a
  );
  const clip = end.clip;
  if (clip === 'none') return 'none';
  if (clip === 'non-rectangular') return undefined;
  return clip[1] <= line.baseline && line.baseline <= clip[3] ? clip : undefined;
}

/** The clip rectangle of the run at the `side` end of a line (`endClip`), or undefined. */
export function clipAt(
  runs: readonly TextRun[],
  line: PdfLine,
  side: 'right' | 'left',
  lines: readonly PdfLine[] = []
): PageBox | undefined {
  const clip = endClip(runs, line, side, lines);
  return typeof clip === 'string' ? undefined : clip;
}

/**
 * What the reference clip does at one end of a reference line, with `clipEdgePt` as the
 * tolerance:
 * - `cut`: the reference text runs past the clip edge of its run at that end;
 * - `whole`: the clip of the run at that end is known, a rectangle or none, and the text stays
 *   inside it;
 * - `unknown`: no run lies at that end, or its clip is not one rectangle. Then the clip gives no
 *   evidence at that end, and no clip rule applies there.
 */
export type ReferenceEnd =
  | { readonly kind: 'cut'; readonly cut: ClipCut }
  | { readonly kind: 'whole' }
  | { readonly kind: 'unknown' };

export const SIDES = ['right', 'left'] as const;

export function referenceEnd(
  line: PdfLine,
  evidence: ClipEvidence,
  side: 'right' | 'left'
): ReferenceEnd {
  const tolerance = THRESHOLDS.clipEdgePt;
  const clip = endClip(evidence.reference, line, side, evidence.referenceLines);
  if (clip === undefined) return { kind: 'unknown' };
  if (clip === 'none') return { kind: 'whole' };
  if (side === 'right' && line.xMax > clip[2] + tolerance)
    return { kind: 'cut', cut: { side, edge: clip[2] } };
  if (side === 'left' && line.xMin < clip[0] - tolerance)
    return { kind: 'cut', cut: { side, edge: clip[0] } };
  return { kind: 'whole' };
}

/**
 * The edge at which the reference clip cuts a reference line on one side only, or undefined
 * (`referenceEnd`). A line that a clip cuts on both sides is not cut here.
 */
export function referenceCut(line: PdfLine, evidence: ClipEvidence): ClipCut | undefined {
  const cuts = SIDES.map((side) => referenceEnd(line, evidence, side)).flatMap((end) =>
    end.kind === 'cut' ? [end.cut] : []
  );
  return cuts.length === 1 ? cuts[0] : undefined;
}

/**
 * The edge at which the painted clip cuts the reference text that the painted line `ours` holds
 * only in part, at the `side` end, or undefined: the reference line runs past the clip edge of the
 * painted run at that end by more than `clipEdgePt`. The test takes the painted clip edge, not
 * the end of the painted line, because the painted line ends before an edge that lies in white
 * space.
 */
export function oursClipCuts(
  ours: PdfLine,
  reference: PdfLine,
  evidence: ClipEvidence,
  side: 'right' | 'left'
): ClipCut | undefined {
  const tolerance = THRESHOLDS.clipEdgePt;
  const clip = clipAt(evidence.ours, ours, side, evidence.oursLines);
  if (!clip) return undefined;
  if (side === 'right')
    return reference.xMax > clip[2] + tolerance ? { side, edge: clip[2] } : undefined;
  return reference.xMin < clip[0] - tolerance ? { side, edge: clip[0] } : undefined;
}

/**
 * The edge at which the painted clip cuts the painted line at the `side` end, or undefined: the
 * painted line runs past the right edge of the clip of its last run, or past the left edge of the
 * clip of its first run, by more than `clipEdgePt`. Used at an end that the reference clip
 * holds whole (`LineScore.extraCut`).
 */
export function oursCut(
  ours: PdfLine,
  evidence: ClipEvidence,
  side: 'right' | 'left'
): ClipCut | undefined {
  const tolerance = THRESHOLDS.clipEdgePt;
  const clip = clipAt(evidence.ours, ours, side, evidence.oursLines);
  if (!clip) return undefined;
  if (side === 'right')
    return ours.xMax > clip[2] + tolerance ? { side, edge: clip[2] } : undefined;
  return ours.xMin < clip[0] - tolerance ? { side, edge: clip[0] } : undefined;
}

/**
 * Whether white space of `text` stands right after its first `n` characters that are not white
 * space. Iterates the characters of the line text only.
 */
function spaceAfter(text: string, n: number): boolean {
  if (n <= 0) return false;
  let seen = 0;
  for (const ch of text) {
    const blank = /\s/.test(ch);
    if (seen === n) return blank;
    if (!blank) seen += 1;
  }
  return false;
}

/**
 * The advance extents of the words of a reference line, left to right, from the text runs on its
 * baseline that lie inside its items (`TextRun.words`), or from its items. Undefined when they do
 * not count the words of the line text, so that each extent belongs to one word.
 */
function wordAdvances(
  line: PdfLine,
  runs: readonly TextRun[]
): readonly (readonly [number, number])[] | undefined {
  const count = line.text.split(/\s+/).filter(Boolean).length;
  const tolerance = THRESHOLDS.clipEdgePt;
  const spans = line.items?.length ? line.items : [line];
  const fromRuns = runs
    .filter(
      (r) =>
        Math.abs(r.baseline - line.baseline) <= SAME_BASELINE_PT &&
        spans.some((s) => r.x0 >= s.xMin - tolerance && r.x1 <= s.xMax + tolerance)
    )
    .flatMap((r) => r.words ?? [[r.x0, r.x1] as const])
    .sort((p, q) => p[0] - q[0]);
  if (fromRuns.length === count) return fromRuns;
  const fromItems = (line.items ?? []).map((i) => [i.xMin, i.xMax] as const);
  return fromItems.length === count ? fromItems : undefined;
}

/**
 * For a painted line that holds the start (`right`) or the end (`left`) of the reference text and
 * stops at white space of it: the advance edge of the reference word at the end of the text
 * (`end`), and the advance edge of the next reference word past that white space (`next`), from
 * `wordAdvances`. Undefined when the painted text does not stop at white space of the reference
 * text, or when the word extents are unknown. Paint draws no glyph for a space, so a clip edge
 * inside a space leaves the painted line ending at the glyph before it.
 */
function advancePastSpace(
  ours: PdfLine,
  reference: PdfLine | undefined,
  side: 'right' | 'left',
  runs: readonly TextRun[]
): { readonly end: number; readonly next: number } | undefined {
  if (!reference) return undefined;
  const total = reference.text.replace(/\s+/g, '').length;
  const own = ours.text.replace(/\s+/g, '').length;
  const n = side === 'right' ? own : total - own;
  if (n <= 0 || n >= total || !spaceAfter(reference.text, n)) return undefined;
  const words = wordAdvances(reference, runs);
  if (!words) return undefined;
  // The words of the reference text before the white space: the first `n` characters that are
  // not white space end with the `k`-th word.
  let k = 0;
  let seen = 0;
  for (const word of reference.text.split(/\s+/).filter(Boolean)) {
    if (seen >= n) break;
    seen += word.length;
    k += 1;
  }
  if (seen !== n || k <= 0 || k >= words.length) return undefined;
  return side === 'right'
    ? { end: words[k - 1]![1], next: words[k]![0] }
    : { end: words[k]![0], next: words[k - 1]![1] };
}

/**
 * Whether the painted line reaches the painted clip edge `edge` on `side` within
 * `clipEdgePt`. A painted line that drops a glyph that the clip shows ends before it, and does
 * not reach it. When the reference text has white space at the end of the painted line, the
 * painted line may end before the edge (`advancePastSpace`): then the painted line ends within
 * `clipEdgePt` of the advance end of the reference word before that white space, and the edge
 * lies at most `clipEdgePt` past the advance start of the next reference word.
 */
export function reachesEdge(
  ours: PdfLine,
  side: 'right' | 'left',
  edge: number,
  evidence: ClipEvidence,
  reference?: PdfLine
): boolean {
  const tolerance = THRESHOLDS.clipEdgePt;
  if (side === 'right' ? ours.xMax >= edge - tolerance : ours.xMin <= edge + tolerance) return true;
  const space = advancePastSpace(ours, reference, side, evidence.reference);
  if (!space) return false;
  const ourEnd = side === 'right' ? ours.xMax : ours.xMin;
  if (Math.abs(ourEnd - space.end) > tolerance) return false;
  return side === 'right'
    ? edge >= space.end - tolerance && edge <= space.next + tolerance
    : edge <= space.end + tolerance && edge >= space.next - tolerance;
}

/**
 * Whether the painted line, from `xMin` to `xMax` on `baseline`, is cut at the reference edge
 * `cut`: the clip of the painted run at that end is within `clipEdgePt` of the edge, and the
 * painted line reaches the painted clip edge (`reachesEdge`). Which glyphs the line holds is the
 * test of `holdsCut`.
 */
export function cutsOurs(
  ours: PdfLine,
  cut: ClipCut,
  evidence: ClipEvidence,
  reference?: PdfLine
): boolean {
  const o = clipAt(evidence.ours, ours, cut.side, evidence.oursLines);
  if (!o) return false;
  const edge = cut.side === 'right' ? o[2] : o[0];
  if (Math.abs(edge - cut.edge) > THRESHOLDS.clipEdgePt) return false;
  return reachesEdge(ours, cut.side, edge, evidence, reference);
}

/**
 * Whether the visible text of the painted line `ours` starts (`right`) or ends (`left`) the
 * visible text of `reference`. Both sides keep every symbol, so text that `lineKey` keys on one
 * side only, such as `[h]:mm` cut to `[h]:`, or `(a) Note` cut to `(a) N`, compares alike.
 */
export function holdsEnd(reference: string, ours: string, side: 'right' | 'left'): boolean {
  const visible = (text: string): string => normalize(text).replace(/\s+/g, '');
  const [r, o] = [visible(reference), visible(ours)];
  return o !== '' && (side === 'right' ? r.startsWith(o) : r.endsWith(o));
}

/** The visible characters of a line: its text without white space. */
const visibleChars = (text: string): string[] => Array.from(normalize(text).replace(/\s+/g, ''));

/**
 * How many visible glyphs of `reference` the clip edge `cut` leaves, as the least and the most
 * within `clipEdgePt`: on a cut on the right, the glyphs whose advance starts before the
 * edge; on the left, those whose advance ends past it. The glyph advances come from the reference
 * runs on the baseline of the line that the cutting clip holds (`TextRun.glyphs`), so a run of the
 * next cell under another clip does not count. `unknown` when the cutting clip has no rectangle
 * (`clipAt`), or when those runs do not hold exactly the visible characters of the line, as with
 * a ligature or a glyph without Unicode. `untracked` when a run carries no glyph advances, as only
 * a run that a test builds does.
 */
function glyphsKept(
  reference: PdfLine,
  evidence: ClipEvidence,
  cut: ClipCut
): readonly [number, number] | 'unknown' | 'untracked' {
  const tolerance = THRESHOLDS.clipEdgePt;
  const runs = evidence.reference;
  const box = clipAt(runs, reference, cut.side, evidence.referenceLines);
  if (!box) return 'unknown';
  const own = runs.filter(
    (r) =>
      Math.abs(r.baseline - reference.baseline) <= SAME_BASELINE_PT &&
      sameClip(r.clip, box) &&
      r.x0 >= reference.xMin - tolerance &&
      r.x1 <= reference.xMax + tolerance
  );
  if (!own.length || own.some((r) => !r.glyphs)) return own.length ? 'untracked' : 'unknown';
  const glyphs = own.flatMap((r) => r.glyphs!).sort((a, b) => a[0] - b[0]);
  if (glyphs.length !== visibleChars(reference.text).length) return 'unknown';
  const count = (keep: (g: readonly [number, number]) => boolean): number =>
    glyphs.filter(keep).length;
  return cut.side === 'right'
    ? [count((g) => g[0] < cut.edge - tolerance), count((g) => g[0] < cut.edge + tolerance)]
    : [count((g) => g[1] > cut.edge + tolerance), count((g) => g[1] > cut.edge - tolerance)];
}

/**
 * Whether the painted line `ours` holds the reference text that the clip edge `cut` leaves: its
 * visible text is the start of the reference text, for a cut on the right, or the end, for a cut
 * on the left, without white space. With the glyph advances of the reference runs, it holds the
 * glyphs that the edge leaves (`glyphsKept`), and no glyph more or less. When the reference runs
 * do not account for the glyphs of the line, the text rule alone holds here, and `compareLines`
 * names the end as unknown (`LineScore.unknownReference`), which fails the page.
 */
export function holdsCut(
  reference: PdfLine,
  ours: PdfLine,
  cut: ClipCut,
  evidence: ClipEvidence
): boolean {
  return (
    holdsEnd(reference.text, ours.text, cut.side) &&
    keepsGlyphs(reference, ours, cut, evidence) !== false
  );
}

/**
 * Whether the painted line `ours` holds as many visible glyphs as the clip edge `cut` leaves of
 * `reference` (`glyphsKept`). Also a painted line with the whole reference text must: it shows a
 * glyph that the reference clip hides when the edge leaves fewer. `unknown` when the reference
 * glyphs are unknown: no rule checks the glyphs then.
 */
export function keepsGlyphs(
  reference: PdfLine,
  ours: PdfLine,
  cut: ClipCut,
  evidence: ClipEvidence
): boolean | 'unknown' {
  const kept = glyphsKept(reference, evidence, cut);
  if (kept === 'untracked') return true;
  if (kept === 'unknown') return 'unknown';
  const n = visibleChars(ours.text).length;
  return n >= kept[0] && n <= kept[1];
}
