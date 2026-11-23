// A test helper of the line checks: reference text runs with glyph advances, as `textRuns` of
// `pdf-clips.ts` records them for every run of a real PDF. A run that a test builds has none, and
// `compareLines` then skips the glyph rule of a clipped line (`untracked`). The tests compare with
// glyph advances by default, through `compareTracked`, and call `compareLines` without them only
// to test that path.

import { compareLines, type ClipEvidence, type LineScore } from './compare.ts';
import type { TextRun } from './pdf-clips.ts';
import type { PdfLine } from './poppler.ts';

type Extent = readonly [number, number];

/** `n` advances of equal width from `x0` to `x1`. */
const spread = (x0: number, x1: number, n: number): Extent[] =>
  Array.from({ length: n }, (_, i) => [x0 + (i * (x1 - x0)) / n, x0 + ((i + 1) * (x1 - x0)) / n]);

/**
 * The words of the reference text that the run `r` draws, as their visible character counts: the
 * words of a line that the run holds whole, or the items of a line that lie inside the run. Empty
 * when no line on its baseline holds the run.
 */
function wordsOf(r: TextRun, lines: readonly PdfLine[]): number[] {
  const near = (a: number, b: number): boolean => a >= b - 0.5;
  for (const l of lines) {
    if (Math.abs(l.baseline - r.baseline) > 0.25) continue;
    if (!near(r.x0, l.xMin) || !near(l.xMax, r.x1)) continue;
    const own = l.items?.filter((s) => near(s.xMin, r.x0) && near(r.x1, s.xMax)) ?? [];
    const whole = !l.items?.length || (own.length === l.items.length && own.length === 1);
    if (whole) return l.text.split(/\s+/).flatMap((w) => (w ? [w.length] : []));
    return own.map((s) => s.length);
  }
  return [];
}

/**
 * The runs `runs` of the reference lines `lines`, each with a glyph advance for each visible
 * character that it draws. A run with word extents (`TextRun.words`) spreads the glyphs of each
 * word over its extent; any other run spreads them over its advance. A run that already has glyph
 * advances, or that draws no character of a line, stays as it is.
 */
export function withGlyphs(runs: readonly TextRun[], lines: readonly PdfLine[]): TextRun[] {
  return runs.map((r) => {
    const words = r.glyphs ? [] : wordsOf(r, lines);
    const n = words.reduce((a, b) => a + b, 0);
    if (!n) return r;
    const glyphs =
      r.words?.length === words.length
        ? words.flatMap((k, i) => spread(r.words![i]![0], r.words![i]![1], k))
        : spread(r.x0, r.x1, n);
    return { ...r, glyphs };
  });
}

/** `compareLines` with the glyph advances of the reference runs (`withGlyphs`). */
export function compareTracked(
  ours: readonly PdfLine[],
  reference: readonly PdfLine[],
  clipping?: ClipEvidence
): LineScore {
  if (!clipping) return compareLines(ours, reference);
  return compareLines(ours, reference, {
    ...clipping,
    reference: withGlyphs(clipping.reference, reference),
  });
}
