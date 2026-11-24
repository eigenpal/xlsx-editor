// Fidelity thresholds and known gaps: the ONE place they live (design D8). Tighten over time;
// never loosen to pass. `bun run fidelity:check-gaps` compares this file with the merge base and
// fails when a threshold or a recorded gap value becomes looser by any amount. A threshold that is
// looser in either direction (`looser: 'either'`) does not change at all: a new threshold replaces
// it, and it retires in `RETIRED_LIMITS`.
//
// The thresholds are copied from pptx-editor/scripts/fidelity/thresholds.ts at 1943d2c9
// (EigenPal), comparison A. Differences:
// - `compareDpi` replaces `renderWidthPx`: a printed page is Letter, so the raster size follows
//   from the density (`poppler.ts`), and each raster is supersampled. The painted page is a
//   screenshot of the canvas paint (design D7). The lines of the painted page come from a PDF that
//   Chromium prints of the same paint (`run.ts`).
// - Each known gap records its measured value, as design D8 states. A run compares its scores
//   with the record inside one tolerance for each check (`GAP_TOLERANCE`): a score that improves
//   beyond the tolerance moves the recorded value in the same change, and a score that worsens
//   beyond it fails. The record itself never grows (`check-known-gaps.ts`).
// - Known gaps are kept for each font environment: the reference fonts from the machine, or the
//   metric-compatible substitutes when the machine has none. Only a measured list is recorded.
//   The `reference` environment has no measured list yet, so a run in it fails as a gap.
// - Each font environment holds one list for each raster toolchain (the Poppler and Chromium
//   versions), because the same page rasterizes differently with other tools. A run uses the
//   list of its own tools. In CI, a run with no such list fails. Outside CI, it compares with
//   the tightest recorded value of each check across the lists of its environment
//   (`gateRun` in `known-gaps.ts`) and fails on any worse score.
// - The editor's painted page and the reference PDF are both Letter by construction, so the page
//   size tolerance and the save-and-reopen self-consistency check are not ported yet.
// - A list names the lines of each fixture that a cell clip cuts (`fixtureClipped`), and a run
//   or a change that clips other lines fails. The source has no clipped lines.
// - The eight matching tolerances of the line checks (`sameBaselinePt` to `joinGapPt`) are
//   thresholds with the direction `either`, fixed at the values that the code used. The source
//   keeps its matching tolerances as module constants.

export const THRESHOLDS = {
  /** The density at which both pages are compared, in pixels per inch. */
  compareDpi: 144,
  /** Anti-aliasing-tolerant pixel mismatch per page, as a fraction of all pixels. */
  maxPixelMismatch: 0.005,
  /** pixelmatch per-pixel color threshold (0–1). */
  pixelThreshold: 0.1,
  /** Each matched line's baseline must be within this many points of the reference line. */
  maxBaselineDeltaPt: 0.5,
  /**
   * Each matched line must start or end within this many points of the reference line (the
   * nearer of its left and right edges, so left, right, and centered text all count).
   */
  maxLineShiftPt: 0.5,
  /** Line breaks must be identical: the share of reference lines that the painted page holds. */
  minLineBreakMatch: 1,
  /** Lines of the painted page that `lineKey` checks and that the reference page does not hold. */
  maxExtraLines: 0,
  /**
   * The clip-match tolerances of the line checks, in points. A larger value can pass a wrong
   * render, and a smaller value can fail a correct one, so neither direction is tighter. Painted
   * lines and text runs within `sameBaselinePt` share a baseline (`clip-rules.ts`).
   */
  sameBaselinePt: 0.25,
  /** Two clip rectangles of text runs whose edges lie this close are one clip (`clip-rules.ts`). */
  sameClipPt: 0.01,
  /** A run that starts this close to the start of an item draws that item (`clip-rules.ts`). */
  onItemPt: 0.5,
  /** Two clip rectangles of text items whose edges lie this close are one clip (`pdf-clips.ts`). */
  sameEdgePt: 0.01,
  /** A point this close to the start of a run is on that run (`clipOfItem`, `pdf-clips.ts`). */
  onRunPt: 0.5,
  /** Text items whose baselines lie this close share a row and a clip (`pdf-lines.ts`). */
  sameRowPt: 0.25,
  /**
   * The tolerance of the clip rules at a line end, in points (`clip-rules.ts`): how far a line end
   * may lie from a clip edge, and two clip edges from each other, and still count as one edge. It
   * is apart from `maxLineShiftPt`, so a tighter position threshold does not change which ends
   * count as cut or whole.
   */
  clipEdgePt: 0.5,
  /**
   * The slack of the gap at which the pieces of one page join to hold one line of the other, in
   * points, past one em (`joinable`, `compare.ts`): once at each end of the gap where the items of
   * the whole line split at the pieces, and once everywhere else. It is the value that the code
   * used; no measured PDF needs it. It is apart from `maxLineShiftPt`, so a tighter position
   * threshold does not change which pieces join.
   */
  joinGapPt: 0.5,
} as const;

/**
 * The direction in which each threshold becomes looser, and the tightest value that a threshold
 * takes when the merge base does not have it. A threshold that is looser when larger starts at 0
 * unless `tightest` says otherwise; one that is looser when smaller needs `tightest` to be new.
 * A threshold that is looser in `either` direction is fixed: the check refuses any change to it,
 * and a new one starts at its `tightest` value. A threshold without an entry fails the check.
 */
export const THRESHOLD_DIRECTIONS: Readonly<
  Record<string, { readonly looser: 'larger' | 'smaller' | 'either'; readonly tightest?: number }>
> = {
  compareDpi: { looser: 'smaller' },
  maxPixelMismatch: { looser: 'larger' },
  pixelThreshold: { looser: 'larger' },
  maxBaselineDeltaPt: { looser: 'larger' },
  maxLineShiftPt: { looser: 'larger' },
  minLineBreakMatch: { looser: 'smaller', tightest: 1 },
  maxExtraLines: { looser: 'larger' },
  // The matching tolerances are fixed at the values that the code used.
  sameBaselinePt: { looser: 'either', tightest: 0.25 },
  sameClipPt: { looser: 'either', tightest: 0.01 },
  onItemPt: { looser: 'either', tightest: 0.5 },
  sameEdgePt: { looser: 'either', tightest: 0.01 },
  onRunPt: { looser: 'either', tightest: 0.5 },
  sameRowPt: { looser: 'either', tightest: 0.25 },
  clipEdgePt: { looser: 'either', tightest: 0.5 },
  joinGapPt: { looser: 'either', tightest: 0.5 },
};

/**
 * Thresholds that the harness no longer has, each with the reason, as PPTX Editor's
 * `RETIRED_LIMITS`. `check-known-gaps.ts` refuses a threshold that goes away unless it is listed
 * here, so no threshold goes away unseen.
 */
export const RETIRED_LIMITS: Readonly<Record<string, string>> = {};

/**
 * Tighter pixel limits for single fixtures, as a fraction of all pixels. One fill case of
 * `fills-patterns` covers about 0.3 % of its page, so the page limit cannot see one wrongly
 * painted case. Measured on its page, the wrong cases tried differ in 0.024 % to 0.29 %: 0.024 %
 * for one moved dot of one `gray0625` tile in one cell, 0.039 % for a `gray0625` tile 1 point out
 * of phase, 0.062 % for `gray125` in place of `gray0625`, and 0.24 % to 0.29 % for a mirrored
 * linear gradient. Its limit is five times its measured score (0.004 % with both raster
 * toolchains), so the smallest wrong case tried is about 1.2 times the limit. An entry only
 * tightens, stays while its fixture exists, and is never looser than the page limit
 * (`check-known-gaps.ts`).
 */
export const FIXTURE_PIXEL_MISMATCH: Readonly<Record<string, number>> = {
  'fills-patterns': 0.0002,
  // One wrong solid fill case changes the whole cell of its case: column B is 96 points wide
  // (width 16 at a maximum digit width of 6 pixels, ECMA-376 Part 1, 18.3.1.13) and a row prints
  // 14 points high, so about 192 × 28 pixels at 144 dpi, 0.28% of the page, whatever the face. The
  // score is 0.106% locally (Poppler 26.10.0) and 0.105% in CI (Poppler 24.02.0), so this limit
  // fails such a case. Font and border cases change too few pixels for a page limit: the measured
  // test in `cell-format.test.ts` checks that the cascade resolves them case by case.
  'cell-styles': 0.002,
};

/** The pixel limit of one fixture. */
export const pixelLimit = (fixture: string): number =>
  Object.hasOwn(FIXTURE_PIXEL_MISMATCH, fixture)
    ? Math.min(FIXTURE_PIXEL_MISMATCH[fixture]!, THRESHOLDS.maxPixelMismatch)
    : THRESHOLDS.maxPixelMismatch;

/** A threshold that a page misses. */
export type FailureKind = 'pixels' | 'line-breaks' | 'baseline' | 'line-position' | 'extra-text';

export const FAILURE_KINDS: readonly FailureKind[] = [
  'pixels',
  'line-breaks',
  'baseline',
  'line-position',
  'extra-text',
];

/**
 * How far a score may move from its recorded gap value before the gate asks for a new record:
 * the noise of one check between machines. Pixels are a fraction of all pixels; line breaks are
 * the fraction of reference lines matched; positions are points; extra text is a count of
 * lines.
 */
export const GAP_TOLERANCE: Readonly<Record<FailureKind, number>> = {
  pixels: 0.001,
  'line-breaks': 0,
  baseline: 0.1,
  'line-position': 0.1,
  'extra-text': 0,
};

/** The font environment of a run. */
export type FontEnvironment = 'reference' | 'substitute';

/** One known gap: the measured value of each check that a fixture fails today. */
export type KnownGap = Readonly<Partial<Record<FailureKind, number>>>;

/** The known gaps of one font environment and one raster toolchain. */
export interface KnownGapList {
  /** The Poppler version that draws the reference PDF, and the pdf.js version that reads it. */
  readonly poppler: string;
  /** Playwright's Chromium, which draws the painted page. */
  readonly chromium: string;
  /** How many reference lines each fixture checks. Fewer lines fail the gate. */
  readonly fixtureLines: Readonly<Record<string, number>>;
  /**
   * The reference lines of each fixture that a cell clip shortens on both pages, where the painted
   * page holds only part of the reference text (`LineScore.clipped` of `compareLines`),
   * by `clippedKey` (the whole text, the baseline, and the reference clip edge), once for each
   * time a clip cuts the line.
   * A fixture that is not listed has none. A clipped line that is not listed fails the gate. A
   * list from before this record has no `fixtureClipped`; the gate then allows none.
   */
  readonly fixtureClipped?: Readonly<Record<string, readonly string[]>>;
  /** By fixture name. */
  readonly fixtures: Readonly<Record<string, KnownGap>>;
}

/**
 * The approximation (`packages/core/src/layout/approximations.ts`) that causes each known gap,
 * by fixture and check. A test (`gap-causes.test.ts`) fails when a known gap has no cause, and
 * when the public item of its approximation does not name the fixture and the recorded value.
 */
export const GAP_CAUSES: Readonly<Record<string, Partial<Record<FailureKind, string>>>> = {
  'conditional-rules': { 'line-position': 'measured-text-offsets' },
  'grid-structure': { baseline: 'measured-text-offsets' },
};

/**
 * Known gaps, a ratchet, for each font environment and raster toolchain. `bun run fidelity`
 * fails when a fixture fails a check that is not listed, or when a listed value moves beyond
 * `GAP_TOLERANCE`. An environment with no list has no measurement: a run in it fails, never
 * passes. `bun run fidelity --update-gaps` rewrites the list of the run's environment and
 * toolchain, or adds one; review the diff. The first `substitute` list is measured on a
 * developer machine, the second on the CI runner.
 *
 * The passing fixture nearest to the page limit: `number-formats`, against the page limit of
 * 0.5%. It scores 0.490% with Poppler 26.10.0 on a developer machine and 0.428% with Poppler 24.02.0
 * on the CI runner, both with Chromium 158.0.8127.18 and Carlito 1.104
 * (`@docx-editor.dev/fonts` 2.35.0). Chromium 156.0.8078.4 with `@docx-editor.dev/fonts` 2.33.0
 * gave the same known gaps and the same `number-formats` scores. With Poppler 26.10.0, all
 * 9499 of its differing pixels lie inside the glyph boxes of the reference text (`bun
 * scripts/fidelity/glyph-share.ts number-formats` after `bun run fidelity`). Both sides draw
 * Carlito, so no face difference adds to them; the share shows where they are, not their
 * cause (approximation "Glyph rasterization"). A tool update that moves it past the limit
 * fails the gate with no known gap. If it does, run `bun scripts/fidelity/rasterizer-control.ts
 * number-formats` first: it shows whether the editor page or the rasterizers moved.
 *
 * Why the `substitute` gaps remain:
 * - `conditional-rules` line position: the reference PDF draws each right-aligned italic or
 *   bold-italic number 2 points left of a regular one in the same column (x 247 against 249);
 *   layout does not (approximation "Measured text offsets"). The italic comes from differential
 *   formats only: no reference PDF has italic text from a cell format, so the cause has no
 *   measurement.
 * - `grid-structure` baseline: approximation "Measured text offsets", a line centered in a
 *   merged range sits 1 point low. A test keeps the two equal (`gap-causes.test.ts`).
 */
export const KNOWN_GAPS: Readonly<Partial<Record<FontEnvironment, readonly KnownGapList[]>>> = {
  substitute: [
    {
      poppler: 'poppler 26.10.0, pdf.js 6.5.438',
      chromium: 'chromium 159.0.8154.6',
      fixtureLines: {
        'basic-values': 11,
        'cell-styles': 37,
        'clipboard-edit': 16,
        'colors-fonts': 68,
        'conditional-rules': 49,
        'dates-1904': 12,
        'fills-patterns': 0,
        'formula-overflow': 5,
        formulas: 12,
        'grid-structure': 8,
        'macro-enabled': 8,
        'number-formats': 152,
        'number-formats-more': 22,
        'preserved-parts': 2,
        'resize-edit': 7,
        'sheet-add': 4,
        'sheet-ops': 12,
        structure: 31,
        'text-layout': 41,
        'wrapped-formula': 5,
        'wrapped-only-formula': 3,
      },
      fixtureClipped: {
        'formula-overflow': [
          'This text is longer than its cell @ 47.0 | right clip 108.0',
          'This text is longer than its cell @ 75.0 | right clip 108.0',
        ],
        'number-formats': ['_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_) @ 327.0 | right clip 414.0'],
        'preserved-parts': ['Kept content @ 47.0 | right clip 89.0'],
      },
      fixtures: {
        'conditional-rules': {
          'line-position': 2,
        },
        'grid-structure': {
          baseline: 1,
        },
      },
    },
    {
      poppler: 'poppler 24.02.0, pdf.js 6.5.438',
      chromium: 'chromium 159.0.8154.6',
      fixtureLines: {
        'basic-values': 11,
        'cell-styles': 37,
        'clipboard-edit': 16,
        'colors-fonts': 68,
        'conditional-rules': 49,
        'dates-1904': 12,
        'fills-patterns': 0,
        'formula-overflow': 5,
        formulas: 12,
        'grid-structure': 8,
        'macro-enabled': 8,
        'number-formats': 152,
        'number-formats-more': 22,
        'preserved-parts': 2,
        'resize-edit': 7,
        'sheet-add': 4,
        'sheet-ops': 12,
        structure: 31,
        'text-layout': 41,
        'wrapped-formula': 5,
        'wrapped-only-formula': 3,
      },
      fixtureClipped: {
        'formula-overflow': [
          'This text is longer than its cell @ 47.0 | right clip 108.0',
          'This text is longer than its cell @ 75.0 | right clip 108.0',
        ],
        'number-formats': ['_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_) @ 327.0 | right clip 414.0'],
        'preserved-parts': ['Kept content @ 47.0 | right clip 89.0'],
      },
      fixtures: {
        'conditional-rules': {
          'line-position': 2,
        },
        'grid-structure': {
          baseline: 1,
        },
      },
    },
  ],
};
