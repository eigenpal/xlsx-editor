// The cases of `cases.ts` that the browser hang test (`run.ts`) runs, and why each one needs a
// real browser. The engine run (`engine-run.ts`) covers every other file case.

import { HANG_CASES, type HangCase } from './cases.ts';

/**
 * The file cases that the browser run keeps although the engine run covers them, and why. The
 * engine run cannot see a request of the page, the paint, or the chrome, and cannot read the
 * memory of the browser process tree or the blocks of its main thread.
 */
export const BROWSER_RUN: ReadonlyMap<string, string> = new Map([
  // The page must ask for no URL outside the demo's origin (the FETCH verdict).
  ['relationship-target-traversal', 'no request for a target outside the package'],
  ['active-content-inert', 'no request or run for DDE, OLE, a control, or a query table'],
  ['macro-part', 'no request or run for a macro part'],
  ['external-references', 'no request for a remote picture, workbook, link, or connection'],
  // The paint and the chrome of the page.
  ['sheets-many', 'the sheet tabs of 2000 worksheets in the chrome'],
  ['view-hostile', 'the paint and the scroll of a hostile view and selection'],
  ['frozen-split-huge', 'the paint and the scroll of a frozen pane split past the grid'],
  ['conditional-many', 'the paint of 20000 conditional formatting rules'],
  ['font-size-hostile', 'the paint of hostile font sizes, indents, and rotations'],
  ['merge-whole-sheet', 'the paint and the scroll of a merge over the whole grid'],
  // Memory on the browser process tree against the cap, and main-thread blocks.
  ['sheet-dense', 'memory and main-thread blocks of 500000 cells'],
  ['sst-50mb', 'memory and main-thread blocks of a 50 MB shared string table'],
  ['wrapped-rows-many', 'memory and main-thread blocks of 200000 wrapped rows'],
  ['text-long-cell', 'memory and main-thread blocks of 10 MB of text in one cell'],
  ['text-many-runs', 'memory and main-thread blocks of 200000 rich text runs'],
  ['style-many-formats', 'memory and main-thread blocks of 64000 fonts and cell formats'],
  ['drawing-anchors-hostile', 'memory and main-thread blocks of 20000 drawing anchors'],
]);

/** Why the case needs a real browser, or null when the engine run covers it alone. */
export function browserReason(c: HangCase): string | null {
  if (c.network) return `the network fault ${c.network}`;
  if (c.race) return `the race ${c.race} in the demo`;
  return BROWSER_RUN.get(c.id) ?? null;
}

/** The cases of the browser run, in the order of `HANG_CASES`. */
export const BROWSER_CASES: readonly HangCase[] = HANG_CASES.filter(
  (c) => browserReason(c) !== null
);
