// The hostile workbooks of `cases.ts`, through the engine with no browser and no DOM. Each file
// case runs the open path of the demo (`whyNotOpen` of `examples/vite/src/open-check.ts` over
// `readXlsx`), the store, the layout of the visible window, a save, an edit of cell A1, and a
// reopen. The read must model each XML part that it models in a plain workbook, and name the
// hostile sheet first. A refused edit fails the case unless `EDIT_REFUSED` names the case.
//
// The work counters of the engine measure the cost, so a bound is exact on every machine. The
// bounds of a window come only from `WINDOW` and from a reference workbook: the work of the same
// window over a full sheet of short text and one conditional formatting rule, times
// `WORK_HEADROOM`. They never come from the layout under test.
// - Cell lookups, key reads, key searches, and rule checks (`layoutWorkTestRecorder`) stay under
//   the reference cap, at the top and after a scroll into the middle of the stored rows. Each
//   key search takes at most ceil(log2(n + 1)) steps over the n stored cells.
// - The scroll builds no sheet index and reads no row height, measured or cached
//   (`sheetIndexTestRecorder`, `rowHeightTestRecorder`).
// - The first layout reads each stored row once for its height (`rowHeightTestRecorder`). That
//   is the current geometry: its cost grows with the stored rows, not with the window.
// `WIDE_WINDOW` names the cases whose pane holds up to `MAX_WINDOW` rows, with a cap that is
// still a constant. The merge search and the drawing test read every merged range above the
// window and every drawing of the sheet. The runs record those counts, and no bound covers them yet.
//
// `engine-cases-*.test.ts` run these checks on a share of the cases each, so the parallel test
// runner can spread them. `engine-coverage.test.ts` keeps every case in the engine run or in the
// browser run with a reason (`browserReason` of `browser-cases.ts`).

import {
  cellFormat,
  cellValue,
  sheetStale,
  shownText,
  EN_US,
} from '../../packages/core/src/cascade/index.ts';
import {
  layoutSheet,
  layoutWorkTestRecorder,
  type SheetLayout,
  type Viewport,
} from '../../packages/core/src/layout/sheet-layout.ts';
import { MAX_WINDOW } from '../../packages/core/src/layout/index.ts';
import { SCREEN_PROFILE } from '../../packages/core/src/layout/metrics.ts';
import { rowHeightTestRecorder } from '../../packages/core/src/layout/row-height.ts';
import type { TextMeasurer } from '../../packages/core/src/layout/text/measurer.ts';
import { readZip } from '../../packages/ooxml-core/src/zip.ts';
import {
  attr,
  cellKey,
  readXlsx,
  saveXlsx,
  sheetByName,
  sheetIndex,
  workbookIndex,
  WorkbookDocumentStore,
  xmlPartOf,
  type XlsxPackage,
} from '../../packages/core/src/store/index.ts';
import { sheetIndexTestRecorder } from '../../packages/core/src/store/workbook/sheet-index.ts';
import { whyNotOpen } from '../../examples/vite/src/open-check.ts';
import { buildXlsx, c, row } from '../fixtures/builder.ts';
import { HANG_CASES, HOSTILE_SHEET, type HangCase } from './cases.ts';
import { EDIT_REFUSED, LAST_CELL, SAVE_KEEPS_CONTENT } from './first-cells.ts';
import { showsException } from './verdict.ts';

/** The grid of the demo in a window of 1440 by 900 pixels, in points. */
export const WINDOW: Viewport = { width: 1080, height: 600, scrollX: 0, scrollY: 0 };

/**
 * The headroom over the work of the reference window. The reference fills every cell, so no
 * text shows across a neighbor and each row reads one key on each side. A hostile sheet can
 * show text across empty neighbors, split the window into frozen panes, and test cells against
 * rules, so it may do somewhat more work for the same window. Twice the reference holds every
 * case of `cases.ts` except those of `WIDE_WINDOW`; a window of all the rows that a pane allows
 * (`MAX_WINDOW`) needs far more.
 */
export const WORK_HEADROOM = 2;

/**
 * The cases whose window holds a pane of up to `MAX_WINDOW` rows, and why. Their cap is the
 * reference cap times the rows of four full panes (`4 * MAX_WINDOW`, for the four panes of a
 * frozen split) over the rows of the reference window, which still does not depend on the file.
 * It is loose: it bounds the work by the pane limit, not by the window. A known follow-up: a pane should hold only the rows that
 * the window shows.
 */
export const WIDE_WINDOW: ReadonlyMap<string, string> = new Map([
  ['view-hostile', 'a frozen pane split past the grid makes a pane of 2000 by 2000 cells'],
]);

/** The counters of the work of a window that the reference bounds. */
const BOUNDED = ['looked', 'keySteps', 'searches', 'ruleChecks'] as const;
type BoundedWork = Record<(typeof BOUNDED)[number], number>;

/** The rows and the columns of the reference sheet: more than the window shows. */
const REFERENCE_ROWS = 200;
const REFERENCE_COLS = 120;

/** The reference workbook: every cell holds short text, and one rule covers the cells. */
function referenceWorkbook(): XlsxPackage {
  const name = (k: number) =>
    (k < 26 ? '' : String.fromCharCode(64 + Math.floor(k / 26))) +
    String.fromCharCode(65 + (k % 26));
  const rows = Array.from({ length: REFERENCE_ROWS }, (_, r) =>
    row(
      r + 1,
      Array.from({ length: REFERENCE_COLS }, (_, k) => c(`${name(k)}${r + 1}`, 'x'))
    )
  ).join('');
  const read = readXlsx(
    buildXlsx({
      styles: { dxfs: ['<dxf><font><b/></font></dxf>'] },
      sheets: [
        {
          name: HOSTILE_SHEET,
          rows,
          afterMerges:
            `<conditionalFormatting sqref="A1:${name(REFERENCE_COLS - 1)}${REFERENCE_ROWS}">` +
            '<cfRule type="containsText" dxfId="0" priority="1" operator="containsText" text="y">' +
            '<formula>NOT(ISERROR(SEARCH("y",A1)))</formula></cfRule></conditionalFormatting>',
        },
      ],
    })
  );
  if (!read.ok) throw new Error('the reference workbook does not read');
  return read.pkg;
}

/** The work of the counters of `BOUNDED` since the last reset. */
function boundedWork(): BoundedWork {
  const work = layoutWorkTestRecorder();
  return Object.fromEntries(BOUNDED.map((k) => [k, work[k]])) as BoundedWork;
}

/** The scroll position of the middle stored row and column of the hostile sheet, or the top. */
function middleOf(pkg: XlsxPackage, layout: SheetLayout): Viewport {
  const index = workbookIndex(pkg)!;
  const part = xmlPartOf(pkg, sheetByName(index, HOSTILE_SHEET)?.partName);
  if (!part) return WINDOW;
  const rows = [...sheetIndex(part).rows.keys()].sort((a, b) => a - b);
  const cols = [...sheetIndex(part).cells.values()].map((e) => e.col).sort((a, b) => a - b);
  const g = layout.geometry;
  return {
    ...WINDOW,
    scrollY: rows.length > 0 ? g.rows.offsetOf(rows[rows.length >> 1]!) : 0,
    scrollX: cols.length > 0 ? g.columns.offsetOf(cols[cols.length >> 1]!) : 0,
  };
}

/** The caps of one measurer: for the window, and for a pane of `MAX_WINDOW` rows. */
export interface WindowCaps {
  readonly window: BoundedWork;
  readonly wide: BoundedWork;
}

const capCache = new WeakMap<TextMeasurer, WindowCaps>();

/**
 * The cap of each bounded counter for one layout of `WINDOW`: the work of the reference window
 * at the top and in the middle of the reference sheet, whichever is more, times the headroom.
 */
export function windowCap(measurer: TextMeasurer): WindowCaps {
  const hit = capCache.get(measurer);
  if (hit) return hit;
  const pkg = referenceWorkbook();
  const options = { measurer, profile: SCREEN_PROFILE };
  const work = layoutWorkTestRecorder();
  work.reset();
  const top = layoutSheet(pkg, HOSTILE_SHEET, WINDOW, options)!;
  const atTop = boundedWork();
  work.reset();
  layoutSheet(pkg, HOSTILE_SHEET, middleOf(pkg, top), options);
  const inMiddle = boundedWork();
  const capOf = (scale: number) =>
    Object.fromEntries(
      BOUNDED.map((k) => [k, Math.ceil(scale * WORK_HEADROOM * Math.max(atTop[k], inMiddle[k]))])
    ) as BoundedWork;
  const rows = top.panes.reduce((sum, pane) => sum + pane.rows.length, 0);
  const caps = { window: capOf(1), wide: capOf((4 * MAX_WINDOW) / rows) };
  capCache.set(measurer, caps);
  return caps;
}

/** The parts that the read models as XML: a hostile value must not leave one as bytes. */
const MODELED = /\/(sharedStrings|styles|workbook|sheet\d+|theme\d*|drawing\d+)\.xml$/i;

/** These parts stay as bytes on purpose: one passes the depth limit, and one declares a DTD. */
const STAYS_BINARY = new Set(['style-xml-deep', 'xml-entity-expansion']);

/** The cases that need no browser: each file case with no network fault and no race. */
export const ENGINE_CASES: readonly HangCase[] = HANG_CASES.filter(
  (c) => c.bytes && !c.network && !c.race
);

/** The cases of share `part` of `of`, by their place in the list. */
export const engineShare = (part: number, of: number): readonly HangCase[] =>
  ENGINE_CASES.filter((_, i) => i % of === part - 1);

/** The hostile sheet's index: its stored rows and cells. */
function hostileIndex(pkg: XlsxPackage) {
  const index = workbookIndex(pkg)!;
  const part = xmlPartOf(pkg, sheetByName(index, HOSTILE_SHEET)?.partName);
  return part ? sheetIndex(part) : undefined;
}

/** The text that the grid shows in the cell at `row`, `col` of the hostile sheet, or null. */
export function shownAt(pkg: XlsxPackage, row: number, col: number): string | null {
  const index = workbookIndex(pkg)!;
  const part = xmlPartOf(pkg, sheetByName(index, HOSTILE_SHEET)?.partName);
  const entry = part ? sheetIndex(part).cells.get(cellKey(row, col)) : undefined;
  if (!part || !entry) return null;
  const { value } = cellValue(pkg, index, entry.el, sheetStale(index, part.root));
  const { format } = cellFormat(pkg, index, attr(entry.el, 's'));
  return shownText(value, format, { locale: EN_US, date1904: index.date1904 });
}

/** The text of the last used cell (the last row and the last column), as Control+End goes. */
function lastShown(pkg: XlsxPackage): string | null {
  let lastRow = -1;
  let lastCol = -1;
  for (const entry of hostileIndex(pkg)?.cells.values() ?? []) {
    lastRow = Math.max(lastRow, entry.row);
    lastCol = Math.max(lastCol, entry.col);
  }
  return shownAt(pkg, lastRow, lastCol);
}

/** The inflated size of the parts of `bytes`. */
function inflatedBytes(bytes: Uint8Array): number {
  const read = readZip(bytes);
  if (!read.ok) return 0;
  let total = 0;
  for (const data of read.entries.values()) total += data.length;
  return total;
}

/** The work of one layout of a window. */
export interface WindowWork extends BoundedWork {
  readonly searchSteps: number;
  readonly mergesRead: number;
  readonly drawingsRead: number;
  readonly indexBuilds: number;
  readonly rowsMeasured: number;
  readonly rowsCached: number;
}

/** What one case did. Each field is a fact that a test compares with its bound. */
export interface EngineRun {
  readonly outcome: 'open' | 'refuse';
  /** The words of a refusal. */
  readonly refusal?: string;
  readonly firstCell?: string | null;
  readonly lastCell?: string | null;
  /** The stored rows and cells of the hostile sheet. */
  readonly storedRows?: number;
  readonly storedCells?: number;
  /** The first layout of the window, at the top, and the layout after the scroll. */
  readonly top?: WindowWork;
  readonly scrolled?: WindowWork;
  /** Whether the store took the edit, and the text of A1 after a save and a reopen. */
  readonly edited?: boolean;
  readonly editRefusal?: string;
  readonly editRefusalCode?: string;
  readonly reopenedFirst?: string | null;
  /** The inflated bytes of the opened file and of the saved copy, for `SAVE_KEEPS_CONTENT`. */
  readonly openedContent?: number;
  readonly savedContent?: number;
  /** Whether the saved copy of the opened workbook, before the edit, reads again. */
  readonly saveReads?: boolean;
  /** The parts of a modeled kind that the read kept as bytes, and the name of the first sheet. */
  readonly unreadParts?: readonly string[];
  readonly firstSheet?: string;
}

/** Lays out `viewport`, and returns the layout with its work. */
function measured(
  pkg: XlsxPackage,
  viewport: Viewport,
  measurer: TextMeasurer
): { layout: SheetLayout; work: WindowWork } {
  const work = layoutWorkTestRecorder();
  const rows = rowHeightTestRecorder();
  const indexes = sheetIndexTestRecorder();
  work.reset();
  rows.reset();
  indexes.reset();
  const layout = layoutSheet(pkg, HOSTILE_SHEET, viewport, {
    measurer,
    profile: SCREEN_PROFILE,
  });
  if (!layout) throw new Error(`no sheet ${HOSTILE_SHEET} to lay out`);
  return {
    layout,
    work: {
      ...boundedWork(),
      searchSteps: work.searchSteps,
      mergesRead: work.mergesRead,
      drawingsRead: work.drawingsRead,
      indexBuilds: indexes.builds,
      rowsMeasured: rows.measured,
      rowsCached: rows.cached,
    },
  };
}

/**
 * Runs one case through the engine. `window` replaces `WINDOW` for a probe of the bounds. It
 * throws only for a defect of the test itself.
 */
export function runEngineCase(
  testCase: HangCase,
  measurer: TextMeasurer,
  window: Viewport = WINDOW
): EngineRun {
  const bytes = testCase.bytes!();
  const read = readXlsx(bytes);
  const problem = whyNotOpen(read);
  if (problem || !read.ok) return { outcome: 'refuse', refusal: problem ?? 'no reason' };
  const unreadParts = [...read.pkg.parts.values()]
    .filter((p) => p.kind === 'binary' && MODELED.test(p.partName))
    .map((p) => p.partName);
  const firstSheet = workbookIndex(read.pkg)?.sheets[0]?.name;
  const store = new WorkbookDocumentStore(read.pkg);
  const sheet = hostileIndex(store.package);

  const top = measured(store.package, window, measurer);
  const middle = {
    ...middleOf(store.package, top.layout),
    width: window.width,
    height: window.height,
  };
  const scrolled = measured(store.package, middle, measurer);

  const saved = saveXlsx(store.package);
  const saveReads = readXlsx(saved).ok;

  const edit = store.transact([
    { op: 'setCellValue', sheet: HOSTILE_SHEET, ref: 'A1', value: { type: 'string', value: 'x' } },
  ]);
  let reopenedFirst: string | null = null;
  if (edit.ok) {
    const again = readXlsx(saveXlsx(store.package));
    reopenedFirst = again.ok ? shownAt(again.pkg, 0, 0) : null;
  }
  return {
    outcome: 'open',
    firstCell: shownAt(read.pkg, 0, 0),
    lastCell: LAST_CELL.has(testCase.id) ? lastShown(read.pkg) : undefined,
    storedRows: sheet?.rows.size ?? 0,
    storedCells: sheet?.cells.size ?? 0,
    top: top.work,
    scrolled: scrolled.work,
    edited: edit.ok,
    ...(edit.ok
      ? {}
      : {
          editRefusal: `${edit.rejection.code}: ${edit.rejection.detail ?? ''}`,
          editRefusalCode: edit.rejection.code,
        }),
    reopenedFirst,
    openedContent: inflatedBytes(bytes),
    savedContent: inflatedBytes(saved),
    saveReads,
    unreadParts,
    firstSheet,
  };
}

/** Each way that the work of one window passes its bounds, in words. */
function windowProblems(
  name: string,
  work: WindowWork,
  cap: BoundedWork,
  storedCells: number
): string[] {
  const problems: string[] = [];
  for (const k of BOUNDED)
    if (work[k] > cap[k]) problems.push(`${name}: ${k} ${work[k]} passes the cap of ${cap[k]}`);
  const steps = work.searches * Math.ceil(Math.log2(storedCells + 1));
  if (work.searchSteps > steps)
    problems.push(`${name}: ${work.searchSteps} search steps for ${work.searches} searches`);
  return problems;
}

/** Each way that `run` breaks the rules of its case, in words, or an empty list. */
export function engineProblems(testCase: HangCase, run: EngineRun, caps: WindowCaps): string[] {
  const cap = WIDE_WINDOW.has(testCase.id) ? caps.wide : caps.window;
  const problems: string[] = [];
  if (run.outcome !== testCase.expect)
    problems.push(`expected ${testCase.expect}, got ${run.outcome} ${run.refusal ?? ''}`);
  if (run.outcome === 'refuse') {
    if (run.refusal && showsException(run.refusal))
      problems.push(`the refusal shows an exception: ${run.refusal}`);
    return problems;
  }
  if (testCase.firstCell !== undefined && !run.firstCell?.includes(testCase.firstCell))
    problems.push(`cell A1 shows ${run.firstCell}, not ${testCase.firstCell}`);
  const last = LAST_CELL.get(testCase.id);
  if (last !== undefined && run.lastCell !== last)
    problems.push(`the last cell shows ${run.lastCell}, not ${last}`);
  problems.push(...windowProblems('the window at the top', run.top!, cap, run.storedCells!));
  problems.push(...windowProblems('the scrolled window', run.scrolled!, cap, run.storedCells!));
  const reads = run.top!.rowsMeasured + run.top!.rowsCached;
  if (reads > run.storedRows!)
    problems.push(`the layout read ${reads} row heights of ${run.storedRows} stored rows`);
  const { indexBuilds, rowsMeasured, rowsCached } = run.scrolled!;
  if (indexBuilds + rowsMeasured + rowsCached > 0)
    problems.push(
      `the scroll built ${indexBuilds} sheet indexes and read ${rowsMeasured} measured and ` +
        `${rowsCached} cached row heights`
    );
  if (!run.saveReads) problems.push('the saved copy does not read');
  if (run.firstSheet !== HOSTILE_SHEET) problems.push(`the first sheet is ${run.firstSheet}`);
  const unread = STAYS_BINARY.has(testCase.id) ? [] : (run.unreadParts ?? []);
  if (unread.length > 0) problems.push(`the read kept modeled parts as bytes: ${unread.join()}`);
  const refusal = EDIT_REFUSED.get(testCase.id);
  if (run.edited === (refusal !== undefined))
    problems.push(
      refusal
        ? 'the store took the edit of A1, which it must refuse'
        : `the store refused the edit of A1: ${run.editRefusal}`
    );
  else if (refusal && run.editRefusalCode !== refusal.code)
    problems.push(`the store refused the edit of A1 with ${run.editRefusal}, not ${refusal.code}`);
  if (run.edited && run.reopenedFirst !== 'x')
    problems.push(`the edited copy reopens with ${run.reopenedFirst} in A1`);
  if (SAVE_KEEPS_CONTENT.has(testCase.id) && run.savedContent! * 2 < run.openedContent!)
    problems.push(`the save kept ${run.savedContent} of ${run.openedContent} bytes of content`);
  return problems;
}
