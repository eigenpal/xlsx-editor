// The height of a row without `x:row/@ht` (`CT_Row`, sml.xsd; the approximation "Row auto height").
// The standard does not state this height; the rule is measured. It is the tallest of the default
// row height, the row height of the font of each cell and of each rich text run, and, for a cell
// that wraps its text, the row height of its font for each line that the text breaks into at the
// width of its column. The lines come from the same text layout that paints the cell
// (`cell-text.ts`), with the same measurer, so the row fits the lines that layout draws.
//
// A cell inside a merged range does not fit its row, because its text wraps at the width of the
// range. The cost is the cells of the rows without `@ht`, which is in proportion to the content
// of the file. The line count of a wrapped cell is cached per cell element, so an edit
// lays out only the wrapped cells that it changed.
//
// One sheet measures at most `WRAP_MEASURE_BUDGET_CHARS` UTF-16 code units of wrapped text, in
// row order (the approximation "Row height of very long wrapped text"). Each text run also counts
// as one unit, so empty runs spend the budget too. A later wrapped cell takes its line count
// from its line breaks and length, so a hostile sheet cannot block the page. A cached cell counts
// against the budget too, so a height depends on the text before it, and not on an earlier layout.
//
// The height of a row with no cell that wraps is cached per row element, with what it read: the
// style lists, the theme, the measurer, the values of the profile, the default height, and the
// shared string item of each cell. An edit copies only the rows that it changes (`tree-edit.ts`),
// so the other rows keep their elements and their heights, and a value edit measures only the
// rows that it changed. An edit that writes a new styles part, for example a new cell format,
// measures each row again. A row with a wrapped cell is measured again at each layout, because
// its height depends on the wrapped text before it. A cached row, and a cached wrapped cell
// (`fitted`), keep the item elements that they read, not the item list: a new string makes a new
// list, and an old list must not stay in memory for each row or cell that read it. For the same
// reason, each one keeps a token of the measurer, not the measurer: a font change makes a new
// measurer, and the old one must not stay in memory.

import type { Element } from '../store/package/tree.ts';
import { attr } from '../store/package/query.ts';
import { parseCellRef } from '../store/cell-ref.ts';
import { styleLists, type WorkbookIndex } from '../store/workbook/workbook-index.ts';
import type { XlsxPackage } from '../store/package/package.ts';
import {
  cellFormat,
  cellValue,
  displayValue,
  EN_US,
  type ResolvedCellFormat,
  richTextRuns,
  workbookTheme,
} from '../cascade/index.ts';
import type { AxisIndex } from './axis.ts';
import {
  layoutCellText,
  MAX_CELL_LINES,
  type TextContext,
  type TextRunInput,
} from './cell-text.ts';
import { cellSharedStringIndex } from '../cascade/values.ts';
import { autoRowHeight, fontPixels, type GridProfile } from './metrics.ts';
import { mergeAt, type MergeIndex } from './panes.ts';

/**
 * Whether a cell format breaks its text into lines at the cell width: `x:alignment/@wrapText`, or
 * `@horizontal` `justify` or `distributed` (`CT_CellAlignment`), with no `@textRotation`.
 */
export function wraps(format: ResolvedCellFormat): boolean {
  const a = format.alignment;
  return (
    a.textRotation === 0 &&
    (a.wrapText || a.horizontal === 'justify' || a.horizontal === 'distributed')
  );
}

/**
 * The UTF-16 code units of wrapped text that one sheet measures. Later cells estimate their
 * lines.
 */
export const WRAP_MEASURE_BUDGET_CHARS = 4_000_000;

interface Fitted {
  /** The characters and the runs of the text of the cell. */
  readonly chars: number;
  readonly format: ResolvedCellFormat;
  readonly width: number;
  /** The shared string item that the cell read, or undefined when it read none. */
  readonly item: Element | undefined;
  /**
   * The measurer, the values of the profile, and the maximum digit width, as one token
   * (`tokenOf`). The token does not keep the measurer in memory.
   */
  readonly text: object;
  readonly height: number;
}

/**
 * The lines of wrapped `runs` past the budget. A line break (LF, or CR with no LF after it) starts
 * a line, as in layout, and each line takes its lines from its length and the maximum digit width.
 * `digit` and `width` are in points.
 */
function estimatedLines(runs: readonly TextRunInput[], digit: number, width: number): number {
  let lines = 0;
  let length = 0;
  // A CR breaks when the next character, also in the next run, is not LF.
  let cr = false;
  for (const { text } of runs) {
    for (let i = 0; i < text.length; i++) {
      const c = text.charCodeAt(i);
      if (c === 10 || cr) {
        lines += Math.max(1, Math.ceil((length * digit) / width));
        length = 0;
        if (lines >= MAX_CELL_LINES) return MAX_CELL_LINES;
      }
      if (c === 10) cr = false;
      else if (c === 13) cr = true;
      else {
        cr = false;
        length += 1;
      }
    }
  }
  if (cr) {
    lines += Math.max(1, Math.ceil((length * digit) / width));
    length = 0;
  }
  // As in layout, a final line break starts no line.
  if (length === 0 && lines > 0) return Math.min(MAX_CELL_LINES, lines);
  return Math.min(MAX_CELL_LINES, lines + Math.max(1, Math.ceil((length * digit) / width)));
}

/** The runs of a rich string with one base font, with what the row height reads from them. */
interface RichRuns {
  readonly runs: readonly TextRunInput[];
  /** The characters of all runs. */
  readonly chars: number;
  /** The row height of the tallest font of the runs. */
  readonly line: number;
}

/** The list of runs that a rich string holds. */
type RichList = Parameters<typeof richTextRuns>[0];

/** The fitted height of a wrapped cell, by cell element. */
const fitted = new WeakMap<Element, Fitted>();

/** The height of a row with no wrapped cell, with what it read. */
interface RowHeight {
  /** The style lists, the theme, the measurer, the profile, and the default height. */
  readonly token: object;
  /** The item list that the row last matched, by its token, so that one check serves a list. */
  itemsToken: object;
  /** The index that each shared string cell of the row names, or -1 for no valid index. */
  readonly indexes: readonly number[];
  /** The item at each of `indexes`, or undefined past the end of the list. */
  readonly read: readonly (Element | undefined)[];
  readonly height: number;
}

const NONE: readonly never[] = [];

/** The height of a row with no wrapped cell, by row element. */
const rowHeights = new WeakMap<Element, RowHeight>();

interface TokenNode {
  readonly next: WeakMap<object, TokenNode>;
  value?: Map<string, object>;
}

const tokens = new WeakMap<object, TokenNode>();

/**
 * The values of a profile that a row height reads. Layout makes a new profile object often. Each
 * field is named, so a new field of `GridProfile` fails to compile until it is here.
 */
const profileKey = (p: GridProfile): string => {
  const fields: Record<keyof GridProfile, number> = {
    pixelsPerInch: p.pixelsPerInch,
    rowPadding: p.rowPadding,
    rowScale: p.rowScale,
  };
  return Object.values(fields).join('|');
};

/** One object for each combination of what a row height reads, other than the row. */
function tokenOf(keys: readonly object[], key: string): object {
  let map = tokens;
  let node: TokenNode | undefined;
  for (const k of keys) {
    node = map.get(k);
    if (!node) map.set(k, (node = { next: new WeakMap() }));
    map = node.next;
  }
  const byKey = (node!.value ??= new Map());
  let token = byKey.get(key);
  if (!token) {
    if (byKey.size > 8) byKey.clear();
    byKey.set(key, (token = {}));
  }
  return token;
}

/** One object for each item list. It does not keep the list in memory. */
const itemTokens = new WeakMap<readonly Element[], object>();

function itemsTokenOf(items: readonly Element[]): object {
  let token = itemTokens.get(items);
  if (!token) itemTokens.set(items, (token = {}));
  return token;
}

/** Whether each shared string cell of a cached row names the same item in `items`. */
function sameItems(hit: RowHeight, items: readonly Element[]): boolean {
  for (let k = 0; k < hit.indexes.length; k++) {
    const i = hit.indexes[k]!;
    if (i >= 0 && items[i] !== hit.read[k]) return false;
  }
  return true;
}

// Work counters for tests, as PPTX Editor's `nodeIndexTestRecorder`
// (`packages/core/src/store/store/node-index.ts` at d4502552): a count is exact on every machine,
// a wall-clock limit is not.
let rowsMeasured = 0;
let rowsCached = 0;

/** @internal Counts the rows whose height layout measures, and the rows that it takes cached. */
export function rowHeightTestRecorder(): {
  readonly measured: number;
  readonly cached: number;
  reset(): void;
} {
  return {
    get measured() {
      return rowsMeasured;
    },
    get cached() {
      return rowsCached;
    },
    reset() {
      rowsMeasured = 0;
      rowsCached = 0;
    },
  };
}

/**
 * The function that gives the height in points of a row without `@ht`, before the row scale of
 * the profile. A row with no cell, or with no cell taller than the default, gets `defaultHeight`.
 */
export function rowAutoHeight(
  pkg: XlsxPackage,
  index: WorkbookIndex,
  items: readonly Element[],
  ctx: TextContext,
  defaultHeight: number,
  merges: MergeIndex
): (row: Element, rowIndex: number, columns: AxisIndex) => number | undefined {
  const formats = new Map<string | undefined, { format: ResolvedCellFormat; height: number }>();
  // What a wrapped cell reads from `ctx`, for the cache of fitted heights.
  const text = tokenOf([ctx.measurer], `${profileKey(ctx.profile)}|${ctx.maxDigitWidth}`);
  const fontHeight = (font: Parameters<typeof fontPixels>[1]) =>
    autoRowHeight(fontPixels(ctx.measurer, font, ctx.profile), ctx.profile);
  let measured = 0;
  // One shared string can serve every row, so its runs are walked once for each base font.
  const richRuns = new Map<object, Map<ResolvedCellFormat['font'], RichRuns>>();
  const richRunsOf = (list: NonNullable<RichList>, font: ResolvedCellFormat['font']): RichRuns => {
    let byFont = richRuns.get(list);
    if (!byFont) richRuns.set(list, (byFont = new Map()));
    let known = byFont.get(font);
    if (!known) {
      const runs = richTextRuns(list, font, pkg, index);
      let chars = 0;
      let line = 0;
      for (const run of runs) {
        chars += run.text.length;
        line = Math.max(line, fontHeight(run.font));
      }
      known = { runs, chars, line };
      byFont.set(font, known);
    }
    return known;
  };
  const wrappedHeight = (
    cell: Element,
    format: ResolvedCellFormat,
    width: number
  ): number | undefined => {
    const hit = fitted.get(cell);
    const itemIndex = attr(cell, 't') === 's' ? (cellSharedStringIndex(cell) ?? -1) : -1;
    const item = itemIndex >= 0 ? items[itemIndex] : undefined;
    if (
      hit &&
      measured + hit.chars <= WRAP_MEASURE_BUDGET_CHARS &&
      hit.format === format &&
      hit.width === width &&
      hit.item === item &&
      hit.text === text
    ) {
      measured += hit.chars;
      return hit.height;
    }
    const v = cellValue(pkg, index, cell, false).value;
    if (v.kind !== 'string') return undefined;
    let rich: RichRuns;
    if (v.runs && v.runs.length > 0) rich = richRunsOf(v.runs, format.font);
    else {
      const text = displayValue(v, format, { locale: EN_US, date1904: index.date1904 }).text;
      rich = {
        runs: [{ text, font: format.font }],
        chars: text.length,
        line: fontHeight(format.font),
      };
    }
    const { runs, chars, line } = rich;
    measured += chars + runs.length;
    if (measured > WRAP_MEASURE_BUDGET_CHARS)
      return (
        estimatedLines(runs, (ctx.maxDigitWidth * 72) / ctx.profile.pixelsPerInch, width) * line
      );
    const laid = layoutCellText(
      {
        runs,
        kind: 'text',
        alignment: format.alignment,
        box: { x: 0, y: 0, width, height: 0 },
      },
      ctx
    );
    const height = Math.max(1, laid.lines.length) * line;
    fitted.set(cell, { chars: chars + runs.length, format, width, item, text, height });
    return height;
  };
  const token = tokenOf(
    [styleLists(pkg, index), workbookTheme(pkg, index), ctx.measurer],
    `${profileKey(ctx.profile)}|${defaultHeight}`
  );
  const itemsToken = itemsTokenOf(items);
  return (row, rowIndex, columns) => {
    const hit = rowHeights.get(row);
    if (hit && hit.token === token && (hit.itemsToken === itemsToken || sameItems(hit, items))) {
      hit.itemsToken = itemsToken;
      rowsCached += 1;
      return hit.height;
    }
    rowsMeasured += 1;
    let tallest = defaultHeight;
    let col = -1;
    let indexes: number[] | undefined;
    let read: (Element | undefined)[] | undefined;
    let wrapped = false;
    for (const node of row.children) {
      if (node.type !== 'element' || node.local !== 'c') continue;
      const r = attr(node, 'r');
      col = r === undefined ? col + 1 : (parseCellRef(r)?.col ?? col + 1);
      const s = attr(node, 's');
      let known = formats.get(s);
      if (!known) {
        const format = cellFormat(pkg, index, s).format;
        known = { format, height: fontHeight(format.font) };
        formats.set(s, known);
      }
      const format = known.format;
      tallest = Math.max(tallest, known.height);
      // The runs of a rich string take their own fonts (text-layout.pdf, row 14).
      if (attr(node, 't') === 's') {
        const i = cellSharedStringIndex(node) ?? -1;
        (indexes ??= []).push(i);
        (read ??= []).push(i >= 0 ? items[i] : undefined);
        const v = cellValue(pkg, index, node, false).value;
        if (v.kind === 'string' && v.runs)
          tallest = Math.max(tallest, richRunsOf(v.runs, format.font).line);
      }
      if (wraps(format)) {
        wrapped = true;
        if (mergeAt(merges, rowIndex, col)) continue;
        const width = columns.sizeOf(col);
        const height = width > 0 ? wrappedHeight(node, format, width) : undefined;
        if (height !== undefined) tallest = Math.max(tallest, height);
      }
    }
    if (!wrapped)
      rowHeights.set(row, {
        token,
        itemsToken,
        indexes: indexes ?? NONE,
        read: read ?? NONE,
        height: tallest,
      });
    return tallest;
  };
}
