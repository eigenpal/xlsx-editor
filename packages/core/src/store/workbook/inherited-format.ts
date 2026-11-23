// The cell format that a cell with no `x:c` element shows: the format of its row
// (`x:row/@s` when `@customFormat` is set, ECMA-376 Part 1, 18.3.1.73), else the format of its
// column (`x:col/@style`, 18.3.1.13), else the default cell format 0. The store starts a new
// cell from it, the editor reads the format state from it, and layout paints the fill and the
// borders of an empty cell with it, so all three agree. An index past the cell formats counts
// as 0, as a cell's own `@s` does (`cellFormat`), so a new cell never takes an index that the
// part check refuses.
//
// Schema types (sml.xsd): `x:row/@s` (`CT_Row`) and `x:col/@style` (`CT_Col`) are
// `xsd:unsignedInt` with default 0, and `x:row/@customFormat` is an `xsd:boolean` with default
// false. `x:col/@min` and `@max` are required `xsd:unsignedInt` column numbers, 1-based.
//
// Nothing here loops over a range that the file states: an `x:col` range is clipped to the
// columns that the caller asks for. The column spans of a sheet are sorted once and cached by
// its root element; the tree is immutable, so a changed sheet has a new root and new spans.

import type { Element } from '../package/tree.ts';
import { attr, boolAttr, child, children } from '../package/query.ts';
import type { SheetIndex } from './sheet-index.ts';

/**
 * A non-negative decimal integer of up to 9 digits, or undefined. The 9-digit cap is an engine
 * limit and is less than the `xsd:unsignedInt` range of the schema.
 */
export function formatIndex(raw: string | undefined): number | undefined {
  if (raw === undefined || raw.length === 0 || raw.length > 9) return undefined;
  let n = 0;
  for (let i = 0; i < raw.length; i += 1) {
    const c = raw.charCodeAt(i);
    if (c < 48 || c > 57) return undefined;
    n = n * 10 + (c - 48);
  }
  return n;
}

/** The cell format of a row (`x:row/@s` with `@customFormat`), or undefined. */
export function rowFormat(row: Element | null | undefined): number | undefined {
  if (!row || !boolAttr(row, 'customFormat')) return undefined;
  return formatIndex(attr(row, 's')) ?? 0;
}

/** Disjoint column spans `[from, to, style]` of a sheet, ascending, by its `x:worksheet` root. */
const spanCache = new WeakMap<Element, readonly (readonly [number, number, number])[]>();

/**
 * The column format spans of a sheet: the `x:col` ranges with a style other than 0, 0-based,
 * sorted by `min` and made disjoint. Where styled ranges overlap, the first one by `min` wins.
 * Built once for each sheet root (the tree is immutable, so a changed sheet has a new root).
 * The cost is the number of `x:col` elements and a sort, never their widths.
 */
function columnSpans(root: Element): readonly (readonly [number, number, number])[] {
  const hit = spanCache.get(root);
  if (hit) return hit;
  const all: [number, number, number][] = [];
  for (const el of children(child(root, 'x:cols'), 'x:col')) {
    const min = formatIndex(attr(el, 'min'));
    const max = formatIndex(attr(el, 'max'));
    const style = formatIndex(attr(el, 'style'));
    if (min === undefined || max === undefined || style === undefined || style === 0) continue;
    if (max >= 1 && min <= max) all.push([Math.max(min - 1, 0), max - 1, style]);
  }
  // A stable sort keeps the document order of ranges with the same `min`.
  all.sort((a, b) => a[0] - b[0]);
  const spans: [number, number, number][] = [];
  let next = 0;
  for (const [from, to, style] of all) {
    const start = Math.max(from, next);
    if (start > to) continue;
    spans.push([start, to, style]);
    next = to + 1;
  }
  spanCache.set(root, spans);
  return spans;
}

/** The position of the first span that ends at or after `col`. */
function firstSpanFrom(spans: readonly (readonly [number, number, number])[], col: number): number {
  let lo = 0;
  let hi = spans.length;
  while (lo < hi) {
    const mid = (lo + hi) >> 1;
    if (spans[mid]![1] < col) lo = mid + 1;
    else hi = mid;
  }
  return lo;
}

/**
 * The column formats (`x:col/@style`) other than 0 of the 0-based columns `left` to `right`, as
 * a sparse map. Where styled `x:col` ranges overlap, the first one by `min` wins, for any
 * window, so every caller reads one rule (`columnSpans`). The cost is a binary search and the
 * width of the range, never the number of `x:col` elements.
 */
export function columnFormats(root: Element, left: number, right: number): Map<number, number> {
  const spans = columnSpans(root);
  const out = new Map<number, number>();
  for (let i = firstSpanFrom(spans, left); i < spans.length; i += 1) {
    const [from, to, style] = spans[i]!;
    if (from > right) break;
    for (let c = Math.max(from, left); c <= Math.min(to, right); c += 1) out.set(c, style);
  }
  return out;
}

/**
 * The column format (`x:col/@style`) of the 0-based column `col`, or undefined for 0 or none,
 * by the rule of `columnFormats`. A lookup for each cell of a paste costs a binary search, not a
 * walk over every `x:col`.
 */
export function columnFormatAt(root: Element, col: number): number | undefined {
  const spans = columnSpans(root);
  const span = spans[firstSpanFrom(spans, col)];
  return span && span[0] <= col ? span[2] : undefined;
}

/**
 * The cell format that the empty cell at 0-based `row`, `col` shows. With `cellXfs`, the number
 * of cell formats, an index past them is 0.
 */
export function inheritedFormat(
  sheet: SheetIndex,
  row: number,
  col: number,
  cellXfs = Number.POSITIVE_INFINITY
): number {
  const s = rowFormat(sheet.rows.get(row)) ?? columnFormatAt(sheet.root, col) ?? 0;
  return s < Math.max(cellXfs, 1) ? s : 0;
}
