// The cell index of one worksheet (design D2): cell reference to cell node, and row number to
// row node, built from the `x:row` and `x:c` elements that exist.
//
// Nothing here allocates or loops by a value that the file states: `x:dimension/@ref`,
// `x:row/@spans`, `x:col/@min` and `@max`, and merged ranges are never read. A sheet with three
// cells and a dimension of `A1:XFD1048576` costs three cells of work.
//
// A cell or row without `@r` follows the one before it (quick reference, "Cells and values").
// The schema makes both optional (sml.xsd): `x:row/@r` is an `xsd:unsignedInt` (`CT_Row`), and
// `x:c/@r` is an `ST_CellRef` (`CT_Cell`).
// Two cells with one reference, cells out of order, or a reference outside the grid are invalid
// input: the index records each problem and keeps the first cell for a reference. A row that
// repeats or goes back is left out of the index with its cells. Every consumer reads
// `problems`: the store refuses edits to a sheet that has one, and the viewer reports it. The
// engine never guesses an order. A row skipped for a bad number does not advance the row
// counter, so a following row without `@r` takes the number after the last good row.

import type { Element, XmlPart } from '../package/tree.ts';
import { attr, child } from '../package/query.ts';
import { MAX_COLS, MAX_ROWS, parseCellRef } from '../cell-ref.ts';
import { NS } from '../package/namespaces.ts';

/** One cell of the index. */
export interface CellEntry {
  /** 0-based row. */
  readonly row: number;
  /** 0-based column. */
  readonly col: number;
  readonly el: Element;
  /** The row element that holds the cell. */
  readonly rowEl: Element;
}

/** A problem with the cell structure of a sheet, recorded with the reference where it is. */
export interface SheetProblem {
  readonly code:
    | 'bad-row-number'
    | 'row-out-of-order'
    | 'bad-cell-reference'
    | 'cell-in-wrong-row'
    | 'cell-out-of-order'
    | 'duplicate-cell'
    /** Rows or cells inside a markup-compatibility wrapper: kept verbatim, not indexed. */
    | 'wrapped-content';
  /** The cell reference or the row number, as the file wrote it, when there is one. */
  readonly at?: string;
}

/** The index of one worksheet. */
export interface SheetIndex {
  readonly partName: string;
  readonly root: Element;
  /** `x:sheetData`, when the sheet has one in place. */
  readonly sheetData: Element | undefined;
  /** Cells by {@link cellKey}. */
  readonly cells: ReadonlyMap<number, CellEntry>;
  /** Row elements by 0-based row. */
  readonly rows: ReadonlyMap<number, Element>;
  /** Cell keys in document order. */
  readonly order: readonly number[];
  /**
   * The number of indexed cells with a direct `x:f` child, counted in the same pass. It skips
   * the cells that the index skips and an `x:f` inside `mc:AlternateContent` in a cell. So the
   * stale flag reads `workbookHasFormula`, and the notice reads `sheetMayHoldFormula`.
   */
  readonly formulas: number;
  readonly problems: readonly SheetProblem[];
}

/** The map key of a cell: unique for every cell of the grid, and ordered row by row. */
export const cellKey = (row: number, col: number): number => row * MAX_COLS + col;

/** The row and column of a {@link cellKey}. */
export function keyAddress(key: number): { row: number; col: number } {
  return { row: Math.floor(key / MAX_COLS), col: key % MAX_COLS };
}

/** A positive decimal integer of up to seven digits, or undefined. No regex, no `Number()`. */
function rowNumber(raw: string | undefined): number | undefined {
  if (raw === undefined || raw.length === 0 || raw.length > 7) return undefined;
  let n = 0;
  for (let i = 0; i < raw.length; i += 1) {
    const c = raw.charCodeAt(i);
    if (c < 48 || c > 57) return undefined;
    n = n * 10 + (c - 48);
  }
  return n >= 1 && n <= MAX_ROWS ? n : undefined;
}

const cache = new WeakMap<XmlPart, SheetIndex>();

// Work counters for tests, as PPTX Editor's `nodeIndexTestRecorder`
// (`packages/core/src/store/store/node-index.ts` at d4502552): a count is exact on every machine,
// a wall-clock limit is not.
let sheetIndexBuilds = 0;
let sheetIndexCells = 0;

/** @internal Counts the builds of `sheetIndex` and the cells that they index. */
export function sheetIndexTestRecorder(): {
  readonly builds: number;
  readonly cells: number;
  reset(): void;
} {
  return {
    get builds() {
      return sheetIndexBuilds;
    },
    get cells() {
      return sheetIndexCells;
    },
    reset() {
      sheetIndexBuilds = 0;
      sheetIndexCells = 0;
    },
  };
}

/** The cell index of a worksheet part. Cached per part revision. */
export function sheetIndex(part: XmlPart): SheetIndex {
  const hit = cache.get(part);
  if (hit) return hit;
  const root = part.root;
  const sheetData = root.kind === 'worksheet' ? child(root, 'x:sheetData') : undefined;
  const cells = new Map<number, CellEntry>();
  const rows = new Map<number, Element>();
  const order: number[] = [];
  const problems: SheetProblem[] = [];
  let formulas = 0;
  let lastRow = -1;
  for (const rowNode of sheetData?.children ?? []) {
    if (rowNode.type !== 'element') continue;
    if (rowNode.kind !== 'row') {
      if (rowNode.ns === NS.mc) problems.push({ code: 'wrapped-content' });
      continue;
    }
    const rawR = attr(rowNode, 'r');
    let row: number;
    if (rawR === undefined) row = lastRow + 1;
    else {
      const parsed = rowNumber(rawR);
      if (parsed === undefined) {
        problems.push({ code: 'bad-row-number', at: rawR.slice(0, 16) });
        continue;
      }
      row = parsed - 1;
    }
    if (row <= lastRow || row >= MAX_ROWS) {
      problems.push({ code: 'row-out-of-order', at: String(row + 1) });
      continue;
    }
    lastRow = row;
    rows.set(row, rowNode);
    let lastCol = -1;
    for (const cellNode of rowNode.children) {
      if (cellNode.type !== 'element') continue;
      if (cellNode.kind !== 'cell') {
        if (cellNode.ns === NS.mc) problems.push({ code: 'wrapped-content', at: String(row + 1) });
        continue;
      }
      const ref = attr(cellNode, 'r');
      let col: number;
      if (ref === undefined) col = lastCol + 1;
      else {
        const address = parseCellRef(ref);
        if (!address) {
          problems.push({ code: 'bad-cell-reference', at: ref.slice(0, 16) });
          continue;
        }
        if (address.row !== row) {
          problems.push({ code: 'cell-in-wrong-row', at: ref });
          continue;
        }
        col = address.col;
      }
      if (col >= MAX_COLS) {
        problems.push({ code: 'bad-cell-reference', at: ref ?? '' });
        continue;
      }
      const key = cellKey(row, col);
      if (cells.has(key)) {
        problems.push({ code: 'duplicate-cell', at: ref ?? '' });
        continue;
      }
      if (col <= lastCol) problems.push({ code: 'cell-out-of-order', at: ref ?? '' });
      lastCol = Math.max(lastCol, col);
      cells.set(key, { row, col, el: cellNode, rowEl: rowNode });
      order.push(key);
      if (child(cellNode, 'x:f')) formulas += 1;
    }
  }
  const index: SheetIndex = {
    partName: part.partName,
    root,
    sheetData,
    cells,
    rows,
    order,
    formulas,
    problems,
  };
  sheetIndexBuilds += 1;
  sheetIndexCells += order.length;
  cache.set(part, index);
  return index;
}
