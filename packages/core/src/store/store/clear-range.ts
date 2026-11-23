// The `clearValues` op: clear the values and formulas of a range, keep the cell formats.
//
// A range can name the whole sheet. The op never walks the range: it reads the rows and cells
// that exist (AGENTS.md, "Security"), rebuilds each changed row once, and records one
// `restoreRow` inverse for each, so one Delete is one undo step. A cell inside an array or data
// table range refuses the op; the master of a shared formula inside the range is unshared first,
// so its children outside the range keep their formulas (design D4).

import type { Element, Node, XmlPart } from '../package/tree.ts';
import type { XlsxPackage } from '../package/package.ts';
import { attr, child } from '../package/query.ts';
import { formatCellRef, parseCellRef } from '../cell-ref.ts';
import { sheetIndex, type CellEntry, type SheetIndex } from '../workbook/sheet-index.ts';
import { inheritedFormat } from '../workbook/inherited-format.ts';
import { styleLists } from '../workbook/workbook-index.ts';
import {
  formulaRangeSubject,
  isSharedMaster,
  unshareBlocker,
  unshareChildren,
} from './shared-formulas.ts';
import type { OpRejection } from './ops.ts';
import {
  arrayFormulaIn,
  buildCell,
  cellChildBlocker,
  holdsAlternateValue,
  holdsWrappedValue,
  isBareRow,
  pinCellRefs,
  pinRowRefs,
  sheetTarget,
  withPart,
} from './cells.ts';
import { childIndex, IdSource, withAttr, withChild } from './tree-edit.ts';
import { refuse, type OpResult, type HistoryOp } from './ops.ts';
import { withoutFittedHeight } from './format-range.ts';
import { cellOrNestedHasFormula } from './calc.ts';
import { rangeMasterBlocker, rangeMasterRef } from '../workbook/cell-formula.ts';

/**
 * A range of cells as its 0-based, inclusive bounds: top row, left column, bottom row, and right
 * column. The store exports take and return ranges in this one shape.
 */
export type RangeBox = readonly [top: number, left: number, bottom: number, right: number];

/** The 0-based bounds of `A1` or `A1:C3`, in any corner order, or undefined. */
export function rangeBounds(ref: string): RangeBox | undefined {
  const [a, b = a, extra] = ref.split(':');
  if (extra !== undefined) return undefined;
  const from = parseCellRef(a!);
  const to = parseCellRef(b!);
  if (!from || !to) return undefined;
  return [
    Math.min(from.row, to.row),
    Math.min(from.col, to.col),
    Math.max(from.row, to.row),
    Math.max(from.col, to.col),
  ];
}

/**
 * Whether a cell holds a value or a formula that clearing removes: an `x:v` or `x:is` (children of
 * `CT_Cell`, sml.xsd), or a formula by the scan rule (`cellOrNestedHasFormula`), so a clear never
 * skips a formula cell. A clear also removes a value in `mc:AlternateContent` (`clearsValue`).
 */
export function holdsContent(cell: Element): boolean {
  return (
    child(cell, 'x:v') !== undefined ||
    child(cell, 'x:is') !== undefined ||
    cellOrNestedHasFormula(cell)
  );
}

/**
 * Whether a clear removes a value from a cell: direct value content, or a value in an
 * `mc:AlternateContent` child that holds only value content (`holdsAlternateValue`). The sheet
 * protection check reads the same cells.
 */
export const clearsValue = (cell: Element): boolean =>
  holdsContent(cell) || holdsAlternateValue(cell);

/**
 * Why a clear of the values in `box` of a worksheet is refused for a child of a cell, or
 * undefined (`cellChildBlocker`). It reads each cell that holds a value in any form: one that
 * the clear removes, and one in a child that the clear can neither keep nor replace, such as a
 * formula in `mc:AlternateContent`. A cell with no value in any form is not read, so a clear
 * never refuses a cell that it leaves as it is.
 */
export function clearBlocker(part: XmlPart, sheet: string, box: RangeBox): OpRejection | undefined {
  const [top, left, bottom, right] = box;
  for (const e of blockedCells(sheetIndex(part))) {
    if (e.row < top || e.row > bottom || e.col < left || e.col > right) continue;
    return cellChildBlocker(sheet, formatCellRef(e), e.el);
  }
  return undefined;
}

/** The cells that refuse a clear (`clearBlocker`), in index order, for each sheet index. */
const blockedOf = new WeakMap<SheetIndex, readonly CellEntry[]>();

/**
 * The cells of a sheet index that refuse a clear: one pass for each revision of the sheet, so
 * the editor state can read a selection with no scan of the cells that it does not block.
 */
function blockedCells(index: SheetIndex): readonly CellEntry[] {
  const hit = blockedOf.get(index);
  if (hit) return hit;
  const found: CellEntry[] = [];
  for (const e of index.cells.values()) {
    if (!clearsValue(e.el) && !holdsWrappedValue(e.el)) continue;
    if (cellChildBlocker('', formatCellRef(e), e.el)) found.push(e);
  }
  blockedOf.set(index, found);
  return found;
}

/**
 * The first cell in `box` of a worksheet that holds its value only in `mc:AlternateContent`
 * (`holdsAlternateValue`), as a cell reference, or undefined. The copy reads the direct value
 * content of a cell, so a cut of that cell would copy nothing and the clear of its paste would
 * remove the value: the editor refuses such a cut with `cut-wrapped-value`.
 */
export function wrappedValueCell(part: XmlPart, box: RangeBox): string | undefined {
  const [top, left, bottom, right] = box;
  for (const e of sheetIndex(part).cells.values()) {
    if (e.row < top || e.row > bottom || e.col < left || e.col > right) continue;
    // The copy reads the direct `x:v`, `x:is`, and `x:f` of the cell, not a nested cell.
    const direct = ['x:v', 'x:is', 'x:f'].some((n) => child(e.el, n) !== undefined);
    if (!direct && holdsAlternateValue(e.el)) return formatCellRef(e);
  }
  return undefined;
}

/** The formula range masters of a sheet index, in index order (`clearRangeBlocker`). */
const mastersOf = new WeakMap<SheetIndex, readonly CellEntry[]>();

/** The cells of a sheet index that hold a formula range master, direct or wrapped, once each. */
function masterCells(index: SheetIndex): readonly CellEntry[] {
  const hit = mastersOf.get(index);
  if (hit) return hit;
  const found: CellEntry[] = [];
  for (const e of index.cells.values())
    if (isSharedMaster(e.el) || rangeMasterBlocker(e.el)) found.push(e);
  mastersOf.set(index, found);
  return found;
}

/**
 * Why a clear of the values of `box` must be refused before it changes a cell, or undefined. It
 * runs the checks of `clearRange` in its order: a cell that the clear can neither keep nor
 * replace (`clearBlocker`, `cell-alternate-content`); an array or data table range that the box
 * meets; a formula range master inside `mc:AlternateContent`; and a shared master that cannot
 * unshare (`shared-formula`). The clear state of the toolbar reads it, so Clear contents and
 * Clear all are disabled with the reason of the refusal. Only the move of a child's references
 * out of the grid is checked at the clear alone. The master cells are found once for each
 * revision of the sheet. Each call filters them by the box, so its cost grows with the number
 * of masters in the sheet, not with the size of the box.
 */
export function clearRangeBlocker(
  part: XmlPart,
  sheet: string,
  box: RangeBox
): OpRejection | undefined {
  const [top, left, bottom, right] = box;
  const blocker = clearBlocker(part, sheet, box);
  if (blocker) return blocker;
  const inArray = arrayFormulaIn(part, top, left, bottom, right);
  if (inArray)
    return {
      code: 'shared-formula',
      detail: `${formatRange(box)} meets the array or data table range ${inArray}`,
      subject: `${sheet}!${inArray}`,
    };
  const masters: CellEntry[] = [];
  for (const e of masterCells(sheetIndex(part))) {
    if (e.row < top || e.row > bottom || e.col < left || e.col > right) continue;
    // A formula range master inside `mc:AlternateContent` refuses the clear, as `clearBlocker`
    // does for its cell, because removing it would leave its range undefined.
    const master = rangeMasterBlocker(e.el);
    if (master)
      return {
        code: 'shared-formula',
        detail: `${formatCellRef(e)} ${master}`,
        subject: formulaRangeSubject(sheet, rangeMasterRef(e.el), e),
      };
    masters.push(e);
  }
  return unshareBlocker(part, sheet, masters);
}

const formatRange = ([top, left, bottom, right]: RangeBox): string => {
  const from = formatCellRef({ row: top, col: left });
  const to = formatCellRef({ row: bottom, col: right });
  return from === to ? from : `${from}:${to}`;
};

export function clearRange(pkg: XlsxPackage, sheet: string, ref: string): OpResult {
  const box = rangeBounds(ref);
  if (!box) return refuse('bad-reference', `${JSON.stringify(ref)} is not a range in the grid`);
  const [top, left, bottom, right] = box;
  const target = sheetTarget(pkg, sheet);
  if ('code' in target) return { ok: false, rejection: target };
  // A cell that the clear can neither keep nor replace, such as a formula in
  // `mc:AlternateContent`, is named first: the edit of that cell is refused the same way.
  const blocker = clearRangeBlocker(target.part, sheet, box);
  if (blocker) return { ok: false, rejection: blocker };
  const xfs = styleLists(pkg, target.index).cellXfs.length;
  const inside = (row: number, col: number) =>
    row >= top && row <= bottom && col >= left && col <= right;
  // Unshare every shared master in the range. `clearRangeBlocker` checked each one.
  const masters = masterCells(sheetIndex(target.part)).filter(
    (e) => isSharedMaster(e.el) && inside(e.row, e.col)
  );
  // A wrapped child inside the range refuses the clear (`clearBlocker`), so no child is left
  // with an `@si` and no master.
  const unshared = unshareChildren(target.part, sheet, masters);
  if ('code' in unshared) return { ok: false, rejection: unshared };
  const part: XmlPart = unshared.part;
  const index = sheetIndex(part);
  const ids = new IdSource(part);
  const rows = new Map<Element, Element | null>();
  const inverse: HistoryOp[] = [];
  let formula = masters.length > 0;
  // The column of each indexed cell in the rows of the range. A cell that the index left out
  // (a sheet with a problem refuses the op earlier) is not cleared.
  const colOf = new Map<Element, number>();
  for (const e of index.cells.values()) if (e.row >= top && e.row <= bottom) colOf.set(e.el, e.col);
  for (const [row, rowEl] of index.rows) {
    if (row < top || row > bottom) continue;
    let changed = false;
    let dropped = false;
    const kids: Node[] = [];
    // The cells after a removed cell keep their columns: each kept cell carries its `@r`.
    const pinned = pinCellRefs(rowEl, row).children;
    for (let i = 0; i < rowEl.children.length; i += 1) {
      const node = rowEl.children[i]!;
      if (node.type !== 'element' || node.kind !== 'cell') {
        kids.push(node);
        continue;
      }
      const at = colOf.get(node);
      if (at === undefined || !inside(row, at) || !clearsValue(node)) {
        kids.push(dropped ? pinned[i]! : node);
        continue;
      }
      if (cellOrNestedHasFormula(node)) formula = true;
      const cleared = buildCell(
        ids,
        formatCellRef({ row, col: at }),
        node,
        { type: 'empty' },
        undefined,
        undefined,
        inheritedFormat(index, row, at, xfs)
      );
      if (cleared) kids.push(cleared);
      else dropped = true;
      changed = true;
    }
    if (!changed) continue;
    // A clear fits a row without `@customHeight` to its content again.
    const cleared = { ...withoutFittedHeight(rowEl), children: kids };
    // A row that the clear leaves bare is dropped, as a single clear does. A kept row carries
    // its `@r`, so it keeps its number when a row before it is dropped.
    rows.set(
      rowEl,
      isBareRow(cleared)
        ? null
        : attr(cleared, 'r') === undefined
          ? withAttr(cleared, 'r', String(row + 1))
          : cleared
    );
    inverse.push({ op: 'restoreRow', sheet, row: row + 1, node: rowEl });
  }
  if (rows.size === 0 && masters.length === 0)
    return { ok: true, pkg, inverse: [], value: false, formula: false };
  const original = index.sheetData!;
  // A removed row keeps the numbers of the rows after it (`pinRowRefs`).
  const removes = [...rows.values()].some((n) => n === null);
  const pinnedRows = removes ? pinRowRefs(original).children : original.children;
  const root = withChild(part.root, childIndex(part.root, original), {
    ...original,
    children: original.children.flatMap((n, i) => {
      if (n.type !== 'element' || !rows.has(n)) return [pinnedRows[i]!];
      const next = rows.get(n);
      return next ? [next] : [];
    }),
  });
  const next: XmlPart = { partName: part.partName, root, nextId: ids.next };
  for (const op of unshared.inverse) inverse.push(op);
  return {
    ok: true,
    pkg: withPart(pkg, { kind: 'xml', partName: next.partName, xml: next }),
    inverse,
    value: true,
    formula,
  };
}
