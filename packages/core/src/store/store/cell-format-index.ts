// The `setCellFormatIndex` op (design D2 of the change `workbook-structure-editing`): give a cell
// the cell format with a known index (`x:c/@s`, an index into `x:cellXfs`, ECMA-376 Part 1,
// 18.3.1.4; in `CT_Cell`, sml.xsd, `@s` is an `xsd:unsignedInt` with default 0). An internal paste
// sets the format of each target cell to the format of its source cell with it, so the paste never
// appends a style record.
//
// The index is the format that the cell shows. A cell with no `@s` shows format 0, and a cell
// with no `x:c` element shows the format of its row or column (`inheritedFormat`). So:
//
// - A cell that exists gets `@s`, or loses it for format 0. When it then holds nothing and
//   its row or column shows the same format, it goes away. When it holds nothing and the row or
//   column shows another format, it keeps an explicit `@s`, also `s="0"`, as `clearFormats` does.
// - A cell that does not exist is created only when its row or column shows another format.
//
// An index outside `x:cellXfs` is refused with `bad-format`. A run of these ops on one sheet
// applies in one pass (`applyOps`), so a paste of N cells costs the rows it changes, not N
// times the sheet. Each changed row gets one `restoreRow` inverse, and a cell created outside
// `x:dimension/@ref` widens it in the same undo step. The op changes no value, so it never
// drops the calculation chain (`calc.ts`). But a cell format changes the result of
// `CELL("format")`, `CELL("color")`, `CELL("parentheses")`, `CELL("prefix")`,
// `CELL("protect")`, and `GET.CELL`. So when a row changed and a formula of a cell or a defined
// name calls `CELL` or a `GET.*` macro (`CELL_STATE_FUNCTIONS`), the op marks the cached values
// stale, as a value edit and `applyFormat` do. A change of format that can change the row
// height (another font, or wrapping) removes a fitted `x:row/@ht`, as `applyFormat` does.

import type { Element, Node } from '../package/tree.ts';
import type { XlsxPackage } from '../package/package.ts';
import { attr } from '../package/query.ts';
import { formatCellRef, parseCellRef } from '../cell-ref.ts';
import { sheetIndex } from '../workbook/sheet-index.ts';
import { styleLists } from '../workbook/workbook-index.ts';
import { columnFormats, formatIndex, rowFormat } from '../workbook/inherited-format.ts';
import { sheetTarget, withPart } from './cells.ts';
import { cellColumn } from './cell-batch.ts';
import { heightFormats, rowInverses, withoutFittedHeight, withRows } from './format-range.ts';
import { widenDimension } from './dimension.ts';
import { unreadStyles } from './styles-part.ts';
import { IdSource, plainAttr, smlElement, withAttr } from './tree-edit.ts';
import { refuse, unchanged, type OpResult, type WorkbookOp } from './ops.ts';
import { CELL_STATE_FUNCTIONS, workbookCallsSheetFunction } from '../structure/sheet-functions.ts';

type SetCellFormatIndex = Extract<WorkbookOp, { op: 'setCellFormatIndex' }>;

/** Whether a cell holds nothing but its reference: no value, formula, or other attribute. */
const holdsNothing = (cell: Element): boolean =>
  cell.children.length === 0 && cell.attrs.every((a) => a.ns === '' && a.local === 'r');

/**
 * Apply a run of `setCellFormatIndex` ops on one sheet in one pass. The last op for a cell
 * wins. Atomic: one refused op refuses the run.
 */
export function setCellFormatIndexes(
  pkg: XlsxPackage,
  ops: readonly SetCellFormatIndex[]
): OpResult {
  const sheet = ops[0]!.sheet;
  const target = sheetTarget(pkg, sheet);
  if ('code' in target) return { ok: false, rejection: target };
  // While the styles part cannot be read, only the default format 0 can be given.
  const unread = unreadStyles(pkg, target.index);
  // A workbook with no styles part has the implicit default format 0 only.
  const count = Math.max(1, styleLists(pkg, target.index).cellXfs.length);
  const wanted = new Map<number, Map<number, number>>();
  let left = Infinity;
  let right = -1;
  for (const op of ops) {
    if (op.sheet !== sheet) return refuse('bad-value', 'one run of ops names one sheet');
    const a = parseCellRef(op.ref);
    if (!a) return refuse('bad-reference', `${JSON.stringify(op.ref)} is not a cell in the grid`);
    if (unread && op.index !== 0) return { ok: false, rejection: unread };
    if (!Number.isInteger(op.index) || op.index < 0 || op.index >= count)
      return refuse(
        'bad-format',
        `${op.ref}: cell format ${op.index} does not exist; the workbook has ${count}`
      );
    let row = wanted.get(a.row);
    if (!row) wanted.set(a.row, (row = new Map()));
    row.set(a.col, op.index);
    left = Math.min(left, a.col);
    right = Math.max(right, a.col);
  }
  const index = sheetIndex(target.part);
  const colStyles = columnFormats(target.part.root, left, right);
  const refits = heightFormats(pkg, target.index);
  const ids = new IdSource(target.part);
  const changed = new Map<number, Element>();
  // The bounds of the created cells, to widen `x:dimension/@ref` once.
  let made: [number, number, number, number] | undefined;
  for (const r of [...wanted.keys()].sort((x, y) => x - y)) {
    const list = [...wanted.get(r)!.entries()].sort((x, y) => x[0] - y[0]);
    const rowEl = index.rows.get(r) ?? null;
    const fromRow = rowFormat(rowEl);
    const inherited = (col: number) => fromRow ?? colStyles.get(col) ?? 0;
    const kids = rowEl?.children ?? [];
    let lastCellAt = -1;
    for (let i = 0; i < kids.length; i += 1) {
      const n = kids[i]!;
      if (n.type === 'element' && n.kind === 'cell') lastCellAt = i;
    }
    const out: Node[] = [];
    let dirty = false;
    let refit = false;
    const create = (col: number, xf: number) => {
      if (xf === inherited(col)) return;
      const attrs = [plainAttr('r', formatCellRef({ row: r, col })), plainAttr('s', String(xf))];
      out.push(smlElement(ids, 'c', 'row', attrs));
      dirty = true;
      refit ||= refits(inherited(col)) || refits(xf);
      made = made
        ? [
            Math.min(made[0], r),
            Math.min(made[1], col),
            Math.max(made[2], r),
            Math.max(made[3], col),
          ]
        : [r, col, r, col];
    };
    const update = (col: number, xf: number, cell: Element) => {
      const raw = attr(cell, 's');
      const old = formatIndex(raw) ?? 0;
      // An unchanged format writes nothing; a malformed `@s` is written again.
      if (old === xf && (raw === undefined ? xf === 0 : formatIndex(raw) === xf)) {
        out.push(cell);
        return;
      }
      dirty = true;
      refit ||= refits(old) || refits(xf);
      const plain = withAttr(cell, 's', undefined);
      if (!holdsNothing(plain)) out.push(xf === 0 ? plain : withAttr(cell, 's', String(xf)));
      else if (xf !== inherited(col)) out.push(withAttr(cell, 's', String(xf)));
    };
    let j = 0;
    let lastCol = -1;
    for (let i = 0; i < kids.length; i += 1) {
      const node = kids[i]!;
      if (node.type === 'element' && node.kind === 'cell') {
        const col = cellColumn(node, lastCol);
        lastCol = col;
        while (j < list.length && list[j]![0] < col) create(list[j]![0], list[j++]![1]);
        if (j < list.length && list[j]![0] === col) update(col, list[j++]![1], node);
        else out.push(node);
        continue;
      }
      // A new cell goes before any trailing non-cell child (`x:extLst`).
      if (i > lastCellAt) while (j < list.length) create(list[j]![0], list[j++]![1]);
      out.push(node);
    }
    while (j < list.length) create(list[j]![0], list[j++]![1]);
    if (!dirty) continue;
    const base = rowEl ?? smlElement(ids, 'row', 'sheetData', [plainAttr('r', String(r + 1))]);
    // `spans` is a hint that a new cell can make wrong; dropping it is always valid.
    const edited = withAttr({ ...base, children: out }, 'spans', undefined);
    changed.set(r, refit ? withoutFittedHeight(edited) : edited);
  }
  if (changed.size === 0) return unchanged(pkg);
  let part = withRows(target.part, index, changed, ids);
  const inverse = rowInverses(sheet, index, changed);
  // New cells outside `x:dimension/@ref` widen it, in the same op and undo step.
  const box = made as [number, number, number, number] | undefined;
  const first = box && widenDimension(part, box[0], box[1]);
  if (first) part = first.part;
  const last = box && widenDimension(part, box[2], box[3]);
  if (last) part = last.part;
  const oldDimension = first?.old ?? last?.old;
  if (oldDimension !== undefined) inverse.push({ op: 'setDimension', sheet, ref: oldDimension });
  return {
    ok: true,
    pkg: withPart(pkg, { kind: 'xml', partName: part.partName, xml: part }),
    inverse,
    // A cell format changes the result of `CELL("format")`: a call marks the values stale.
    value: workbookCallsSheetFunction(pkg, target.index, CELL_STATE_FUNCTIONS),
    formula: false,
  };
}
