// A run of `setCellValue` ops on one worksheet, applied in one pass.
//
// One op at a time copies the row list, re-reads the sheet index, and rebuilds the shared string
// table for every write, so a transaction of N writes costs N times the sheet. Here the run reads
// the sheet once, appends every new shared string in one rebuild, builds each changed row once
// with its writes in column order, and rebuilds `x:sheetData` once (`restoreRows`). The cost is
// O(N log N) in the writes plus the size of the sheet.
//
// The package is the one that the ops give one after another: the last write to a cell wins,
// each cell keeps its format and its other children (`buildCell`), a write that leaves a cell as
// it was changes nothing, and a new cell outside `x:dimension` widens it. The inverse restores
// each changed row as it was before the run, then the old dimension. A run that one op at a time
// would treat differently returns undefined, and the caller applies its ops singly: an op that
// it would refuse (so the refusal is the same), the master of a shared formula (unsharing reads
// the rows that earlier writes changed), and a write that needs a shared string table that it
// cannot create.

import type { Element, Node } from '../package/tree.ts';
import type { XlsxPackage } from '../package/package.ts';
import { attr } from '../package/query.ts';
import { formatCellRef, parseCellRef } from '../cell-ref.ts';
import { cellKey, sheetIndex } from '../workbook/sheet-index.ts';
import { inheritedFormat } from '../workbook/inherited-format.ts';
import { styleLists } from '../workbook/workbook-index.ts';
import { readFormulaInput } from '../formula/format.ts';
import { formulaScope } from '../formula/references.ts';
import { encodeXstring } from '../workbook/xstring.ts';
import { cellOrNestedHasFormula, sameNode } from './calc.ts';
import { isSharedMaster } from './shared-formulas.ts';
import { rangeMasterBlocker } from '../workbook/cell-formula.ts';
import { sharedStringIndices, sharedStringsUndo } from './shared-strings.ts';
import {
  arrayFormulaAt,
  buildCell,
  cellChildBlocker,
  isBareRow,
  isIsoDate,
  sheetTarget,
  withPart,
} from './cells.ts';
import { restoreRows } from './restore-rows.ts';
import { withoutFittedHeight } from './format-range.ts';
import { widenDimension } from './dimension.ts';
import { IdSource, plainAttr, smlElement, withAttr } from './tree-edit.ts';
import {
  MAX_CELL_TEXT,
  type CellInput,
  type HistoryOp,
  type OpResult,
  type WorkbookOp,
} from './ops.ts';

type SetCellValue = Extract<WorkbookOp, { op: 'setCellValue' }>;
type RestoreRow = Extract<HistoryOp, { op: 'restoreRow' }>;

interface Write {
  readonly ref: string;
  readonly value: CellInput;
  readonly formulaText: string | undefined;
  /** A write of the run put a formula in the cell. */
  readonly sawFormula: boolean;
}

/** The column of a cell, by the rule of `cellSlot`: a cell without `@r` follows the one before. */
export function cellColumn(node: Element, lastCol: number): number {
  const ref = attr(node, 'r');
  return ref === undefined ? lastCol + 1 : (parseCellRef(ref)?.col ?? lastCol + 1);
}

/** The writes of a run by row and column, or undefined when one op would act differently. */
function collectWrites(
  pkg: XlsxPackage,
  ops: readonly SetCellValue[]
): { writes: Map<number, Map<number, Write>>; texts: string[] } | undefined {
  const sheet = ops[0]!.sheet;
  const target = sheetTarget(pkg, sheet);
  if ('code' in target) return undefined;
  const cells = sheetIndex(target.part).cells;
  const writes = new Map<number, Map<number, Write>>();
  const texts: string[] = [];
  for (const op of ops) {
    if (op.sheet !== sheet) return undefined;
    const a = parseCellRef(op.ref);
    if (!a) return undefined;
    const v = op.value;
    if (v.type === 'number' && !Number.isFinite(v.value)) return undefined;
    if (v.type === 'date' && !isIsoDate(v.iso)) return undefined;
    // The single-op path refuses a text that is too long, with its reason.
    if (v.type === 'string' && v.value.length > MAX_CELL_TEXT) return undefined;
    let formulaText: string | undefined;
    if (v.type === 'formula') {
      // A refused formula falls back to the single-op path, which returns its reason.
      const read = readFormulaInput(v.text, formulaScope(pkg));
      if (!read.ok) return undefined;
      formulaText = encodeXstring(read.text);
    }
    if (arrayFormulaAt(target.part, a.row, a.col)) return undefined;
    const old = cells.get(cellKey(a.row, a.col))?.el;
    if (old) {
      if (isSharedMaster(old) || rangeMasterBlocker(old)) return undefined;
    }
    if (v.type === 'string') texts.push(v.value);
    let row = writes.get(a.row);
    if (!row) writes.set(a.row, (row = new Map()));
    const sawFormula = (row.get(a.col)?.sawFormula ?? false) || v.type === 'formula';
    row.set(a.col, { ref: op.ref, value: v, formulaText, sawFormula });
  }
  return { writes, texts };
}

/** Apply a run of `setCellValue` ops on one sheet in one pass, or undefined to apply them singly. */
export function setCellValues(
  pkg: XlsxPackage,
  ops: readonly SetCellValue[]
): OpResult | undefined {
  const sheet = ops[0]!.sheet;
  const collected = collectWrites(pkg, ops);
  const target = sheetTarget(pkg, sheet);
  if (!collected || 'code' in target) return undefined;
  // A cell whose children an edit can neither keep nor replace goes to the single-op path, which
  // refuses it with its reason (`cellChildBlocker`).
  const existing = sheetIndex(target.part).cells;
  for (const [r, cols] of collected.writes)
    for (const [col, w] of cols) {
      const old = existing.get(cellKey(r, col))?.el;
      if (old && cellChildBlocker(sheet, w.ref, old)) return undefined;
    }
  const { writes, texts } = collected;
  let next = pkg;
  let strings: ReadonlyMap<string, number> = new Map();
  let sstUndo: HistoryOp | null = null;
  if (texts.length > 0) {
    const r = sharedStringIndices(pkg, target.index, texts);
    if ('code' in r) return undefined;
    sstUndo = sharedStringsUndo(pkg, target.index, r.pkg);
    next = r.pkg;
    strings = r.indices;
  }
  const ids = new IdSource(target.part);
  const targetIndex = sheetIndex(target.part);
  const rowsIndex = targetIndex.rows;
  // The format that a new cell takes from its row or column, and that a cleared cell keeps out.
  const xfs = styleLists(pkg, target.index).cellXfs.length;
  const inherited = (row: number, col: number) => inheritedFormat(targetIndex, row, col, xfs);
  const restores: RestoreRow[] = [];
  let value = false;
  let formula = false;
  // The bounds of the cells that the run creates, for `x:dimension`.
  let box: [number, number, number, number] | undefined;
  for (const r of [...writes.keys()].sort((x, y) => x - y)) {
    const list = [...writes.get(r)!.entries()].sort((x, y) => x[0] - y[0]);
    const orig = rowsIndex.get(r) ?? null;
    const kids = orig?.children ?? [];
    let lastCellAt = -1;
    for (let i = 0; i < kids.length; i += 1) {
      const n = kids[i]!;
      if (n.type === 'element' && n.kind === 'cell') lastCellAt = i;
    }
    const out: Node[] = [];
    let changed = false;
    const emit = (col: number, w: Write, old: Element | null) => {
      const sst = w.value.type === 'string' ? strings.get(w.value.value) : undefined;
      const cell = buildCell(ids, w.ref, old, w.value, sst, w.formulaText, inherited(r, col));
      if (cell === null ? old === null : old !== null && sameNode(old, cell)) {
        if (old) out.push(keep(old, col));
        return;
      }
      changed = true;
      value = true;
      formula ||= w.sawFormula || (old !== null && cellOrNestedHasFormula(old));
      if (cell) out.push(cell);
      if (!old && cell)
        box = box
          ? [Math.min(box[0], r), Math.min(box[1], col), Math.max(box[2], r), Math.max(box[3], col)]
          : [r, col, r, col];
    };
    // A kept cell carries its `@r`, so it keeps its column when a cell before it is removed.
    const keep = (node: Element, col: number): Element =>
      attr(node, 'r') === undefined ? withAttr(node, 'r', formatCellRef({ row: r, col })) : node;
    let j = 0;
    let lastCol = -1;
    for (let i = 0; i < kids.length; i += 1) {
      const node = kids[i]!;
      if (node.type === 'element' && node.kind === 'cell') {
        const c = cellColumn(node, lastCol);
        lastCol = c;
        while (j < list.length && list[j]![0] < c) emit(list[j]![0], list[j++]![1], null);
        if (j < list.length && list[j]![0] === c) emit(c, list[j++]![1], node);
        else out.push(keep(node, c));
        continue;
      }
      // A new cell goes before any trailing non-cell child (`x:extLst`).
      if (i > lastCellAt) while (j < list.length) emit(list[j]![0], list[j++]![1], null);
      out.push(node);
    }
    while (j < list.length) emit(list[j]![0], list[j++]![1], null);
    if (!changed) continue;
    let rowEl = orig ?? smlElement(ids, 'row', 'sheetData', [plainAttr('r', String(r + 1))]);
    // `spans` is a hint that a new cell can make wrong; dropping it is always valid.
    // A value edit fits a row without `@customHeight` to its content again.
    rowEl = withoutFittedHeight(withAttr(rowEl, 'spans', undefined));
    if (attr(rowEl, 'r') === undefined) rowEl = withAttr(rowEl, 'r', String(r + 1));
    const node = { ...rowEl, children: out };
    // A row that the run leaves bare is dropped, as the single-op path does.
    restores.push({ op: 'restoreRow', sheet, row: r + 1, node: isBareRow(node) ? null : node });
  }
  // A run that changes no cell changes nothing, as one write that changes nothing does.
  if (restores.length === 0) return { ok: true, pkg, inverse: [], value: false, formula: false };
  // The new ids continue the part's counter, so the rebuilt part keeps it.
  const counted = withPart(next, {
    kind: 'xml',
    partName: target.part.partName,
    xml: { ...target.part, nextId: ids.next },
  });
  const rows = restoreRows(counted, restores);
  if (!rows?.ok) return undefined;
  let result = rows.pkg;
  const inverse: HistoryOp[] = [...rows.inverse];
  if (box) {
    const after = sheetTarget(result, sheet);
    if ('code' in after) return undefined;
    const first = widenDimension(after.part, box[0], box[1]);
    const second = widenDimension(first?.part ?? after.part, box[2], box[3]);
    const widened = second?.part ?? first?.part;
    const old = first?.old ?? second?.old;
    if (widened && old !== undefined) {
      result = withPart(result, { kind: 'xml', partName: widened.partName, xml: widened });
      inverse.push({ op: 'setDimension', sheet, ref: old });
    }
  }
  // The removal of a shared string table that the run created comes last.
  if (sstUndo) inverse.push(sstUndo);
  return { ok: true, pkg: result, inverse, value, formula };
}
