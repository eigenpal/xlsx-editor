// A run of `restoreRow` ops on one sheet, applied in one pass: the inverse of one command that
// changed many rows (a cleared range and the rows that unsharing changed) puts every row back
// with one rebuild of `x:sheetData`, so undo costs the rows that exist, not rows times ops.
//
// Applying the run in one pass gives the same package as applying its ops one after another:
// the last op for a row wins, and the inverse restores each row as it was before the run. A run
// with a node that has no `@r` falls back to the ops one at a time (`undefined`).

import type { Element, Node } from '../package/tree.ts';
import type { XlsxPackage } from '../package/package.ts';
import { NS } from '../package/namespaces.ts';
import { attr } from '../package/query.ts';
import { MAX_ROWS } from '../cell-ref.ts';
import { rowChange } from './calc.ts';
import { sheetTarget, withPart, withSheetData } from './cells.ts';
import { IdSource, withAttr } from './tree-edit.ts';
import type { HistoryOp, OpResult } from './ops.ts';

type RestoreRow = Extract<HistoryOp, { op: 'restoreRow' }>;

/** Apply `ops` (restoreRow ops on one sheet) in one pass, or undefined to apply them singly. */
export function restoreRows(pkg: XlsxPackage, ops: readonly RestoreRow[]): OpResult | undefined {
  const sheet = ops[0]!.sheet;
  const wanted = new Map<number, Element | null>();
  for (const op of ops) {
    if (op.sheet !== sheet || !Number.isInteger(op.row) || op.row < 1 || op.row > MAX_ROWS)
      return undefined;
    const node = op.node;
    if (
      node !== null &&
      (node.kind !== 'row' ||
        node.ns !== NS.x ||
        node.local !== 'row' ||
        Number(attr(node, 'r')) !== op.row)
    )
      return undefined;
    wanted.set(op.row, node);
  }
  const target = sheetTarget(pkg, sheet);
  if ('code' in target) return undefined;
  const before = new Map<number, Element | null>();
  const children: Node[] = [];
  let lastRow = 0;
  let removed = false;
  let value = false;
  let formula = false;
  const insertUpTo = (limit: number) => {
    for (const [row, node] of pending) {
      if (row >= limit) break;
      if (node) children.push(node);
      pending.delete(row);
    }
  };
  // The rows to insert, in row order: the wanted rows that do not exist yet.
  const pending = new Map<number, Element | null>();
  const existing = new Set<number>();
  for (const node of target.sheetData.children) {
    if (node.type !== 'element' || node.kind !== 'row') continue;
    const raw = attr(node, 'r');
    lastRow = raw === undefined ? lastRow + 1 : Number(raw);
    existing.add(lastRow);
  }
  for (const row of [...wanted.keys()].sort((a, b) => a - b))
    if (!existing.has(row)) pending.set(row, wanted.get(row)!);
  lastRow = 0;
  for (const node of target.sheetData.children) {
    if (node.type !== 'element' || node.kind !== 'row') {
      children.push(node);
      continue;
    }
    const raw = attr(node, 'r');
    const row = raw === undefined ? lastRow + 1 : Number(raw);
    lastRow = row;
    insertUpTo(row);
    if (!wanted.has(row)) {
      // After a removed row, a row without `@r` takes its number as `@r`, so it keeps it.
      children.push(removed && raw === undefined ? withAttr(node, 'r', String(row)) : node);
      continue;
    }
    const next = wanted.get(row)!;
    if (next === null) removed = true;
    before.set(row, node);
    const change = rowChange(node, next);
    value ||= change.value;
    formula ||= change.formula;
    if (next) children.push(next);
  }
  insertUpTo(Infinity);
  for (const [row, node] of wanted)
    if (!before.has(row)) {
      before.set(row, null);
      const change = rowChange(null, node);
      value ||= change.value;
      formula ||= change.formula;
    }
  const sheetData = { ...target.sheetData, children };
  const part = withSheetData(target.part, sheetData, new IdSource(target.part));
  const inverse: HistoryOp[] = [];
  for (const [row, node] of before) inverse.push({ op: 'restoreRow', sheet, row, node });
  return {
    ok: true,
    pkg: withPart(pkg, { kind: 'xml', partName: part.partName, xml: part }),
    inverse,
    value,
    formula,
  };
}
