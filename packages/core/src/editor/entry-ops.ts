// The ops that write typed entries, DOM-free: the value of each cell, and the number format of
// its typed form (`typedInput`). A paste groups the cells that take the same format into
// rectangles, so it writes one format op for each rectangle, not one for each cell.
//
// This module and `typed-forms.ts` are in the editor lane next to `cell-input.ts`: this module
// uses its entry type, and `cell-input.ts` uses the readers of `typed-forms.ts`. They import no
// DOM. When a headless caller needs typed input (the automation API), they move to the
// automation lane with `cell-input.ts`.

import {
  child,
  children,
  formatCellRef,
  lowerNumberFormat,
  MAX_STYLE_RECORDS,
  protectionRefusal,
  styleLists,
  workbookIndex,
  xmlPartOf,
  type TransactResult,
  type WorkbookOp,
  type XlsxPackage,
} from '../store/index.ts';
import type { TypedEntry } from './cell-input.ts';
import type { NumberFormatRef } from './typed-forms.ts';
import type { GridPoint } from '../contracts/index.ts';

/** A format op, or undefined when the protection of the sheet does not allow it. */
function formatOp(
  pkg: XlsxPackage,
  sheet: string,
  ref: string,
  numberFormat: NumberFormatRef
): WorkbookOp | undefined {
  const op: WorkbookOp = { op: 'applyFormat', sheet, ref, patch: { numberFormat } };
  return protectionRefusal(pkg, op) ? undefined : op;
}

/**
 * The ops that write a typed entry to the cell `ref`: the value, and the number format of its
 * form. A sheet whose protection does not allow formatting cells takes the value only, so a
 * typed `12%` in an unlocked cell of a protected sheet is stored as the number 0.12 with the
 * cell's own format. The caller checks `formatsFit` before it writes the format op: a full style
 * part takes the value only too, so a typed entry is never refused for its format.
 */
export function entryOps(
  pkg: XlsxPackage,
  sheet: string,
  ref: string,
  entry: TypedEntry
): { readonly value: WorkbookOp; readonly format?: WorkbookOp } {
  const value: WorkbookOp = { op: 'setCellValue', sheet, ref, value: entry.value };
  const format = entry.numberFormat && formatOp(pkg, sheet, ref, entry.numberFormat);
  return format ? { value, format } : { value };
}

/**
 * Whether the format ops of typed entries apply. When they do not, the caller writes the values
 * with no formats instead of refusing the entry. A cell format list or the number format list
 * holds at most `MAX_STYLE_RECORDS` records. The ops format `cells` cells, and each cell adds at
 * most one record to each list. Far from that cap, with a styles part and format codes that the
 * store can read, the ops apply with no check. Otherwise a dry run of the format ops decides:
 * near the cap, with no styles part or one that the store cannot read, or with a code from the
 * number locale that the store cannot read.
 */
export function formatsFit(
  pkg: XlsxPackage,
  ops: readonly WorkbookOp[],
  cells: number,
  dryRun: (ops: readonly WorkbookOp[]) => TransactResult
): boolean {
  const index = workbookIndex(pkg);
  if (!index || ops.length === 0) return true;
  // The cap counts the records of each list as the file holds them, also a repeated numFmt id.
  const root = xmlPartOf(pkg, index.stylesPart)?.root;
  // A code from the host locale can be one that the store cannot read: the dry run decides.
  const readable = ops.every(
    (op) => op.op !== 'applyFormat' || !('problem' in lowerNumberFormat(op.patch))
  );
  if (root?.kind === 'styleSheet' && readable) {
    const numFmts = children(child(root, 'x:numFmts'), 'x:numFmt').length;
    const used = Math.max(styleLists(pkg, index).cellXfs.length, numFmts);
    if (used + cells < MAX_STYLE_RECORDS - 1) return true;
  }
  // A refusal that also applies to the values (protection, the sheet) refuses them in turn.
  return dryRun(ops).ok;
}

/** One cell of a paste that takes the number format of its typed form. */
export interface FormattedCell {
  readonly at: GridPoint;
  readonly numberFormat: NumberFormatRef;
}

/**
 * The format ops of a paste: for each number format, the runs of cells in each column, with
 * the runs of neighboring columns that span the same rows joined into one rectangle. A pasted
 * block or column of `12%` is one op. The grouping is linear in the pasted cells after one sort,
 * and the store applies the run of format ops in one pass over the sheet (`applyFormatRanges`).
 * A sheet whose protection does not allow formatting cells takes none. The caller checks
 * `formatsFit`: a full styles part takes none either.
 */
export function pastedFormatOps(
  pkg: XlsxPackage,
  sheet: string,
  cells: readonly FormattedCell[]
): WorkbookOp[] {
  const byFormat = new Map<string, { format: NumberFormatRef; points: GridPoint[] }>();
  for (const { at, numberFormat } of cells) {
    const key = 'id' in numberFormat ? `#${numberFormat.id}` : `"${numberFormat.code}`;
    let group = byFormat.get(key);
    if (!group) byFormat.set(key, (group = { format: numberFormat, points: [] }));
    group.points.push(at);
  }
  const ops: WorkbookOp[] = [];
  for (const { format, points } of byFormat.values()) {
    points.sort((a, b) => a.col - b.col || a.row - b.row);
    // The runs of each column, in column order: [col, top, bottom].
    const runs: [number, number, number][] = [];
    for (const p of points) {
      const last = runs[runs.length - 1];
      if (last && last[0] === p.col && last[2] === p.row - 1) last[2] = p.row;
      else runs.push([p.col, p.row, p.row]);
    }
    // A run joins the open rectangle of the same rows that ends in the column before it.
    const open = new Map<string, [number, number, number, number]>();
    const boxes: [number, number, number, number][] = [];
    for (const [col, top, bottom] of runs) {
      const key = `${top}:${bottom}`;
      const box = open.get(key);
      if (box && box[3] === col - 1) box[3] = col;
      else {
        const next: [number, number, number, number] = [top, col, bottom, col];
        open.set(key, next);
        boxes.push(next);
      }
    }
    for (const [top, left, bottom, right] of boxes) {
      const a = formatCellRef({ row: top, col: left });
      const b = formatCellRef({ row: bottom, col: right });
      const op = formatOp(pkg, sheet, a === b ? a : `${a}:${b}`, format);
      if (!op) return [];
      ops.push(op);
    }
  }
  return ops;
}
