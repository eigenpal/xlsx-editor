// `x:dimension/@ref`, the used range of a worksheet (ECMA-376 Part 1, 18.3.1.35;
// `CT_SheetDimension`, sml.xsd, where `@ref` is a required `ST_Ref`). The store never adds the
// element and never narrows it; a write that creates a cell outside it widens it in the same op, so
// a consumer that trusts the range sees every cell. The op's inverse sets the old range back, so
// undo restores it.
//
// Security (AGENTS.md, "Security"): the range is a file value. It is parsed with the bounded cell
// reference reader and never drives a loop or an allocation.

import type { XmlPart } from '../package/tree.ts';
import { attr, child } from '../package/query.ts';
import { formatCellRef, parseCellRef } from '../cell-ref.ts';
import { childIndex, withAttr, withChild } from './tree-edit.ts';

/** The 0-based bounds of `A1` or `A1:C3` as `[top, left, bottom, right]`, or undefined. */
export function dimensionBounds(ref: string): [number, number, number, number] | undefined {
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

/** The part with `x:dimension/@ref` set, and the old value; undefined when nothing changes. */
export function withDimensionRef(
  part: XmlPart,
  ref: string
): { readonly part: XmlPart; readonly old: string } | undefined {
  const dimension = child(part.root, 'x:dimension');
  const old = attr(dimension, 'ref');
  if (!dimension || old === undefined || old === ref) return undefined;
  const root = withChild(
    part.root,
    childIndex(part.root, dimension),
    withAttr(dimension, 'ref', ref)
  );
  return { part: { ...part, root }, old };
}

/**
 * The part with `x:dimension/@ref` widened to hold the 0-based cell, and the old value;
 * undefined when the sheet has no readable dimension, or the cell is inside it already.
 */
export function widenDimension(
  part: XmlPart,
  row: number,
  col: number
): { readonly part: XmlPart; readonly old: string } | undefined {
  const old = attr(child(part.root, 'x:dimension'), 'ref');
  const box = old === undefined ? undefined : dimensionBounds(old);
  if (!box) return undefined;
  const [top, left, bottom, right] = box;
  if (row >= top && row <= bottom && col >= left && col <= right) return undefined;
  const from = formatCellRef({ row: Math.min(top, row), col: Math.min(left, col) });
  const to = formatCellRef({ row: Math.max(bottom, row), col: Math.max(right, col) });
  return withDimensionRef(part, from === to ? from : `${from}:${to}`);
}
