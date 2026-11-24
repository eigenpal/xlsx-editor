// The worksheet elements outside `x:sheetData` that a structural edit moves (design D4):
// merged ranges, column entries, hyperlinks, the dimension, and the sheet views.
//
// - `x:mergeCell`: moves; grows when an insertion lands inside it; shrinks when a deletion
//   takes part of it; goes when the deletion takes all of it or leaves one cell.
// - `x:col`: moves; an insertion inside an entry splits it, so the new columns have no entry;
//   an entry pushed past the last column is cut there, because no cell is lost.
// - `x:hyperlink`: `@ref` moves like a range and goes with its cells. `@location` is a place in
//   the workbook, written as a reference, so it moves on every sheet.
// - `x:dimension`, `x:sheetView/@topLeftCell`, `x:pane/@topLeftCell`, and `x:selection` move. A
//   view position that the edit deletes goes to the first row or column after the deletion.
// - A frozen pane (`x:pane/@state` of `frozen` or `frozenSplit`, `ST_PaneState`; `CT_Pane`,
//   sml.xsd, default `split`) keeps its band: `ySplit` rows and `xSplit` columns from the cell of
//   `x:sheetView/@topLeftCell` (`CT_SheetView`, ECMA-376 Part 1, 18.3.1.87; approximation "Frozen
//   band origin"), as layout reads them (`layout/panes.ts`). An insertion inside the band grows
//   it, and the top-left cell of the view keeps its place on that axis; a deletion inside it
//   shrinks it. An edit before the band moves the band with the top-left cell of the view. A
//   split that reaches 0 goes, and the pane goes when both go. `x:pane/@topLeftCell` stays below
//   and right of the band, and `@activePane` and each `x:selection/@pane` name only panes that
//   remain (18.3.1.66 and 18.18.52). A frozen view whose `@topLeftCell` is not one cell
//   reference keeps it: layout then starts the band at row 1 and column A, and a moved reference
//   would move the band.
//
// A container left with no item is removed, because its schema needs one (`x:mergeCells`,
// `x:cols`, `x:hyperlinks`). File-supplied counts are written from the items that remain.

import type { Element, Node, XmlPart } from '../package/tree.ts';
import { attr, intAttr } from '../package/query.ts';
import { formatCellRef, parseCellRef } from '../cell-ref.ts';
import { IdSource, withAttr } from '../store/tree-edit.ts';
import type { OpRejection } from '../store/ops.ts';
import { axisLimit, movePosition, moveSpan, pushedOut, splits } from './axis.ts';
import { moveReferences, type ReferenceContext } from './references.ts';
import {
  axisSpan,
  blocked,
  boxRef,
  gridFull,
  moveBox,
  parseBox,
  type CellBox,
} from './sheet-cells.ts';

type Rewrite = Element | null | OpRejection;
const isRefusal = (r: unknown): r is OpRejection =>
  typeof r === 'object' && r !== null && 'code' in r;

/** Map the element children of `parent`: keep, replace, or drop (null). */
function mapChildren(
  parent: Element,
  fn: (el: Element) => Rewrite
): { readonly el: Element; readonly count: number } | OpRejection {
  const out: Node[] = [];
  let changed = false;
  let count = 0;
  for (const node of parent.children) {
    if (node.type !== 'element') {
      out.push(node);
      continue;
    }
    const r = fn(node);
    if (isRefusal(r)) return r;
    if (r !== node) changed = true;
    if (r) {
      out.push(r);
      count += 1;
    }
  }
  return { el: changed ? { ...parent, children: out } : parent, count };
}

/** Move the merged ranges of the edited sheet. */
function moveMerge(ctx: ReferenceContext, name: string, el: Element): Rewrite {
  const ref = attr(el, 'ref') ?? '';
  const box = parseBox(ref);
  if (!box)
    return blocked(`${name}!${ref}`, 'the merged range cannot be read', {
      kind: 'merged-range',
      name: `${name}!${ref}`,
    });
  if (pushedOut(ctx.edit, axisSpan(ctx, box)[1]))
    return gridFull(`${name}!${ref}`, { kind: 'merged-range', name: `${name}!${ref}` });
  const next = moveBox(ctx, box);
  if (!next || (next.r0 === next.r1 && next.c0 === next.c1)) return null;
  const text = boxRef(next);
  return text === ref ? el : withAttr(el, 'ref', text);
}

/** Move one column entry; an insertion inside it gives two entries. */
function moveCol(
  ctx: ReferenceContext,
  name: string,
  el: Element,
  ids: IdSource
): Element[] | OpRejection {
  const min = intAttr(el, 'min');
  const max = intAttr(el, 'max');
  const limit = axisLimit('cols');
  if (min === undefined || max === undefined || min < 1 || max < min || max > limit)
    return blocked(`${name}!x:col`, 'a column entry has no valid min and max', {
      kind: 'columns',
      name,
    });
  const e = ctx.edit;
  const set = (a: number, b: number, base: Element): Element =>
    withAttr(withAttr(base, 'min', String(a + 1)), 'max', String(b + 1));
  if (splits(e, min - 1, max - 1)) {
    const after = moveSpan(e, e.at, max - 1);
    const first = set(min - 1, e.at - 1, el);
    return after ? [first, set(after[0], after[1], { ...el, id: ids.id() })] : [first];
  }
  const span = moveSpan(e, min - 1, max - 1);
  if (!span) return [];
  return span[0] === min - 1 && span[1] === max - 1 ? [el] : [set(span[0], span[1], el)];
}

/** Move a hyperlink: its cells on the edited sheet, and its place in the workbook. */
function moveHyperlink(ctx: ReferenceContext, name: string, el: Element, target: boolean): Rewrite {
  let out = el;
  const location = attr(el, 'location');
  if (location !== undefined && location !== '') {
    const m = moveReferences(location, ctx);
    if (!m.ok)
      return blocked(`${name}!x:hyperlink`, `the location ${m.reason}`, {
        kind: 'hyperlink',
        name,
      });
    if (m.changed) out = withAttr(out, 'location', m.text);
  }
  if (!target) return out;
  const ref = attr(el, 'ref') ?? '';
  const box = parseBox(ref);
  if (!box)
    return blocked(`${name}!x:hyperlink`, `the hyperlink range ${ref} cannot be read`, {
      kind: 'hyperlink',
      name,
    });
  if (pushedOut(ctx.edit, axisSpan(ctx, box)[1]))
    return gridFull(`${name}!${ref}`, { kind: 'hyperlink', name });
  const next = moveBox(ctx, box);
  if (!next) return null;
  return boxRef(next) === ref ? out : withAttr(out, 'ref', boxRef(next));
}

/** A cell reference for a view: moved, or the first position after a deletion. */
function viewCell(ctx: ReferenceContext, ref: string | undefined): string | undefined {
  const box = ref === undefined ? undefined : parseBox(ref);
  if (!box) return ref;
  const e = ctx.edit;
  const rows = e.axis === 'rows';
  const p = rows ? box.r0 : box.c0;
  const q = movePosition(e, p) ?? Math.min(e.kind === 'delete' ? e.at : p, axisLimit(e.axis) - 1);
  return formatCellRef(rows ? { row: q, col: box.c0 } : { row: box.r0, col: q });
}

/**
 * Move a selection (`x:selection`) with the edit. A selection is view state, so it never refuses
 * the edit: an unreadable `@sqref` becomes the active cell, and an active cell that falls
 * outside every kept range moves to the first cell of the first kept range, so the selection
 * stays valid: the active cell lies in the selection.
 */
function moveSelection(ctx: ReferenceContext, el: Element): Element {
  let out = el;
  const active = attr(el, 'activeCell');
  let nextActive = viewCell(ctx, active);
  const sqref = attr(el, 'sqref');
  if (sqref === undefined) {
    return nextActive !== active ? withAttr(out, 'activeCell', nextActive) : out;
  }
  const parts = sqref.split(' ').filter((s) => s !== '');
  const kept: CellBox[] = [];
  let unreadable = false;
  for (const part of parts) {
    const box = parseBox(part);
    if (!box) {
      unreadable = true;
      break;
    }
    const next = moveBox(ctx, box);
    if (next) kept.push(next);
  }
  if (unreadable) kept.length = 0;
  const at = nextActive === undefined ? undefined : parseBox(nextActive);
  const inside = (b: CellBox) =>
    at !== undefined && at.r0 >= b.r0 && at.r0 <= b.r1 && at.c0 >= b.c0 && at.c0 <= b.c1;
  if (kept.length > 0 && !kept.some(inside))
    nextActive = boxRef({ r0: kept[0]!.r0, c0: kept[0]!.c0, r1: kept[0]!.r0, c1: kept[0]!.c0 });
  if (nextActive !== active) out = withAttr(out, 'activeCell', nextActive);
  const text = kept.length ? kept.map(boxRef).join(' ') : (nextActive ?? 'A1');
  if (text !== sqref) out = withAttr(out, 'sqref', text);
  if (unreadable || kept.length !== parts.length) out = withAttr(out, 'activeCellId', undefined);
  return out;
}

type PaneName = 'topLeft' | 'topRight' | 'bottomLeft' | 'bottomRight';

/** Whether the edit inserts lines inside a frozen band `[first, first + n)`. */
const insertsInBand = (ctx: ReferenceContext, first: number, n: number): boolean =>
  ctx.edit.kind === 'insert' && n > 0 && ctx.edit.at >= first && ctx.edit.at < first + n;

/** The new size of a frozen band `[first, first + n)` after the edit. */
function frozenBand(ctx: ReferenceContext, first: number, n: number): number {
  const e = ctx.edit;
  if (n <= 0) return 0;
  if (e.kind === 'insert') {
    const grown = Math.min(n + e.count, axisLimit(e.axis) - 1 - first);
    return insertsInBand(ctx, first, n) ? Math.max(n, grown) : n;
  }
  const lost = Math.max(0, Math.min(first + n, e.at + e.count) - Math.max(first, e.at));
  return n - lost;
}

/** The first line of a frozen band on one axis: the cell of `x:sheetView/@topLeftCell`, or 0. */
function bandStart(top: string | undefined, rows: boolean): number {
  const cell = top === undefined ? null : parseCellRef(top);
  return cell ? (rows ? cell.row : cell.col) : 0;
}

/** The pane that a pane name becomes when the vertical (`x`) or horizontal (`y`) split goes. */
function remainingPane(name: string | undefined, x: boolean, y: boolean): PaneName {
  const p = (name ?? 'topLeft') as PaneName;
  const right = x && (p === 'topRight' || p === 'bottomRight');
  const bottom = y && (p === 'bottomLeft' || p === 'bottomRight');
  return bottom ? (right ? 'bottomRight' : 'bottomLeft') : right ? 'topRight' : 'topLeft';
}

/** The frozen split of one axis, or 0 when the pane is not frozen or the value is not a count. */
function splitOf(pane: Element | undefined, name: 'xSplit' | 'ySplit'): number {
  const state = attr(pane, 'state');
  if (!pane || (state !== 'frozen' && state !== 'frozenSplit')) return 0;
  const n = Number(attr(pane, name) ?? '0');
  return Number.isInteger(n) && n > 0 ? n : 0;
}

/** The selections of a view with their panes mapped, one for each pane that remains. */
function mapSelections(view: Element, x: boolean, y: boolean, active: string | undefined): Element {
  const kept = new Map<PaneName, Element>();
  const out: Node[] = [];
  for (const node of view.children) {
    if (node.type !== 'element' || node.kind !== 'selection') {
      out.push(node);
      continue;
    }
    const from = attr(node, 'pane') ?? 'topLeft';
    const to = remainingPane(from, x, y);
    const prior = kept.get(to);
    // Two selections for one pane: keep the one of the active pane.
    if (prior && from !== (active ?? 'topLeft')) continue;
    const el = withAttr(node, 'pane', to === 'topLeft' ? undefined : to);
    if (prior) out[out.indexOf(prior)] = el;
    else out.push(el);
    kept.set(to, el);
  }
  return { ...view, children: out };
}

function moveSheetView(ctx: ReferenceContext, name: string, view: Element): Rewrite {
  let out = view;
  const pane = view.children.find((n): n is Element => n.type === 'element' && n.kind === 'pane');
  const rows = ctx.edit.axis === 'rows';
  const splitName = rows ? 'ySplit' : 'xSplit';
  const top = attr(view, 'topLeftCell');
  const first = bandStart(top, rows);
  // The band as layout reads it (`frozenPanes`): cut so that the body keeps a line in the grid.
  const split = splitOf(pane, splitName);
  const band = Math.min(split, Math.max(0, axisLimit(ctx.edit.axis) - 1 - first));
  const nextBand = frozenBand(ctx, first, band);
  // The band grows from its top: the top-left cell of the view keeps its place, also for an
  // insertion at the first line of the band, as for a band at row 1 or column A. A frozen view
  // whose top-left cell does not parse starts its band at row 1 and column A, so it keeps it.
  // A pane with no split is not frozen, so its top-left cell still moves with the edit.
  const frozen = splitOf(pane, 'xSplit') > 0 || splitOf(pane, 'ySplit') > 0;
  const keep =
    insertsInBand(ctx, first, band) || (frozen && top !== undefined && parseCellRef(top) === null);
  const nextTop = keep ? top : viewCell(ctx, top);
  if (nextTop !== top) out = withAttr(out, 'topLeftCell', nextTop);
  // The first line after the band, kept inside the grid: a split can name more lines than it has.
  const nextEnd = Math.min(bandStart(nextTop, rows) + nextBand, axisLimit(ctx.edit.axis) - 1);
  let removedX = false;
  let removedY = false;
  let active = attr(pane, 'activePane');
  const r = mapChildren(out, (el) => {
    if (el.kind === 'pane') {
      const t = attr(el, 'topLeftCell');
      let n = viewCell(ctx, t);
      let p = el;
      if (band > 0) {
        // A split past the grid is written as the count of the band, so the band stays put.
        if (nextBand !== split) p = withAttr(p, splitName, nextBand ? String(nextBand) : undefined);
        // The first scrolling row or column is the first one after the band.
        const box = n === undefined ? undefined : parseBox(n);
        if (box && nextBand > 0 && (rows ? box.r0 : box.c0) < nextEnd)
          n = formatCellRef(rows ? { row: nextEnd, col: box.c0 } : { row: box.r0, col: nextEnd });
        removedY = rows && nextBand === 0;
        removedX = !rows && nextBand === 0;
      }
      if (n !== t) p = withAttr(p, 'topLeftCell', n);
      const noX = removedX || splitOf(p, 'xSplit') === 0;
      const noY = removedY || splitOf(p, 'ySplit') === 0;
      if (band > 0 && noX && noY) return null;
      if (removedX || removedY) {
        active = remainingPane(active, !noX, !noY);
        p = withAttr(p, 'activePane', active === 'topLeft' ? undefined : active);
      }
      return p;
    }
    if (el.kind === 'selection') return moveSelection(ctx, el);
    return blocked(`${name}!x:sheetView`, `the view holds ${el.local}, which can refer to cells`, {
      kind: 'sheet-view',
      name,
    });
  });
  if (isRefusal(r)) return r;
  if (!removedX && !removedY) return r.el;
  const after = r.el.children.find((n): n is Element => n.type === 'element' && n.kind === 'pane');
  const x = after ? splitOf(after, 'xSplit') > 0 : false;
  const y = after ? splitOf(after, 'ySplit') > 0 : false;
  return mapSelections(r.el, x, y, attr(pane, 'activePane'));
}

/** The worksheet part with its elements outside `x:sheetData` moved, or the refusal. */
export function moveSheetElements(
  part: XmlPart,
  name: string,
  ctx: ReferenceContext,
  target: boolean
): XmlPart | OpRejection {
  const ids = new IdSource(part);
  const r = mapChildren(part.root, (el): Rewrite => {
    if (el.kind === 'hyperlinks') {
      const h = mapChildren(el, (x) => moveHyperlink(ctx, name, x, target));
      return isRefusal(h) ? h : h.count ? h.el : null;
    }
    if (!target) return el;
    switch (el.kind) {
      case 'mergeCells': {
        const m = mapChildren(el, (x) => moveMerge(ctx, name, x));
        if (isRefusal(m)) return m;
        if (m.count === 0) return null;
        return m.el === el ? el : withAttr(m.el, 'count', String(m.count));
      }
      case 'cols': {
        if (ctx.edit.axis !== 'cols') return el;
        const out: Node[] = [];
        for (const node of el.children) {
          if (node.type !== 'element' || node.kind !== 'col') {
            out.push(node);
            continue;
          }
          const moved = moveCol(ctx, name, node, ids);
          if (isRefusal(moved)) return moved;
          for (const m of moved) out.push(m);
        }
        return out.some((n) => n.type === 'element') ? { ...el, children: out } : null;
      }
      case 'dimension': {
        const ref = attr(el, 'ref');
        const box = ref === undefined ? undefined : parseBox(ref);
        if (!box) return el;
        const next = moveBox(ctx, box);
        const text = next ? boxRef(next) : 'A1';
        return text === ref ? el : withAttr(el, 'ref', text);
      }
      case 'sheetViews': {
        const v = mapChildren(el, (x) =>
          x.kind === 'sheetView' ? moveSheetView(ctx, name, x) : x
        );
        return isRefusal(v) ? v : v.el;
      }
      default:
        return el;
    }
  });
  if (isRefusal(r)) return r;
  if (r.el === part.root) return part;
  return { partName: part.partName, root: r.el, nextId: ids.next };
}
