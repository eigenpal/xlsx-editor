// The drawing anchors of the edited sheet under a structural edit (design D4). The anchors are
// those of dml-spreadsheetDrawing.xsd: `CT_TwoCellAnchor`, `CT_OneCellAnchor`, and
// `CT_AbsoluteAnchor`. A corner (`xdr:from`, `xdr:to`) is a `CT_Marker`: `xdr:col`, `xdr:colOff`,
// `xdr:row`, and `xdr:rowOff`, with 0-based cell indexes.
//
// - `xdr:twoCellAnchor` with `@editAs` (`ST_EditAs`) `twoCell` (the default): `xdr:from` and
//   `xdr:to` move with their cells, so the object moves and resizes.
// - `xdr:oneCellAnchor`: `xdr:from` moves with its cell; `xdr:ext` keeps the size.
// - `xdr:absoluteAnchor`: no cell position, so it does not change.
// - `xdr:twoCellAnchor` with `@editAs` `oneCell` or `absolute` keeps its size or position in
//   points, which needs column widths in the store lane; the edit is refused when it would move
//   such an anchor.
//
// A corner in a deleted band goes to the start of the first row or column after it. An
// insertion that pushes a corner past the grid is refused. An anchor inside a
// markup-compatibility wrapper moves like any other. A shape linked to a cell (`@textlink`) and
// a relationship to a part other than an image or a hyperlink refuse the edit.
//
// A drawing part that the reader keeps as bytes, or that is not a `xdr:wsDr`, has anchors that
// the edit cannot read, so it refuses every structural edit and every resize of the sheet
// (`sheetDrawing`). So does a drawing whose relationships part is kept as bytes, for a
// structural edit, because its relationships cannot be checked. An edit never leaves such a
// drawing on its old cells. The refusals of a structural edit that do not depend on its rows or
// columns are `drawingBlocker`, which `structureBlocker` reads as the state of the commands.
//
// A resize (design D3) keeps every anchor in its cells, so a two-cell anchor stretches with
// them. A `xdr:twoCellAnchor` with `@editAs` `oneCell` keeps its size in points, and one with
// `absolute` keeps its position and size, so a resize that would change them is refused
// (`resizeBlocker`): for `oneCell`, a range that meets the anchor's columns or rows; for
// `absolute`, also a range before them. The anchor of an embedded control or object
// (`CT_ObjectAnchor`, `sml.xsd`) keeps its size unless `@sizeWithCells` is true, and its
// position unless `@moveWithCells` is true, with the same rules.

import type { Element, Node, XmlPart } from '../package/tree.ts';
import { attr, child, textOf } from '../package/query.ts';
import { NS, REL } from '../package/namespaces.ts';
import {
  partOf,
  relationshipsOf,
  relsPartName,
  relTarget,
  type XlsxPackage,
} from '../package/package.ts';
import { IdSource, textNode } from '../store/tree-edit.ts';
import type { OpRejection } from '../store/ops.ts';
import { axisLimit, movePosition, type AxisEdit } from './axis.ts';
import { blocked, gridFull } from './sheet-cells.ts';

/** What a drawing refusal names, for user text. */
const DRAWING = { kind: 'drawing' } as const;

/** The deepest markup-compatibility nesting that the anchor walk follows. */
const MAX_WRAP_DEPTH = 4;

type Moved = Element | OpRejection;
const isRefusal = (r: unknown): r is OpRejection =>
  typeof r === 'object' && r !== null && 'code' in r;

/** A non-negative integer from element text, or undefined. */
function indexOf(el: Element | undefined): number | undefined {
  const raw = textOf(el).trim();
  if (raw.length === 0 || raw.length > 8) return undefined;
  let n = 0;
  for (let i = 0; i < raw.length; i += 1) {
    const c = raw.charCodeAt(i);
    if (c < 48 || c > 57) return undefined;
    n = n * 10 + (c - 48);
  }
  return n;
}

const setText = (ids: IdSource, el: Element, text: string): Element => ({
  ...el,
  children: [textNode(ids, text)],
});

/** Move one anchor corner (`xdr:from` or `xdr:to`, `CT_Marker`). */
function movePoint(ids: IdSource, e: AxisEdit, point: Element, subject: string): Moved {
  const key = e.axis === 'rows' ? 'row' : 'col';
  const idxEl = child(point, `xdr:${key}`);
  const p = indexOf(idxEl);
  if (!idxEl || p === undefined || p >= axisLimit(e.axis))
    return blocked(subject, `the anchor ${key} cannot be read`, DRAWING);
  let q = movePosition(e, p);
  let reset = false;
  if (q === null) {
    if (e.kind === 'insert') return gridFull(subject, DRAWING);
    q = e.at;
    reset = true;
  }
  if (q === p && !reset) return point;
  const offEl = child(point, `xdr:${key}Off`);
  return {
    ...point,
    children: point.children.map((n) => {
      if (n === idxEl) return setText(ids, idxEl, String(q));
      if (reset && n === offEl && offEl) return setText(ids, offEl, '0');
      return n;
    }),
  };
}

/** Whether an anchor holds a shape whose text is linked to a cell. */
function linksCell(el: Element, depth = 0): boolean {
  if (depth > 32) return true;
  if (el.local === 'sp' && el.ns === NS.xdr && (attr(el, 'textlink') ?? '') !== '') return true;
  return el.children.some((c) => c.type === 'element' && linksCell(c, depth + 1));
}

function moveAnchor(ids: IdSource, e: AxisEdit, el: Element, subject: string): Moved {
  if (el.kind === 'absoluteAnchor') return el;
  if (linksCell(el)) return blocked(subject, 'a shape takes its text from a cell', DRAWING);
  const editAs = attr(el, 'editAs') ?? 'twoCell';
  const out: Node[] = [];
  let changed = false;
  for (const c of el.children) {
    if (c.type === 'element' && (c.kind === 'anchorFrom' || c.kind === 'anchorTo')) {
      const m = movePoint(ids, e, c, subject);
      if (isRefusal(m)) return m;
      if (m !== c) changed = true;
      out.push(m);
    } else out.push(c);
  }
  if (!changed) return el;
  if (el.kind === 'twoCellAnchor' && editAs !== 'twoCell')
    return blocked(
      subject,
      `an anchor that keeps its ${editAs === 'oneCell' ? 'size' : 'position'} in points`,
      DRAWING
    );
  return { ...el, children: out };
}

/** Walk the anchors of a drawing root, through markup-compatibility wrappers. */
function moveChildren(
  ids: IdSource,
  e: AxisEdit,
  el: Element,
  subject: string,
  depth: number
): Moved {
  const out: Node[] = [];
  let changed = false;
  for (const c of el.children) {
    if (c.type !== 'element') {
      out.push(c);
      continue;
    }
    let m: Moved = c;
    if (c.kind === 'twoCellAnchor' || c.kind === 'oneCellAnchor' || c.kind === 'absoluteAnchor')
      m = moveAnchor(ids, e, c, subject);
    else if (c.ns === NS.mc && depth < MAX_WRAP_DEPTH)
      m = moveChildren(ids, e, c, subject, depth + 1);
    if (isRefusal(m)) return m;
    if (m !== c) changed = true;
    out.push(m);
  }
  return changed ? { ...el, children: out } : el;
}

/**
 * The drawing part of a worksheet: its tree, the name of a part that cannot be read (kept as
 * bytes, or not a `xdr:wsDr`), or undefined when the sheet relates to no drawing part.
 */
function sheetDrawing(
  pkg: XlsxPackage,
  sheet: XmlPart
): { readonly part: XmlPart } | { readonly unread: string } | undefined {
  const relId = attr(child(sheet.root, 'x:drawing'), 'r:id');
  const found = partOf(pkg, relId ? relTarget(pkg, sheet.partName, relId) : undefined);
  if (!found) return undefined;
  if (found.kind !== 'xml' || found.xml.root.kind !== 'wsDr') return { unread: found.partName };
  return { part: found.xml };
}

/**
 * Why the drawing part of a worksheet refuses every structural edit of the sheet, whatever rows
 * or columns it names, or undefined: the part cannot be read, its relationships part is kept as
 * bytes, or it relates to a part other than an image or a hyperlink.
 */
export function drawingBlocker(pkg: XlsxPackage, sheet: XmlPart): OpRejection | undefined {
  return drawingRefusal(pkg, sheetDrawing(pkg, sheet));
}

/** `drawingBlocker` for the drawing that `sheetDrawing` gives. */
function drawingRefusal(
  pkg: XlsxPackage,
  drawing: ReturnType<typeof sheetDrawing>
): OpRejection | undefined {
  if (!drawing) return undefined;
  if ('unread' in drawing)
    return blocked(drawing.unread, 'the drawing part cannot be read', DRAWING);
  const name = drawing.part.partName;
  // A relationships part kept as bytes hides the relationships that the loop below checks.
  const rels = partOf(pkg, relsPartName(name));
  if (rels && rels.kind !== 'xml')
    return blocked(rels.partName, 'the relationships of the drawing part cannot be read', DRAWING);
  for (const rel of relationshipsOf(pkg, name))
    if (rel.type !== REL.image && rel.type !== REL.hyperlink)
      return blocked(rel.target ?? name, 'a drawing object can refer to cells', DRAWING);
  return undefined;
}

/** The moved drawing part of a worksheet, the part unchanged, undefined, or the refusal. */
export function moveDrawing(
  pkg: XlsxPackage,
  sheet: XmlPart,
  e: AxisEdit
): XmlPart | undefined | OpRejection {
  const drawing = sheetDrawing(pkg, sheet);
  const refusal = drawingRefusal(pkg, drawing);
  if (refusal) return refusal;
  if (!drawing || 'unread' in drawing) return undefined;
  const part = drawing.part;
  const ids = new IdSource(part);
  const root = moveChildren(ids, e, part.root, part.partName, 0);
  if (isRefusal(root)) return root;
  return root === part.root ? part : { partName: part.partName, root, nextId: ids.next };
}

/**
 * The first and last cell index that an anchor with `xdr:from` and `xdr:to` covers on one axis,
 * or undefined. A `xdr:to` with offset 0 ends at the start of its cell, so the span ends at the
 * cell before it.
 */
function anchorSpan(el: Element, key: 'row' | 'col'): [number, number] | undefined {
  let from: number | undefined;
  let to: number | undefined;
  let toOff: number | undefined;
  for (const c of el.children) {
    if (c.type !== 'element' || c.ns !== NS.xdr) continue;
    if (c.local === 'from') from = indexOf(child(c, `xdr:${key}`));
    else if (c.local === 'to') {
      to = indexOf(child(c, `xdr:${key}`));
      toOff = indexOf(child(c, `xdr:${key}Off`));
    }
  }
  if (from === undefined || to === undefined) return undefined;
  return [from, toOff === 0 && to > from ? to - 1 : to];
}

/** What a resize must keep of an anchor: its size, its position, or both. */
interface Kept {
  readonly size: boolean;
  readonly position: boolean;
}

/**
 * Why a resize of `first` to `last` would change an anchor that keeps its size or its position,
 * or undefined. A kept size refuses a range that meets the anchor; a kept position refuses a
 * range that meets or comes before its first cell.
 */
function anchorProblem(
  el: Element,
  key: 'row' | 'col',
  first: number,
  last: number,
  kept: Kept
): string | undefined {
  if (!kept.size && !kept.position) return undefined;
  const span = anchorSpan(el, key);
  // An anchor that cannot be read can lie anywhere.
  if (!span) return `the anchor ${key} cannot be read`;
  const meets = first <= span[1] && last >= span[0];
  if (kept.size && meets) return 'the anchor keeps its size in points';
  if (kept.position && first <= span[0]) return 'the anchor keeps its position in points';
  return undefined;
}

/** A refusal of a resize, with the object name as its subject when the file gives one. */
interface AnchorHit {
  readonly name: string | undefined;
  readonly reason: string;
  /** The drawing part cannot be read, so no anchor of it can be checked. */
  readonly unread?: true;
}

/** The `@name` of the first `xdr:cNvPr` under an anchor, or undefined. Bounded by depth. */
function shapeName(el: Element, depth = 0): string | undefined {
  if (depth > 16) return undefined;
  for (const c of el.children) {
    if (c.type !== 'element') continue;
    if (c.ns === NS.xdr && c.local === 'cNvPr') return attr(c, 'name') || undefined;
    const n = shapeName(c, depth + 1);
    if (n !== undefined) return n;
  }
  return undefined;
}

/** The first drawing anchor of `el`, through wrappers, that the resize would change. */
function resizedAnchor(
  el: Element,
  key: 'row' | 'col',
  first: number,
  last: number,
  depth: number
): AnchorHit | undefined {
  for (const c of el.children) {
    if (c.type !== 'element') continue;
    if (c.ns === NS.mc && depth < MAX_WRAP_DEPTH) {
      const hit = resizedAnchor(c, key, first, last, depth + 1);
      if (hit) return hit;
      continue;
    }
    if (c.kind !== 'twoCellAnchor') continue;
    // `@editAs` is an `xsd:token`: surrounding white space does not count.
    const editAs = (attr(c, 'editAs') ?? 'twoCell').trim();
    const kept = {
      size: editAs === 'oneCell' || editAs === 'absolute',
      position: editAs === 'absolute',
    };
    const reason = anchorProblem(c, key, first, last, kept);
    if (reason) return { name: shapeName(c), reason };
  }
  return undefined;
}

const truthy = (v: string | undefined) => {
  const t = v?.trim();
  return t === '1' || t === 'true';
};

/**
 * The first embedded control or object anchor (`x:controlPr/x:anchor`, `x:objectPr/x:anchor`,
 * `CT_ObjectAnchor`) under `el` that the resize would change. `@moveWithCells` and
 * `@sizeWithCells` default to false: such an anchor keeps its position and size in points.
 */
function resizedObject(
  el: Element,
  key: 'row' | 'col',
  first: number,
  last: number,
  depth: number,
  name: string | undefined
): AnchorHit | undefined {
  if (depth > MAX_WRAP_DEPTH * 3) return { name, reason: 'the object is nested too deep' };
  for (const c of el.children) {
    if (c.type !== 'element') continue;
    const own = c.ns === NS.x && c.local === 'control' ? attr(c, 'name') || undefined : name;
    if (c.ns === NS.x && c.local === 'anchor') {
      const kept = {
        size: !truthy(attr(c, 'sizeWithCells')),
        position: !truthy(attr(c, 'moveWithCells')),
      };
      const reason = anchorProblem(c, key, first, last, kept);
      if (reason) return { name, reason };
      continue;
    }
    const hit = resizedObject(c, key, first, last, depth + 1, own);
    if (hit) return hit;
  }
  return undefined;
}

/** Worksheet children that hold the anchors of embedded controls and objects. */
const OBJECT_LISTS = new Set(['controls', 'oleObjects']);

/**
 * Why a resize of the 0-based columns or rows `first` to `last` of a worksheet would change an
 * anchor that keeps its size or position in points, or undefined: a drawing anchor
 * (`xdr:twoCellAnchor` with `@editAs` `oneCell` or `absolute`), or the anchor of an embedded
 * control or object. The refusal is `resize-anchor`, with the shape or control name as its
 * subject when the file gives one. A drawing part that cannot be read refuses every resize of
 * the sheet, with no subject and with `causeUnreadable`. The anchors of a legacy drawing (VML
 * `x:ClientData`) are not read: a resize keeps their cells (approximation
 * `legacy-drawing-anchors`).
 */
export function resizeBlocker(
  pkg: XlsxPackage,
  sheet: XmlPart,
  axis: 'rows' | 'columns',
  first: number,
  last: number
): OpRejection | undefined {
  const key = axis === 'rows' ? 'row' : 'col';
  let hit: AnchorHit | undefined;
  const drawing = sheetDrawing(pkg, sheet);
  // A drawing that cannot be read can hold an anchor that keeps its size or position anywhere.
  if (drawing && 'unread' in drawing)
    hit = {
      name: undefined,
      reason: `the drawing part ${drawing.unread} cannot be read`,
      unread: true,
    };
  else if (drawing) hit = resizedAnchor(drawing.part.root, key, first, last, 0);
  for (const c of sheet.root.children) {
    if (hit) break;
    if (c.type !== 'element') continue;
    if ((c.ns === NS.x && OBJECT_LISTS.has(c.local)) || c.ns === NS.mc)
      hit = resizedObject(c, key, first, last, 0, undefined);
  }
  if (!hit) return undefined;
  return {
    code: 'resize-anchor',
    ...(hit.name !== undefined ? { subject: hit.name } : {}),
    ...(hit.unread ? { causeUnreadable: true as const } : {}),
    detail: `${sheet.partName}: ${hit.reason}`,
  };
}
