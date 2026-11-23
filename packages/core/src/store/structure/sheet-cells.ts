// The cells of one worksheet under a structural edit (design D4): the rows and cells of the
// edited sheet move, and the formula text of every sheet moves its references.
//
// - Rows and cells: only the elements that exist are read (design D2); nothing iterates the
//   grid. A moved row and every cell of it get an explicit `@r`. Cells that the edit deletes go
//   with their rows or columns. An insertion that pushes a cell past the last row or column is
//   refused (`grid-full`); a row with no cells there is cut, as a column entry is (design D4).
// - Formulas: each `x:f` text moves its references (`references.ts`). Array formula ranges move
//   as a whole; an insertion that splits one, or a deletion that removes part of one, is refused.
//   A data table formula on the edited sheet is refused: its `@r1` and `@r2` inputs are
//   references that this change does not move.
// - Shared formulas: a child's formula is its master's text moved by the child's offset. When
//   the edit moves the master's references and a child's references differently, or deletes
//   the master, the child is written as its own formula (unshared). A master with no shared
//   child left becomes a normal formula. The master's `@ref` moves as a span. The store never
//   writes a new shared formula, and a master that stays shared never gains a sheet-qualified
//   reference (`gainsSheetPrefix`): a shared formula with one fails to open in the save
//   check (`scripts/fixtures/structure-case.ts`).
// - Formulas inside `mc:AlternateContent` in a cell (`wrapped-formulas.ts`): the edit is refused
//   when it changes the text of a wrapped copy, and when it changes the formula of any member of
//   a shared formula with a wrapped member, as each master copy gives it. On the edited sheet, it
//   is also refused when it moves, resizes, or deletes the range of a wrapped array or shared
//   formula master, when a wrapped data table formula is on the sheet, and when it moves or
//   deletes a cell of a shared formula that has a wrapped member. A cell whose wrapped formulas
//   the edit leaves as they are moves as it is, and a cell that the edit deletes goes with them.

import type { Element, Node, XmlPart } from '../package/tree.ts';
import { attr, child, textOf } from '../package/query.ts';
import { formatCellRef, parseCellRef } from '../cell-ref.ts';
import { sheetIndex, type CellEntry } from '../workbook/sheet-index.ts';
import { tokenizeFormula, type FormulaToken } from '../formula/tokenizer.ts';
import { formatTokens, moveFormula, normalizeFormula } from '../formula/format.ts';
import { IdSource, textNode, withAttr } from '../store/tree-edit.ts';
import type { OpRejection, StructureSubject } from '../store/ops.ts';
import { cutsPart, movePosition, moveSpan, pushedOut, splits } from './axis.ts';
import { moveReferences, moveTokens, type MovedText, type ReferenceContext } from './references.ts';
import { wrappedFormulas } from '../workbook/cell-formula.ts';
import { wrappedTextProblem } from './wrapped-formulas.ts';

/** A move of the references in formula text. */
type MoveText = (text: string) => MovedText;

/**
 * A refusal of a structural edit that names what holds the reference: as the file names it in
 * `subject` and `detail`, and as a kind and display name in `about`, for user text.
 */
export const blocked = (subject: string, detail: string, about: StructureSubject): OpRejection => ({
  code: 'structure-blocked',
  subject,
  detail: `${subject}: ${detail}`,
  about,
});

/** A refusal that a cell holds, such as the formula of `Data!B5`. */
const cellBlocked = (ref: string, detail: string): OpRejection =>
  blocked(ref, detail, { kind: 'cell', name: ref });

/** A refusal for a formula that a cell keeps inside `mc:AlternateContent`. */
const wrappedBlocked = (ref: string, detail: string): OpRejection =>
  blocked(ref, detail, { kind: 'wrapped-formula', name: ref });

/** A refusal of an insertion that would push content out of the grid, named as `blocked` is. */
export const gridFull = (subject: string, about: StructureSubject): OpRejection => ({
  code: 'grid-full',
  subject,
  about,
  detail: `${subject} would move past the last row or column of the sheet`,
});

/** A range of cells, 0-based and inclusive. */
export interface CellBox {
  readonly r0: number;
  readonly c0: number;
  readonly r1: number;
  readonly c1: number;
}

/** `A1` or `A1:B2` as a box, or undefined. */
export function parseBox(ref: string): CellBox | undefined {
  const colon = ref.indexOf(':');
  const a = parseCellRef(colon < 0 ? ref : ref.slice(0, colon));
  const b = colon < 0 ? a : parseCellRef(ref.slice(colon + 1));
  if (!a || !b) return undefined;
  return {
    r0: Math.min(a.row, b.row),
    c0: Math.min(a.col, b.col),
    r1: Math.max(a.row, b.row),
    c1: Math.max(a.col, b.col),
  };
}

/** The text of a box: one cell, or a range. */
export function boxRef(b: CellBox): string {
  const from = formatCellRef({ row: b.r0, col: b.c0 });
  return b.r0 === b.r1 && b.c0 === b.c1
    ? from
    : `${from}:${formatCellRef({ row: b.r1, col: b.c1 })}`;
}

/** The box moved along the edit's axis, or null when nothing of it is left. */
export function moveBox(ctx: ReferenceContext, b: CellBox): CellBox | null {
  const rows = ctx.edit.axis === 'rows';
  const span = moveSpan(ctx.edit, rows ? b.r0 : b.c0, rows ? b.r1 : b.c1);
  if (!span) return null;
  return rows ? { ...b, r0: span[0], r1: span[1] } : { ...b, c0: span[0], c1: span[1] };
}

/** The span of a box along the edit's axis. */
export const axisSpan = (ctx: ReferenceContext, b: CellBox): readonly [number, number] =>
  ctx.edit.axis === 'rows' ? [b.r0, b.r1] : [b.c0, b.c1];

/** A formula element with new text, and with attributes removed. */
function rewriteFormula(
  ids: IdSource,
  f: Element,
  text: string | undefined,
  drop: readonly string[] = [],
  set: readonly (readonly [string, string])[] = []
): Element {
  let out: Element = {
    ...f,
    attrs: f.attrs.filter((a) => !(a.ns === '' && drop.includes(a.local))),
  };
  for (const [name, value] of set) out = withAttr(out, name, value);
  if (text !== undefined) out = { ...out, children: [textNode(ids, text)] };
  return out;
}

/** The worksheet part after the edit, or the refusal. `target` is the edited sheet. */
export function moveSheetCells(
  part: XmlPart,
  name: string,
  ctx: ReferenceContext,
  target: boolean
): XmlPart | OpRejection {
  const index = sheetIndex(part);
  const sheetData = index.sheetData;
  if (!sheetData) return part;
  const e = ctx.edit;
  const rows = e.axis === 'rows';
  const ids = new IdSource(part);
  const at = (entry: { row: number; col: number }) => `${name}!${formatCellRef(entry)}`;
  /** The new position of a cell, or null when the edit deletes it or pushes it out. */
  const moved = (row: number, col: number): { row: number; col: number } | null => {
    if (!target) return { row, col };
    const q = movePosition(e, rows ? row : col);
    return q === null ? null : rows ? { row: q, col } : { row, col: q };
  };
  // New `x:f` elements by the cell element that holds them.
  const formulas = new Map<Element, Element>();
  const masters = new Map<string, CellEntry>();
  const shared = new Map<string, CellEntry[]>();
  const wrapped: CellEntry[] = [];
  for (const key of index.order) {
    const entry = index.cells.get(key)!;
    if (wrappedFormulas(entry.el).length > 0) wrapped.push(entry);
    const f = child(entry.el, 'x:f');
    if (!f) continue;
    const t = attr(f, 't');
    const si = attr(f, 'si');
    if (t === 'dataTable') {
      if (target) return cellBlocked(at(entry), 'a data table formula refers to input cells');
      continue;
    }
    if (t === 'shared' && si !== undefined) {
      if (attr(f, 'ref') !== undefined) {
        if (masters.has(si)) return cellBlocked(at(entry), `two masters of shared formula ${si}`);
        masters.set(si, entry);
      } else {
        let list = shared.get(si);
        if (!list) shared.set(si, (list = []));
        list.push(entry);
      }
      continue;
    }
    const text = textOf(f);
    let ref: string | undefined;
    const fRef = attr(f, 'ref');
    if (t === 'array' && fRef !== undefined && target) {
      const box = parseBox(fRef);
      if (!box) return cellBlocked(at(entry), `the array formula range ${fRef} cannot be read`);
      const [a, b] = axisSpan(ctx, box);
      if (splits(e, a, b) || cutsPart(e, a, b))
        return cellBlocked(`${name}!${fRef}`, 'the edit splits an array formula range');
      if (pushedOut(e, b))
        return gridFull(`${name}!${fRef}`, { kind: 'cell', name: `${name}!${fRef}` });
      const next = moveBox(ctx, box);
      if (next && boxRef(next) !== fRef) ref = boxRef(next);
    }
    if (text === '' && ref === undefined) continue;
    const m = text === '' ? { ok: true as const, text, changed: false } : moveReferences(text, ctx);
    if (!m.ok) return cellBlocked(at(entry), m.reason);
    if (m.changed || ref !== undefined)
      formulas.set(
        entry.el,
        rewriteFormula(ids, f, m.changed ? m.text : undefined, [], ref ? [['ref', ref]] : [])
      );
  }
  const w = wrappedBlocker(ctx, name, wrapped, masters, shared, moved, target);
  if (w) return w;
  for (const [si, master] of masters) {
    const r = moveShared(ids, ctx, name, master, shared.get(si) ?? [], moved, target);
    if ('code' in r) return r;
    for (const [cell, f] of r) formulas.set(cell, f);
  }
  if (!target && formulas.size === 0) return part;
  return rewriteRows(part, ids, sheetData, formulas, ctx, target, name);
}

/**
 * Why the edit is refused for a formula that a cell keeps inside `mc:AlternateContent`, or
 * undefined (`wrapped-formulas.ts`). `wrapped` lists the cells with such a formula; `masters`
 * and `shared` are the direct shared formula cells by `@si`.
 */
function wrappedBlocker(
  ctx: ReferenceContext,
  name: string,
  wrapped: readonly CellEntry[],
  masters: ReadonlyMap<string, CellEntry>,
  shared: ReadonlyMap<string, readonly CellEntry[]>,
  moved: (row: number, col: number) => { row: number; col: number } | null,
  target: boolean
): OpRejection | undefined {
  const at = (entry: CellEntry) => `${name}!${formatCellRef(entry)}`;
  const stays = (entry: CellEntry): boolean => {
    const to = moved(entry.row, entry.col);
    return to !== null && to.row === entry.row && to.col === entry.col;
  };
  const move: MoveText = (text) => moveReferences(text, ctx);
  // The cells of each shared formula with a wrapped member, and its wrapped master copies, by
  // `@si`. The loop reads each wrapped formula one time, so a cell with copies in many groups
  // costs no more than its copies.
  const groups = new Map<string, CellEntry[]>();
  const copies = new Map<string, MasterCopy[]>();
  for (const entry of wrapped) {
    // A cell that the edit deletes takes its formula with it: only its range and its shared
    // formula, which other cells read, are checked.
    const gone = target && moved(entry.row, entry.col) === null;
    for (const f of wrappedFormulas(entry.el)) {
      const t = attr(f, 't');
      const si = attr(f, 'si');
      const fRef = attr(f, 'ref');
      if (t === 'dataTable' && target)
        return wrappedBlocked(
          at(entry),
          'a data table formula inside mc:AlternateContent refers to input cells'
        );
      if (t === 'shared' && si !== undefined) {
        let list = groups.get(si);
        if (!list) groups.set(si, (list = []));
        // The cells come in order, so a cell with two such copies is the last one listed.
        if (list[list.length - 1] !== entry) list.push(entry);
        if (fRef !== undefined) {
          let own = copies.get(si);
          if (!own) copies.set(si, (own = []));
          own.push({ entry, text: textOf(f) });
        }
      }
      // An array or shared formula master: its range keeps its text only when it stays.
      if ((t === 'array' || t === 'shared') && fRef !== undefined && target) {
        const box = parseBox(fRef);
        const next = box ? moveBox(ctx, box) : null;
        if (!box || !next || boxRef(next) !== boxRef(box) || !stays(entry))
          return wrappedBlocked(
            at(entry),
            `the edit changes the range ${fRef} of a formula inside mc:AlternateContent`
          );
      }
      const problem = gone ? undefined : wrappedTextProblem(textOf(f), move);
      if (problem) return wrappedBlocked(at(entry), problem.reason);
    }
  }
  for (const [si, list] of groups) {
    const master = masters.get(si);
    const direct = shared.get(si) ?? [];
    const moves = !list.every(stays) || (master && !stays(master)) || !direct.every(stays);
    if (target && moves)
      return wrappedBlocked(
        at(list[0]!),
        `the edit moves a cell of shared formula ${si}, which has a member inside mc:AlternateContent`
      );
    const problem = sharedGroupProblem(si, master, direct, list, copies.get(si) ?? [], move);
    if (problem) return wrappedBlocked(at(problem.entry), problem.reason);
  }
  return undefined;
}

/** The most distinct master texts that one shared formula can have in its master cell. */
const MAX_MASTER_TEXTS = 4;

/** A wrapped master copy of a shared formula: the cell and the formula text. */
interface MasterCopy {
  readonly entry: CellEntry;
  readonly text: string;
}

/**
 * Why the edit cannot keep a shared formula with a member inside `mc:AlternateContent`, and the
 * cell that the reason is about, or undefined. Each member's formula is the master text moved by
 * the member's offset. The direct master and each wrapped master copy give the text, so every
 * member is checked against every distinct copy (`copies` lists the wrapped ones, in cell
 * order). A member whose text the edit would change, or
 * cannot read, refuses the edit: the store would write that member as its own formula, and it
 * does not rewrite a wrapped copy. A group with no master is left as it is, as a direct one is.
 * Master copies in two cells are refused, as two direct masters are. More than
 * `MAX_MASTER_TEXTS` distinct texts are refused too, so one group costs a fixed number of
 * passes over its members. `wrappedBlocker` reads each wrapped copy one time, and each cell is
 * a member of a group for each `@si` that it holds, so all groups together cost a fixed number
 * of passes over the formulas of the sheet, whatever the file holds.
 */
function sharedGroupProblem(
  si: string,
  master: CellEntry | undefined,
  direct: readonly CellEntry[],
  wrapped: readonly CellEntry[],
  copies: readonly MasterCopy[],
  move: MoveText
): { readonly entry: CellEntry; readonly reason: string } | undefined {
  let origin = master;
  const texts = new Set<string>();
  if (master) texts.add(textOf(child(master.el, 'x:f')!));
  for (const { entry, text } of copies) {
    if (origin && origin !== entry) return { entry, reason: `two masters of shared formula ${si}` };
    origin = entry;
    texts.add(text);
    if (texts.size > MAX_MASTER_TEXTS)
      return {
        entry,
        reason: `shared formula ${si} has more master copies than the editor reads`,
      };
  }
  if (!origin) return undefined;
  // The wrapped members come first, so the refusal names one when it can.
  const members = new Set<CellEntry>(wrapped);
  if (master) members.add(master);
  for (const entry of direct) members.add(entry);
  const group = `shared formula ${si}, which has a member inside mc:AlternateContent`;
  for (const text of texts) {
    const read = tokenizeFormula(text);
    if (!read.ok)
      return { entry: origin, reason: `the master of ${group} cannot be read (${read.problem})` };
    for (const entry of members) {
      const own = formatTokens(read.tokens, {
        rows: entry.row - origin.row,
        cols: entry.col - origin.col,
      });
      if (!own.ok) return { entry, reason: `the formula of ${group} moves out of the grid` };
      const r = move(own.text);
      if (!r.ok) return { entry, reason: `the formula of ${group} (${r.reason})` };
      if (r.changed) return { entry, reason: `the edit changes the formula of ${group}` };
    }
  }
  return undefined;
}

/** The number of sheet-qualified references and names in a token list. */
function qualifiedCount(tokens: readonly FormulaToken[]): number {
  let n = 0;
  for (const t of tokens) if ('sheet' in t && t.sheet) n += 1;
  return n;
}

/**
 * Whether formula text `after` holds more sheet-qualified references than `before`, or cannot be
 * read. A moved reference keeps its prefix as written, so a move never adds one; this check
 * keeps that true for every shared formula that the store writes.
 */
export function gainsSheetPrefix(before: string, after: string): boolean {
  const a = tokenizeFormula(before);
  const b = tokenizeFormula(after);
  if (!a.ok || !b.ok) return true;
  return qualifiedCount(b.tokens) > qualifiedCount(a.tokens);
}

/** The new `x:f` elements of one shared formula group, or the refusal. */
function moveShared(
  ids: IdSource,
  ctx: ReferenceContext,
  name: string,
  master: CellEntry,
  children: readonly CellEntry[],
  moved: (row: number, col: number) => { row: number; col: number } | null,
  target: boolean
): Map<Element, Element> | OpRejection {
  const out = new Map<Element, Element>();
  const at = (entry: CellEntry) => `${name}!${formatCellRef(entry)}`;
  const f = child(master.el, 'x:f')!;
  const text = textOf(f);
  const read = tokenizeFormula(text);
  if (!read.ok)
    return cellBlocked(at(master), `the shared formula text cannot be read (${read.problem})`);
  const m = moveTokens(text, read.tokens, ctx);
  if (!m.ok) return cellBlocked(at(master), m.reason);
  const masterNorm = normalizeFormula(m.text);
  if (!masterNorm.ok) return cellBlocked(at(master), 'the moved shared formula cannot be read');
  const masterTo = moved(master.row, master.col);
  let kept = 0;
  for (const entry of children) {
    const to = moved(entry.row, entry.col);
    if (!to) continue;
    const own = formatTokens(read.tokens, {
      rows: entry.row - master.row,
      cols: entry.col - master.col,
    });
    if (!own.ok) return cellBlocked(at(entry), 'the shared formula moves out of the grid');
    const ownMoved = moveReferences(own.text, ctx);
    if (!ownMoved.ok) return cellBlocked(at(entry), ownMoved.reason);
    if (masterTo) {
      const derived = moveFormula(masterNorm.text, {
        rows: to.row - masterTo.row,
        cols: to.col - masterTo.col,
      });
      const ownNorm = normalizeFormula(ownMoved.text);
      if (derived.ok && ownNorm.ok && derived.text === ownNorm.text) {
        kept += 1;
        continue;
      }
    }
    const cf = child(entry.el, 'x:f')!;
    out.set(entry.el, rewriteFormula(ids, cf, ownMoved.text, ['t', 'si', 'ref']));
  }
  if (!masterTo) return out;
  const ref = attr(f, 'ref')!;
  let newRef: string | undefined;
  // A master that had children and keeps none becomes a normal formula.
  const lone = kept === 0 && children.length > 0;
  if (target && !lone) {
    const box = parseBox(ref);
    const next = box ? moveBox(ctx, box) : null;
    if (!next) return cellBlocked(at(master), `the shared formula range ${ref} cannot be moved`);
    if (boxRef(next) !== ref) newRef = boxRef(next);
  }
  if (!lone && m.changed && gainsSheetPrefix(text, m.text))
    return cellBlocked(
      at(master),
      'the moved shared formula would gain a sheet-qualified reference'
    );
  if (lone) out.set(master.el, rewriteFormula(ids, f, m.text, ['t', 'si', 'ref']));
  else if (m.changed || newRef)
    out.set(
      master.el,
      rewriteFormula(ids, f, m.changed ? m.text : undefined, [], newRef ? [['ref', newRef]] : [])
    );
  return out;
}

/** The part with its rows and cells moved and its formulas replaced. */
function rewriteRows(
  part: XmlPart,
  ids: IdSource,
  sheetData: Element,
  formulas: ReadonlyMap<Element, Element>,
  ctx: ReferenceContext,
  target: boolean,
  name: string
): XmlPart | OpRejection {
  const index = sheetIndex(part);
  const e = ctx.edit;
  const rows = e.axis === 'rows';
  const rowOf = new Map<Element, number>();
  for (const [r, el] of index.rows) rowOf.set(el, r);
  const cellOf = new Map<Element, CellEntry>();
  for (const entry of index.cells.values()) cellOf.set(entry.el, entry);
  const out: Node[] = [];
  let changed = false;
  for (const node of sheetData.children) {
    const r = node.type === 'element' ? rowOf.get(node) : undefined;
    if (node.type !== 'element' || r === undefined) {
      out.push(node);
      continue;
    }
    const band = `${name}!${r + 1}:${r + 1}`;
    const hasCells = node.children.some((n) => n.type === 'element' && n.kind === 'cell');
    let nr = r;
    if (target && rows) {
      const q = movePosition(e, r);
      if (q === null) {
        // Deleted, or pushed past the last row: a row with cells refuses, a row with only
        // properties is cut, as a column entry is (design D4).
        if (e.kind === 'insert' && hasCells) return gridFull(band, { kind: 'row', name: band });
        changed = true;
        continue;
      }
      nr = q;
    }
    const kids: Node[] = [];
    let rowChanged = nr !== r;
    for (const c of node.children) {
      const entry = c.type === 'element' ? cellOf.get(c) : undefined;
      if (c.type !== 'element' || !entry) {
        if (target && c.type === 'element')
          return blocked(band, 'the row holds content that is not a cell', {
            kind: 'row',
            name: band,
          });
        kids.push(c);
        continue;
      }
      const ref = `${name}!${formatCellRef(entry)}`;
      // A formula is named as a formula (`cell`); any other content as the cell's content.
      const content: StructureSubject = {
        kind: c.children.some((n) => n.type === 'element' && n.local === 'f')
          ? 'cell'
          : 'cell-content',
        name: ref,
      };
      if (target && c.children.some((n) => n.type === 'element' && n.local === 'extLst'))
        return blocked(ref, 'the cell holds extension content', content);
      let nc = entry.col;
      if (target && !rows) {
        const q = movePosition(e, entry.col);
        if (q === null) {
          if (e.kind === 'insert') return gridFull(ref, content);
          rowChanged = true;
          continue;
        }
        nc = q;
      }
      const f = formulas.get(c);
      if (nr === r && nc === entry.col && !f) {
        kids.push(c);
        continue;
      }
      let cell: Element = withAttr(c, 'r', formatCellRef({ row: nr, col: nc }));
      if (f)
        cell = {
          ...cell,
          children: cell.children.map((n) => (n.type === 'element' && n.local === 'f' ? f : n)),
        };
      kids.push(cell);
      rowChanged = true;
    }
    if (!rowChanged) {
      out.push(node);
      continue;
    }
    changed = true;
    // Every cell of a changed row gets an explicit reference (the part checks read it).
    const withRefs = kids.map((k) => {
      const entry = k.type === 'element' ? cellOf.get(k) : undefined;
      return entry && attr(entry.el, 'r') === undefined
        ? withAttr(entry.el, 'r', formatCellRef({ row: nr, col: entry.col }))
        : k;
    });
    let row: Element = withAttr({ ...node, children: withRefs }, 'r', String(nr + 1));
    // `spans` is a hint that a column edit makes wrong; dropping it is always valid.
    if (target && !rows) row = withAttr(row, 'spans', undefined);
    out.push(row);
  }
  if (!changed) return part;
  const root = {
    ...part.root,
    children: part.root.children.map((n) =>
      n === sheetData ? { ...sheetData, children: out } : n
    ),
  };
  return { partName: part.partName, root, nextId: ids.next };
}
