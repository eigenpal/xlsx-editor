// The sheet ops of design D1 of the change `workbook-structure-editing`: add, rename, delete,
// and move a worksheet.
//
// - `addSheet` writes a worksheet part `sheetN.xml` with the smallest free `N` next to the
//   other worksheets, with only an empty `x:sheetData`, a workbook relationship with a new `Id`,
//   an `x:sheet` with `sheetId` one more than the largest, and a content type override.
//   `x:sheetData` is the one required child of `CT_Worksheet`, and `x:sheet/@sheetId` is a
//   required `xsd:unsignedInt` (`CT_Sheet`, sml.xsd).
// - `renameSheet` changes `x:sheet/@name`, and the sheet prefix of each reference to the sheet
//   in every formula (`x:f`, shared formula masters included), defined name, internal
//   hyperlink location, and hyperlink relationship target that is a fragment (`#Data!A1`,
//   `hyperlink-rels.ts`) (`sheet-refs.ts`). A formula or a defined name inside
//   `mc:AlternateContent` that the op would change refuses the op (`wrapped-formulas.ts`), in a
//   rename, a delete, or a move.
// - `deleteSheet` removes the part, its relationship part, its workbook relationship, its
//   override, its `x:sheet`, and the printer settings, the Drawings part with its relationship
//   part, and the images that no other part relates to. Each
//   reference to it becomes `#REF!`. A defined name scoped to it goes, and the other
//   `@localSheetId` values move down.
// - `moveSheet` moves the `x:sheet` element.
//
// An applied add, rename, or move marks the cached values stale when a formula calls a function
// that reads sheet names or positions (`sheet-functions.ts`); a delete always marks them.
//
// `@localSheetId` (`CT_DefinedName`) and `x:workbookView/@activeTab` (`CT_BookView`) follow their
// sheets in every op. An add, a delete, or a move that changes the position of the sheet that the
// `@localSheetId` of a defined name inside `mc:AlternateContent` names is refused
// (`wrapped-formulas.ts`). `@firstSheet` is a scroll position of the tab strip: it keeps its index,
// clamped after a delete or a move. An op is refused, with the subject named, when content that the
// op does not rewrite names the sheet (the blockers of design D4, `blockers.ts`), when a move
// changes the sheets that a sheet range of a formula or defined name spans, or a delete removes an
// end of one (`sheet-range-blocked`), when the sheet to delete has a relationship other than a
// hyperlink, its own printer settings, images, or a Drawings part that holds only pictures and
// shapes, when a custom workbook view shows the sheet to delete, and
// when a delete would leave no visible sheet (`last-sheet`). A chart or pivot part, a sheet that
// is not a worksheet, or a workbook relationship outside the allowlist refuses a rename and a
// delete of each sheet of the opened file, and of an added sheet only when it names that sheet
// (`sheetEditBlocker`).
//
// Undo: the inverse is one `restoreParts` op of every part that the op changed, added, or
// removed, as for the structural edits of design D4 (`structure.ts`).

import type { Element, Node, XmlPart } from '../package/tree.ts';
import { parseXmlPart } from '../package/tree.ts';
import {
  CONTENT_TYPES_PART,
  equivalentPartName,
  relationshipsOf,
  relsPartName,
  xmlPartOf,
  type PackagePart,
  type XlsxPackage,
} from '../package/package.ts';
import { asciiFold, normalizePartName } from '@eigenpal/ooxml-core';
import { CONTENT_TYPE, NS, REL } from '../package/namespaces.ts';
import { attr, child, intAttr, textOf } from '../package/query.ts';
import {
  sheetByName,
  sheetNameKey,
  sheetsByKey,
  workbookIndex,
  type SheetEntry,
  type WorkbookIndex,
} from '../workbook/workbook-index.ts';
import { workbookHasFormula } from '../store/calc.ts';
import {
  IdSource,
  makeElement,
  plainAttr,
  smlElement,
  textNode,
  withAttr,
  withChild,
} from '../store/tree-edit.ts';
import {
  refuse,
  type OpRejection,
  type OpResult,
  type StructureSubject,
  type WorkbookOp,
} from '../store/ops.ts';
import type { PartRestore } from './structure.ts';
import { blocked } from './sheet-cells.ts';
import {
  cellsOrNamesMention,
  longSheetNameBlocker,
  sheetBlocker,
  sheetEditBlocker,
  sheetMentions,
  workbookMentionBlocker,
  macroBlocker,
  type SheetMentions,
} from './blockers.ts';
import { rewriteSheetPrefixes, sheetNameProblem, type PrefixEdit } from './sheet-refs.ts';
import { workbookCallsSheetFunction } from './sheet-functions.ts';
import { moveHyperlinkTargets } from './hyperlink-rels.ts';
import {
  isAlternateContent,
  wrappedCellProblem,
  wrappedDefinedNames,
  wrappedTextProblem,
} from './wrapped-formulas.ts';

/** The names of the sheet ops. */
export type SheetOpName = 'addSheet' | 'renameSheet' | 'deleteSheet' | 'moveSheet';

/** The sheet ops. */
export type SheetOp = Extract<WorkbookOp, { readonly op: SheetOpName }>;

const isRefusal = (r: unknown): r is OpRejection =>
  typeof r === 'object' && r !== null && 'code' in r;

/** Whether a node is the element `local` in `ns`. */
const isEl = (n: Node, ns: string, local: string): boolean =>
  n.type === 'element' && n.ns === ns && n.local === local;

// ---------------------------------------------------------------- content

/**
 * The refusal of formula text that a sheet op cannot rewrite. A sheet range refusal keeps
 * `about` only for a wrapped formula, so the message says that the user cannot edit that
 * formula, and does not ask the user to change the reference.
 */
function textRefusal(
  subject: string,
  r: { reason: string; range?: true },
  about: StructureSubject
): OpRejection {
  if (!r.range) return blocked(subject, r.reason, about);
  const detail = `${subject}: ${r.reason}`;
  return about.kind === 'wrapped-formula'
    ? { code: 'sheet-range-blocked', subject, detail, about }
    : { code: 'sheet-range-blocked', subject, detail };
}

/** A worksheet with the sheet prefixes of its formulas and hyperlink locations rewritten. */
function rewriteWorksheet(part: XmlPart, name: string, edit: PrefixEdit): XmlPart | OpRejection {
  const ids = new IdSource(part);
  const kids = part.root.children.slice();
  let changed = false;
  for (let k = 0; k < kids.length; k += 1) {
    const node = kids[k]!;
    if (node.type !== 'element') continue;
    if (isEl(node, NS.x, 'sheetData')) {
      const rows = node.children.slice();
      let rowsChanged = false;
      for (let i = 0; i < rows.length; i += 1) {
        const row = rows[i]!;
        if (row.type !== 'element' || row.kind !== 'row') continue;
        let cells = row.children;
        for (let j = 0; j < cells.length; j += 1) {
          const cell = cells[j]!;
          if (cell.type !== 'element' || cell.kind !== 'cell') continue;
          const inside = wrappedCellProblem(cell, (t) => rewriteSheetPrefixes(t, edit));
          if (inside) {
            const at = `${name}!${attr(cell, 'r') ?? '?'}`;
            return textRefusal(at, inside, { kind: 'wrapped-formula', name: at });
          }
          const f = child(cell, 'x:f');
          const text = f ? textOf(f) : '';
          if (!f || text === '') continue;
          const r = rewriteSheetPrefixes(text, edit);
          if (!r.ok) {
            const at = `${name}!${attr(cell, 'r') ?? '?'}`;
            return textRefusal(at, r, { kind: 'cell', name: at });
          }
          if (!r.changed) continue;
          const nf: Element = { ...f, children: [textNode(ids, r.text)] };
          if (cells === row.children) cells = cells.slice();
          (cells as Node[])[j] = withChild(cell, cell.children.indexOf(f), nf);
        }
        if (cells !== row.children) {
          rows[i] = { ...row, children: cells };
          rowsChanged = true;
        }
      }
      if (rowsChanged) {
        kids[k] = { ...node, children: rows };
        changed = true;
      }
    } else if (isEl(node, NS.x, 'hyperlinks')) {
      const links = node.children.slice();
      let linksChanged = false;
      for (let i = 0; i < links.length; i += 1) {
        const el = links[i]!;
        const location = el.type === 'element' ? attr(el, 'location') : undefined;
        if (el.type !== 'element' || !location) continue;
        const r = rewriteSheetPrefixes(location, edit);
        if (!r.ok)
          return textRefusal(
            `${name}!x:hyperlink`,
            { ...r, reason: `the location ${r.reason}` },
            { kind: 'hyperlink', name }
          );
        if (!r.changed) continue;
        links[i] = withAttr(el, 'location', r.text);
        linksChanged = true;
      }
      if (linksChanged) {
        kids[k] = { ...node, children: links };
        changed = true;
      }
    }
  }
  if (!changed) return part;
  return { partName: part.partName, root: { ...part.root, children: kids }, nextId: ids.next };
}

/**
 * Every worksheet but `skip`, rewritten, by part name, with the relationship parts whose
 * hyperlink targets changed (`moveHyperlinkTargets`), or the refusal.
 */
function rewriteSheets(
  pkg: XlsxPackage,
  index: WorkbookIndex,
  edit: PrefixEdit,
  skip?: string
): Map<string, XmlPart> | OpRejection {
  const out = new Map<string, XmlPart>();
  for (const sheet of index.sheets) {
    if (sheet.kind !== 'worksheet' || sheet.partName === skip) continue;
    const part = xmlPartOf(pkg, sheet.partName);
    if (!part) continue;
    const r = rewriteWorksheet(part, sheet.name, edit);
    if (isRefusal(r)) return r;
    const links = moveHyperlinkTargets(pkg, r, sheet.name, (t) => rewriteSheetPrefixes(t, edit));
    if (isRefusal(links)) return links;
    if (links.part !== part) out.set(part.partName, links.part);
    if (links.rels) out.set(links.rels.partName, links.rels);
  }
  return out;
}

/** The first element of a worksheet or the workbook that names the sheet and is not rewritten. */
function mentionBlocker(
  pkg: XlsxPackage,
  index: WorkbookIndex,
  needles: SheetMentions,
  skip?: string
): OpRejection | undefined {
  const wb = workbookMentionBlocker(pkg, index, needles, skip);
  if (wb) return wb;
  for (const sheet of index.sheets) {
    if (sheet.kind !== 'worksheet' || sheet.partName === skip) continue;
    const part = xmlPartOf(pkg, sheet.partName);
    const b = part ? sheetBlocker(pkg, part, sheet.name, false, needles) : undefined;
    if (b) return b;
  }
  return undefined;
}

/** Spellings that find any sheet range of the workbook (`A:B!`, `'A:B'!`). */
function sheetRangeMentions(index: WorkbookIndex): SheetMentions {
  const starts: string[] = [];
  const ends: string[] = [];
  for (const s of index.sheets) {
    const n = sheetNameKey(s.name);
    const q = n.replace(/'/g, "''");
    starts.push(`${n}:`, `'${q}:`);
    ends.push(`:${n}!`, `:${q}'!`);
  }
  return { direct: [], starts, ends };
}

// ---------------------------------------------------------------- workbook part

/** How an op changes the workbook part. */
interface WorkbookEdit {
  /** The `x:sheet` elements after the op, in tab order. */
  readonly sheets: readonly Element[];
  /** The position after the op of the sheet at an old position, or null when it is gone. */
  readonly position: (old: number) => number | null;
  /** The position for `@activeTab` when its sheet is gone. */
  readonly fallback?: number;
  /**
   * True when `@firstSheet` is clamped to the last position and to `@activeTab` (a delete or a
   * move). Otherwise the op keeps it.
   */
  readonly clampFirstSheet?: boolean;
  readonly prefixes?: PrefixEdit;
}

/**
 * `@activeTab` moved with its sheet. An absent attribute reads as its schema default, 0
 * (`CT_BookView`), and is written when its position changes.
 */
function movedActiveTab(el: Element, count: number, edit: WorkbookEdit): Element {
  const v = intAttr(el, 'activeTab') ?? 0;
  if (v < 0 || v >= count) return el;
  const p = edit.position(v) ?? edit.fallback ?? 0;
  return p === v ? el : withAttr(el, 'activeTab', String(p));
}

/**
 * A workbook view with its active tab moved with its sheet. `@firstSheet` is the first tab that
 * the tab strip shows, a scroll position and not a sheet: an absent one stays absent, and a
 * present one keeps its index, clamped after a delete or a move to the last position and to
 * the active tab.
 */
function movedView(view: Element, count: number, edit: WorkbookEdit): Element {
  const out = movedActiveTab(view, count, edit);
  const first = intAttr(out, 'firstSheet');
  if (!edit.clampFirstSheet || first === undefined || first < 0) return out;
  const last = edit.sheets.length - 1;
  const active = intAttr(out, 'activeTab') ?? 0;
  const limit = active >= 0 && active <= last ? active : last;
  const p = Math.min(first, limit);
  return p === first ? out : withAttr(out, 'firstSheet', String(p));
}

/** The workbook part after the op, or the refusal. */
function editWorkbook(part: XmlPart, count: number, edit: WorkbookEdit): XmlPart | OpRejection {
  const ids = new IdSource(part);
  const kids: Node[] = [];
  const hasViews = part.root.children.some((n) => isEl(n, NS.x, 'bookViews'));
  for (const node of part.root.children) {
    if (node.type !== 'element') {
      kids.push(node);
      continue;
    }
    // A defined name list inside a wrapper of the workbook element.
    const inside = wrappedNameRefusal(node, count, edit);
    if (inside) return inside;
    if (isEl(node, NS.x, 'sheets')) {
      if (!hasViews) {
        // No view: the active tab reads as 0, so a view goes in when it moves.
        const views = smlElement(ids, 'bookViews', part.root.kind);
        const view = movedView(smlElement(ids, 'workbookView', views.kind), count, edit);
        if (view.attrs.length > 0) kids.push({ ...views, children: [view] });
      }
      const other = node.children.filter((n) => !isEl(n, NS.x, 'sheet'));
      kids.push({ ...node, children: [...edit.sheets, ...other] });
    } else if (isEl(node, NS.x, 'bookViews')) {
      const views = node.children.map((v) => {
        if (v.type !== 'element' || !isEl(v, NS.x, 'workbookView')) return v;
        return movedView(v, count, edit);
      });
      kids.push({ ...node, children: views });
    } else if (isEl(node, NS.x, 'definedNames')) {
      const names: Node[] = [];
      let removed = 0;
      for (const el of node.children) {
        if (el.type !== 'element' || !isEl(el, NS.x, 'definedName')) {
          const inside = el.type === 'element' ? wrappedNameRefusal(el, count, edit) : undefined;
          if (inside) return inside;
          names.push(el);
          continue;
        }
        let out = el;
        const local = intAttr(el, 'localSheetId');
        if (local !== undefined && local >= 0 && local < count) {
          const p = edit.position(local);
          if (p === null) {
            removed += 1;
            continue;
          }
          if (p !== local) out = withAttr(out, 'localSheetId', String(p));
        }
        const text = textOf(el);
        if (edit.prefixes && text !== '') {
          const r = rewriteSheetPrefixes(text, edit.prefixes);
          if (!r.ok)
            return textRefusal(`x:definedName ${attr(el, 'name') ?? ''}`, r, {
              kind: 'defined-name',
              name: attr(el, 'name') ?? '',
            });
          if (r.changed) out = { ...out, children: [textNode(ids, r.text)] };
        }
        names.push(out);
      }
      // A name list that the op emptied goes with its last name.
      if (removed === 0 || names.length > 0) kids.push({ ...node, children: names });
    } else kids.push(node);
  }
  return { partName: part.partName, root: { ...part.root, children: kids }, nextId: ids.next };
}

/**
 * Why the op is refused for a defined name inside `mc:AlternateContent`, in the name list or in
 * the workbook element, or undefined (`wrapped-formulas.ts`): the op would move the sheet that
 * its `@localSheetId` names, or change its text. Any other node gives undefined.
 */
function wrappedNameRefusal(el: Node, count: number, edit: WorkbookEdit): OpRejection | undefined {
  if (!isAlternateContent(el)) return undefined;
  for (const name of wrappedDefinedNames(el)) {
    const label = attr(name, 'name') ?? '';
    const subject = `x:definedName ${label}`;
    const about: StructureSubject = { kind: 'defined-name', name: label };
    const local = intAttr(name, 'localSheetId');
    if (local !== undefined && local >= 0 && local < count && edit.position(local) !== local)
      return blocked(
        subject,
        'the edit moves the sheet of a defined name inside mc:AlternateContent, and the editor does not rewrite each branch',
        about
      );
    const prefixes = edit.prefixes;
    const problem = prefixes
      ? wrappedTextProblem(textOf(name), (t) => rewriteSheetPrefixes(t, prefixes), 'a defined name')
      : undefined;
    if (problem) return textRefusal(subject, problem, about);
  }
  return undefined;
}

/** The `x:sheet` elements of the workbook part, in tab order. */
function sheetElements(main: XmlPart): Element[] | undefined {
  const sheets = child(main.root, 'x:sheets');
  return sheets?.children.filter(
    (n): n is Element => n.type === 'element' && isEl(n, NS.x, 'sheet')
  );
}

// ---------------------------------------------------------------- result

/** The op's result: the parts it changed, and the `restoreParts` op that puts them back. */
function result(
  pkg: XlsxPackage,
  changes: ReadonlyMap<string, PackagePart | null>,
  flags: { value: boolean; formula: boolean }
): OpResult {
  const parts = new Map(pkg.parts);
  const restore: PartRestore[] = [];
  for (const [name, part] of changes) {
    const old = pkg.parts.get(name) ?? null;
    if (old === part) continue;
    restore.push({ partName: name, part: old });
    if (part) parts.set(name, part);
    else parts.delete(name);
  }
  if (restore.length === 0) return { ok: true, pkg, inverse: [], value: false, formula: false };
  const next: XlsxPackage = { parts };
  return {
    ok: true,
    pkg: next,
    inverse: [{ op: 'restoreParts', parts: restore }],
    value: flags.value,
    formula: flags.formula && (workbookHasFormula(next) || workbookHasFormula(pkg)),
  };
}

const xml = (part: XmlPart): PackagePart => ({ kind: 'xml', partName: part.partName, xml: part });

/** The common context of an op: the index, the workbook part, and its sheet elements. */
function context(
  pkg: XlsxPackage
): { index: WorkbookIndex; main: XmlPart; sheets: Element[] } | OpRejection {
  const index = workbookIndex(pkg);
  const main = xmlPartOf(pkg, index?.mainPart);
  const sheets = main ? sheetElements(main) : undefined;
  if (!index || !main || !sheets || sheets.length !== index.sheets.length)
    return { code: 'unsupported-package', detail: 'the workbook part has no sheet list' };
  return { index, main, sheets };
}

function target(index: WorkbookIndex, name: string): SheetEntry | OpRejection {
  const entry = typeof name === 'string' ? sheetByName(index, name) : undefined;
  if (entry) return entry;
  const many = typeof name === 'string' && sheetsByKey(index, name).length > 1;
  const detail = many
    ? `${JSON.stringify(name)} matches more than one sheet name without case`
    : `no sheet named ${JSON.stringify(name)}`;
  return { code: 'unknown-sheet', detail };
}

/**
 * Why a rename or a delete cannot rewrite the sheet prefixes of `entry`, or undefined: another
 * sheet, which only a file can hold, has the same name without case, so a prefix such as
 * `DATA!A1` can name either sheet (`sheet-name-ambiguous`, with the other sheet as the
 * subject).
 */
function twinBlocker(index: WorkbookIndex, entry: SheetEntry): OpRejection | undefined {
  const twin = sheetsByKey(index, entry.name).find((s) => s !== entry);
  return twin
    ? {
        code: 'sheet-name-ambiguous',
        subject: twin.name,
        detail: `the sheets ${JSON.stringify(entry.name)} and ${JSON.stringify(twin.name)} have the same name without case, so formula text cannot tell them apart`,
      }
    : undefined;
}

/** Whether the op changed a defined name of the workbook part. */
function namesChanged(before: XmlPart, after: XmlPart): boolean {
  const a = child(before.root, 'x:definedNames')?.children ?? [];
  const b = child(after.root, 'x:definedNames')?.children ?? [];
  return a.length !== b.length || a.some((n, i) => n !== b[i]);
}

// ---------------------------------------------------------------- ops

/**
 * The folded part names that a content type override or a relationship of the workbook part
 * names, whether the part exists or not.
 */
function namedPartKeys(pkg: XlsxPackage, types: XmlPart, mainPart: string): Set<string> {
  const out = new Set<string>();
  for (const n of types.root.children) {
    if (n.type !== 'element' || !isEl(n, NS.ct, 'Override')) continue;
    const norm = normalizePartName(attr(n, 'PartName') ?? '');
    if (norm.ok) out.add(asciiFold(norm.partName));
  }
  for (const rel of relationshipsOf(pkg, mainPart)) if (rel.target) out.add(asciiFold(rel.target));
  return out;
}

/** Why a sheet add or move is refused in a workbook that holds macros. */
const ORDER_REASON =
  'the macros of the workbook can read a sheet by its position, and the editor does not change them';

function addSheet(pkg: XlsxPackage, name: string, at: number): OpResult {
  const c = context(pkg);
  if (isRefusal(c)) return { ok: false, rejection: c };
  const { index, main, sheets } = c;
  if (!Number.isInteger(at) || at < 0 || at > sheets.length)
    return refuse('bad-position', `position ${at} is outside the sheet list`);
  const macros = macroBlocker(pkg, index, ORDER_REASON);
  if (macros) return { ok: false, rejection: macros };
  const problem = sheetNameProblem(name, index);
  if (problem) return { ok: false, rejection: problem };
  const relsName = equivalentPartName(pkg, relsPartName(main.partName));
  const rels = xmlPartOf(pkg, relsName);
  const types = xmlPartOf(pkg, CONTENT_TYPES_PART);
  if (!rels || !types)
    return refuse('unsupported-package', 'the workbook has no relationship or content type part');
  // The smallest free `sheetN.xml` next to the workbook part, with no relationship part, and
  // that no content type override or workbook relationship names (a dangling one included).
  const dir = main.partName.slice(0, main.partName.lastIndexOf('/'));
  const named = namedPartKeys(pkg, types, main.partName);
  let partName = '';
  const limit = pkg.parts.size + named.size + 1;
  for (let n = 1; n <= limit; n += 1) {
    const candidate = `${dir}/worksheets/sheet${n}.xml`;
    if (
      equivalentPartName(pkg, candidate) ||
      equivalentPartName(pkg, relsPartName(candidate)) ||
      named.has(asciiFold(candidate))
    )
      continue;
    partName = candidate;
    break;
  }
  const relIds = new Set<string>();
  for (const n of rels.root.children)
    if (n.type === 'element' && isEl(n, NS.rel, 'Relationship')) relIds.add(attr(n, 'Id') ?? '');
  let relId = '';
  for (let n = 1; n <= relIds.size + 1 && !relId; n += 1)
    if (!relIds.has(`rId${n}`)) relId = `rId${n}`;
  const sheetId = index.sheets.reduce((m, s) => Math.max(m, s.sheetId ?? 0), 0) + 1;
  if (sheetId > 0xffffffff) return refuse('bad-value', 'no sheetId is free');
  const created = parseXmlPart(partName, `<worksheet xmlns="${NS.x}"><sheetData/></worksheet>`);
  if (!created.ok) return refuse('bad-value', 'the new worksheet part cannot be built');
  const ids = new IdSource(main);
  const sheetEl = smlElement(ids, 'sheet', 'sheets', [
    plainAttr('name', name),
    plainAttr('sheetId', String(sheetId)),
    { ns: NS.r, local: 'id', prefix: 'r', value: relId },
  ]);
  const list = sheets.slice();
  list.splice(at, 0, sheetEl);
  const wb = editWorkbook({ ...main, nextId: ids.next }, sheets.length, {
    sheets: list,
    position: (old) => (old >= at ? old + 1 : old),
  });
  if (isRefusal(wb)) return { ok: false, rejection: wb };
  const relIdsSource = new IdSource(rels);
  const rel = makeElement(relIdsSource, NS.rel, 'Relationship', rels.root.kind, [
    plainAttr('Id', relId),
    plainAttr('Type', REL.worksheet),
    plainAttr('Target', partName.slice(dir.length + 1)),
  ]);
  const ctIds = new IdSource(types);
  const override = makeElement(ctIds, NS.ct, 'Override', types.root.kind, [
    plainAttr('PartName', partName),
    plainAttr('ContentType', CONTENT_TYPE.worksheet),
  ]);
  const changes = new Map<string, PackagePart | null>([
    [wb.partName, xml(wb)],
    [partName, xml(created.part)],
    [
      rels.partName,
      xml({
        ...rels,
        root: { ...rels.root, children: [...rels.root.children, rel] },
        nextId: relIdsSource.next,
      }),
    ],
    [
      types.partName,
      xml({
        ...types,
        root: { ...types.root, children: [...types.root.children, override] },
        nextId: ctIds.next,
      }),
    ],
  ]);
  // A formula that already names the new sheet reads it now: its cached result is stale.
  const stale =
    cellsOrNamesMention(pkg, index, sheetMentions(name)) || workbookCallsSheetFunction(pkg, index);
  return result(pkg, changes, { value: stale, formula: false });
}

function renameSheet(
  pkg: XlsxPackage,
  sheet: string,
  name: string,
  openedParts?: ReadonlySet<string>
): OpResult {
  const c = context(pkg);
  if (isRefusal(c)) return { ok: false, rejection: c };
  const { index, main, sheets } = c;
  const entry = target(index, sheet);
  if (isRefusal(entry)) return { ok: false, rejection: entry };
  const problem = sheetNameProblem(name, index, entry.index);
  if (problem) return { ok: false, rejection: problem };
  if (name === entry.name) return { ok: true, pkg, inverse: [], value: false, formula: false };
  const blocker =
    sheetEditBlocker(pkg, index, entry, openedParts) ??
    mentionBlocker(pkg, index, sheetMentions(entry.name));
  if (blocker) return { ok: false, rejection: blocker };
  // Another sheet with the same name without case refuses the rename only when a reference
  // names one of the two, so a rename of a twin that nothing names ends the ambiguity.
  const twin = twinBlocker(index, entry);
  const edit: PrefixEdit = {
    kind: 'rename',
    sheet: sheetNameKey(entry.name),
    name,
    sheets: new Set(index.sheets.map((s) => sheetNameKey(s.name))),
  };
  const rewritten = rewriteSheets(pkg, index, edit);
  if (isRefusal(rewritten)) return { ok: false, rejection: twin ?? rewritten };
  const list = sheets.slice();
  list[entry.index] = withAttr(list[entry.index]!, 'name', name);
  const wb = editWorkbook(main, sheets.length, {
    sheets: list,
    position: (old) => old,
    prefixes: edit,
  });
  if (isRefusal(wb)) return { ok: false, rejection: twin ?? wb };
  if (twin && (rewritten.size > 0 || namesChanged(main, wb))) return { ok: false, rejection: twin };
  const changes = new Map<string, PackagePart | null>([[wb.partName, xml(wb)]]);
  for (const part of rewritten.values()) changes.set(part.partName, xml(part));
  // A formula that already named the new name reads this sheet now: its cached result is stale.
  const stale =
    cellsOrNamesMention(pkg, index, sheetMentions(name)) || workbookCallsSheetFunction(pkg, index);
  return result(pkg, changes, { value: stale, formula: rewritten.size > 0 });
}

const PRINTER_SETTINGS =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings';

/** Whether a part other than the `owners` relates to the part `target`. */
function sharedPart(pkg: XlsxPackage, target: string, owners: ReadonlySet<string>): boolean {
  for (const name of pkg.parts.keys()) {
    if (owners.has(name) || name.endsWith('.rels')) continue;
    if (relationshipsOf(pkg, name).some((r) => r.target === target)) return true;
  }
  return false;
}

/** The relationship types of a drawing part whose targets go with the drawing, or stay. */
const DRAWING_PART_RELS = new Set<string>([REL.image, REL.hyperlink]);

/**
 * The parts that go with the Drawings part `drawing` of a deleted sheet `sheetPart`: the drawing
 * and the Image parts that only the sheet and the drawing relate to. A Drawings part is the
 * target of an explicit relationship from a worksheet (ECMA-376 Part 1, 12.3.8; root `xdr:wsDr`,
 * 20.5.2.35), and an Image part is the target of an explicit relationship (15.2.14), so neither
 * may stay without its source. A drawing that another part relates to stays. Undefined when the
 * drawing relates to a part that is not an image or a hyperlink, such as a chart, because such a
 * part can hold content that the editor cannot remove.
 */
function drawingRemovals(
  pkg: XlsxPackage,
  drawing: string,
  sheetPart: string
): string[] | undefined {
  if (sharedPart(pkg, drawing, new Set([sheetPart]))) return [];
  const owners = new Set([sheetPart, drawing]);
  const out = [drawing];
  for (const rel of relationshipsOf(pkg, drawing)) {
    if (rel.external) continue;
    if (!DRAWING_PART_RELS.has(rel.type)) return undefined;
    const t = rel.target;
    if (rel.type === REL.image && t && pkg.parts.has(t) && !sharedPart(pkg, t, owners))
      if (!out.includes(t)) out.push(t);
  }
  return out;
}

/** The visible sheet nearest to position `at` of a sheet list, or 0. */
function nearestVisible(list: readonly SheetEntry[], at: number): number {
  const ok = (s: SheetEntry | undefined) => s?.state === 'visible';
  const start = Math.min(at, list.length - 1);
  for (let d = 0; d < list.length; d += 1) {
    if (ok(list[start + d])) return start + d;
    if (ok(list[start - d])) return start - d;
  }
  return 0;
}

function deleteSheet(pkg: XlsxPackage, sheet: string, openedParts?: ReadonlySet<string>): OpResult {
  const c = context(pkg);
  if (isRefusal(c)) return { ok: false, rejection: c };
  const { index, main, sheets } = c;
  const entry = target(index, sheet);
  if (isRefusal(entry)) return { ok: false, rejection: entry };
  const at = entry.index;
  const rest = index.sheets.filter((s) => s.index !== at);
  if (!rest.some((s) => s.state === 'visible'))
    return {
      ok: false,
      rejection: {
        code: 'last-sheet',
        subject: entry.name,
        detail: 'no visible sheet would be left',
      },
    };
  const partName = entry.partName;
  if (!partName || entry.kind !== 'worksheet')
    return {
      ok: false,
      rejection: blocked(entry.name, 'the sheet is not a worksheet that the editor can remove', {
        kind: 'not-worksheet',
        name: entry.name,
      }),
    };
  const blocker =
    twinBlocker(index, entry) ??
    sheetEditBlocker(pkg, index, entry, openedParts) ??
    longSheetNameBlocker(index) ??
    mentionBlocker(
      pkg,
      index,
      sheetMentions(
        entry.name,
        index.sheets.map((s) => s.name)
      ),
      partName
    );
  if (blocker) return { ok: false, rejection: blocker };
  // A custom workbook view that shows the sheet keeps its `sheetId`, which no sheet would have.
  const views = child(main.root, 'x:customWorkbookViews');
  for (const v of views?.children ?? [])
    if (
      v.type === 'element' &&
      entry.sheetId !== undefined &&
      intAttr(v, 'activeSheetId') === entry.sheetId
    )
      return {
        ok: false,
        rejection: blocked(
          `x:customWorkbookView ${attr(v, 'name') ?? ''}`.trim(),
          'the custom view shows the sheet, and the editor does not change the view',
          { kind: 'workbook-content' }
        ),
      };
  // Printer settings, drawings, and images that only this sheet uses go with it, and shared
  // ones stay. Any other related part refuses.
  const removed: string[] = [];
  const add = (name: string) => {
    if (!removed.includes(name)) removed.push(name);
  };
  for (const rel of relationshipsOf(pkg, partName)) {
    if (rel.type === REL.hyperlink) continue;
    const t = rel.target;
    if (rel.type === PRINTER_SETTINGS || rel.type === REL.image) {
      if (t && pkg.parts.has(t) && !sharedPart(pkg, t, new Set([partName]))) add(t);
      continue;
    }
    const drawn = rel.type === REL.drawing && t && pkg.parts.has(t);
    const parts = drawn ? drawingRemovals(pkg, t, partName) : undefined;
    if (parts) {
      parts.forEach(add);
      continue;
    }
    return {
      ok: false,
      rejection: blocked(
        rel.target ?? partName,
        'the sheet has a related part that the editor cannot remove',
        { kind: 'sheet-content', name: entry.name }
      ),
    };
  }
  const edit: PrefixEdit = { kind: 'delete', sheet: sheetNameKey(entry.name) };
  const rewritten = rewriteSheets(pkg, index, edit, partName);
  if (isRefusal(rewritten)) return { ok: false, rejection: rewritten };
  const list = sheets.filter((_, i) => i !== at);
  const wb = editWorkbook(main, sheets.length, {
    sheets: list,
    position: (old) => (old === at ? null : old > at ? old - 1 : old),
    fallback: nearestVisible(rest, at),
    clampFirstSheet: true,
    prefixes: edit,
  });
  if (isRefusal(wb)) return { ok: false, rejection: wb };
  const changes = new Map<string, PackagePart | null>([[wb.partName, xml(wb)]]);
  for (const part of rewritten.values()) changes.set(part.partName, xml(part));
  for (const name of [partName, ...removed]) {
    changes.set(name, null);
    const ownRels = equivalentPartName(pkg, relsPartName(name));
    if (ownRels) changes.set(ownRels, null);
  }
  // The workbook relationship of the sheet, and its content type override.
  const relsName = equivalentPartName(pkg, relsPartName(main.partName));
  const rels = xmlPartOf(pkg, relsName);
  if (rels && entry.relId !== undefined) {
    const kept = rels.root.children.filter(
      (n) =>
        !(n.type === 'element' && isEl(n, NS.rel, 'Relationship') && attr(n, 'Id') === entry.relId)
    );
    changes.set(rels.partName, xml({ ...rels, root: { ...rels.root, children: kept } }));
  }
  const types = xmlPartOf(pkg, CONTENT_TYPES_PART);
  if (types) {
    const keys = new Set([partName, ...removed].map(asciiFold));
    const kept = types.root.children.filter((n) => {
      if (n.type !== 'element' || !isEl(n, NS.ct, 'Override')) return true;
      const norm = normalizePartName(attr(n, 'PartName') ?? '');
      return !(norm.ok && keys.has(asciiFold(norm.partName)));
    });
    if (kept.length !== types.root.children.length)
      changes.set(types.partName, xml({ ...types, root: { ...types.root, children: kept } }));
  }
  return result(pkg, changes, { value: true, formula: true });
}

function moveSheet(pkg: XlsxPackage, sheet: string, to: number): OpResult {
  const c = context(pkg);
  if (isRefusal(c)) return { ok: false, rejection: c };
  const { index, main, sheets } = c;
  const entry = target(index, sheet);
  if (isRefusal(entry)) return { ok: false, rejection: entry };
  if (!Number.isInteger(to) || to < 0 || to >= sheets.length)
    return refuse('bad-position', `position ${to} is outside the sheet list`);
  const from = entry.index;
  // A move to the same place changes nothing: no revision and no undo step.
  if (to === from) return { ok: true, pkg, inverse: [], value: false, formula: false };
  const macros = macroBlocker(pkg, index, ORDER_REASON);
  if (macros) return { ok: false, rejection: macros };
  // Content that the op does not read refuses any sheet range: its span is not known.
  const long = longSheetNameBlocker(index);
  if (long) return { ok: false, rejection: long };
  const blocker = mentionBlocker(pkg, index, sheetRangeMentions(index));
  if (blocker) return { ok: false, rejection: { ...blocker, code: 'sheet-range-blocked' } };
  const order = sheets.map((_, i) => i);
  order.splice(from, 1);
  order.splice(to, 0, from);
  const newPosition = new Map(order.map((old, i) => [old, i]));
  const edit: PrefixEdit = { kind: 'move', spanProblem: spanCheck(index, newPosition) };
  const checked = rewriteSheets(pkg, index, edit);
  if (isRefusal(checked)) return { ok: false, rejection: checked };
  const wb = editWorkbook(main, sheets.length, {
    sheets: order.map((i) => sheets[i]!),
    position: (old) => newPosition.get(old) ?? old,
    clampFirstSheet: true,
    prefixes: edit,
  });
  if (isRefusal(wb)) return { ok: false, rejection: wb };
  // A call such as `SHEET()` or `INDIRECT("Sheet1!A1")` can read the new order.
  const stale = workbookCallsSheetFunction(pkg, index);
  return result(pkg, new Map([[wb.partName, xml(wb)]]), { value: stale, formula: false });
}

/**
 * The check of a move for a sheet range: why the set of sheets between its ends changes, or
 * why its ends change order in the tab list, or undefined. A range whose end names no sheet
 * spans nothing before and after. A range whose end names two sheets without case
 * (`twinBlocker`) is refused, because its span is not known.
 */
function spanCheck(
  index: WorkbookIndex,
  newPosition: ReadonlyMap<number, number>
): (first: string, last: string) => string | undefined {
  const at = (key: string): number | null | undefined => {
    const list = sheetsByKey(index, key);
    return list.length > 1 ? null : list[0]?.index;
  };
  const seen = new Map<string, string | undefined>();
  return (first, last) => {
    const key = `${first}\u0000${last}`;
    if (seen.has(key)) return seen.get(key);
    const a = at(first);
    const b = at(last);
    let reason: string | undefined;
    if (a === null || b === null)
      reason = 'an end of the sheet range names more than one sheet without case';
    else if (a !== undefined && b !== undefined) {
      const before = (p: number) => p >= Math.min(a, b) && p <= Math.max(a, b);
      const na = newPosition.get(a)!;
      const nb = newPosition.get(b)!;
      const after = (p: number) => p >= Math.min(na, nb) && p <= Math.max(na, nb);
      // A swap of the ends in tab order keeps the span, but no evidence shows that a workbook
      // with the swapped text opens, so it is refused too.
      if (a !== b && a < b !== na < nb)
        reason = `the move changes which end of the sheet range ${first}:${last} comes first`;
      else
        for (const s of index.sheets)
          if (before(s.index) !== after(newPosition.get(s.index)!)) {
            reason = `the move changes the sheets that the sheet range ${first}:${last} spans`;
            break;
          }
    }
    seen.set(key, reason);
    return reason;
  };
}

/**
 * Apply one sheet op. `openedParts` holds the part names of the package that the store opened
 * (`partNamesOf`): a rename or a delete of a sheet that an add op of the session made is not
 * refused for the whole workbook (`sheetEditBlocker`). Without it, every sheet counts as a sheet
 * of the file.
 */
export function applySheetOp(
  pkg: XlsxPackage,
  op: SheetOp,
  openedParts?: ReadonlySet<string>
): OpResult {
  switch (op.op) {
    case 'addSheet':
      return addSheet(pkg, op.name, op.index);
    case 'renameSheet':
      return renameSheet(pkg, op.sheet, op.name, openedParts);
    case 'deleteSheet':
      return deleteSheet(pkg, op.sheet, openedParts);
    case 'moveSheet':
      return moveSheet(pkg, op.sheet, op.index);
  }
}
