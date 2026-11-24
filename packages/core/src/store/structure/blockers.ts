// What refuses a structural edit before anything moves (design D4): every element or part that
// can hold a reference, and that this change does not move, refuses the edit when the edit can
// change a position it refers to. The refusal names the element and its part, and its kind of
// content (`about`) for user text. It never locks
// other editing: values, clearing, and formats stay available.
//
// - The edited sheet: every worksheet child outside an allowlist of elements that hold no cell
//   reference, or that the structural edit moves (`sheet-cells.ts`, `sheet-elements.ts`). So
//   conditional formatting, data validation, `x:autoFilter`, `x:sortState`,
//   `x:protectedRanges`, `x:ignoredErrors`, page breaks, `x:cellWatches`, scenarios, custom
//   sheet views, tables, comments with their legacy drawing, embedded objects and controls,
//   `x:extLst`, and any element that this list does not name refuse the edit. So does a
//   relationship of the sheet to a part other than a hyperlink, a drawing, an image, a
//   printer setting, or a header and footer drawing.
// - Every other sheet: an element outside the same allowlist refuses the edit when its text or
//   attributes name the edited sheet (`Sheet1!`, `'My sheet'!`), because it can refer to it.
//   So does a child of `x:sheetData` or of a row that is not a typed row or cell (a row or a
//   cell inside `mc:AlternateContent`), because the edit rewrites typed cells only.
// - The workbook: macros first (a macro project or a macro sheet, `macroBlocker`), then chart
//   and pivot parts (they can refer to any sheet), a sheet part that is not
//   a worksheet, a relationship of the workbook or of the package to a part outside an
//   allowlist (`WORKBOOK_RELS`, `PACKAGE_RELS`), for example connections, volatile
//   dependencies, or revisions, and a workbook element outside an allowlist that names the
//   edited sheet.
// - The drawing, VML drawing, and control property parts of every other sheet: a shape linked
//   to a cell (`xdr:sp/@textlink`), a control linked to a cell (`@fmlaLink`, `@fmlaRange`), or
//   a legacy control (`x:FmlaLink`, `x:FmlaRange`) can name the edited sheet. Such a part
//   refuses the edit when its text or attributes name it. On the edited sheet, the drawing
//   rules (`drawing.ts`) and the relationship rule refuse it.
// - Table and query table parts, by content type, on any sheet: a table on another sheet holds
//   formulas (`x:calculatedColumnFormula`, `x:totalsRowFormula`) that can name the edited
//   sheet. Such a part refuses the edit when its text or attributes name it. A table on the
//   edited sheet refuses it through the relationship rule of that sheet.

import { asciiFold } from '@eigenpal/ooxml-core';
import type { Element, XmlPart } from '../package/tree.ts';
import {
  contentTypeOf,
  equivalentPartName,
  relationshipsOf,
  relsPartName,
  xmlPartOf,
  type PackagePart,
  type XlsxPackage,
} from '../package/package.ts';
import { NS, REL } from '../package/namespaces.ts';
import { child } from '../package/query.ts';
import { sheetNameKey, type WorkbookIndex } from '../workbook/workbook-index.ts';
import type { OpRejection, StructureSubject, StructureSubjectKind } from '../store/ops.ts';
import { blocked } from './sheet-cells.ts';
import { MAX_SHEET_NAME } from './sheet-refs.ts';
import { isAlternateContent } from './wrapped-formulas.ts';

/**
 * Worksheet children that hold no cell reference, or that a structural edit moves. Each name is
 * a child of the `CT_Worksheet` sequence (sml.xsd).
 */
const SHEET_SAFE = new Set([
  'sheetPr',
  'dimension',
  'sheetViews',
  'sheetFormatPr',
  'cols',
  'sheetData',
  'sheetCalcPr',
  'sheetProtection',
  'mergeCells',
  'hyperlinks',
  'printOptions',
  'pageMargins',
  'pageSetup',
  'headerFooter',
  'drawing',
  'legacyDrawingHF',
  'picture',
  'phoneticPr',
]);

/**
 * Workbook children that hold no cell reference, or that a structural edit moves. Each name is
 * a child of the `CT_Workbook` sequence (sml.xsd).
 */
const WORKBOOK_SAFE = new Set([
  'fileVersion',
  'fileSharing',
  'workbookPr',
  'workbookProtection',
  'bookViews',
  'sheets',
  'functionGroups',
  'externalReferences',
  'definedNames',
  'calcPr',
  'oleSize',
  'customWorkbookViews',
  'smartTagPr',
  'smartTagTypes',
  'fileRecoveryPr',
]);

/** Sheet relationships whose targets hold no cell reference. */
const SHEET_RELS = new Set<string>([
  REL.hyperlink,
  REL.drawing,
  REL.image,
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings',
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing',
]);

/**
 * Workbook relationships whose targets hold no reference to a cell of a worksheet, or that a
 * structural edit moves or drops: the worksheets, the shared string table, the styles, the
 * theme, the external links (each holds a cache of another workbook), the calculation chain
 * (dropped by design D5), custom XML data, and cell metadata.
 */
const WORKBOOK_RELS = new Set<string>([
  REL.worksheet,
  REL.sharedStrings,
  REL.styles,
  REL.theme,
  REL.calcChain,
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink',
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml',
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata',
]);

/**
 * The relationship from a macro-enabled workbook to its macro part. ECMA-376 does not define it.
 * The macro code can name any sheet or cell, and the editor never reads or rewrites it, so the
 * part refuses every edit that moves or renames a reference, with the subject kind `macros`.
 */
const MACRO_REL = 'http://schemas.microsoft.com/office/2006/relationships/vbaproject';
/** The relationship types of a macro sheet, folded: the sheet and its international form. */
const MACRO_SHEET_RELS: ReadonlySet<string> = new Set(
  ['xlmacrosheet', 'xlintlmacrosheet'].map(
    (t) => `http://schemas.microsoft.com/office/2006/relationships/${t}`
  )
);

/**
 * Package relationships (ECMA-376 Part 2) whose targets hold no cell reference: the properties
 * and thumbnail.
 */
const PACKAGE_RELS = new Set<string>([
  REL.officeDocument,
  'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties',
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties',
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties',
  'http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail',
]);

/** Content types of parts that a sheet relates to and that can name any sheet in a formula. */
const FORMULA_PART_TYPES = new Set<string>([
  'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml',
  'application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml',
]);

/** The kind of user content of a worksheet child; any other child is `sheet-content`. */
const SHEET_KINDS: ReadonlyMap<string, StructureSubjectKind> = new Map([
  ['conditionalFormatting', 'conditional-formatting'],
  ['dataValidations', 'data-validation'],
  ['autoFilter', 'filter'],
  ['sortState', 'filter'],
  ['tableParts', 'table'],
  ['legacyDrawing', 'comment'],
]);

/**
 * The related parts of another sheet that can link a shape or a control to a cell of the
 * edited sheet, with the kind of user content that the refusal names.
 */
const LINKING_RELS: ReadonlyMap<string, StructureSubjectKind> = new Map([
  [REL.drawing, 'linked-drawing'],
  ['http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing', 'control'],
  ['http://schemas.openxmlformats.org/officeDocument/2006/relationships/ctrlProp', 'control'],
]);

/** The kind of user content of a related part, by relationship type. */
const REL_KINDS: ReadonlyMap<string, StructureSubjectKind> = new Map([
  ['http://schemas.openxmlformats.org/officeDocument/2006/relationships/table', 'table'],
  ['http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments', 'comment'],
  ['http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable', 'pivot-table'],
]);

/** The deepest element that the mention scan reads. */
const MAX_SCAN_DEPTH = 64;

/** The name of an element as a refusal shows it: `x:conditionalFormatting`. */
export function elementLabel(el: Element): string {
  const prefix = el.ns === NS.x ? 'x' : (el.prefix ?? '?');
  return `${prefix}:${el.local}`;
}

/**
 * The spellings that can name a sheet in formula text, folded as `sheetNameKey` folds a sheet name. `direct` names it
 * (`Data!`, `'My sheet'!`, `Data:Last!`). A sheet range that only spans it (`First:Last!`,
 * with the sheet between the two in tab order) is a `starts` spelling of a sheet before it with
 * an `ends` spelling of a sheet after it, in the same text.
 */
export interface SheetMentions {
  readonly direct: readonly string[];
  readonly starts: readonly string[];
  readonly ends: readonly string[];
  /**
   * A sheet range that starts at the sheet, unquoted (`DATA:LAST!`). A pattern, not a spelling:
   * `A:` alone is a whole column (`A:A`), so it needs the rest of the prefix up to its `!`.
   */
  readonly rangeStart?: RegExp;
}

const folded = (name: string) => {
  const n = sheetNameKey(name);
  return { n, q: n.replace(/'/g, "''") };
};

/** The spellings that name the sheet `name`, of the sheets in tab order. */
export function sheetMentions(name: string, sheets: readonly string[] = []): SheetMentions {
  const { n, q } = folded(name);
  const at = sheets.findIndex((s) => sheetNameKey(s) === n);
  const starts: string[] = [];
  const ends: string[] = [];
  sheets.forEach((s, i) => {
    if (at < 0 || i === at) return;
    const f = folded(s);
    if (i < at) starts.push(`${f.n}:`, `'${f.q}:`);
    else ends.push(`:${f.n}!`, `:${f.q}'!`);
  });
  const literal = n.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  return {
    direct: [`${n}!`, `'${q}'!`, `'${q}:`, `:${q}'!`],
    starts,
    ends,
    rangeStart: new RegExp(`${literal}:[^\\s!:'"(),;=+\\-*/&^<>{}\\[\\]]+!`),
  };
}

/**
 * The work of the mention scan, for the cost tests (design D4, "cost"). `chars` counts each
 * folded text that the scan reads, once for each read; the direct spellings and the sheet
 * range pattern are a fixed number of searches of it. `steps` counts each character that a
 * walk of a sheet range trie reads. A scan that searches a text once for each sheet or each
 * spelling makes `chars` or `steps` grow with the sheets; this scan keeps both linear in the
 * text. Not document state: tests reset it and read it.
 */
export const mentionScanWork = { chars: 0, steps: 0 };

/**
 * A trie of spellings, one character for each step. A walk from a `:` reads at most as many
 * characters as the longest spelling holds, never one comparison for each spelling.
 */
interface Trie {
  readonly next: Map<string, Trie>;
  end: boolean;
}

function trieOf(list: readonly string[], reverse: boolean): Trie {
  const root: Trie = { next: new Map(), end: false };
  for (const word of list) {
    let node = root;
    for (let i = 0; i < word.length; i += 1) {
      const ch = word[reverse ? word.length - 1 - i : i]!;
      let step = node.next.get(ch);
      if (!step) node.next.set(ch, (step = { next: new Map(), end: false }));
      node = step;
    }
    node.end = true;
  }
  return root;
}

/**
 * Whether a spelling of the trie reads from position `p` of `u`, forward or backward. The walk
 * stops at the first character that no spelling holds.
 */
function walk(u: string, p: number, trie: Trie, step: 1 | -1): boolean {
  let node = trie;
  for (let i = p; i >= 0 && i < u.length; i += step) {
    mentionScanWork.steps += 1;
    const nextNode = node.next.get(u[i]!);
    if (!nextNode) return false;
    if (nextNode.end) return true;
    node = nextNode;
  }
  return false;
}

/**
 * The `starts` and `ends` spellings of a `SheetMentions` as tries. Each spelling holds one `:`:
 * a `starts` spelling ends with it, an `ends` spelling starts with it. So a text is read once:
 * at each `:`, one walk back and one walk forward, never one search for each sheet (design D4,
 * "cost"). A walk reads at most the longest spelling, which `longSheetNameBlocker` bounds.
 */
interface RangeSets {
  readonly starts: Trie;
  readonly ends: Trie;
  readonly nonAscii: boolean;
}

const rangeSetCache = new WeakMap<SheetMentions, RangeSets>();

function rangeSets(m: SheetMentions): RangeSets {
  let sets = rangeSetCache.get(m);
  if (sets) return sets;
  const any = (list: readonly string[]) => list.some((x) => /[^\x00-\x7f]/.test(x));
  sets = {
    starts: trieOf(m.starts, true),
    ends: trieOf(m.ends, false),
    nonAscii: any(m.direct) || any(m.starts) || any(m.ends),
  };
  rangeSetCache.set(m, sets);
  return sets;
}

/**
 * Whether the folded text holds a `starts` spelling and an `ends` spelling, each as a
 * substring, as `starts.some(includes) && ends.some(includes)` would find them.
 */
function rangeMentions(u: string, m: SheetMentions): boolean {
  if (m.starts.length === 0 || m.ends.length === 0) return false;
  const sets = rangeSets(m);
  let start = false;
  let end = false;
  for (let p = u.indexOf(':'); p >= 0 && !(start && end); p = u.indexOf(':', p + 1)) {
    if (!start) start = walk(u, p, sets.starts, -1);
    if (!end) end = walk(u, p, sets.ends, 1);
  }
  return start && end;
}

/**
 * Why an op that searches the sheet ranges of the workbook is refused, or undefined: a stored
 * sheet name holds more than `MAX_SHEET_NAME` characters. A file can store such a name, and the
 * search cost grows with the longest name, so the op is refused before the search. A rename
 * does not search the sheet ranges, so it stays available to shorten the name.
 */
export function longSheetNameBlocker(index: WorkbookIndex): OpRejection | undefined {
  const long = index.sheets.find((s) => s.name.length > MAX_SHEET_NAME);
  return long
    ? {
        code: 'sheet-name-too-long',
        subject: long.name,
        detail: `a sheet name holds more than ${MAX_SHEET_NAME} characters, so the editor does not search the references to sheet ranges`,
      }
    : undefined;
}

/** Whether a text names the sheet, directly or through a sheet range that spans it. */
function textMentions(text: string, m: SheetMentions): boolean {
  const u = sheetNameKey(text);
  mentionScanWork.chars += u.length;
  if (m.direct.some((x) => u.includes(x))) return true;
  if (m.rangeStart?.test(u)) return true;
  return rangeMentions(u, m);
}

const ENTITY: Readonly<Record<string, string>> = {
  amp: '&',
  lt: '<',
  gt: '>',
  quot: '"',
  apos: "'",
};

/** Whether a spelling of `needles` holds a character outside ASCII. */
const nonAscii = (needles: SheetMentions): boolean => rangeSets(needles).nonAscii;

/** Whether decoded XML text declares an encoding other than UTF-8 or UTF-16. */
function foreignEncoding(text: string): boolean {
  const decl = /^\uFEFF?\s*<\?xml\s[^>]{0,200}?encoding\s*=\s*["']([^"']{1,40})["']/.exec(text);
  if (!decl) return false;
  const name = decl[1]!.toLowerCase();
  return name !== 'utf-8' && name !== 'utf8' && name !== 'utf-16';
}

/**
 * Whether the text of a part that the reader keeps as bytes (a table, a query table, or a
 * legacy drawing) names the sheet. The bytes are decoded as XML text, and the predefined and
 * numeric character references are resolved, so `&apos;My sheet&apos;!` reads as
 * `'My sheet'!`. When the sheet name holds non-ASCII characters, bytes that do not decode
 * cleanly or that declare another encoding count as a mention, so the edit is refused.
 */
function bytesMention(bytes: Uint8Array, needles: SheetMentions): boolean {
  const utf16 =
    bytes[0] === 0xff && bytes[1] === 0xfe
      ? 'utf-16le'
      : bytes[0] === 0xfe && bytes[1] === 0xff
        ? 'utf-16be'
        : 'utf-8';
  let text: string;
  try {
    text = new TextDecoder(utf16, { fatal: true }).decode(bytes);
  } catch {
    // Bytes that are not valid in the encoding: a name with non-ASCII characters cannot be
    // found with certainty, so the part counts as a mention.
    if (nonAscii(needles)) return true;
    text = new TextDecoder(utf16).decode(bytes);
  }
  // A declared encoding other than UTF-8 or UTF-16 (a legacy VML part can declare one) spells
  // non-ASCII characters in other bytes, so such a name counts as a mention.
  if (nonAscii(needles) && foreignEncoding(text)) return true;
  const plain = text.replace(
    /&(#x[0-9a-fA-F]{1,6}|#[0-9]{1,7}|[a-z]{2,4});/g,
    (all, ref: string) => {
      if (ref[0] !== '#') return ENTITY[ref] ?? all;
      const code = ref[1] === 'x' ? parseInt(ref.slice(2), 16) : parseInt(ref.slice(1), 10);
      return code <= 0x10ffff ? String.fromCodePoint(code) : all;
    }
  );
  return textMentions(plain, needles);
}

/** Whether the text or the attribute values of an element name a sheet. Bounded by depth. */
function mentions(el: Element, needles: SheetMentions, depth = 0): boolean {
  if (depth > MAX_SCAN_DEPTH) return true;
  const hit = (s: string) => textMentions(s, needles);
  if (el.attrs.some((a) => hit(a.value))) return true;
  for (const c of el.children) {
    if (c.type === 'text' ? hit(c.value) : mentions(c, needles, depth + 1)) return true;
  }
  return false;
}

/**
 * Whether a cell of a worksheet or a defined name names the sheet of `needles`. A sheet add or
 * rename reads it before the op: a formula that names the new name had no sheet to read, so its
 * cached result is out of date after the op.
 */
export function cellsOrNamesMention(
  pkg: XlsxPackage,
  index: WorkbookIndex,
  needles: SheetMentions
): boolean {
  for (const sheet of index.sheets) {
    const data = child(xmlPartOf(pkg, sheet.partName)?.root, 'x:sheetData');
    if (data && mentions(data, needles)) return true;
  }
  const root = xmlPartOf(pkg, index.mainPart)?.root;
  const names = child(root, 'x:definedNames');
  if (names !== undefined && mentions(names, needles)) return true;
  // A name list in a wrapper of `x:workbook` (`wrapped-formulas.ts`) reads the sheets too.
  return (root?.children ?? []).some((n) => isAlternateContent(n) && mentions(n, needles));
}

/** Whether a part names the sheet: an XML tree, or bytes read as XML text. */
function partMentions(part: PackagePart, needles: SheetMentions): boolean {
  return part.kind === 'xml' ? mentions(part.xml.root, needles) : bytesMention(part.bytes, needles);
}

/**
 * The drawing, legacy drawing, and control property parts that the worksheets relate to, with
 * the kind of content and the sheet that the refusal names. The parts of the sheet `skip` (a
 * deleted sheet) are left out, because they go with it; a part that another sheet relates to
 * stays in through that sheet.
 */
function linkedParts(
  pkg: XlsxPackage,
  index: WorkbookIndex,
  skip?: string
): Map<string, StructureSubject> {
  const out = new Map<string, StructureSubject>();
  for (const sheet of index.sheets) {
    if (!sheet.partName || sheet.partName === skip) continue;
    for (const rel of relationshipsOf(pkg, sheet.partName)) {
      const kind = LINKING_RELS.get(rel.type);
      if (rel.target && kind) out.set(rel.target, { kind, name: sheet.name });
    }
  }
  return out;
}

/**
 * A child of `x:sheetData` or of a row that is not a typed row or cell (for example a row or a
 * cell inside `mc:AlternateContent`) and names the sheet. The sheet ops rewrite the formulas of
 * typed cells only, so such content refuses the op.
 */
function untypedCellMention(
  part: XmlPart,
  sheetName: string,
  needles: SheetMentions
): OpRejection | undefined {
  const data = part.root.children.find(
    (n): n is Element => n.type === 'element' && n.ns === NS.x && n.local === 'sheetData'
  );
  for (const row of data?.children ?? []) {
    if (row.type !== 'element') continue;
    const candidates = row.kind === 'row' ? row.children : [row];
    const untyped = candidates.find(
      (c): c is Element =>
        c.type === 'element' && c.kind !== 'row' && c.kind !== 'cell' && mentions(c, needles)
    );
    if (untyped)
      return blocked(
        `${elementLabel(untyped)} (${part.partName})`,
        'it holds cells that refer to the edited sheet, and the editor does not rewrite them',
        { kind: 'sheet-content', name: sheetName }
      );
  }
  return undefined;
}

/** The first element or part of a worksheet that refuses the edit, or undefined. */
export function sheetBlocker(
  pkg: XlsxPackage,
  part: XmlPart,
  sheetName: string,
  target: boolean,
  needles: SheetMentions
): OpRejection | undefined {
  for (const node of part.root.children) {
    if (node.type !== 'element') continue;
    if (node.ns === NS.x && SHEET_SAFE.has(node.local)) continue;
    if (!target && !mentions(node, needles)) continue;
    const kind = (node.ns === NS.x && SHEET_KINDS.get(node.local)) || 'sheet-content';
    return blocked(
      `${elementLabel(node)} (${part.partName})`,
      target
        ? 'it can refer to cells that the edit moves, and the editor does not move it'
        : 'it refers to the edited sheet, and the editor does not move it',
      { kind, name: sheetName }
    );
  }
  if (!target)
    return (
      untypedCellMention(part, sheetName, needles) ?? linkingBlocker(pkg, part, sheetName, needles)
    );
  for (const rel of relationshipsOf(pkg, part.partName)) {
    if (SHEET_RELS.has(rel.type)) continue;
    return blocked(
      rel.target ?? part.partName,
      'a related part can refer to cells that the edit moves',
      { kind: REL_KINDS.get(rel.type) ?? 'sheet-content', name: sheetName }
    );
  }
  return undefined;
}

/**
 * The first drawing, VML drawing, or control property part of another sheet that names the
 * edited sheet, or undefined. A part that the reader keeps as bytes is read as XML text.
 */
function linkingBlocker(
  pkg: XlsxPackage,
  part: XmlPart,
  sheetName: string,
  needles: SheetMentions
): OpRejection | undefined {
  for (const rel of relationshipsOf(pkg, part.partName)) {
    const kind = LINKING_RELS.get(rel.type);
    const linked = kind && rel.target ? pkg.parts.get(rel.target) : undefined;
    if (!kind || !linked) continue;
    const hit =
      linked.kind === 'xml'
        ? mentions(linked.xml.root, needles)
        : bytesMention(linked.bytes, needles);
    if (hit)
      return blocked(
        `${rel.target} (${part.partName})`,
        'it links a shape or a control to the edited sheet, and the editor does not move the link',
        { kind, name: sheetName }
      );
  }
  return undefined;
}

/** The first workbook element or part that refuses the edit, or undefined. */
export function workbookBlocker(
  pkg: XlsxPackage,
  index: WorkbookIndex,
  needles: SheetMentions
): OpRejection | undefined {
  return workbookPartBlocker(pkg, index) ?? workbookMentionBlocker(pkg, index, needles);
}

/**
 * The refusal of an edit in a workbook that holds macros, with `reason`, or undefined. Macro code
 * can name any cell or sheet, and can read a sheet by its position; the editor never rewrites it.
 * A workbook holds macros when it has a workbook relationship of the macro project type, or a
 * macro sheet: a sheet whose relationship has a macro sheet type, or whose part has a content
 * type of the `macrosheet` family. The schema defines `CT_Macrosheet` (sml.xsd), but no part
 * clause defines its content type or these relationship types, so they are matched by name.
 * ECMA-376 Part 2, 6.5.3.4 compares relationship types with case; the match here folds ASCII
 * case on purpose, so it refuses more, never less.
 */
export function macroBlocker(
  pkg: XlsxPackage,
  index: WorkbookIndex,
  reason: string
): OpRejection | undefined {
  const rels = relationshipsOf(pkg, index.mainPart);
  const rel = rels.find((r) => asciiFold(r.type) === MACRO_REL);
  if (rel) return blocked(rel.target ?? rel.type, reason, { kind: 'macros' });
  const sheetRels = new Set(
    rels.filter((r) => MACRO_SHEET_RELS.has(asciiFold(r.type))).map((r) => r.id)
  );
  const sheet = index.sheets.find(
    (s) =>
      (s.relId !== undefined && sheetRels.has(s.relId)) ||
      (s.partName && asciiFold(contentTypeOf(pkg, s.partName) ?? '').includes('macrosheet'))
  );
  return sheet ? blocked(sheet.name, reason, { kind: 'macros' }) : undefined;
}

/**
 * The first part that refuses every edit that moves or renames a reference, whatever it names:
 * a chart or pivot part, a sheet that is not a worksheet, or a relationship outside the
 * allowlists. Undefined when there is none.
 */
export function workbookPartBlocker(
  pkg: XlsxPackage,
  index: WorkbookIndex
): OpRejection | undefined {
  // Macros come first, so the reason names them whatever else the workbook holds.
  const macros = macroBlocker(
    pkg,
    index,
    'the macros of the workbook can refer to the cells and sheets that the edit changes, and the editor does not change them'
  );
  if (macros) return macros;
  // A chartsheet is named by its sheet name, as its tab shows it; another part by its name.
  const sheetNames = new Map<string, string>();
  for (const sheet of index.sheets) if (sheet.partName) sheetNames.set(sheet.partName, sheet.name);
  for (const name of pkg.parts.keys()) {
    const type = contentTypeOf(pkg, name) ?? '';
    if (type.includes('chart') || type.includes('pivot'))
      return blocked(sheetNames.get(name) ?? name, 'a chart or pivot part can refer to any sheet', {
        kind: type.includes('chart') ? 'chart' : 'pivot-table',
      });
  }
  for (const sheet of index.sheets)
    if (sheet.kind === 'other' && sheet.partName)
      return blocked(sheet.name, 'a sheet that is not a worksheet can refer to any sheet', {
        kind: 'other-sheet',
        name: sheet.name,
      });
  for (const [source, allowed] of [
    [index.mainPart, WORKBOOK_RELS],
    ['/', PACKAGE_RELS],
  ] as const)
    for (const rel of relationshipsOf(pkg, source)) {
      if (allowed.has(rel.type)) continue;
      return blocked(
        rel.target ?? rel.type,
        'a part of the workbook can refer to cells that the edit moves, and the editor does not move it',
        { kind: 'workbook-content' }
      );
    }
  return undefined;
}

/**
 * The first table or query table part, linked drawing part, or workbook element outside the
 * allowlist, whose text names the sheet of `needles`, or undefined. `skip` is the part of a
 * deleted sheet: its own drawings go with it, so they refuse nothing.
 */
export function workbookMentionBlocker(
  pkg: XlsxPackage,
  index: WorkbookIndex,
  needles: SheetMentions,
  skip?: string
): OpRejection | undefined {
  const reason = 'it refers to the edited sheet, and the editor does not move it';
  for (const [name, part] of pkg.parts) {
    if (!FORMULA_PART_TYPES.has(contentTypeOf(pkg, name) ?? '')) continue;
    if (partMentions(part, needles))
      return blocked(name, reason, {
        kind: contentTypeOf(pkg, name)!.includes('queryTable') ? 'workbook-content' : 'table',
      });
  }
  for (const [name, about] of linkedParts(pkg, index, skip))
    for (const scanned of [name, relsPartName(name)]) {
      const part = pkg.parts.get(equivalentPartName(pkg, scanned) ?? scanned);
      if (part && partMentions(part, needles)) return blocked(scanned, reason, about);
    }
  const root = xmlPartOf(pkg, index.mainPart)?.root;
  for (const node of root?.children ?? []) {
    if (node.type !== 'element') continue;
    if (node.ns === NS.x && WORKBOOK_SAFE.has(node.local)) continue;
    if (mentions(node, needles))
      return blocked(
        `${elementLabel(node)} (${index.mainPart})`,
        'it refers to the edited sheet, and the editor does not move it',
        { kind: 'workbook-content' }
      );
  }
  return undefined;
}
