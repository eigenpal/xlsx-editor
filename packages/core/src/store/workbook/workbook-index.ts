// The derived indexes of a workbook (design D2): sheet name to part, the shared string table,
// the style lists, and the workbook properties that the rest of the engine reads.
//
// Every index is built from the elements that exist, never from a count or a range that the
// file states (`x:sst/@count`, `x:dimension/@ref`). File-keyed lookups use `Map`, so a sheet
// named `__proto__` is an ordinary key. The indexes are caches of immutable parts: a part that
// an edit replaces gets a new index, and nothing here holds document state in a global.

import type { Element, XmlPart } from '../package/tree.ts';
import {
  mainWorkbookPart,
  relationshipsOf,
  relTarget,
  xmlPartOf,
  type XlsxPackage,
} from '../package/package.ts';
import { attr, boolAttr, child, children, intAttr } from '../package/query.ts';
import { NS, REL } from '../package/namespaces.ts';

/** The visibility of a sheet tab: `x:sheet/@state` (`ST_SheetState`; default `visible`). */
export type SheetState = 'visible' | 'hidden' | 'veryHidden';

/** One `x:sheet` of the workbook, in tab order. */
export interface SheetEntry {
  /** 0-based position in `x:sheets`: the tab order, and the `localSheetId` of defined names. */
  readonly index: number;
  readonly name: string;
  readonly sheetId: number | undefined;
  readonly state: SheetState;
  /** The relationship id (`r:id`). */
  readonly relId: string | undefined;
  /** The part the relationship names, when it is internal and present. */
  readonly partName: string | undefined;
  /** `worksheet` when the part is a modeled worksheet; `other` for a chartsheet or a gap. */
  readonly kind: 'worksheet' | 'other';
}

/** The workbook-level facts the engine reads. */
export interface WorkbookIndex {
  readonly mainPart: string;
  readonly sheets: readonly SheetEntry[];
  /** `x:workbookPr/@date1904` (`CT_WorkbookPr`, sml.xsd; default false). */
  readonly date1904: boolean;
  /**
   * `x:calcPr/@fullCalcOnLoad` (`CT_CalcPr`, sml.xsd; default false): every formula result of the
   * workbook is stale.
   */
  readonly fullCalcOnLoad: boolean;
  /**
   * `x:workbookView/@activeTab` of the first view (`CT_BookView`, sml.xsd; default 0), 0 when
   * absent or out of range.
   */
  readonly activeTab: number;
  readonly sharedStringsPart: string | undefined;
  readonly stylesPart: string | undefined;
  readonly themePart: string | undefined;
}

const workbookCache = new WeakMap<XlsxPackage, WorkbookIndex>();

function sheetState(raw: string | undefined): SheetState {
  return raw === 'hidden' || raw === 'veryHidden' ? raw : 'visible';
}

/**
 * The workbook index of a package. Cached per package, so repeated reads of one revision
 * cost one lookup. Undefined when the package has no workbook part.
 */
export function workbookIndex(pkg: XlsxPackage): WorkbookIndex | undefined {
  const mainPart = mainWorkbookPart(pkg);
  const part = xmlPartOf(pkg, mainPart);
  if (!mainPart || !part) return undefined;
  const hit = workbookCache.get(pkg);
  if (hit) return hit;
  const root = part.root;
  const sheets: SheetEntry[] = [];
  const sheetsEl = child(root, 'x:sheets');
  for (const el of children(sheetsEl, 'x:sheet')) {
    const relId = attr(el, 'r:id');
    const partName = relId ? relTarget(pkg, mainPart, relId) : undefined;
    const target = xmlPartOf(pkg, partName);
    sheets.push({
      index: sheets.length,
      name: attr(el, 'name') ?? '',
      sheetId: intAttr(el, 'sheetId'),
      state: sheetState(attr(el, 'state')),
      relId,
      partName: partName && pkg.parts.has(partName) ? partName : undefined,
      kind: target?.root.kind === 'worksheet' ? 'worksheet' : 'other',
    });
  }
  const rels = relationshipsOf(pkg, mainPart);
  const firstOf = (type: string): string | undefined => {
    const target = rels.find((r) => r.type === type && !r.external)?.target;
    return target && pkg.parts.has(target) ? target : undefined;
  };
  const view = child(child(root, 'x:bookViews'), 'x:workbookView');
  const activeTab = intAttr(view, 'activeTab') ?? 0;
  const index: WorkbookIndex = {
    mainPart,
    sheets,
    date1904: boolAttr(child(root, 'x:workbookPr'), 'date1904') ?? false,
    fullCalcOnLoad: boolAttr(child(root, 'x:calcPr'), 'fullCalcOnLoad') ?? false,
    activeTab: activeTab >= 0 && activeTab < sheets.length ? activeTab : 0,
    sharedStringsPart: firstOf(REL.sharedStrings),
    stylesPart: firstOf(REL.styles),
    themePart: firstOf(REL.theme),
  };
  workbookCache.set(pkg, index);
  return index;
}

/**
 * Fold a sheet name for lookup. Sheet names compare without case. The fold keeps the length:
 * each code point becomes its uppercase only when that uppercase is one code point, so `ß`
 * stays `ß` and `Maße` and `MASSE` stay two names. It is a lookup key only, and the stored name
 * never changes.
 */
export function sheetNameKey(name: string): string {
  if (/^[\x00-\x7f]*$/.test(name)) return name.toUpperCase();
  let out = '';
  for (const ch of name) {
    const up = ch.toUpperCase();
    out += [...up].length === 1 ? up : ch;
  }
  return out;
}

const nameMaps = new WeakMap<WorkbookIndex, ReadonlyMap<string, readonly SheetEntry[]>>();

/** The sheets whose names fold (`sheetNameKey`) to the key of `name`, in tab order. */
export function sheetsByKey(index: WorkbookIndex, name: string): readonly SheetEntry[] {
  let map = nameMaps.get(index);
  if (!map) {
    const built = new Map<string, SheetEntry[]>();
    for (const sheet of index.sheets) {
      const key = sheetNameKey(sheet.name);
      const list = built.get(key);
      if (list) list.push(sheet);
      else built.set(key, [sheet]);
    }
    nameMaps.set(index, (map = built));
  }
  return map.get(sheetNameKey(name)) ?? [];
}

/**
 * The sheet with a name, or undefined. The one sheet with exactly that name wins; else the one
 * sheet whose name matches without case. A name that matches more than one sheet, which only a
 * file can hold, finds none, so an op on it is refused.
 */
export function sheetByName(index: WorkbookIndex, name: string): SheetEntry | undefined {
  const matches = sheetsByKey(index, name);
  if (matches.length === 1) return matches[0];
  const exact = matches.filter((s) => s.name === name);
  return exact.length === 1 ? exact[0] : undefined;
}

/**
 * The names of the defined names that the sheet at `position` scopes (`@localSheetId`), as
 * `sheetNameKey` folds them. Defined names compare without case.
 */
export function localNameKeys(
  pkg: XlsxPackage,
  index: WorkbookIndex,
  position: number
): ReadonlySet<string> {
  const out = new Set<string>();
  const root = xmlPartOf(pkg, index.mainPart)?.root;
  for (const el of children(child(root, 'x:definedNames'), 'x:definedName')) {
    const name = attr(el, 'name');
    if (name && intAttr(el, 'localSheetId') === position) out.add(sheetNameKey(name));
  }
  return out;
}

// ---------------------------------------------------------------- shared strings

const sstCache = new WeakMap<XmlPart, readonly Element[]>();
const NO_ITEMS: readonly Element[] = Object.freeze([]);

/**
 * The `x:si` items of the shared string table, in order. Empty when the part is absent. The same
 * array for the same part revision, so callers can key caches on it.
 */
export function sharedStringItems(pkg: XlsxPackage, index: WorkbookIndex): readonly Element[] {
  const part = xmlPartOf(pkg, index.sharedStringsPart);
  if (!part || part.root.kind !== 'sst') return NO_ITEMS;
  const hit = sstCache.get(part);
  if (hit) return hit;
  const items = children(part.root, 'x:si');
  sstCache.set(part, items);
  return items;
}

// ---------------------------------------------------------------- styles

/** The lists of the styles part that a cell's `@s` reaches, in order. */
export interface StyleLists {
  /** `x:numFmt` by `numFmtId`. The first definition of an id wins. */
  readonly numFmts: ReadonlyMap<number, string>;
  readonly fonts: readonly Element[];
  readonly fills: readonly Element[];
  readonly borders: readonly Element[];
  readonly cellStyleXfs: readonly Element[];
  readonly cellXfs: readonly Element[];
  /** `x:cellStyle` elements. */
  readonly cellStyles: readonly Element[];
  /** `x:dxfs/x:dxf`: the differential formats that conditional formatting rules name. */
  readonly dxfs: readonly Element[];
  /** `x:colors/x:indexedColors/x:rgbColor/@rgb`, when the file redefines the palette. */
  readonly indexedColors: readonly string[] | undefined;
}

const EMPTY_STYLES: StyleLists = Object.freeze({
  numFmts: new Map<number, string>(),
  fonts: [],
  fills: [],
  borders: [],
  cellStyleXfs: [],
  cellXfs: [],
  cellStyles: [],
  dxfs: [],
  indexedColors: undefined,
});

const styleCache = new WeakMap<XmlPart, StyleLists>();

/** The style lists of the workbook. Empty lists when the styles part is absent. */
export function styleLists(pkg: XlsxPackage, index: WorkbookIndex): StyleLists {
  const part = xmlPartOf(pkg, index.stylesPart);
  if (!part || part.root.kind !== 'styleSheet') return EMPTY_STYLES;
  const hit = styleCache.get(part);
  if (hit) return hit;
  const root = part.root;
  const numFmts = new Map<number, string>();
  for (const el of children(child(root, 'x:numFmts'), 'x:numFmt')) {
    const id = intAttr(el, 'numFmtId');
    const code = attr(el, 'formatCode');
    if (id !== undefined && code !== undefined && !numFmts.has(id)) numFmts.set(id, code);
  }
  const indexed = child(child(root, 'x:colors'), 'x:indexedColors');
  const lists: StyleLists = {
    numFmts,
    fonts: children(child(root, 'x:fonts'), 'x:font'),
    fills: children(child(root, 'x:fills'), 'x:fill'),
    borders: children(child(root, 'x:borders'), 'x:border'),
    cellStyleXfs: children(child(root, 'x:cellStyleXfs'), 'x:xf'),
    cellXfs: children(child(root, 'x:cellXfs'), 'x:xf'),
    cellStyles: children(child(root, 'x:cellStyles'), 'x:cellStyle'),
    dxfs: children(child(root, 'x:dxfs'), 'x:dxf'),
    indexedColors: indexed
      ? children(indexed, 'x:rgbColor').map((el) => attr(el, 'rgb') ?? '')
      : undefined,
  };
  styleCache.set(part, lists);
  return lists;
}

/** The theme part's root (`a:theme`), or undefined. */
export function themeRoot(pkg: XlsxPackage, index: WorkbookIndex): Element | undefined {
  const root = xmlPartOf(pkg, index.themePart)?.root;
  return root && root.ns === NS.a && root.local === 'theme' ? root : undefined;
}
