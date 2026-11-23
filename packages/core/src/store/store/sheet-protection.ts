// Sheet protection (`x:sheetProtection`, ECMA-376 Part 1, 18.3.1.85), enforced in the store as
// DOCX Editor enforces document protection: one reader of the element, one reader of the cell
// lock (`cellLocked`), and one refusal with the code `locked` before an op applies.
//
// When `@sheet` is true, the sheet is protected:
//
// - A value edit (`setCellValue`, `clearValues`) is refused on a cell whose protection is locked.
//   The lock is `x:protection/@locked` of the cell format (18.8.33), true by default. A cell
//   format with no `x:protection` and no `@applyProtection` takes the protection of its cell
//   style (`x:cellStyleXfs`, by `@xfId`). A cell that does not exist takes the format of its row
//   or column (`inheritedFormat`).
// - A format edit (`applyFormat`, `clearFormats`, `setCellFormatIndex`) is refused unless
//   `@formatCells` is false. A `setCellFormatIndex` that gives a cell the format it already
//   shows changes nothing, so a paste of values in the same format stays allowed.
// - A `setCellFormatIndex` whose new format changes the `@locked` or `@hidden` protection of the
//   cell is refused with the rule `cell`, whatever `@formatCells` states. So a script or a paste
//   cannot unlock a locked cell by a new cell format and then write it.
// - A `clearFormats` keeps the `@locked` and `@hidden` protection of each cell: a cell gets
//   format 0 with its own protection (`format-range.ts`), whatever format 0 states.
// - A column width (`setColumnWidth`, `setColumnWidths`) is refused unless `@formatColumns` is
//   false, and a row height (`setRowHeight`, `setFittedRowHeights`) unless `@formatRows` is
//   false.
// - An insert or a delete of rows or columns is refused unless `@insertRows`, `@deleteRows`,
//   `@insertColumns`, or `@deleteColumns` is false. When a delete is allowed, it deletes locked
//   cells too: the attribute allows the action on the whole sheet.
//
// Each of those attributes is true by default in sml.xsd (`CT_SheetProtection`), and true means
// that the action is locked. The public `SheetProtection` states the opposite, what is allowed,
// so that true means allowed. The password hash is never read, checked, or changed: the store
// keeps the element as the file holds it.
//
// The workbook protection (`x:workbookProtection`, ECMA-376 Part 1, 18.2.29) locks the sheet
// list when `@lockStructure` is true (false by default in sml.xsd, `CT_WorkbookProtection`): a
// sheet op (`addSheet`, `renameSheet`, `deleteSheet`, `moveSheet`) is then refused with
// `locked` and the rule `workbook`. The store does not enforce `@lockWindows` or
// `@lockRevision`, and keeps its password hash as the file holds it.
//
// A transaction checks each of its ops (`applyOps`), so a paste or a script of many cell writes
// does the checks once, not for each op (`protectionChecks`): it resolves the worksheet of each
// sheet name (`x:sheet/@name`, `CT_Sheet`, 18.2.19) once, reads the `x:sheetProtection` of each
// worksheet part once, and reads the protection of each cell format (`x:cellXfs`, 18.8.10) and
// of each cell style format (`x:cellStyleXfs`, 18.8.9) once. The caches live as long as the
// checks of one transaction, and an edit that replaces a part makes a new key.

import type { Element, XmlPart } from '../package/tree.ts';
import {
  partOf,
  relsPartName,
  xmlPartOf,
  type PackagePart,
  type XlsxPackage,
} from '../package/package.ts';
import { attr, boolAttr, child } from '../package/query.ts';
import { cellKey, sheetIndex, type SheetIndex } from '../workbook/sheet-index.ts';
import {
  sheetByName,
  styleLists,
  workbookIndex,
  type SheetEntry,
  type StyleLists,
  type WorkbookIndex,
} from '../workbook/workbook-index.ts';
import { formatIndex, inheritedFormat } from '../workbook/inherited-format.ts';
import { formatCellRef } from '../cell-ref.ts';
import { clearsValue, rangeBounds, type RangeBox } from './clear-range.ts';
import type { OpRejection, StructuralOpName, WorkbookOp } from './ops.ts';

/**
 * What the protection of a worksheet allows (`x:sheetProtection`, ECMA-376 Part 1, 18.3.1.85).
 * True means allowed. Each field is the opposite of its attribute in the file, where true means
 * locked and the default is true, so a protection that states nothing allows none of them.
 *
 * @public
 */
export interface SheetProtection {
  /**
   * Format edits on any cell (`applyFormat`, `clearFormats`, `setCellFormatIndex`, and the
   * formats of a paste): `@formatCells` is false. Otherwise they are refused with `locked`.
   */
  readonly allowFormatCells: boolean;
  /**
   * Setting the width of columns (`setColumnWidth`, `setColumnWidths`, a resize drag, a fit):
   * `@formatColumns` is false. Otherwise it is refused with `locked`.
   */
  readonly allowFormatColumns: boolean;
  /**
   * Setting the height of rows (`setRowHeight`, `setFittedRowHeights`, a resize drag, a fit):
   * `@formatRows` is false. Otherwise it is refused with `locked`.
   */
  readonly allowFormatRows: boolean;
  /** Inserting whole rows: `@insertRows` is false. */
  readonly allowInsertRows: boolean;
  /** Deleting whole rows, locked cells included: `@deleteRows` is false. */
  readonly allowDeleteRows: boolean;
  /** Inserting whole columns: `@insertColumns` is false. */
  readonly allowInsertColumns: boolean;
  /** Deleting whole columns, locked cells included: `@deleteColumns` is false. */
  readonly allowDeleteColumns: boolean;
}

/**
 * What the protection of a workbook allows (`x:workbookProtection`, ECMA-376 Part 1, 18.2.29).
 * True means allowed.
 *
 * @public
 */
export interface WorkbookProtection {
  /**
   * Adding, renaming, deleting, and moving sheets (`addSheet`, `renameSheet`, `deleteSheet`,
   * `moveSheet`, and the sheet tabs): `@lockStructure` is not true. Otherwise they are refused
   * with `locked`, and the lock rule is `workbook`.
   */
  readonly allowStructure: boolean;
}

/** What the protection of a workbook allows, or null when it has no `x:workbookProtection`. */
export function readWorkbookProtection(
  pkg: XlsxPackage,
  index: WorkbookIndex
): WorkbookProtection | null {
  const el = child(xmlPartOf(pkg, index.mainPart)?.root, 'x:workbookProtection');
  // `@lockStructure` is false by default (`CT_WorkbookProtection`).
  return el ? { allowStructure: boolAttr(el, 'lockStructure') !== true } : null;
}

/** The ops that change the sheet list, which `@lockStructure` locks. */
const SHEET_LIST_OPS: ReadonlySet<string> = new Set([
  'addSheet',
  'renameSheet',
  'deleteSheet',
  'moveSheet',
]);

/** The field of `SheetProtection` that allows each structural op. */
export const STRUCTURE_PERMIT: Readonly<Record<StructuralOpName, keyof SheetProtection>> = {
  insertRows: 'allowInsertRows',
  deleteRows: 'allowDeleteRows',
  insertColumns: 'allowInsertColumns',
  deleteColumns: 'allowDeleteColumns',
};

// Work counters for tests, as `sheetIndexTestRecorder` (`workbook/sheet-index.ts`): a count is
// exact on every machine, a wall-clock limit is not.
const work = { sheetResolves: 0, protectionReads: 0, formatReads: 0, styleReads: 0 };

/**
 * @internal Counts the work of the protection checks: the sheet names resolved to a worksheet
 * (`sheetResolves`), the `x:sheetProtection` reads (`protectionReads`), and the protection reads
 * of a cell format (`formatReads`) and of a cell style format (`styleReads`).
 */
export function protectionWorkTestRecorder(): Readonly<typeof work> & { reset(): void } {
  return {
    get sheetResolves() {
      return work.sheetResolves;
    },
    get protectionReads() {
      return work.protectionReads;
    },
    get formatReads() {
      return work.formatReads;
    },
    get styleReads() {
      return work.styleReads;
    },
    reset() {
      work.sheetResolves = 0;
      work.protectionReads = 0;
      work.formatReads = 0;
      work.styleReads = 0;
    },
  };
}

/**
 * What the protection of a worksheet allows, or null when `@sheet` is not true
 * (`CT_SheetProtection`, ECMA-376 Part 1, 18.3.1.85).
 */
export function readSheetProtection(root: Element | undefined): SheetProtection | null {
  work.protectionReads += 1;
  const el = child(root, 'x:sheetProtection');
  if (boolAttr(el, 'sheet') !== true) return null;
  // An attribute that is true, or absent (its default), locks its action.
  const allows = (name: string) => boolAttr(el, name) === false;
  return {
    allowFormatCells: allows('formatCells'),
    allowFormatColumns: allows('formatColumns'),
    allowFormatRows: allows('formatRows'),
    allowInsertRows: allows('insertRows'),
    allowDeleteRows: allows('deleteRows'),
    allowInsertColumns: allows('insertColumns'),
    allowDeleteColumns: allows('deleteColumns'),
  };
}

/** The lock and the hidden state of cell formats, by cell format index. */
type LockReader = ReturnType<typeof lockReaderOf>;

/**
 * The lock and the hidden state of the cell formats of one styles part, by cell format index.
 * The `x:protection` (`CT_CellProtection`, 18.8.33) of each cell format (`x:cellXfs`, 18.8.10)
 * and of each cell style format (`x:cellStyleXfs`, 18.8.9) is read once for each reader.
 */
function lockReaderOf(lists: StyleLists) {
  const xfs = lists.cellXfs;
  const styleXfs = lists.cellStyleXfs;
  // The `x:protection` of each cell format and of each cell style format, by index. A key that
  // is present with the value undefined is a format with no `x:protection`.
  const byFormat = new Map<number, Element | undefined>();
  const byStyle = new Map<number, Element | undefined>();
  const styleProtection = (xfId: number): Element | undefined => {
    if (byStyle.has(xfId)) return byStyle.get(xfId);
    work.styleReads += 1;
    const el = child(styleXfs[xfId], 'x:protection');
    byStyle.set(xfId, el);
    return el;
  };
  // A cell format with no `x:protection` and no `@applyProtection` takes the protection of its
  // cell style (`@xfId`).
  const protection = (s: number): Element | undefined => {
    if (byFormat.has(s)) return byFormat.get(s);
    work.formatReads += 1;
    const xf = xfs[s] ?? xfs[0];
    const own = child(xf, 'x:protection');
    const el =
      own || boolAttr(xf, 'applyProtection') === true
        ? own
        : styleProtection(formatIndex(attr(xf, 'xfId')) ?? 0);
    byFormat.set(s, el);
    return el;
  };
  return {
    count: xfs.length,
    // `@locked` is true by default, and `@hidden` false (`CT_CellProtection`, 18.8.33).
    locked: (s: number) => boolAttr(protection(s), 'locked') !== false,
    hidden: (s: number) => boolAttr(protection(s), 'hidden') === true,
  };
}

/** The lock and the hidden state of the cell formats of a package, by cell format index. */
export function lockReader(pkg: XlsxPackage, index: WorkbookIndex): LockReader {
  return lockReaderOf(styleLists(pkg, index));
}

/** The cell format index of a cell, or of the row or column format when it does not exist. */
function formatAt(sheet: SheetIndex, row: number, col: number, cellXfs: number): number {
  const el = sheet.cells.get(cellKey(row, col))?.el;
  return el ? (formatIndex(attr(el, 's')) ?? 0) : inheritedFormat(sheet, row, col, cellXfs);
}

/**
 * Whether the cell at 0-based `row` and `col` of a worksheet part is locked, whatever the
 * protection of its sheet. The one reader of the cell lock.
 */
export function cellLocked(
  pkg: XlsxPackage,
  index: WorkbookIndex,
  part: XmlPart,
  row: number,
  col: number
): boolean {
  const lock = lockReader(pkg, index);
  return lock.locked(formatAt(sheetIndex(part), row, col, lock.count));
}

const locked = (
  sheet: string,
  what: string,
  lock: NonNullable<OpRejection['lock']>
): OpRejection => ({
  code: 'locked',
  detail: `${sheet} is protected: ${what}`,
  lock,
});

/** The first locked cell that a value edit of `box` changes, or undefined. Sparse. */
function firstLocked(sheet: SheetIndex, box: RangeBox, lock: LockReader): string | undefined {
  const [top, left, bottom, right] = box;
  // One cell: the cell that exists, or the format that a new cell takes.
  if (top === bottom && left === right)
    return lock.locked(formatAt(sheet, top, left, lock.count))
      ? formatCellRef({ row: top, col: left })
      : undefined;
  // A range: the cells that exist in it, never the cells that the range names.
  for (const e of sheet.cells.values()) {
    if (e.row < top || e.row > bottom || e.col < left || e.col > right) continue;
    if (!clearsValue(e.el)) continue;
    if (lock.locked(formatIndex(attr(e.el, 's')) ?? 0)) return formatCellRef(e);
  }
  return undefined;
}

/** The reads of the protection checks, which `protectionChecks` caches. */
interface ProtectionReads {
  /** The sheet with a name (`sheetByName`). */
  entry(pkg: XlsxPackage, index: WorkbookIndex, name: string): SheetEntry | undefined;
  /** What the protection of a worksheet part allows (`readSheetProtection`). */
  protection(part: XmlPart): SheetProtection | null;
  /** The lock reader of the cell formats of the package (`lockReader`). */
  lock(pkg: XlsxPackage, index: WorkbookIndex): LockReader;
}

/**
 * The protection checks of one transaction: a function that gives the refusal of an op by the
 * protection of its sheet, or null, as `protectionRefusal` does. Its caches live as long as the
 * function, so N cell writes on one sheet resolve the sheet name once, read the protection of
 * each worksheet part once, and read the protection of each cell format once.
 *
 * - A sheet name resolves once while the workbook part and its relationship part stay the same
 *   objects: only an edit of the sheet list (`x:sheets`, 18.2.20) or of those relationships can
 *   change the worksheet that a name finds, and that edit replaces the part.
 * - The protection is read once for each worksheet part object. An edit makes a new part object.
 * - The protection of each cell format is read once while the cell format lists (`x:cellXfs`,
 *   18.8.10, and `x:cellStyleXfs`, 18.8.9) stay the same objects.
 */
export function protectionChecks(): (pkg: XlsxPackage, op: WorkbookOp) => OpRejection | null {
  // The names resolved against one workbook part and one relationship part. A Map keeps a
  // file-supplied name such as `__proto__` a plain key.
  let names:
    | {
        readonly main: XmlPart | undefined;
        readonly rels: PackagePart | undefined;
        readonly byName: Map<string, SheetEntry | undefined>;
      }
    | undefined;
  const protections = new Map<XmlPart, SheetProtection | null>();
  let locks: { readonly lists: StyleLists; readonly reader: LockReader } | undefined;
  const reads: ProtectionReads = {
    entry(pkg, index, name) {
      const main = xmlPartOf(pkg, index.mainPart);
      const rels = partOf(pkg, relsPartName(index.mainPart));
      if (!names || names.main !== main || names.rels !== rels)
        names = { main, rels, byName: new Map() };
      if (names.byName.has(name)) return names.byName.get(name);
      work.sheetResolves += 1;
      const entry = sheetByName(index, name);
      names.byName.set(name, entry);
      return entry;
    },
    protection(part) {
      if (protections.has(part)) return protections.get(part)!;
      const protection = readSheetProtection(part.root);
      protections.set(part, protection);
      return protection;
    },
    lock(pkg, index) {
      const lists = styleLists(pkg, index);
      if (locks?.lists !== lists) locks = { lists, reader: lockReaderOf(lists) };
      return locks.reader;
    },
  };
  return (pkg, op) => refusalOf(pkg, op, reads);
}

/** The refusal of an op by the protection of its sheet, or null. */
export function protectionRefusal(pkg: XlsxPackage, op: WorkbookOp): OpRejection | null {
  return protectionChecks()(pkg, op);
}

function refusalOf(pkg: XlsxPackage, op: WorkbookOp, reads: ProtectionReads): OpRejection | null {
  if (SHEET_LIST_OPS.has(op.op)) {
    const index = workbookIndex(pkg);
    const protection = index ? readWorkbookProtection(pkg, index) : null;
    return protection && !protection.allowStructure
      ? {
          code: 'locked',
          detail: `the workbook is protected: ${op.op} is not allowed`,
          lock: 'workbook',
        }
      : null;
  }
  if (!('sheet' in op)) return null;
  const index = workbookIndex(pkg);
  const entry = index ? reads.entry(pkg, index, op.sheet) : undefined;
  const part = xmlPartOf(pkg, entry?.partName);
  if (!index || !part) return null;
  const protection = reads.protection(part);
  if (!protection) return null;
  const name = entry!.name;
  const permit = Object.hasOwn(STRUCTURE_PERMIT, op.op)
    ? STRUCTURE_PERMIT[op.op as StructuralOpName]
    : undefined;
  if (permit)
    return protection[permit] ? null : locked(name, `${op.op} is not allowed`, 'structure');
  if (op.op === 'setCellFormatIndex') {
    // A paste writes the format of each cell: a format that the cell already shows changes
    // nothing, so only a new format is refused.
    const box = rangeBounds(op.ref);
    const lock = reads.lock(pkg, index);
    const now = box ? formatAt(sheetIndex(part), box[0], box[1], lock.count) : undefined;
    if (now === op.index) return null;
    if (!protection.allowFormatCells)
      return locked(name, 'formatting cells is not allowed', 'format');
    // A new format must keep the protection of the cell, or it unlocks a locked cell.
    if (
      now !== undefined &&
      (lock.locked(now) !== lock.locked(op.index) || lock.hidden(now) !== lock.hidden(op.index))
    )
      return locked(name, `the protection of cell ${op.ref} cannot change`, 'cell');
    return null;
  }
  if (op.op === 'applyFormat' || op.op === 'clearFormats')
    return protection.allowFormatCells
      ? null
      : locked(name, 'formatting cells is not allowed', 'format');
  if (op.op === 'mergeCells' || op.op === 'unmergeCells')
    return locked(name, `${op.op} is not allowed on a protected sheet`, 'merge');
  if (op.op === 'setColumnWidth' || op.op === 'setColumnWidths')
    return protection.allowFormatColumns
      ? null
      : locked(name, 'formatting columns is not allowed', 'size');
  if (op.op === 'setRowHeight' || op.op === 'setFittedRowHeights')
    return protection.allowFormatRows
      ? null
      : locked(name, 'formatting rows is not allowed', 'size');
  if (op.op !== 'setCellValue' && op.op !== 'clearValues') return null;
  const box = rangeBounds(op.ref);
  if (!box) return null;
  const cell = firstLocked(sheetIndex(part), box, reads.lock(pkg, index));
  return cell ? locked(name, `cell ${cell} is locked`, 'cell') : null;
}

/**
 * The first locked cell, as `A1`, that a value edit of the 0-based `box` (top, left, bottom,
 * right) would change, or undefined; the rule of `protectionRefusal`, whatever the protection
 * of the sheet. One cell is the cell or the format that a new cell takes; a range is the cells
 * that exist in it and hold content, never the cells that the range names.
 */
export function lockedCellIn(
  pkg: XlsxPackage,
  index: WorkbookIndex,
  part: XmlPart,
  box: RangeBox
): string | undefined {
  return firstLocked(sheetIndex(part), box, lockReader(pkg, index));
}
