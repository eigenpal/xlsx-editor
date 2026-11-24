// The typed ops: the only way a transaction changes the package (design D2).
//
// Ops express intent, never shared-part indexes. An op names a sheet by name and a cell by its
// A1 reference against the revision it applies to, and carries a value or a format patch. The
// store resolves or appends the shared string and the style records when it applies the op.
// Every op returns the ops that restore the content it replaced (its inverse), so undo is a new
// transaction of inverse ops, never a whole-tree snapshot.
//
// A transaction can need side effects that no op names: the stale flag of design D3 and the
// dropped calculation chain of design D5 (`calc.ts`). `applyOps` adds them as ops of their own,
// in the same transaction, with no inverse, so they stay outside the undo history.

import type { Element, XmlPart } from '../package/tree.ts';
import type { XlsxPackage } from '../package/package.ts';
import { NS } from '../package/namespaces.ts';
import { attr } from '../package/query.ts';
import { parseCellRef, MAX_ROWS } from '../cell-ref.ts';
import { cellKey, sheetIndex } from '../workbook/sheet-index.ts';
import { inheritedFormat } from '../workbook/inherited-format.ts';
import { styleLists, workbookIndex } from '../workbook/workbook-index.ts';
import { readFormulaInput } from '../formula/format.ts';
import { formulaScope } from '../formula/references.ts';
import { protectionChecks } from './sheet-protection.ts';
import { encodeXstring } from '../workbook/xstring.ts';
import type { FormulaWriteProblem } from '../formula/format.ts';
import {
  cellOrNestedHasFormula,
  rowChange,
  sameNode,
  withFullCalcOnLoad,
  withoutCalcChain,
  withoutDanglingCalcChain,
  workbookHasFormula,
} from './calc.ts';
import { formulaRangeSubject, isSharedMaster, unshareChildren } from './shared-formulas.ts';
import { rangeMasterBlocker, rangeMasterRef } from '../workbook/cell-formula.ts';
import { sharedStringIndex, sharedStringsUndo } from './shared-strings.ts';
import {
  arrayFormulaAt,
  buildCell,
  cellChildBlocker,
  isBareRow,
  pinCellRefs,
  pinRowRefs,
  isIsoDate,
  cellSlot,
  rowSlot,
  sheetTarget,
  withPart,
  withSheetData,
  type SheetTarget,
} from './cells.ts';
import { clearRange } from './clear-range.ts';
import { applyFormatRange, applyFormatRanges, clearFormatsRange } from './format-range.ts';
import { restoreStyles } from './styles-part.ts';
import { patchProblem, type FormatPatch } from './format-patch.ts';
import { applyStructure, restoreParts, type PartRestore } from '../structure/structure.ts';
import { applySheetOp } from '../structure/sheet-ops.ts';
import { restoreRows } from './restore-rows.ts';
import { applyAxisSize, applyColumnWidths, applyFittedRowHeights } from './axis-size.ts';
import { applyMerge } from './merge-cells.ts';
import { withoutFittedHeight } from './format-range.ts';
import { setCellValues } from './cell-batch.ts';
import { setCellFormatIndexes } from './cell-format-index.ts';
import { dimensionBounds, widenDimension, withDimensionRef } from './dimension.ts';
import { IdSource, plainAttr, smlElement, withAttr, withChild } from './tree-edit.ts';

/** A value to store in a cell. `empty` clears the value and keeps the cell format. */
export type CellInput =
  | { readonly type: 'number'; readonly value: number }
  | { readonly type: 'string'; readonly value: string }
  | { readonly type: 'boolean'; readonly value: boolean }
  /**
   * Formula text without the leading `=`. The store writes it as `x:f` with no cached value,
   * in the form the tokenizer writes back (`normalizeFormula`), and refuses text that the
   * tokenizer cannot read.
   */
  | { readonly type: 'formula'; readonly text: string }
  /**
   * An ISO 8601 date or date-time (`2024-03-01`, `2024-03-01T12:30:00`), written as `t="d"`
   * with the text as `x:v`. The cell keeps its format, which converts the date to a serial for
   * display. Text that is not an ISO 8601 date is refused. The Transitional
   * schema has no `d` value in `ST_CellType` (sml.xsd; ECMA-376 Part 1, 18.18.11).
   */
  | { readonly type: 'date'; readonly iso: string }
  | { readonly type: 'empty' };

type SetCellFormatIndex = Extract<WorkbookOp, { op: 'setCellFormatIndex' }>;

/** The names of the structural ops. */
export type StructuralOpName = 'insertRows' | 'deleteRows' | 'insertColumns' | 'deleteColumns';

/** One structural op for each name, so that a check of `op` narrows the union. */
type StructuralOps = {
  [K in StructuralOpName]: {
    readonly op: K;
    readonly sheet: string;
    readonly start: number;
    readonly count: number;
  };
}[StructuralOpName];

/** One typed op. */
export type WorkbookOp =
  | {
      readonly op: 'setCellValue';
      readonly sheet: string;
      readonly ref: string;
      readonly value: CellInput;
    }
  /** Set `x:calcPr/@fullCalcOnLoad`. No inverse: undo never clears it. */
  | { readonly op: 'setFullCalcOnLoad' }
  /** Drop the calculation chain. No inverse: the chain is optional. */
  | { readonly op: 'dropCalcChain' }
  /**
   * Set `x:dimension/@ref` of a worksheet that has an `x:dimension`. A write that creates a
   * cell outside it widens it in the same op; this op is its inverse.
   */
  | { readonly op: 'setDimension'; readonly sheet: string; readonly ref: string }
  /**
   * Clear the values and formulas of the cells that exist in a range (`A1` or `A1:C3`), and keep
   * their cell formats. Cells inside an array formula range refuse it; shared formula masters
   * inside it are unshared first. The other refusals are the `shared-formula` cases
   * of {@link OpRejection}.
   */
  | { readonly op: 'clearValues'; readonly sheet: string; readonly ref: string }
  /**
   * Apply a format patch to every cell of a range (`A1` or `A1:C3`): each cell gets its own cell
   * format with the patch applied. The store reuses or appends the style records; cells that do
   * not exist are created, at most `MAX_FORMAT_NEW_CELLS` of them (`format-range.ts`).
   */
  | {
      readonly op: 'applyFormat';
      readonly sheet: string;
      readonly ref: string;
      readonly patch: FormatPatch;
    }
  /**
   * Remove the cell format (`@s`) of the cells that exist in a range, and keep their values. On a
   * protected sheet, a cell keeps its protection: it gets format 0 with that protection. When a
   * formula or a defined name calls `CELL` or a `GET.*` macro, the cached values become stale.
   * When the record of that protection would pass the record limit of `x:cellXfs`, the op is
   * refused with `style-limit`. When the styles part cannot be read, such as a part past a read
   * limit that stays as bytes, the op still gives the default format 0. On a protected sheet, it
   * is then refused with `unsupported-package`, because the protection of a cell is in the
   * unread part.
   */
  | { readonly op: 'clearFormats'; readonly sheet: string; readonly ref: string }
  /**
   * Give one cell (`A1`) the cell format at `index` of `x:cellXfs`, and keep its value (design
   * D2 of the change `workbook-structure-editing`): the format that an internal paste copies.
   * An index that the workbook does not have is refused with `bad-format`. A cell with no
   * `x:c` that already shows the format (from its row or column) stays absent, and an emptied
   * cell that shows the format of its row or column goes away (`cell-format-index.ts`). On a
   * protected sheet that does not allow formatting cells, a new format is refused with
   * `locked` (`sheet-protection.ts`). On any protected sheet, a new format that changes the
   * `@locked` or `@hidden` protection of the cell is refused with `locked` and the rule `cell`.
   * When a formula or a defined name calls `CELL` or a `GET.*` macro, a changed cell format
   * makes the cached values stale (`CELL("format")` reads it). When the styles part cannot be
   * read, the op gives format 0 only, and any other index is refused with `unsupported-package`.
   */
  | {
      readonly op: 'setCellFormatIndex';
      readonly sheet: string;
      readonly ref: string;
      readonly index: number;
    }
  /**
   * Insert or delete whole rows or columns of a worksheet: `start` is the 1-based
   * number of the first row or column (`A` is 1), and `count` how many, at most the size of the
   * grid. Every reference of the workbook moves with them, or the op is refused with a reason
   * that names what holds the reference (`structure-blocked`) or the content that would leave
   * the grid (`grid-full`). It is refused with `sheet-name-ambiguous`, with the other sheet as
   * the subject, when a file gives the name of the sheet to another sheet too, in any case,
   * and with `sheet-name-too-long`, with that name as the subject, when a stored sheet name has
   * more than 31 characters. A hyperlink relationship target that is a fragment (`#Data!A5`)
   * moves with the cells that it names. A formula or a defined name inside
   * `mc:AlternateContent` whose text the edit would change refuses it with `structure-blocked`,
   * because the editor does not rewrite each branch.
   */
  | StructuralOps
  /**
   * Add an empty worksheet named `name` at position `index` of the sheet list (0 is the first
   * tab; the sheet count adds it at the end). A name that breaks the rules of a sheet name is
   * refused with `sheet-name-invalid` or `sheet-name-taken`, and a position outside 0 to the sheet
   * count with `bad-position`. Each sheet op is refused with `locked` (rule `workbook`) when
   * `x:workbookProtection/@lockStructure` is true. Refused with `structure-blocked` when the add
   * changes the position of the sheet that the `@localSheetId` of a defined name inside
   * `mc:AlternateContent` names, because the editor does not rewrite each branch, and when the
   * workbook holds macros (`about` kind `macros`), because macro code can read a sheet by its
   * position.
   */
  | { readonly op: 'addSheet'; readonly name: string; readonly index: number }
  /**
   * Rename a sheet, and rewrite each reference to it in every formula, defined name, and
   * internal hyperlink location. Refused with `unknown-sheet` when no sheet has the name, with
   * `sheet-name-invalid` or `sheet-name-taken` for the new name, with `structure-blocked`
   * when other content names the sheet and the store does not rewrite it, and with
   * `sheet-name-ambiguous` when another sheet has the same name without case and a reference
   * names one of the two, and with `locked` when the workbook protection locks its structure.
   * A formula or a defined name inside `mc:AlternateContent` that names the sheet is content
   * that the store does not rewrite, so it refuses the rename with `structure-blocked`.
   */
  | { readonly op: 'renameSheet'; readonly sheet: string; readonly name: string }
  /**
   * Delete a sheet with its part, relationships, and override. Each reference to it becomes
   * `#REF!`, and its sheet-scoped defined names go. Refused with `unknown-sheet` when no sheet
   * has the name, with `last-sheet` when no visible sheet would be left, with
   * `sheet-range-blocked` for a sheet range that starts or ends at the sheet, with
   * `structure-blocked` for a related part, with `sheet-name-ambiguous` when another sheet
   * has the same name without case, and with `locked` when the workbook protection locks its
   * structure. Refused with `structure-blocked` when the delete would change a formula or a
   * defined name inside `mc:AlternateContent`, or when it changes the position of the sheet that
   * the `@localSheetId` of a defined name inside `mc:AlternateContent` names.
   */
  | { readonly op: 'deleteSheet'; readonly sheet: string }
  /**
   * Move a sheet to position `index` of the sheet list. Sheet-scoped defined names and the
   * active tab follow their sheets. A move to the same place changes nothing. Refused with
   * `unknown-sheet` when no sheet has the name, with `bad-position` for a position outside
   * the sheet list, with `sheet-range-blocked` for a sheet range whose span the move
   * changes, and with `locked` when the workbook protection locks its structure. Refused with
   * `structure-blocked` when the move changes the position of the sheet that the `@localSheetId`
   * of a defined name inside `mc:AlternateContent` names, and when the workbook holds macros
   * (`about` kind `macros`).
   */
  | { readonly op: 'moveSheet'; readonly sheet: string; readonly index: number }
  /**
   * Set the width of the columns `first` to `last` (1-based, `A` is 1) to `width` characters,
   * from 0 to 255, with `@customWidth`.
   * The `x:col` elements that the range crosses split, and keep their other attributes. A
   * width outside 0 to 255 is refused with `bad-size`, a range outside the grid with
   * `bad-reference`, an unknown sheet with `unknown-sheet`, a resize that would change a drawing,
   * control, or object anchor that keeps its size or position, or a resize of a sheet whose
   * drawing part cannot be read, with `resize-anchor`, and a cell structure that the editor
   * cannot change with `invalid-sheet` (`axis-size.ts`). On a protected sheet, it is refused
   * with `locked` unless the protection allows formatting columns (`@formatColumns` is false).
   */
  | {
      readonly op: 'setColumnWidth';
      readonly sheet: string;
      readonly first: number;
      readonly last: number;
      readonly width: number;
    }
  /**
   * Set the height of the rows `first` to `last` (1-based) to `height` points, from 0 to 409,
   * with `@customHeight`. A row with no `x:row` gets an empty one. A height outside 0 to 409 is
   * refused with `bad-size`, a range outside the grid with `bad-reference`, an unknown sheet
   * with `unknown-sheet`, a range that would create more than `MAX_RESIZE_NEW_ROWS` rows with
   * `too-many-rows`, a resize that would change an anchor that keeps its size or position, or a
   * resize of a sheet whose drawing part cannot be read, with `resize-anchor`, and rows out of
   * order in `x:sheetData` with `invalid-sheet`. On a protected sheet, it is refused with
   * `locked` unless the protection allows formatting rows (`@formatRows` is false).
   */
  | {
      readonly op: 'setRowHeight';
      readonly sheet: string;
      readonly first: number;
      readonly last: number;
      readonly height: number;
    }
  /**
   * Merge the range `ref` of a sheet into one cell (`x:mergeCell`, ECMA-376 Part 1, 18.3.1.54).
   * The top-left cell keeps its value. A range where another cell holds a value or a formula is
   * refused with `merge-drops-values`, so a caller clears those cells first in the same
   * transaction; a range of one cell with `merge-one-cell`; and a range that meets a merged
   * range with `merge-overlap` (`merge-cells.ts`). On a protected sheet, it is refused with
   * `locked`.
   */
  | { readonly op: 'mergeCells'; readonly sheet: string; readonly ref: string }
  /**
   * Unmerge each merged range of a sheet that meets the range `ref`. A range that meets no
   * merged range changes nothing. On a protected sheet, it is refused with `locked`.
   */
  | { readonly op: 'unmergeCells'; readonly sheet: string; readonly ref: string }
  /**
   * Set the width of each listed column (1-based `column`, ascending) to `width` characters, in
   * one pass over `x:cols`, as `setColumnWidth` sets one range. An empty list, columns out of
   * ascending order, or a column outside the grid is refused with `bad-reference`, a width
   * outside 0 to 255 with `bad-size`, an unknown sheet with `unknown-sheet`, an `x:col` without
   * a valid range with `invalid-sheet`, and a width change that meets an anchor that keeps its
   * size or position with `resize-anchor`. Protection and stale values follow `setColumnWidth`.
   * A fit to content uses it: the file stores a fitted width as a custom width (`@customWidth`),
   * so no fitted form exists, unlike `setFittedRowHeights`. A list that changes no `x:col`
   * attribute is no change. A column at its width that gains `@customWidth` or loses `@bestFit`
   * changes the file, and no anchor refuses it.
   */
  | {
      readonly op: 'setColumnWidths';
      readonly sheet: string;
      readonly columns: readonly { readonly column: number; readonly width: number }[];
    }
  /**
   * Set each listed row (1-based `row`, ascending, each with an `x:row`) to the given height as
   * a fitted height: `@ht` without `@customHeight`. An existing `@customHeight` goes, because a
   * fitted height is not a height set by hand (18.3.1.73). One pass over `x:sheetData` sets every
   * row. A row with no `x:row`, a row outside the grid, listed rows out of ascending order, or an
   * empty list is refused with `bad-reference`, a height outside 0 to 409 with `bad-size`, an
   * unknown sheet with `unknown-sheet`, rows out of order in `x:sheetData` with `invalid-sheet`,
   * and a change that meets an anchor that keeps its size or position with `resize-anchor`: a new
   * height, or the loss of a true `@customHeight` (`1` or `true`), after which a value edit gives
   * the row its automatic height. Protection and stale values follow `setRowHeight`. A list that
   * changes no `x:row` attribute is no change. A row at its height that loses a true
   * `@customHeight` changes the file, and an anchor can refuse it.
   */
  | {
      readonly op: 'setFittedRowHeights';
      readonly sheet: string;
      readonly rows: readonly { readonly row: number; readonly height: number }[];
    };

/**
 * The inverse ops: only undo and redo apply them, because they put back content of an earlier
 * revision as it was, with no check of what it holds. `transact` refuses them, so no caller can
 * write a part, a style list, or a row through them.
 */
export type InverseOp =
  | {
      /**
       * Put back one row as it was: the inverse of an op that changed cells in it. `node` is a
       * row element of an earlier revision, or `null` to remove the row.
       */
      readonly op: 'restoreRow';
      readonly sheet: string;
      /** 1-based row number. */
      readonly row: number;
      readonly node: Element | null;
    }
  /**
   * Put back the styles part of an earlier revision, or remove the styles part that a later edit
   * created (`null`): the inverse of a format edit that appended style records.
   */
  | { readonly op: 'restoreStyles'; readonly part: XmlPart | null }
  /**
   * Put back parts of an earlier revision, or remove a part (`null`): the inverse of a
   * structural edit, which rewrites references across the workbook.
   */
  | { readonly op: 'restoreParts'; readonly parts: readonly PartRestore[] };

/** An op that a history entry holds: a typed op or an inverse op. */
export type HistoryOp = WorkbookOp | InverseOp;

const INVERSE_ONLY: ReadonlySet<string> = new Set<InverseOp['op']>([
  'restoreRow',
  'restoreStyles',
  'restoreParts',
]);

/** Whether an op is an inverse op, which only undo and redo apply. */
export const isInverseOp = (op: { readonly op: string }): boolean => INVERSE_ONLY.has(op.op);

/** Why an op or a transaction was refused. */
export interface OpRejection {
  readonly code:
    | 'unknown-sheet'
    | 'bad-reference'
    | 'bad-value'
    /** A text longer than a cell holds (`MAX_CELL_TEXT`). */
    | 'text-too-long'
    /** A column width or a row height outside its allowed values. */
    | 'bad-size'
    /** A merge of one cell: there is nothing to merge. */
    | 'merge-one-cell'
    /** A merge that meets a merged range. */
    | 'merge-overlap'
    /** A merge would drop a value or a formula outside the top-left cell of its range. */
    | 'merge-drops-values'
    /** A sheet position outside the sheet list. */
    | 'bad-position'
    | 'invalid-sheet'
    /**
     * The edit would change a formula range that the store keeps whole: the cell is the master of,
     * or inside, an array or data table formula range (its master can be inside
     * `mc:AlternateContent`); or it is a shared formula master whose text cannot be read; or it is
     * a shared master with a child whose `x:f` is inside `mc:AlternateContent` in its cell, which
     * the unshare cannot rewrite; or the edit would move the formula of a shared child out of the
     * grid. An edit to any other shared master unshares its children. The `subject` names the
     * formula range that blocks the edit (`Sheet1!A2:A3`), or the cell when the range cannot be
     * read; `causeCell` names the cell whose formula the editor cannot rewrite, for a shared
     * master. A batch writes its cells in order, so a batch that writes a shared master before its
     * wrapped child is refused. A cell with a formula inside `mc:AlternateContent` refuses its own
     * edit or clear with `cell-alternate-content` first (`cellChildBlocker`).
     */
    | 'shared-formula'
    /** Formula text that the reference tokenizer cannot read. */
    | 'unreadable-formula'
    /** Formula text with an operator, an operand, or a separator out of place. */
    | 'formula-syntax'
    /** Formula text in R1C1 notation: `x:f` holds A1 text only. */
    | 'r1c1-formula'
    /**
     * Typed formula text that calls a function outside the predefined functions of ECMA-376
     * Part 1, 18.17.7, such as a future function or a name with a stored-name prefix.
     */
    | 'unsupported-function'
    /**
     * Formula text that names a sheet, an external workbook index `[n]`, or a table that the
     * workbook does not hold, or a sheet name with a space that is not quoted.
     */
    | 'unknown-formula-reference'
    /** Text with a character that XML 1.0 cannot hold, such as a control character. */
    | 'invalid-character'
    | 'unsupported-package'
    /**
     * A value edit, a paste, a clear, or a cut of a cell that holds another form of its value,
     * which the edit can neither keep nor replace (`cellChildBlocker`): non-empty markup
     * compatibility content that is not a stored value (`x:v` or `x:is`), a formula in it
     * included, or an element in another namespace with `x:f`, `x:v`, or `x:is` in it. The
     * subject is the cell with its sheet, such as `Data!B5`.
     */
    | 'cell-alternate-content'
    /** A format patch value outside its allowed values. */
    | 'bad-format'
    /** A custom number format code that the number format parser cannot read. */
    | 'unreadable-number-format'
    /** A format command that would create more cells than one command may. */
    | 'range-too-large'
    /** A row height edit that would create more row records than one edit may. */
    | 'too-many-rows'
    /**
     * A resize would change an anchor that keeps its size or position in points: a drawing
     * anchor, or the anchor of an embedded control or object (`resizeBlocker`). Also a resize of
     * a sheet whose drawing part cannot be read, so its anchors cannot be checked; that refusal
     * has `causeUnreadable`.
     */
    | 'resize-anchor'
    /** A clear of formats that would create more `s="0"` cells than one command may. */
    | 'clear-too-large'
    /** A style list that would hold more records than the store allows. */
    | 'style-limit'
    /** The transaction would leave a part that breaks a structural rule. */
    | 'invalid-part'
    /**
     * The sheet is protected (`x:sheetProtection/@sheet`), and the op changes a locked cell, or
     * an action that the protection does not allow; or the workbook protection locks the sheet
     * list (`x:workbookProtection/@lockStructure`) and the op is a sheet op
     * (`sheet-protection.ts`).
     */
    | 'locked'
    /** `undo` or `redo` with an empty history stack. */
    | 'nothing-to-undo'
    | 'nothing-to-redo'
    /**
     * A structural edit would change a position that an element, a part, or a formula refers
     * to, and the store cannot move that reference with certainty.
     */
    | 'structure-blocked'
    /**
     * A sheet move or delete would change the sheets that a sheet range (`Sheet1:Sheet3!A1`)
     * spans, and the store does not rewrite sheet ranges.
     */
    | 'sheet-range-blocked'
    /**
     * A stored sheet name holds more than 31 characters, so an op that searches the sheet ranges
     * of the workbook is refused (`longSheetNameBlocker`). The subject is the full name.
     */
    | 'sheet-name-too-long'
    /** A structural edit would push content past the last row or column of the grid. */
    | 'grid-full'
    /** The ops were prepared against a revision that is no longer the current one. */
    | 'stale-revision'
    /** A sheet name that is empty, too long, or holds a character that a sheet name cannot. */
    | 'sheet-name-invalid'
    /** A sheet name that another sheet has, compared without case. */
    | 'sheet-name-taken'
    /**
     * Two sheets of a file have the same name without case, so a reference cannot tell them
     * apart: an edit that rewrites or moves references of one of them is refused.
     */
    | 'sheet-name-ambiguous'
    /** A sheet delete that would leave the workbook with no visible sheet. */
    | 'last-sheet';
  readonly detail: string;
  /**
   * What the refusal names, as the file names it: an element and its part
   * (`x:conditionalFormatting (/xl/worksheets/sheet1.xml)`), a part, a cell (`Data!B5`), or a
   * range (`B2:C3`). A merge names a cell or a range of its sheet, with no sheet name.
   */
  readonly subject?: string;
  /**
   * The cell, with its sheet (`Data!A2`), whose formula the editor cannot rewrite, on a
   * `shared-formula` refusal of a shared formula master: a child inside `mc:AlternateContent`, a
   * child whose formula would move out of the grid, or a master whose text cannot be read.
   */
  readonly causeCell?: string;
  /**
   * True when the formula of `causeCell` cannot be read, not only not rewritten. On a
   * `resize-anchor` refusal, true when the drawing part of the sheet cannot be read (the reader
   * keeps it as bytes, or it is not a `xdr:wsDr`), so the resize cannot be checked.
   */
  readonly causeUnreadable?: true;
  /**
   * The shared formula master, with its sheet (`Data!A1`), that the edit would change, beside
   * `causeCell`. It is inside the edited or cleared cells, so a message names it to leave out.
   */
  readonly masterCell?: string;
  /**
   * What a `structure-blocked` or `grid-full` refusal names, as a kind and a display name. A
   * `sheet-range-blocked` refusal has it for content that the op does not rewrite, such as
   * conditional formatting, and for a formula inside `mc:AlternateContent` (kind
   * `wrapped-formula`), which the user cannot edit. A formula, a defined name, or a hyperlink
   * that the op rewrites gives a `sheet-range-blocked` refusal with no `about`.
   */
  readonly about?: StructureSubject;
  /**
   * The protection rule of a `locked` refusal: a locked `cell` that a value edit changes, or a
   * cell whose `@locked` or `@hidden` protection a new cell format index changes; `format` when
   * the protection does not allow formatting cells; `structure` when it does not allow an
   * insert or a delete of rows or columns; `size` when it does not allow a column width or a
   * row height (`@formatColumns`, `@formatRows`); `merge` on a protected sheet, which allows no
   * merge or unmerge; or `workbook` when the workbook protection locks the sheet list
   * (`x:workbookProtection/@lockStructure`).
   */
  readonly lock?: 'cell' | 'format' | 'structure' | 'size' | 'workbook' | 'merge';
}

/**
 * The kinds of content that refuse a structural edit, or that a `grid-full`
 * refusal names, for user text. The part path stays in the rejection's `subject` and `detail`.
 */
export const STRUCTURE_SUBJECT_KINDS = [
  'cell',
  'wrapped-formula',
  'cell-content',
  'row',
  'merged-range',
  'columns',
  'hyperlink',
  'sheet-view',
  'conditional-formatting',
  'data-validation',
  'filter',
  'table',
  'comment',
  'drawing',
  'linked-drawing',
  'control',
  'chart',
  'pivot-table',
  'other-sheet',
  'not-worksheet',
  'defined-name',
  'sheet-content',
  'workbook-content',
  'macros',
] as const;

/** One of {@link STRUCTURE_SUBJECT_KINDS}. A test keeps them equal to the catalog nouns. */
export type StructureSubjectKind = (typeof STRUCTURE_SUBJECT_KINDS)[number];

/**
 * What refuses a structural edit: its kind, and a display name where the kind has one: the
 * sheet name, the defined name, or a reference such as `Data!B5`.
 */
export interface StructureSubject {
  readonly kind: StructureSubjectKind;
  readonly name?: string;
}

export type ApplyResult =
  | {
      readonly ok: true;
      readonly pkg: XlsxPackage;
      readonly inverse: readonly HistoryOp[];
      /** The side-effect ops that the transaction added, already applied. */
      readonly effects: readonly WorkbookOp[];
    }
  | { readonly ok: false; readonly rejection: OpRejection };

/** What one op did. */
export type OpResult =
  | {
      readonly ok: true;
      readonly pkg: XlsxPackage;
      readonly inverse: readonly HistoryOp[];
      /** The op changed a cell value. */
      readonly value: boolean;
      /** The op added, removed, or changed a formula. */
      readonly formula: boolean;
    }
  | { readonly ok: false; readonly rejection: OpRejection };

/**
 * The most UTF-16 code units that the text of one cell holds: 32,767, the limit that consumers of
 * the format apply to a cell value. A longer text is refused with `text-too-long`.
 */
export const MAX_CELL_TEXT = 32_767;

export const refuse = (code: OpRejection['code'], detail: string): OpResult => ({
  ok: false,
  rejection: { code, detail },
});

/** The rejection code of a formula that cannot be written. */
function formulaRejection(problem: FormulaWriteProblem): OpRejection['code'] {
  switch (problem) {
    case 'invalid-character':
      return 'invalid-character';
    case 'r1c1-reference':
      return 'r1c1-formula';
    case 'unsupported-function':
      return 'unsupported-function';
    case 'unknown-sheet':
    case 'unknown-external':
    case 'unknown-table':
    case 'unquoted-sheet-name':
      return 'unknown-formula-reference';
    case 'missing-operand':
    case 'missing-operator':
    case 'misplaced-separator':
      return 'formula-syntax';
    default:
      return 'unreadable-formula';
  }
}

/** The result of an op that changes nothing: no revision and no undo entry. */
export const unchanged = (pkg: XlsxPackage): OpResult => ({
  ok: true,
  pkg,
  inverse: [],
  value: false,
  formula: false,
});

/**
 * Apply ops in order to a package. Atomic: a refused op refuses the whole list. Inverse ops are
 * refused unless `history` is set, which only undo and redo of the document store set.
 * `openedParts` holds the part names of the package that the store opened, for the sheet ops
 * (`applySheetOp`).
 */
export function applyOps(
  pkg: XlsxPackage,
  ops: readonly HistoryOp[],
  history = false,
  openedParts?: ReadonlySet<string>
): ApplyResult {
  if (!history)
    for (const op of ops)
      if (isInverseOp(op))
        return {
          ok: false,
          rejection: {
            code: 'bad-value',
            detail: `${op.op} is an inverse op, which only undo and redo apply`,
          },
        };
  let current = pkg;
  // The end of a run of applyFormat ops that overlap, which apply one by one.
  let unbatched = 0;
  // Each op's inverse ops, in op order; undo applies them newest first.
  const steps: (readonly HistoryOp[])[] = [];
  let value = false;
  let formula = false;
  // The protection checks of the whole list, with one set of caches (`protectionChecks`).
  const protectionRefusal = history ? undefined : protectionChecks();
  for (let i = 0; i < ops.length;) {
    const op = ops[i]!;
    // A run of restoreRow ops on one sheet (an undo of a many-row command), or of setCellValue
    // or setCellFormatIndex ops on one sheet (a paste or a script), is one pass. So is a run of
    // applyFormat ops on one sheet (the format ranges of a paste), each with its own patch.
    let end = i + 1;
    if (op.op === 'restoreRow' || op.op === 'setCellValue' || op.op === 'setCellFormatIndex')
      while (end < ops.length) {
        const next = ops[end]!;
        if (next.op !== op.op || next.sheet !== op.sheet) break;
        end += 1;
      }
    else if (op.op === 'applyFormat' && i >= unbatched) {
      // Only valid patches join a run, so an op that is not valid is refused as it is alone.
      while (validPatch(op.patch) && end < ops.length) {
        const next = ops[end]!;
        if (next.op !== 'applyFormat' || next.sheet !== op.sheet) break;
        if (!validPatch(next.patch)) break;
        end += 1;
      }
    }
    // The protection of the sheet refuses an op before it applies (`sheet-protection.ts`).
    for (let k = i; k < end && protectionRefusal; k += 1) {
      const locked = protectionRefusal(current, ops[k] as WorkbookOp);
      if (locked) return { ok: false, rejection: locked };
    }
    const run = end - i > 1 ? ops.slice(i, end) : undefined;
    const batch = !run
      ? undefined
      : op.op === 'applyFormat'
        ? applyFormatRanges(current, op.sheet, run as Extract<WorkbookOp, { op: 'applyFormat' }>[])
        : op.op === 'restoreRow'
          ? restoreRows(current, run as Extract<HistoryOp, { op: 'restoreRow' }>[])
          : op.op === 'setCellFormatIndex'
            ? setCellFormatIndexes(current, run as SetCellFormatIndex[])
            : setCellValues(current, run as Extract<WorkbookOp, { op: 'setCellValue' }>[]);
    // Ranges that overlap apply one by one, with no new batch for the rest of the run.
    if (run && !batch && op.op === 'applyFormat') unbatched = end;
    const r = batch ?? applyOp(current, op, openedParts);
    i = batch ? end : i + 1;
    if (!r.ok) return r;
    current = r.pkg;
    steps.push(r.inverse);
    value ||= r.value;
    formula ||= r.formula;
  }
  const inverse: HistoryOp[] = [];
  for (let i = steps.length - 1; i >= 0; i -= 1) for (const op of steps[i]!) inverse.push(op);
  const effects: WorkbookOp[] = [];
  {
    // An undo can put back a relationship or override of a chain that an edit dropped.
    const dropped = formula ? withoutCalcChain(current) : withoutDanglingCalcChain(current);
    if (dropped) {
      current = dropped;
      effects.push({ op: 'dropCalcChain' });
    }
  }
  // The flag check comes first: it is one attribute, and the formula scan reads whole parts.
  if (value && !workbookIndex(current)?.fullCalcOnLoad && workbookHasFormula(current)) {
    const flagged = withFullCalcOnLoad(current);
    if (flagged) {
      current = flagged;
      effects.push({ op: 'setFullCalcOnLoad' });
    }
  }
  return { ok: true, pkg: current, inverse, effects };
}

/** Whether a patch is valid, so its op can join a run of applyFormat ops. */
const validPatch = (patch: FormatPatch): boolean => patchProblem(patch) === undefined;

function applyOp(pkg: XlsxPackage, op: HistoryOp, openedParts?: ReadonlySet<string>): OpResult {
  if (op.op === 'setFullCalcOnLoad' || op.op === 'dropCalcChain') {
    const next = op.op === 'setFullCalcOnLoad' ? withFullCalcOnLoad(pkg) : withoutCalcChain(pkg);
    return { ok: true, pkg: next ?? pkg, inverse: [], value: false, formula: false };
  }
  if (op.op === 'clearValues') return clearRange(pkg, op.sheet, op.ref);
  if (op.op === 'applyFormat') return applyFormatRange(pkg, op.sheet, op.ref, op.patch);
  if (op.op === 'clearFormats') return clearFormatsRange(pkg, op.sheet, op.ref);
  if (op.op === 'setCellFormatIndex') return setCellFormatIndexes(pkg, [op]);
  if (op.op === 'restoreStyles') return restoreStyles(pkg, op.part);
  if (op.op === 'restoreParts') return restoreParts(pkg, op.parts);
  if (
    op.op === 'insertRows' ||
    op.op === 'deleteRows' ||
    op.op === 'insertColumns' ||
    op.op === 'deleteColumns'
  )
    return applyStructure(pkg, op);
  if (
    op.op === 'addSheet' ||
    op.op === 'renameSheet' ||
    op.op === 'deleteSheet' ||
    op.op === 'moveSheet'
  )
    return applySheetOp(pkg, op, openedParts);
  if (op.op === 'setColumnWidth' || op.op === 'setRowHeight') return applyAxisSize(pkg, op);
  if (op.op === 'mergeCells' || op.op === 'unmergeCells') return applyMerge(pkg, op);
  if (op.op === 'setColumnWidths') return applyColumnWidths(pkg, op);
  if (op.op === 'setFittedRowHeights') return applyFittedRowHeights(pkg, op);
  const target = sheetTarget(pkg, op.sheet);
  if ('code' in target) return { ok: false, rejection: target };
  if (op.op === 'setDimension') {
    if (!dimensionBounds(op.ref))
      return refuse('bad-reference', `${JSON.stringify(op.ref)} is not a range in the grid`);
    const r = withDimensionRef(target.part, op.ref);
    if (!r) return unchanged(pkg);
    return {
      ok: true,
      pkg: withPart(pkg, { kind: 'xml', partName: r.part.partName, xml: r.part }),
      inverse: [{ op: 'setDimension', sheet: op.sheet, ref: r.old }],
      value: false,
      formula: false,
    };
  }
  if (op.op === 'restoreRow') {
    if (!Number.isInteger(op.row) || op.row < 1 || op.row > MAX_ROWS)
      return refuse('bad-reference', `row ${op.row} is outside the grid`);
    const ids = new IdSource(target.part);
    const slot = rowSlot(target.sheetData, op.row - 1);
    // A row without `@r` is the row after the one before it, so it fits only where that holds.
    // `@r` is an xsd:unsignedInt, so `05` is row 5: compare numbers, not strings.
    const nodeRow = op.node ? attr(op.node, 'r') : undefined;
    const fits =
      op.node === null ||
      (op.node.kind === 'row' &&
        op.node.ns === NS.x &&
        op.node.local === 'row' &&
        (Number(nodeRow) === op.row || (nodeRow === undefined && slot.prevRow === op.row - 2)));
    if (!fits) return refuse('bad-value', `restoreRow ${op.row} needs a row element for that row`);
    const before = slot.el;
    // A removed row keeps the numbers of the rows after it (`pinRowRefs`).
    const sheetData =
      slot.el === null
        ? withChild(target.sheetData, slot.index, op.node, true)
        : withChild(
            op.node === null ? pinRowRefs(target.sheetData) : target.sheetData,
            slot.index,
            op.node
          );
    const part = withSheetData(target.part, sheetData, ids);
    const inverse: HistoryOp = { op: 'restoreRow', sheet: op.sheet, row: op.row, node: before };
    return {
      ok: true,
      pkg: withPart(pkg, { kind: 'xml', partName: part.partName, xml: part }),
      inverse: [inverse],
      ...rowChange(before, op.node),
    };
  }
  const address = parseCellRef(op.ref);
  if (!address)
    return refuse('bad-reference', `${JSON.stringify(op.ref)} is not a cell in the grid`);
  const value = op.value;
  if (value.type === 'number' && !Number.isFinite(value.value))
    return refuse('bad-value', 'a cell number must be finite');
  if (value.type === 'date' && !isIsoDate(value.iso))
    return refuse('bad-value', `${op.ref}: ${JSON.stringify(value.iso)} is not an ISO 8601 date`);
  if (value.type === 'string' && value.value.length > MAX_CELL_TEXT)
    return refuse(
      'text-too-long',
      `${op.ref}: the text has ${value.value.length} UTF-16 code units, more than a cell holds (${MAX_CELL_TEXT})`
    );
  let formulaText: string | undefined;
  if (value.type === 'formula') {
    const read = readFormulaInput(value.text, formulaScope(pkg));
    if (!read.ok)
      return {
        ok: false,
        rejection: {
          code: formulaRejection(read.problem),
          detail: `${op.ref}: the formula text cannot be read (${read.problem} at ${read.at})`,
          // The function name that `unsupported-function` names, for user text.
          ...(read.subject !== undefined ? { subject: read.subject } : {}),
        },
      };
    // `x:f` is an `ST_Xstring`: a literal `_xHHHH_` in the text keeps its meaning as text.
    formulaText = encodeXstring(read.text);
  }
  // A child that the edit can neither keep nor replace refuses the edit (`cellChildBlocker`).
  const present = sheetIndex(target.part).cells.get(cellKey(address.row, address.col));
  const childBlocker = present && cellChildBlocker(op.sheet, op.ref, present.el);
  if (childBlocker) return { ok: false, rejection: childBlocker };
  let next = pkg;
  let sstIndex: number | undefined;
  const inArray = arrayFormulaAt(target.part, address.row, address.col);
  if (inArray)
    return {
      ok: false,
      rejection: {
        code: 'shared-formula',
        detail: `${op.ref} is inside the array or data table range ${inArray}`,
        subject: `${op.sheet}!${inArray}`,
      },
    };
  // Overwriting the master of a shared formula first writes each child as its own formula.
  const unshared: HistoryOp[] = [];
  const master = sheetIndex(target.part).cells.get(cellKey(address.row, address.col));
  // The master of an array or data table formula, or any formula range master inside
  // `mc:AlternateContent`, holds a definition for its whole range: replacing it is refused.
  const blocker = master ? rangeMasterBlocker(master.el) : undefined;
  if (blocker)
    return {
      ok: false,
      rejection: {
        code: 'shared-formula',
        detail: `${op.ref} ${blocker}`,
        subject: formulaRangeSubject(op.sheet, rangeMasterRef(master!.el), address),
      },
    };
  if (master && isSharedMaster(master.el)) {
    const r = unshareChildren(target.part, op.sheet, [master]);
    if ('code' in r) return { ok: false, rejection: r };
    next = withPart(next, { kind: 'xml', partName: r.part.partName, xml: r.part });
    for (const inv of r.inverse) unshared.push(inv);
  }
  let sstUndo: HistoryOp | null = null;
  if (value.type === 'string') {
    const r = sharedStringIndex(next, target.index, value.value);
    if ('code' in r) return { ok: false, rejection: r };
    sstUndo = sharedStringsUndo(next, target.index, r.pkg);
    next = r.pkg;
    sstIndex = r.index;
  }
  // The worksheet part may be the same object after a shared string append; re-read it.
  const sheet = sheetTarget(next, op.sheet) as SheetTarget;
  const ids = new IdSource(sheet.part);
  const slot = rowSlot(sheet.sheetData, address.row);
  const rowEl =
    slot.el ?? smlElement(ids, 'row', 'sheetData', [plainAttr('r', String(address.row + 1))]);
  // `spans` is a hint that a new cell can make wrong; dropping it is always valid.
  // A value edit fits a row without `@customHeight` to its content again.
  let newRow = withoutFittedHeight(withAttr(rowEl, 'spans', undefined));
  const cellAt = cellSlot(newRow, address.col);
  const oldCell = cellAt.el;
  const cell = buildCell(
    ids,
    op.ref,
    oldCell,
    value,
    sstIndex,
    formulaText,
    inheritedFormat(
      sheetIndex(sheet.part),
      address.row,
      address.col,
      styleLists(pkg, target.index).cellXfs.length
    )
  );
  // A write that leaves the cell as it was changes nothing: no row, revision, or undo entry.
  if (cell === null ? oldCell === null : oldCell !== null && sameNode(oldCell, cell))
    return unchanged(pkg);
  newRow =
    cell === null
      ? oldCell
        ? // The cells after the removed one keep their columns (`pinCellRefs`).
          withChild(pinCellRefs(newRow, address.row), cellAt.index, null)
        : newRow
      : oldCell
        ? withChild(newRow, cellAt.index, cell)
        : withChild(newRow, cellAt.index, cell, true);
  // A clear that leaves a bare row drops the row, as the batch path does; the rows after it
  // keep their numbers (`pinRowRefs`).
  const drop = slot.el !== null && cell === null && isBareRow(newRow);
  const sheetData =
    slot.el === null
      ? withChild(sheet.sheetData, slot.index, newRow, true)
      : drop
        ? withChild(pinRowRefs(sheet.sheetData), slot.index, null)
        : withChild(sheet.sheetData, slot.index, newRow);
  let part = withSheetData(sheet.part, sheetData, ids);
  const inverse: HistoryOp[] = [
    { op: 'restoreRow', sheet: op.sheet, row: address.row + 1, node: slot.el },
  ];
  // A new cell outside `x:dimension/@ref` widens it, in the same op and undo step.
  if (!oldCell && cell) {
    const widened = widenDimension(part, address.row, address.col);
    if (widened) {
      part = widened.part;
      inverse.push({ op: 'setDimension', sheet: op.sheet, ref: widened.old });
    }
  }
  return {
    ok: true,
    pkg: withPart(next, { kind: 'xml', partName: part.partName, xml: part }),
    // The edited row first, then the rows that unsharing changed, newest first, and last the
    // removal of a shared string table that the op created.
    inverse: [...inverse, ...unshared.reverse(), ...(sstUndo ? [sstUndo] : [])],
    value: true,
    formula: (oldCell !== null && cellOrNestedHasFormula(oldCell)) || value.type === 'formula',
  };
}
