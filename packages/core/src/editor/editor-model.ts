// The DOM-free part of the workbook editor: the public types, the sheet tabs, the per-sheet view
// that a workbook opens with (`x:sheetView`, `CT_SheetView`, sml.xsd), and the content of the
// active cell.

import type { SubstituteFontFailure } from './substitute-fonts.ts';
import type { GridPoint, GridSelection } from '../contracts/index.ts';
import {
  attr,
  child,
  children,
  formatCellRef,
  parseCellRef,
  sheetIndex,
  sheetMayHoldFormula,
  cellKey,
  inheritedFormat,
  xmlPartOf,
  structureBlocker,
  isIsoDate,
  MAX_FORMAT_NEW_CELLS,
  formatRangeOverflow,
  clearFormatsOverflow,
  type RangeBox,
  readSheetProtection,
  styleLists,
  cellLocked,
  lockedCellIn,
  clearRangeBlocker,
  mergedRefMeeting,
  firstValueLost,
  type CellOverride,
  type SheetProtection,
  type WorkbookProtection,
  type WorkbookIndex,
  type XlsxPackage,
} from '../store/index.ts';
import {
  builtinFormatCode,
  cellFormat,
  cellValue,
  conditionalFormat,
  EN_US,
  sheetHasHiddenRules,
  sheetStale,
  type NumberLocale,
  type ResolvedCellFormat,
} from '../cascade/index.ts';
import { frozenPanes, type FontSource, type TextMeasurer } from '../layout/index.ts';
import type { AccessibleGridLabels } from '../output/index.ts';
import {
  readCellEntry,
  refusalOf,
  type TypedEntry,
  type EditRefusal,
  type EditorTranslate,
} from './cell-input.ts';
import { cellContent } from './cell-content.ts';
import { isWholeSheet } from '../automation/index.ts';
import { builtinOrCode } from './number-format-commands.ts';
import {
  selectionDecimalSteps,
  type ShownGeneralText,
  numberFormatOf,
  NO_DECIMAL_STEPS,
  type DecimalSteps,
  isSingleSelection,
} from './decimal-selection.ts';

export type { EditorTranslate } from './cell-input.ts';
export { cellContent } from './cell-content.ts';

/**
 * One sheet of the workbook, in tab order.
 *
 * @public
 */
export interface EditorSheet {
  /** The position in `x:sheets`. */
  readonly index: number;
  readonly name: string;
  /** Hidden and very hidden sheets have no tab. */
  readonly visible: boolean;
  /**
   * Whether the sheet is a worksheet that the editor can show. A chartsheet, a macro sheet, a
   * dialog sheet, or a sheet whose part is missing is not.
   */
  readonly worksheet: boolean;
  /**
   * Why the workbook refuses a rename or a delete of this sheet (`structure-blocked`), or null.
   * For a sheet of the opened file, it is `sheetEditRefusal`. A sheet that the editor added in
   * this session is refused only when content of the file can refer to it, for example a chart
   * formula that the add made live. `sheetCommandState` disables Rename and Delete with it. The
   * snapshot always sets it; when a host omits it, `sheetCommandState` reads `sheetEditRefusal`.
   */
  readonly editRefusal?: EditRefusal | null;
}

/**
 * How the editor takes input. `edit` (the default) allows every command. `view` refuses every
 * edit command with `view-mode`, and keeps selection, sheet tabs, zoom, and save.
 *
 * @public
 */
export type EditorMode = 'edit' | 'view';

/**
 * Why a workbook did not open.
 *
 * @public
 */
export interface EditorFailure {
  /**
   * `package`: the file is not a workbook the engine reads; `read`: the browser could not read
   * the picked file (the packaged File > Open); `fonts`: the text shaper or the fonts did not load;
   * `no-worksheet`: the workbook opened, but no sheet is a visible worksheet that the editor can
   * show (every sheet is hidden, a chartsheet, or a missing part).
   */
  readonly code: 'package' | 'read' | 'fonts' | 'no-worksheet';
  /** The engine's reason, for logs: never shown as user text. */
  readonly detail?: string;
  /**
   * With the code `package`, the name of the part that did not read, such as
   * `/xl/sharedStrings.xml`. A failure of the whole package, such as a ZIP limit, names no part.
   * The name comes from the file: show it as text only, through `displayName`, which drops
   * control characters and bidirectional format controls and cuts it to 64 characters.
   */
  readonly part?: string;
  /**
   * With the code `package`, the limit of the read that the file passed. A refused file with no
   * `limit` is not a valid package or workbook.
   */
  readonly limit?:
    | 'zip.maxEntries'
    | 'zip.maxTotalBytes'
    | 'zip.maxRatio'
    | 'xml.maxBytes'
    | 'xml.maxElements'
    | 'xml.maxDepth'
    | 'workbook.maxSheetNameLength';
  /**
   * With the code `fonts`, the packaged substitute faces that did not load while other faces
   * did. Only `onFontError` passes such a failure: the workbook opens, and `snapshot().failure`
   * stays `null`.
   */
  readonly substituteFailures?: readonly SubstituteFontFailure[];
}

/**
 * The state of the editor that chrome renders. A new object after each change.
 *
 * @public
 */
export interface EditorSnapshot {
  readonly status: 'empty' | 'loading' | 'ready' | 'failed';
  readonly failure: EditorFailure | null;
  readonly sheets: readonly EditorSheet[];
  /** The position of the shown sheet in `sheets`, or -1. */
  readonly activeSheet: number;
  readonly selection: GridSelection | null;
  /** The selected range as an A1 reference, such as `B2:D2`. */
  readonly selectionRef: string | null;
  /**
   * The active cell's reference and its edit text: formula text, or the value in a form that
   * reads back, such as `2026-11-12` for a date serial or `12%` for 0.12. A text that would read
   * back as another type starts with `'`.
   */
  readonly activeCell: { readonly ref: string; readonly content: string } | null;
  /** 1 is 100%. */
  readonly zoom: number;
  /**
   * The shown values of formulas can be out of date: the stale flag is set, and the
   * shown sheet holds a formula cell. That includes a formula cell that the grid does not show:
   * an `x:c` inside `mc:AlternateContent` (also in a wrapped row), outside a row, or inside
   * another cell.
   */
  readonly stale: boolean;
  /** The cell editor is open over the active cell. */
  readonly editing: boolean;
  /**
   * The text of the open cell edit, as typed, or null when no edit is open. A formula bar shows
   * it in place of `activeCell.content`, so it follows the edit.
   */
  readonly editText: string | null;
  /**
   * The open cell edit has typed input, or holds text that differs from the cell's content.
   * While this is true, redo is refused with `edit-pending`, as the toolbar disables it, and
   * undo cancels the edit.
   */
  readonly editChanged: boolean;
  /**
   * An input method composes in the open cell editor, or in a field that reported it through
   * `setEditText`. Every command that would commit the
   * edit, and undo, is refused with `composing` until the composition ends.
   */
  readonly composing: boolean;
  readonly canUndo: boolean;
  readonly canRedo: boolean;
  /**
   * Why the last edit, save, or command was refused, or null. These events clear it: an edit
   * that the store accepts, a good save, a copy or a cut that succeeds, a change of the
   * selection, the sheet, or the mode, a cancel of the open cell edit, a cell edit that opens
   * (except for `edit-dropped`), and the start or the settle of a load. A new refusal
   * replaces it, and so does `edit-dropped` when a load, `close`, or a closing surface drops
   * an open edit with changed text. When the load settles, `edit-dropped` stays (see
   * `editDropped`). While a load runs, each new refusal is stored with the code `loading`.
   */
  readonly refusal: EditRefusal | null;
  /**
   * An open cell edit with changed text was dropped without a commit (a load, `close`, or a
   * surface that closed while an input method composed). It stays true after the load
   * settles, until the next command after the load, done or refused, an edit that the editor
   * stores, or a selection change, or the next load. A cell edit that opens keeps it. While the load runs, it stays true when a `loading` refusal
   * replaces `edit-dropped` in `refusal`.
   */
  readonly editDropped: boolean;
  /** The number locale of the shown values, which `refusalText` groups a number with. */
  readonly numberLocale: NumberLocale;
  /**
   * The cell format of the active cell, as the grid shows it: the state of the format
   * commands. The Mod+B, Mod+I, and Mod+U toggles read the same format.
   */
  readonly activeFormat: ActiveCellFormat | null;
  /** The decimal steps on the selection, from the stored cells, also while a cell edit is open. */
  readonly decimalSteps: DecimalSteps;
  /**
   * Why the structural commands (insert and delete rows and columns) are refused, or null:
   * `whole-sheet` when the whole sheet is selected, or the refusal of the shown sheet whatever
   * the selection. A command that this does not refuse can still be
   * refused by its own rows or columns, for example with `grid-full`.
   */
  readonly structureRefusal: EditRefusal | null;
  /**
   * Why a format patch is refused for the selection (`range-too-large`: it would create more
   * than `MAX_FORMAT_NEW_CELLS` cells), or null.
   */
  readonly formatRefusal: EditRefusal | null;
  /**
   * The shown sheet has conditional formatting that the grid can leave out, for one of these
   * causes:
   * - a rule kind or operand that the engine does not evaluate, or a malformed rule;
   * - rules past the read limits;
   * - more than 64 shown rules or 1,024 shown ranges, so the caps of one cell or one row can
   *   take effect;
   * - rules in an extension list (`x:extLst`);
   * - a layout of this sheet that ran out of its check budget. This cause depends on the window
   *   that was shown; once set, it stays until the shown sheet changes or another workbook
   *   opens, and a return to the sheet starts without it.
   * The save keeps every rule. Chrome shows a notice.
   */
  readonly conditionalHidden: boolean;
  /**
   * The format state that the grid shows at the active cell when a true conditional formatting
   * rule changes the state of a format control there, or null. A rule that changes only what
   * no control shows (strike, `x:vertAlign`, or the underline style) gives null, though the
   * grid paints it. `activeFormat` stays the cell format, which the format commands read and
   * set, so `toolbarCommandState` gives a hint to each control whose state differs here.
   * Rich text runs are not read, as `activeFormat` does not read them. Null after a layout of
   * the sheet ran out of its check budget, because the grid can then show the cell format.
   */
  readonly conditionalFormat: ActiveCellFormat | null;
  /**
   * Why clearing the formats of the selection is refused (`clear-too-large`: it would create
   * more than `MAX_FORMAT_NEW_CELLS` cells with `s="0"` where a row or column format shows),
   * or null. Clear formats and Clear all read it.
   */
  readonly clearRefusal: EditRefusal | null;
  /**
   * The actions that the protection of the shown sheet allows (`x:sheetProtection`), or null
   * when the sheet is not protected. A locked cell of a protected sheet refuses a value edit
   * with `locked`.
   */
  readonly protection: SheetProtection | null;
  /**
   * Whether the sheet is protected and the active cell is locked, so typing in it, or a value
   * edit of it alone, is refused with `locked`. False on a sheet that is not protected.
   */
  readonly activeCellLocked: boolean;
  /**
   * Whether the sheet is protected and clearing the contents of the selection would change a
   * locked cell: the active cell of a one-cell selection, or a cell of the range that exists
   * and holds content. Clear contents and Clear all are then refused with `locked`. False on a
   * sheet that is not protected.
   */
  readonly selectionContentLocked: boolean;
  /**
   * The first cell of the selection, as `Sheet!A1`, whose value a clear can neither keep nor
   * replace, or null: Clear contents and Clear all are then refused with
   * `cell-alternate-content`. A cell with no value in any form is not named.
   */
  readonly selectionContentBlocked: string | null;
  /**
   * The formula range, as `Sheet!A1:B2`, that a clear of the selection would break, or null:
   * Clear contents and Clear all are then refused with `shared-formula`. It is the subject of
   * that refusal: the formula range, or the blocking cell when that range cannot be read or does
   * not hold the blocking cell. It comes from
   * the checks that the clear runs before it changes a cell (`clearRangeBlocker`). A clear that
   * would move the formula of a shared child out of the grid is refused only when it runs, so
   * this field is null for it. Null when `selectionContentBlocked` names a cell.
   */
  readonly selectionFormulaRangeBlocked: string | null;
  /**
   * The cell, as `Sheet!A2`, whose formula the editor cannot rewrite, when it is why a clear of
   * the selection would break the shared formula range in `selectionFormulaRangeBlocked` (the
   * `causeCell` of that refusal), or null. The disabled reason names it.
   */
  readonly selectionFormulaRangeCauseCell: string | null;
  /**
   * True when the formula of `selectionFormulaRangeCauseCell` cannot be read (the
   * `causeUnreadable` of that refusal), or null. The disabled reason then says so.
   */
  readonly selectionFormulaRangeCauseUnreadable: true | null;
  /**
   * The shared formula master in the selection, as `Sheet!A1`, that the clear would change, when
   * `selectionFormulaRangeCauseCell` is set (the `masterCell` of that refusal), or null. The
   * disabled reason names it as the cell to leave out.
   */
  readonly selectionFormulaRangeMasterCell: string | null;
  /**
   * Whether the selection meets a merged range of the shown sheet: a merge of the selection is
   * then refused with `merge-overlap`, and an unmerge changes it. `mergeCommandState` reads it.
   */
  readonly selectionMerged: boolean;
  /**
   * Whether the selection is one cell, or exactly one merged range. A pick of the number format
   * list that shows the current entry then keeps the format, and a decimal step reason speaks
   * of this cell, not of a selected cell.
   */
  readonly selectionSingle: boolean;
  /**
   * The first cell of the selection, in reading order, that a merge would empty (a cell other
   * than the top-left cell that holds a value or a formula), or null. A host reads it to ask the
   * user before it calls `mergeCells({ dropValues: true })`. An open cell edit counts as it would
   * be after its commit, which the merge makes first: text that is not empty is a value. Null
   * also when the selection is one cell or meets a merged range: the merge then refuses for that
   * reason (`mergeCommandState`).
   */
  readonly selectionDropsValuesAt: string | null;
  /**
   * What the workbook protection allows (`x:workbookProtection`), or null when the workbook
   * has none. When `allowStructure` is false, `addSheet`, `renameSheet`, `deleteSheet`, and
   * `moveSheet` are refused with `locked` (rule `workbook`), and the sheet tabs disable them.
   */
  readonly workbookProtection: WorkbookProtection | null;
  /**
   * Why a rename or a delete of each sheet of the opened file is refused, or null: macros, a
   * chart, a pivot table, or a sheet that is not a worksheet in the workbook, or a workbook or
   * package relationship outside the allowlist (`structure-blocked`). A sheet that the editor
   * added in this session has its own reason in `EditorSheet.editRefusal`, which the sheet menu
   * reads (`sheetCommandState`). When its `about` kind is `macros`, it holds for every sheet, and
   * Add and the moves are disabled too. A rename or a delete that this does not refuse can still
   * be refused by its own sheet.
   */
  readonly sheetEditRefusal: EditRefusal | null;
  /** `edit`, or `view` when every edit command is refused (`setMode`). */
  readonly mode: EditorMode;
  /**
   * The count of `setMode` calls that applied a mode, also the mode that the editor has; a
   * refused call does not count. It is host coordination, kept public on purpose: an adapter
   * that waits to apply a mode of its own (the React root's `mode` prop) reads in it that a
   * host call applied a newer mode, and ends its wait. An internal accessor would hide that
   * rule from a host that builds its own adapter.
   */
  readonly modeRequests: number;
  /**
   * The content revision of the open workbook: 0 after an open, and one more after each change
   * of its content (an edit, an undo, or a redo). A view change (selection, zoom, sheet, mode,
   * an open cell edit) keeps it. Compare it to tell a content change from a view change, for
   * example to track unsaved edits: keep the revision of the last save.
   */
  readonly revision: number;
  /**
   * The kind of the open package by its main part: `workbook` (`.xlsx`) or `template`
   * (`.xltx`), or null with no workbook open. `save()` keeps the kind, so a host names the saved
   * file with the matching extension (`.xlsm` and `.xltm` when `macroEnabled`).
   */
  readonly workbookKind: WorkbookKind | null;
  /**
   * Whether the open package is a macro-enabled workbook (`.xlsm`) or template (`.xltm`); false
   * with no workbook open. No macro runs: `save()` keeps the macro part as bytes, and the content
   * type. Name the saved file with the macro-enabled extension, or it does not open cleanly.
   */
  readonly macroEnabled: boolean;
  /**
   * The parts of the open workbook that passed an XML part limit and stay as bytes, by part
   * name, or an empty list. The list keeps its identity while it does not change. Chrome shows
   * a notice that names them.
   */
  readonly partsPastLimit: readonly PartPastLimit[];
  /**
   * The color scheme of the workbook theme as `RRGGBB`, in theme index order: `lt1`, `dk1`,
   * `lt2`, `dk2`, then the six accents (`a:clrScheme`, `CT_ColorScheme`, ECMA-376 Part 1,
   * 20.1.6.2). An entry that the theme lacks is undefined. The list is empty with no workbook
   * open or no theme part. The color picker reads it through `workbookThemeHexes`.
   */
  readonly themeColors: readonly (string | undefined)[];
}

/**
 * The kind of a package by the content type of its main part.
 *
 * @public
 */
export type WorkbookKind = 'workbook' | 'template';

/**
 * A part that passed an XML part limit and stays in the file as bytes. The workbook opens
 * without it, and a save keeps its bytes.
 *
 * @public
 */
export interface PartPastLimit {
  /**
   * The part name, such as `/xl/styles.xml`. It comes from the file: show it as text only,
   * through `displayName`.
   */
  readonly part: string;
  /**
   * `styles`: every cell shows the default format. `theme`: the theme colors and fonts do not
   * apply. `drawing`: the pictures of the drawing do not show. `relationships`: the parts that
   * the relationship part links to do not show. `other`: a styles or theme part that the
   * workbook part does not link to, so nothing that the workbook shows depends on it.
   */
  readonly kind: 'styles' | 'theme' | 'drawing' | 'relationships' | 'other';
  /** The limit that the part passed. */
  readonly limit: 'xml.maxBytes' | 'xml.maxElements' | 'xml.maxDepth';
}

/**
 * The cell format of the active cell, as the format commands read and set it.
 *
 * @public
 */
export interface ActiveCellFormat {
  readonly bold: boolean;
  readonly italic: boolean;
  /** Any underline. */
  readonly underline: boolean;
  /** The font family, with a theme font resolved to its typeface. */
  readonly fontName: string;
  /** Points. */
  readonly fontSize: number;
  /** `RRGGBB`. */
  readonly fontColor: string;
  /** The color of a solid fill as `RRGGBB`, or `null` for no fill, a pattern, or a gradient. */
  readonly fillColor: string | null;
  /** `ST_HorizontalAlignment`; `general` aligns by value type. */
  readonly horizontal: string;
  /** `ST_VerticalAlignment`. */
  readonly vertical: string;
  readonly wrapText: boolean;
  /**
   * The number format id and its code in the editor's locale. `hasCode` is false for an id with
   * no code here (a built-in id with no locale or ECMA-376 code, or a custom id with no
   * `x:numFmt` record): such a cell shows `General`.
   */
  readonly numberFormat: { readonly id: number; readonly code: string; readonly hasCode: boolean };
}

/**
 * How an editor opens workbooks.
 *
 * @public
 */
export interface XlsxEditorOptions {
  /**
   * The faces that measure and paint text, or a function that loads them once. Omitted, the
   * packaged substitute faces load (`loadSubstituteFonts()`), as in DOCX Editor and PPTX
   * Editor. The engine never fetches a font that a workbook names.
   */
  readonly fonts?: readonly FontSource[] | (() => Promise<readonly FontSource[]>);
  /** The number locale of the shown values. US English separators when absent. */
  readonly numberLocale?: NumberLocale;
  /**
   * Translates the labels that the grid surface shows to assistive technology. English, from
   * the shared catalog, when absent.
   */
  readonly translate?: EditorTranslate;
  /** `edit` (the default) or `view`. Change it later with `setMode`. */
  readonly mode?: EditorMode;
}

/**
 * The snapshot of an editor with no workbook.
 *
 * @public
 */
export const EMPTY_SNAPSHOT: EditorSnapshot = Object.freeze({
  status: 'empty',
  failure: null,
  sheets: [],
  activeSheet: -1,
  selection: null,
  selectionRef: null,
  activeCell: null,
  zoom: 1,
  stale: false,
  editing: false,
  editText: null,
  editChanged: false,
  composing: false,
  canUndo: false,
  canRedo: false,
  refusal: null,
  editDropped: false,
  numberLocale: EN_US,
  activeFormat: null,
  decimalSteps: NO_DECIMAL_STEPS,
  structureRefusal: null,
  formatRefusal: null,
  conditionalHidden: false,
  conditionalFormat: null,
  clearRefusal: null,
  protection: null,
  activeCellLocked: false,
  selectionContentLocked: false,
  selectionContentBlocked: null,
  selectionFormulaRangeBlocked: null,
  selectionFormulaRangeCauseCell: null,
  selectionFormulaRangeCauseUnreadable: null,
  selectionFormulaRangeMasterCell: null,
  selectionMerged: false,
  selectionSingle: false,
  selectionDropsValuesAt: null,
  workbookProtection: null,
  sheetEditRefusal: null,
  mode: 'edit',
  modeRequests: 0,
  revision: 0,
  workbookKind: null,
  macroEnabled: false,
  themeColors: [],
  partsPastLimit: [],
});

/** The zoom limits of `x:sheetView/@zoomScale` (10 to 400%). */
export const MIN_ZOOM = 0.1;
export const MAX_ZOOM = 4;

export const clampZoom = (z: number): number =>
  Number.isFinite(z) ? Math.min(MAX_ZOOM, Math.max(MIN_ZOOM, Math.round(z * 100) / 100)) : 1;

/**
 * The zoom steps of the zoom commands. A fit uses the steps from 50% (`FIT_ZOOMS`,
 * `grid-fit.ts`).
 */
export const ZOOM_STEPS: readonly number[] = [
  0.1, 0.25, 0.5, 0.75, 0.9, 1, 1.1, 1.25, 1.5, 1.75, 2, 2.5, 3, 4,
];

/** The zoom one step in (+1) or out (-1) from `z`, held at the first or the last step. */
export const zoomStepFrom = (z: number, direction: 1 | -1): number =>
  direction > 0
    ? (ZOOM_STEPS.find((s) => s > z + 0.001) ?? ZOOM_STEPS[ZOOM_STEPS.length - 1]!)
    : ([...ZOOM_STEPS].reverse().find((s) => s < z - 0.001) ?? ZOOM_STEPS[0]!);

/** The sheet a workbook opens on: `activeTab` when it can show, else the first that can. */
export function openingSheet(sheets: readonly EditorSheet[], activeTab: number): number {
  const ok = (s: EditorSheet | undefined) => !!s && s.visible && s.worksheet;
  if (ok(sheets[activeTab])) return activeTab;
  return sheets.findIndex(ok);
}

/** What a sheet opens with: its zoom, its active cell, and its top-left cell. */
export interface SheetOpening {
  readonly zoom: number;
  readonly active: GridPoint;
  readonly topLeft: GridPoint;
  readonly headings: boolean;
}

/**
 * The view of a sheet that the file states: the first `x:sheetView` (`CT_SheetView`, sml.xsd,
 * ECMA-376 Part 1, 18.3.1.87). It reads `@zoomScale` (default 100), `@topLeftCell`,
 * `@showRowColHeaders` (default true), and `x:selection/@activeCell` (`CT_Selection`, 18.3.1.78).
 * With a frozen `x:pane` (`CT_Pane`, 18.3.1.66), the scroll comes from `x:pane/@topLeftCell`, the
 * top-left visible cell of the bottom-right pane, and the active cell from the selection whose
 * `@pane` is `x:pane/@activePane` (`ST_Pane`, 18.18.52; both default `topLeft`).
 */
export function sheetOpening(pkg: XlsxPackage, partName: string | undefined): SheetOpening {
  const part = xmlPartOf(pkg, partName);
  const view = part ? child(child(part.root, 'x:sheetViews'), 'x:sheetView') : undefined;
  const zoomRaw = Number(attr(view, 'zoomScale') ?? '100');
  const zoom = Number.isFinite(zoomRaw) && zoomRaw >= 10 && zoomRaw <= 400 ? zoomRaw / 100 : 1;
  const pane = child(view, 'x:pane');
  // A pane that freezes no row or column, and a split pane, do not show: the editor shows the
  // view's top-left cell and the selection of the top-left pane.
  const band = part ? frozenPanes(part.root) : { rows: 0, cols: 0 };
  const frozen = band.rows > 0 || band.cols > 0;
  const shown = frozen ? (attr(pane, 'activePane') ?? 'topLeft') : 'topLeft';
  const sels = children(view, 'x:selection');
  const sel = sels.find((s) => (attr(s, 'pane') ?? 'topLeft') === shown) ?? sels[0];
  const cell = (raw: string | undefined) => (raw ? parseCellRef(raw.split(':')[0]!) : null);
  const point = (a: { row: number; col: number } | null): GridPoint =>
    a ? { row: a.row, col: a.col } : { row: 0, col: 0 };
  const paneTop = frozen ? cell(attr(pane, 'topLeftCell')) : null;
  const headings = attr(view, 'showRowColHeaders');
  return {
    zoom,
    active: point(cell(attr(sel, 'activeCell'))),
    topLeft: point(paneTop ?? cell(attr(view, 'topLeftCell'))),
    headings: headings === undefined || headings === '1' || headings === 'true',
  };
}

/** The A1 reference of a selection's range: one cell, or `start:end`. */
export function selectionRef(sel: GridSelection): string {
  const a = formatCellRef(sel.range.start);
  const b = formatCellRef(sel.range.end);
  return a === b ? a : `${a}:${b}`;
}

/**
 * The font of a cell's format, without conditional formatting (the cell editor shows the cell
 * format, design D4 of `workbook-conditional-formatting`): the face the measurer resolves for
 * the cell format's font (for a cell with no `x:c` element, the row or column format's font),
 * as a CSS family list of its registered id, and the size in points.
 * The id keeps letters, digits, `_`, and `-` only, as paint does, so no file text reaches CSS.
 * With no workbook or sheet, the default face at 11 points.
 */
export function cellFont(
  opened: { readonly pkg: XlsxPackage; readonly index: WorkbookIndex } | null,
  active: number,
  at: GridPoint,
  measurer: TextMeasurer | null
): { readonly family: string; readonly sizePt: number } {
  const shown = opened?.index.sheets[active];
  if (!opened || !shown) return { family: 'sans-serif', sizePt: 11 };
  const { pkg, index } = opened;
  const part = xmlPartOf(pkg, shown.partName);
  const sheet = part ? sheetIndex(part) : undefined;
  const entry = sheet?.cells.get(cellKey(at.row, at.col));
  // A cell that does not exist yet shows, and is typed in, the font of its row or column format.
  const s = entry || !sheet ? attr(entry?.el, 's') : String(inheritedFormat(sheet, at.row, at.col));
  const font = cellFormat(pkg, index, s).format.font;
  const face = measurer?.resolve({ family: font.family, bold: font.bold, italic: font.italic });
  const id = (face?.id ?? '').replace(/[^A-Za-z0-9_-]/g, '');
  return { family: id ? `${id}, sans-serif` : 'sans-serif', sizePt: font.sizePt };
}

/** Whether a cell stores an ISO 8601 date (`t="d"`). */
export function cellIsDate(pkg: XlsxPackage, partName: string | undefined, at: GridPoint): boolean {
  const part = xmlPartOf(pkg, partName);
  const entry = part ? sheetIndex(part).cells.get(cellKey(at.row, at.col)) : undefined;
  return entry !== undefined && attr(entry.el, 't') === 'd';
}

/**
 * Text read as typed input for a cell, as the cell editor reads it (`readCellEntry`). ISO 8601
 * text in a date cell (`t="d"`) stays a date. In a cell with the Text number format (`@`,
 * Part 1, 18.8.31), an entry stays text. A formula, an empty entry, and an entry with a leading
 * `'` are the exceptions; the `'` is not stored. A typed form such as `12%` implies a number
 * format. Only a cell with the General format keeps that format; a cell with another format
 * keeps its own. A plain-text paste reads each cell the same way.
 */
export function typedInput(
  pkg: XlsxPackage,
  index: WorkbookIndex,
  partName: string | undefined,
  at: GridPoint,
  text: string,
  locale: NumberLocale
): TypedEntry {
  if (cellIsDate(pkg, partName, at) && isIsoDate(text))
    return { value: { type: 'date', iso: text } };
  const entry = readCellEntry(text, locale, { date1904: index.date1904 });
  // The format as the file states it. A built-in id other than 0 (General) and 49 (`@`) is
  // another format, also an id with no code in the locale.
  const { numFmtId, formatCode } = resolvedFormatAt(pkg, index, partName, at);
  const code = formatCode ?? (numFmtId === 0 ? 'General' : numFmtId === 49 ? '@' : undefined);
  if (code === '@') {
    const kind = entry.value.type;
    if (kind === 'formula' || kind === 'empty' || text.charCodeAt(0) === 39 /* ' */)
      return { value: entry.value };
    return { value: { type: 'string', value: text } };
  }
  if (entry.numberFormat === undefined || code?.toLowerCase() === 'general')
    return localFormat(pkg, index, entry, locale);
  return { value: entry.value };
}

/**
 * The entry, with the locale code in place of a built-in id when the workbook defines that id
 * with another code. The file code wins for a built-in id, so the store gives the entry a custom
 * id (`lowerNumberFormat`), and the cell shows and copies the typed form in the locale order.
 * The number format list and the decimal steps apply the same rule (`builtinOrCode`).
 */
function localFormat(
  pkg: XlsxPackage,
  index: WorkbookIndex,
  entry: TypedEntry,
  locale: NumberLocale
): TypedEntry {
  const format = entry.numberFormat;
  if (format?.id === undefined) return entry;
  const code = builtinFormatCode(format.id, locale);
  if (code === undefined) return entry;
  return { ...entry, numberFormat: builtinOrCode(format.id, code, styleLists(pkg, index).numFmts) };
}

/** The part of a snapshot that the shown sheet and its selection give. */
export type SheetState = Pick<
  EditorSnapshot,
  | 'activeCell'
  | 'activeFormat'
  | 'decimalSteps'
  | 'stale'
  | 'structureRefusal'
  | 'formatRefusal'
  | 'conditionalHidden'
  | 'conditionalFormat'
  | 'clearRefusal'
  | 'protection'
  | 'activeCellLocked'
  | 'selectionContentLocked'
  | 'selectionContentBlocked'
  | 'selectionFormulaRangeBlocked'
  | 'selectionFormulaRangeCauseCell'
  | 'selectionFormulaRangeCauseUnreadable'
  | 'selectionFormulaRangeMasterCell'
  | 'selectionMerged'
  | 'selectionSingle'
  | 'selectionDropsValuesAt'
>;

/** The sheet state with no sheet shown. */
export const NO_SHEET_STATE: SheetState = Object.freeze({
  activeCell: null,
  activeFormat: null,
  decimalSteps: NO_DECIMAL_STEPS,
  stale: false,
  structureRefusal: null,
  formatRefusal: null,
  conditionalHidden: false,
  conditionalFormat: null,
  clearRefusal: null,
  protection: null,
  activeCellLocked: false,
  selectionContentLocked: false,
  selectionContentBlocked: null,
  selectionFormulaRangeBlocked: null,
  selectionFormulaRangeCauseCell: null,
  selectionFormulaRangeCauseUnreadable: null,
  selectionFormulaRangeMasterCell: null,
  selectionMerged: false,
  selectionSingle: false,
  selectionDropsValuesAt: null,
});

/** The selected range in the store's range shape. */
const boxOf = ({ range: { start, end } }: GridSelection): RangeBox => [
  start.row,
  start.col,
  end.row,
  end.col,
];

/** The snapshot state of the shown sheet: the active cell, its format, and the refusals. */
export function sheetState(
  { pkg, index }: { readonly pkg: XlsxPackage; readonly index: WorkbookIndex },
  entry: { readonly name: string; readonly partName: string | undefined },
  selection: GridSelection,
  locale: NumberLocale,
  edit: CellOverride | null = null,
  shown?: ShownGeneralText
): SheetState {
  const at = selection.active;
  const part = xmlPartOf(pkg, entry.partName);
  const protection = part ? readSheetProtection(part.root) : null;
  const format = resolvedFormatAt(pkg, index, entry.partName, at);
  const merged = part !== undefined && meetsMerge(part.root, selection);
  // The checks that a clear runs before it changes a cell (`clearRangeBlocker`).
  const cleared = part ? clearRangeBlocker(part, entry.name, boxOf(selection)) : undefined;
  return {
    activeCell: {
      ref: formatCellRef(at),
      content: cellContent(pkg, index, entry.partName, at, locale),
    },
    activeFormat: activeFormatOf(format, locale),
    decimalSteps: selectionDecimalSteps(pkg, index, entry.partName, selection, locale, shown),
    // The editor adds the layout budget cause and clears `conditionalFormat` after a cut. The
    // active cell of a merged range is its top-left cell (`selectCell`).
    conditionalHidden: part ? sheetHasHiddenRules(part.root) : false,
    conditionalFormat: conditionalFormatAt(pkg, index, entry.partName, at, format, locale),
    // The notice is about formula results: a sheet with no formula cell shows none, whatever
    // the stale flag says (design D3 keeps the flag). It counts the formula cells that the
    // stale flag counts, also the ones that the sheet index skips (`sheetMayHoldFormula`). The
    // grid does not show those cells, so for them the notice is the only mark. The sheet index
    // count comes first, because it is already built and the scan reads the whole part.
    stale: part
      ? sheetStale(index, part.root) && (sheetIndex(part).formulas > 0 || sheetMayHoldFormula(part))
      : false,
    // The command refuses a whole-sheet selection before it reads the blockers, so the state
    // gives the same refusal first.
    structureRefusal: isWholeSheet(selection)
      ? { code: 'whole-sheet' }
      : structureRefusalOf(pkg, entry.name),
    formatRefusal: formatRefusalOf(pkg, entry.partName, selection),
    clearRefusal:
      part && clearFormatsOverflow(part.root, sheetIndex(part), boxOf(selection)) !== null
        ? { code: 'clear-too-large' }
        : null,
    protection,
    activeCellLocked:
      protection !== null && part !== undefined && cellLocked(pkg, index, part, at.row, at.col),
    selectionContentLocked:
      protection !== null &&
      part !== undefined &&
      lockedCellIn(pkg, index, part, boxOf(selection)) !== undefined,
    selectionContentBlocked:
      cleared?.code === 'cell-alternate-content' ? (cleared.subject ?? null) : null,
    selectionFormulaRangeBlocked:
      cleared?.code === 'shared-formula' ? (cleared.subject ?? null) : null,
    selectionFormulaRangeCauseCell:
      cleared?.code === 'shared-formula' ? (cleared.causeCell ?? null) : null,
    selectionFormulaRangeCauseUnreadable:
      cleared?.code === 'shared-formula' ? (cleared.causeUnreadable ?? null) : null,
    selectionFormulaRangeMasterCell:
      cleared?.code === 'shared-formula' ? (cleared.masterCell ?? null) : null,
    selectionMerged: merged,
    selectionSingle: isSingleSelection(part?.root, selection),
    // A selection of one cell, or one that meets a merged range, never merges: no scan.
    selectionDropsValuesAt:
      part && !merged && !isOneCell(selection)
        ? (firstValueLost(part, boxOf(selection), edit ?? undefined) ?? null)
        : null,
  };
}

const isOneCell = ({ range: { start, end } }: GridSelection): boolean =>
  start.row === end.row && start.col === end.col;

/** Whether `selection` meets a merged range, as the merge op reads them from `root`. */
function meetsMerge(
  root: Parameters<typeof mergedRefMeeting>[0],
  { range: { start, end } }: GridSelection
): boolean {
  return mergedRefMeeting(root, [start.row, start.col, end.row, end.col]) !== undefined;
}

/**
 * The resolved cell format of a cell: the one source of the format state and the toggles. A
 * cell with no `x:c` element shows the format of its row or column (`inheritedFormat`), the
 * format that the store starts a new cell from, so the state and the command agree.
 */
export function resolvedFormatAt(
  pkg: XlsxPackage,
  index: WorkbookIndex,
  partName: string | undefined,
  at: GridPoint
): ResolvedCellFormat {
  const part = xmlPartOf(pkg, partName);
  if (!part) return cellFormat(pkg, index, undefined).format;
  const sheet = sheetIndex(part);
  const entry = sheet.cells.get(cellKey(at.row, at.col));
  const s = entry ? attr(entry.el, 's') : String(inheritedFormat(sheet, at.row, at.col));
  return cellFormat(pkg, index, s).format;
}

const EMPTY_VALUE = { kind: 'empty' } as const;

/**
 * The format state that the grid shows at a cell when a true conditional formatting rule
 * changes it, or null. Layout applies the rules in the same way: a cell with no `x:c` element
 * is empty, and a formula without a cached value tests no rule. The active cell of a merged
 * range is its top-left cell, whose format the range shows. One cell, once for each snapshot.
 */
export function conditionalFormatAt(
  pkg: XlsxPackage,
  index: WorkbookIndex,
  partName: string | undefined,
  at: GridPoint,
  format: ResolvedCellFormat,
  locale: NumberLocale
): ActiveCellFormat | null {
  const part = xmlPartOf(pkg, partName);
  if (!part) return null;
  const sheet = sheetIndex(part);
  const entry = sheet.cells.get(cellKey(at.row, at.col));
  const read = entry ? cellValue(pkg, index, entry.el, false) : undefined;
  const value = read?.value ?? EMPTY_VALUE;
  if (read?.formula && value.kind === 'empty') return null;
  const shown = conditionalFormat(pkg, index, sheet.root, at.row, at.col, value, format);
  if (shown === format) return null;
  const state = activeFormatOf(shown, locale);
  // An overlay can set only values that the controls do not show, or the values that the cell
  // format already has: then no control differs.
  return JSON.stringify(state) === JSON.stringify(activeFormatOf(format, locale)) ? null : state;
}

/** The public format state of a resolved cell format. */
export function activeFormatOf(format: ResolvedCellFormat, locale: NumberLocale): ActiveCellFormat {
  const { font, fill, alignment } = format;
  return {
    bold: font.bold,
    italic: font.italic,
    underline: font.underline !== 'none',
    fontName: font.family,
    fontSize: font.sizePt,
    fontColor: font.color,
    fillColor: fill.kind === 'solid' ? fill.color : null,
    horizontal: alignment.horizontal,
    vertical: alignment.vertical,
    wrapText: alignment.wrapText,
    numberFormat: numberFormatOf(format, locale),
  };
}

/** The structure refusal of each sheet name, for each package revision. */
const structureRefusals = new WeakMap<XlsxPackage, Map<string, EditRefusal | null>>();

/**
 * Why the structural commands are refused on a sheet whatever the selection (design D4), or
 * null. The check reads every part that can hold a reference, so it runs once for each package
 * revision and sheet, never for each selection change.
 */
export function structureRefusalOf(pkg: XlsxPackage, sheet: string): EditRefusal | null {
  let bySheet = structureRefusals.get(pkg);
  if (!bySheet) structureRefusals.set(pkg, (bySheet = new Map()));
  const known = bySheet.get(sheet);
  if (known !== undefined) return known;
  const rejection = structureBlocker(pkg, sheet);
  const refusal = rejection ? refusalOf(rejection) : null;
  bySheet.set(sheet, refusal);
  return refusal;
}

/**
 * Why a format patch is refused for a selection, or null: `range-too-large` when the patch
 * would create more than `MAX_FORMAT_NEW_CELLS` cells, by the store's own count
 * (`formatRangeOverflow`), so the state and the command refuse the same ranges. The count reads
 * the cells inside the range only when the range is larger than the limit, and never walks the
 * range.
 */
export function formatRefusalOf(
  pkg: XlsxPackage,
  partName: string | undefined,
  selection: GridSelection
): EditRefusal | null {
  const { start, end } = selection.range;
  const area = (end.row - start.row + 1) * (end.col - start.col + 1);
  if (area <= MAX_FORMAT_NEW_CELLS) return null;
  const part = xmlPartOf(pkg, partName);
  if (!part) return null;
  return formatRangeOverflow(sheetIndex(part), boxOf(selection))
    ? { code: 'range-too-large' }
    : null;
}

/** The words that the accessible rows of the grid add to a cell text, from the catalog. */
export function accessibleLabels(t: EditorTranslate): AccessibleGridLabels {
  return {
    staleCell: (text) => t('grid.staleResult', { text }),
    staleNoResult: t('grid.staleNoResult'),
  };
}
