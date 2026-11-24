// The public shape of a workbook editor: the methods that the adapters call. `xlsx-editor.ts`
// implements it.

import type { GridPoint } from '../contracts/index.ts';
import type { FormatPatch } from '../store/index.ts';
import type { FormatToggle, MergeOptions } from '../automation/index.ts';
import type { CellEditMode, EditMove, EditTextOptions } from './cell-editor.ts';
import type { EditResult } from './cell-input.ts';
import type { CopyResult, PasteData } from './clipboard.ts';
import type {
  EditorFailure,
  EditorMode,
  EditorSnapshot,
  XlsxEditorOptions,
} from './editor-model.ts';
import type { NumberLocale } from '../cascade/index.ts';
import type { FontReport } from './substitute-fonts.ts';
import type { EditorEvents } from './editor-events.ts';

/**
 * The result of {@link XlsxEditorInstance.saveInTasks}.
 *
 * @public
 */
export interface SavedWorkbook {
  /** The workbook as bytes of the opened package's kind. */
  readonly bytes: Uint8Array<ArrayBuffer>;
  /** The revision that the bytes hold. A later revision has edits that they do not hold. */
  readonly revision: number;
  /**
   * False when another open, `close()`, or `destroy()` ended this workbook during the save. True
   * does not say that the bytes hold the newest edit: compare `revision` for that.
   */
  readonly stillOpen: boolean;
}

/**
 * A workbook editor instance.
 *
 * @public
 */
export interface XlsxEditorInstance {
  /**
   * Open the bytes of a workbook or a template package, in place of the open one. Resolves with
   * the snapshot after the open, or the failure. An open cell edit is dropped when the open
   * starts, without a commit; one with changed text leaves the refusal `edit-dropped`, which
   * the live region announces. The undo history of the old workbook goes when the open
   * settles. Until then, every command, `save` included, is refused with `loading`, `select`
   * and `setActiveSheet` return false, and `setZoom` does nothing. When the open settles, ready
   * or failed, the other refusals clear, and `edit-dropped` (with `editDropped`) stays until
   * the next command after the open, done or refused, an edit that the editor stores, a
   * selection change, or the next open. `close()`
   * drops an edit in the same way. A host that must not lose typed text asks first while
   * `snapshot().editing` and `snapshot().editChanged` are both true.
   *
   * `bytes` is a `Uint8Array` or an `ArrayBuffer`, also an `ArrayBuffer` of another window. A
   * `SharedArrayBuffer` opens from a copy of its bytes. A detached buffer gives no bytes, so the
   * open resolves with the failure `package`. `readXlsx` of the store entry takes a `Uint8Array`
   * only.
   */
  load(bytes: Uint8Array | ArrayBuffer): Promise<EditorSnapshot>;
  /**
   * Close the open workbook, without a save: the grid, the sheet tabs, and the snapshot show no
   * workbook (`status` `'empty'`), an open cell edit and the undo history are dropped, and a
   * `load` that has not settled yet is dropped too.
   */
  close(): void;
  /**
   * The faces in force, and the packaged substitute faces that did not load (`fonts` omitted).
   * Empty before the first open loads the faces. A failed substitute face does not fail the
   * open: its text measures and paints with another face.
   */
  fontReport(): FontReport;
  /** Mount the grid surface in `element`, replacing its children. */
  attach(element: HTMLElement): void;
  /** Unmount the grid surface and keep the workbook; an open edit closes as in `destroy`. */
  detach(): void;
  /** Called after every change of `snapshot()`. Returns the unsubscribe function. */
  subscribe(listener: () => void): () => void;
  /**
   * Call `handler` each time `event` happens, as DOCX Editor's `on`: a content change, a
   * selection change, a refused command, or a failed open (`EditorEvents`). Returns the
   * unsubscribe function. Nothing fires for the state at the call. Unlike `subscribe`, which
   * runs after every state change, each event fires only when its part of the state moves.
   *
   * @example
   * ```ts
   * const off = editor.on('change', (change) => {
   *   if (!change.source) void upload(editor.save());
   * });
   * ```
   */
  on<E extends keyof EditorEvents>(event: E, handler: EditorEvents[E]): () => void;
  snapshot(): EditorSnapshot;
  /**
   * Show a sheet by its position in `snapshot().sheets`. Hidden sheets and chartsheets refuse. An
   * open cell edit is committed first. While an input method composes, it is refused with
   * `composing`.
   */
  setActiveSheet(index: number): boolean;
  /** Zoom the shown sheet; 1 is 100%. Clamped to 10% to 400%. Nothing while an open runs. */
  setZoom(zoom: number): void;
  /** Zoom one step in (+1) or out (-1). */
  zoomStep(direction: 1 | -1): void;
  /**
   * Select a cell, or a range from `anchor` to `focus`, on the shown sheet. A point outside the
   * grid moves to its edge, and a fractional coordinate is cut to an integer. Returns `false`,
   * and changes nothing, when a coordinate is not a finite number or no sheet is shown: before
   * an open settles, after a failed open, or after `close`. An open cell edit is committed
   * first, as for `setActiveSheet`; when the store refuses that commit, `select` returns
   * `false`, the selection stays, and the editor keeps its text and shows the reason.
   * While an input method composes, it is refused with `composing`.
   */
  select(anchor: GridPoint, focus?: GridPoint): boolean;
  /**
   * Select the whole shown sheet, as Control or Command and A in the grid does. An open cell edit
   * is committed first; when the store refuses that commit, the selection stays.
   */
  selectAll(): void;
  /** Give the keyboard to the grid. */
  focus(): void;
  /**
   * Forget the refusal that the snapshot shows: chrome ended the command that it explains, for
   * example Escape ended a sheet rename whose name was refused. Changes no cell.
   */
  dismissRefusal(): void;
  /**
   * Replace the faces that measure and paint text, in place: the workbook, the selection, and
   * the undo history stay. The new faces load first; when they do not load, the old faces stay
   * and the promise resolves with the failure and its cause. It resolves `null` when the faces
   * apply, or when `fonts` holds the same faces as now (`sameFontSource`). Before the first
   * open, the next open reads them.
   */
  setFonts(fonts: XlsxEditorOptions['fonts']): Promise<EditorFailure | null>;
  /**
   * Replace the number locale of the shown values, in place. `undefined` restores US English
   * separators. Typed and pasted values read in the new locale: its separators, its date order, and
   * the number formats that a typed form takes. An open cell edit
   * keeps the number locale it opened with; the next edit reads the new one. A locale equal in
   * value to the current one (`sameNumberLocale`) changes nothing and publishes nothing.
   */
  setNumberLocale(locale: NumberLocale | undefined): void;
  /**
   * Switch between `edit` and `view` without a reload. In `view`, every edit command is
   * refused with `view-mode`; selection, sheet tabs, zoom, and save stay available. An open
   * cell edit is committed first. When the store refuses that commit, the switch is refused:
   * `setMode` returns `false`, the editor stays open with its text, and the reason shows, as
   * `select` and `save` refuse. Returns `true` when the mode is `mode` after the call.
   * Each call that applies, also with the mode that the editor has, counts in
   * `snapshot().modeRequests` and notifies the listeners, so a listener that calls `setMode`
   * compares `snapshot().mode` first. While an input method composes, a switch to `view`
   * returns `false` with no announcement.
   */
  setMode(mode: EditorMode): boolean;
  /**
   * Open the cell editor over the active cell, with `text`, or with the cell's content when
   * `text` is omitted (F2). `mode` is `enter` (arrow keys commit, as after typing) or `edit`
   * (arrow keys move the caret, as after F2); the default is `edit` without `text` and `enter`
   * with it. Needs a mounted surface, and refuses `not-editable` on a sheet without a part.
   * A cell that no commit could change opens no editor: a locked cell of a protected sheet
   * refuses with `locked`, and a cell that holds a second form of its value with
   * `cell-alternate-content` and the sheet and cell as the subject. A cell of an array or data
   * table range, or a shared master that cannot unshare, refuses with `shared-formula`, as a
   * clear of that cell does. A shared master whose child formula would move out of the grid opens
   * an editor, and its commit is refused with `shared-formula`. An edit that opens clears an
   * earlier refusal, except `edit-dropped`.
   */
  beginEdit(text?: string, mode?: CellEditMode): EditResult;
  /**
   * Commit the open cell editor. A refused commit keeps the editor open with its text, and the
   * keyboard where it was: in the cell editor, or in a field outside the grid, such as the
   * formula bar. While an input method composes, it is refused with `composing`. When an open
   * edit commits, the grid takes the keyboard, with or without `move`.
   *
   * With `move`, the active cell then moves one cell, as Enter (`down`), Shift+Enter (`up`),
   * Tab (`right`), and Shift+Tab (`left`) move it in the cell editor: inside a multi-cell
   * selection, it moves in the selection and the selection stays. With `move` and no open
   * edit, only the active cell moves, and the grid takes the keyboard. So a field such as the
   * formula bar maps each of these keys to one call, with or without an open edit. A `move`
   * that is not an `EditMove` throws a `TypeError` before anything changes. With no open edit
   * and no workbook ready, it is refused with `loading` or `not-editable`.
   */
  commitEdit(move?: EditMove): EditResult;
  /** Close the cell editor and keep the cell as it was. The grid takes the keyboard. */
  cancelEdit(): void;
  /**
   * Make `text` the text of the open cell edit, as the formula bar does while the user types in
   * it. The cell editor shows the text in the cell at once, and `snapshot().editText` follows
   * it: the edit has one text, which the bar does not copy. With no open edit, one opens over
   * the active cell first, with the refusals of `beginEdit` (such as `locked` and `view-mode`),
   * in `edit` mode, and without the keyboard: the keyboard stays in the caller's field.
   *
   * `options.composing` is whether an input method composes in the caller's field. While it
   * does, `snapshot().composing` is true, and each command that would commit the edit is
   * refused with `composing`. Call it again with `composing` false when the composition ends,
   * also when the field loses the keyboard or unmounts during a composition. A key typed in the
   * cell editor also ends it. `commitEdit` and `cancelEdit` end the edit as for the cell editor.
   * A `text` that is not a string throws a `TypeError` before anything changes. An edit that
   * opens clears an earlier refusal, except `edit-dropped`, as a new selection does.
   */
  setEditText(text: string, options?: EditTextOptions): EditResult;
  /**
   * Select the range that `reference` names, as the name box does on Enter: a cell (`B2`,
   * `$B$2`), a range (`B2:C5`), whole columns (`C:E`), or whole rows (`3:5`), in any letter
   * case, with an optional sheet name (`Sheet2!A1`, `'My sheet'!A1:B2`) that shows that
   * sheet. A defined name (`x:definedName`) goes to its range when its text is one such
   * reference with a sheet name. A name that the shown sheet scopes (`@localSheetId`) comes
   * before a workbook name, and a sheet prefix (`Data!Rates`) reads the name that that sheet
   * scopes. The active cell is the top-left cell of the range, and the view scrolls to it.
   * An open cell edit commits first, as `select` states.
   *
   * Refused with `go-to-name` for a defined name that names a formula, several ranges, or a
   * range on a sheet that is not a visible worksheet, or that only another sheet scopes;
   * `go-to-sheet` for a sheet name that names no visible worksheet; `go-to-reference` for any
   * other text; `composing` while an input method composes in the open edit; and `loading` or
   * `not-editable` when no workbook is ready. A refusal shows and is announced as an edit
   * refusal is, and changes nothing.
   */
  selectReference(reference: string): EditResult;
  /**
   * Write text to the active cell, or to `at` on the shown sheet, as the cell editor commits it.
   * One undo step.
   *
   * - `=` starts a formula, `TRUE` and `FALSE` are booleans, and a leading apostrophe keeps the
   *   rest as text. The apostrophe is not stored.
   * - A number form stores a number: a decimal, a percent (`12%`), a currency amount (`$5.50`),
   *   a grouped number (`1,234.5`), a negative in parentheses, a fraction with a whole part
   *   (`1 1/2`), a date (`2026-11-12`), a time (`13:45`), an elapsed time (`36:00`), or a date
   *   and a time. The number locale gives the separators and the day and month order. A date
   *   with no year takes the year of the host clock.
   * - A cell with the General format also takes the number format of the form (ECMA-376 Part 1,
   *   18.8.30), in the same undo step. A sheet whose protection does not allow formatting cells
   *   keeps the cell format, and so does a styles part whose cell format list or number format
   *   list is full (64,000 records), or that the store cannot read, and so does a number locale
   *   code that the store cannot read. A cell with another number format keeps it.
   * - A cell with the Text format (`@`) stores the text, but a formula stays a formula, and an
   *   empty entry clears the cell.
   * - A date is a serial number in the date system of the workbook (Part 1, 18.17.4.1). In a
   *   cell that stores a date value (`t="d"`), ISO 8601 text stays a date value.
   *
   * `at` is 0-based: `row` from 0 to 1048575 and `col` from 0 to 16383, both integers; another
   * point is refused with `bad-reference`. An open cell edit is committed first when `at` is
   * omitted or is the cell of that edit, so its text is not lost; a refused commit refuses the
   * write. Another `at` leaves the open edit open. A write to the active cell keeps the selection
   * in its undo step, so undo and redo select it again. A write to another cell through `at`
   * keeps none: undo and redo of it leave the selection as it is. While an input method
   * composes, it is refused with `composing`.
   */
  enterValue(text: string, at?: GridPoint): EditResult;
  /**
   * Clear the values and formulas of the selected cells, and keep their formats. One undo step. An
   * open cell edit is committed first. While an input method composes, it is refused with
   * `composing`. A cell that holds a second form of its value that the clear can neither keep nor
   * replace, such as a formula in `mc:AlternateContent`, refuses it with `cell-alternate-content`
   * and the sheet and cell as the subject. A clear that would break a formula range is refused
   * with `shared-formula` and the blocking range or cell as the subject, and `causeCell` when one
   * cell's formula blocks a shared master. `EDIT_REFUSAL_CODES` lists the cases.
   */
  clearValues(): EditResult;
  /**
   * Apply a format patch to the selected cells, or to the active cell when it is the only cell.
   * Each cell keeps the properties that the patch does not set. One undo step. A range that would
   * create more than `MAX_FORMAT_NEW_CELLS` cells (more than one whole row) is refused with
   * `range-too-large`. An open cell edit is committed first. While an input method composes, it is
   * refused with `composing`.
   */
  applyFormat(patch: FormatPatch): EditResult;
  /**
   * Remove the cell formats of the selected cells, and keep their values. An empty cell whose
   * row or column has a format gets a plain cell (`s="0"`), so it does not show that format.
   * One undo step. A range that would create more than `MAX_FORMAT_NEW_CELLS` such cells is
   * refused with `clear-too-large`, a protected sheet that does not allow formatting with
   * `locked`, and `view` mode with `view-mode`. An open cell edit is committed first.
   * While an input method composes, it is refused with `composing`. On a protected sheet, each
   * cell keeps whether it is locked or hidden: it gets format 0 with its own protection. When a
   * formula or a defined name calls `CELL` or a macro function `GET.*`, which can read a cell
   * format (`CELL("format")`), the saved workbook asks for a full calculation on load. When
   * the record of that protection would pass the record limit of the styles part, the clear is
   * refused with `style-limit`.
   */
  clearFormats(): EditResult;
  /**
   * Clear the values, formulas, and cell formats of the selected cells, as `clearValues` and
   * `clearFormats` in one undo step. Refused with the reasons of both: `clear-too-large`,
   * `style-limit`, `cell-alternate-content`, `shared-formula` when the clear would break a
   * formula range, and `locked` for a locked cell that holds content or for a sheet that does not
   * allow formatting. A `locked`, `cell-alternate-content`, or `shared-formula` refusal has the
   * command `clearAll`. An open cell edit is committed first.
   * While an input method composes, it is refused with `composing`.
   */
  clearAll(): EditResult;
  /**
   * Toggle bold, italic, or underline from the format of the active cell. One undo step. An open
   * cell edit is committed first. While an input method composes, it is refused with `composing`.
   */
  toggleFormat(which: FormatToggle): EditResult;
  /**
   * Insert rows above the selection: as many as a whole-row selection or a range has rows, or
   * one above the active cell. Every reference of the workbook moves with the rows (design
   * D4), or the edit is refused with `structure-blocked`, which names what holds the
   * reference, `grid-full`, or `sheet-name-too-long`, as `deleteSheet` states. It is refused
   * with `sheet-name-ambiguous`, with the other sheet as the subject, when a file gives the
   * name of the shown sheet to another sheet too, in any case: a prefix such as `DATA!A1` can
   * name either sheet, so no reference can move with certainty. A formula or a defined name
   * inside `mc:AlternateContent` whose text the edit would change refuses it with
   * `structure-blocked`, because the editor does not rewrite each branch. The refusal names the
   * cell or the name, or the wrapper when a name list in `x:workbook` names the sheet. An open
   * cell edit commits first. One undo step. While an input method composes, it is refused with
   * `composing`, and in view mode with `view-mode`.
   */
  insertRows(): EditResult;
  /** Delete the rows of the selection, or the active cell's row. As `insertRows`. */
  deleteRows(): EditResult;
  /** Insert columns to the left of the selection, or of the active cell. As `insertRows`. */
  insertColumns(): EditResult;
  /** Delete the columns of the selection, or the active cell's column. As `insertRows`. */
  deleteColumns(): EditResult;
  /**
   * Merge the range of the selection into one cell (`x:mergeCell`), and select it with its
   * top-left cell active. The merged cell keeps the value of the top-left cell only. A range
   * where another cell holds a value or a formula is refused with `merge-drops-values`, so the
   * host can ask the user first and then merge with `dropValues`, which clears those cells in
   * the same undo step. With `center`, the merged cell is also centered, and a range too large to
   * format is refused with `range-too-large`, as `applyFormat`. A selection of one cell
   * is refused with `merge-one-cell`, a range that meets a merged range with `merge-overlap`,
   * and a protected sheet with `locked`. The clear of `dropValues` can refuse with
   * `cell-alternate-content` or `shared-formula`, and that refusal has the command `merge`.
   * An open cell edit commits first. One undo step. While
   * an input method composes, it is refused with `composing`, and in view mode with `view-mode`.
   */
  mergeCells(options?: MergeOptions): EditResult;
  /**
   * Unmerge each merged range that the selection meets. The values stay in the top-left cells.
   * A selection that meets no merged range changes nothing and makes no undo step. As
   * `mergeCells` for the open edit, the protection, and the view mode.
   */
  unmergeCells(): EditResult;
  /**
   * Add an empty worksheet after the shown sheet, and show it. Without `name`, the name is the
   * catalog text `sheetTabs.newSheetName` with the smallest free number `{n}` from 1. A name
   * that breaks the rules of a sheet name is refused with `sheet-name-invalid` (empty, more
   * than 31 characters counted in UTF-16 code units, one of `[ ] : * ? / \`, a control
   * character such as a tab or a line break, a character that XML cannot hold, an `ST_Xstring`
   * escape sequence such as `_x0041_`, or an apostrophe at the start or the end) or
   * `sheet-name-taken` (another sheet has it, compared without case after NFC). An open cell
   * edit commits first. One undo step. When a formula or a defined name already names the new
   * name, the saved workbook asks for a full calculation on load, as after a value edit. An
   * add, a rename, or a move asks for it too when a formula or a defined name calls a function
   * that reads sheet names or positions: `INDIRECT`, `CELL`, `INFO`, `SHEET`, `SHEETS`, or
   * `FORMULATEXT` (stored as `_xlfn.SHEET`, `_xlfn.SHEETS`, and `_xlfn.FORMULATEXT`), or a
   * macro function `GET.*`, such as `GET.WORKBOOK` or `GET.CELL`, in a formula or a defined
   * name. In view mode, it is refused with `view-mode`. When the workbook protection locks the
   * sheet list (`snapshot().workbookProtection.allowStructure` is false), each sheet command
   * (`addSheet`, `renameSheet`, `deleteSheet`, `moveSheet`) is refused with `locked`, and the
   * lock rule is `workbook`. Refused with `structure-blocked`, with the defined name as the
   * subject, when the add changes the position of the sheet that the local scope of a defined
   * name inside `mc:AlternateContent` names.
   */
  addSheet(name?: string): EditResult;
  /**
   * Rename the sheet at position `index` of `snapshot().sheets`. Every formula, defined name,
   * and internal hyperlink location that names the sheet gets the new name. Refused with the
   * name codes of `addSheet`, or with `structure-blocked`, with the blocking content as the
   * subject: when the workbook holds any chart, pivot table, or sheet that is not a worksheet,
   * or a relationship outside the allowlist, whatever it names; and when a table, conditional
   * formatting, data validation, a drawing, a control, or other content that the editor does
   * not rewrite names the sheet. Refused with `unknown-sheet` when no sheet is at `index`, also
   * in view mode, because the index is checked before the mode. Refused with
   * `sheet-name-ambiguous`, with the other sheet as the subject, when a file gives the name of
   * this sheet to another sheet too. A name in other case (`Data` and `DATA`) refuses the rename
   * when a formula, a defined name, or a hyperlink location names one of the two. A name in the
   * same case always refuses it, because the name cannot find this sheet. A formula or a
   * defined name inside `mc:AlternateContent` that names the sheet refuses it with
   * `structure-blocked`, which names the cell or the name, or the wrapper when a name list in
   * `x:workbook` names the sheet. Refused with `view-mode` in view mode, and with `locked` (rule
   * `workbook`) as `addSheet` states.
   * One undo step. Marks the cached results stale as `addSheet` states, for a name that a
   * formula already names and for a call to one of the functions that `addSheet` lists.
   */
  renameSheet(index: number, name: string): EditResult;
  /**
   * Delete the sheet at position `index`. Each reference to it becomes `#REF!`, and its
   * sheet-scoped defined names go. When it is the shown sheet, the nearest visible sheet shows.
   * Printer settings that only this sheet uses go with it. One undo step. Refused with:
   *
   * - `last-sheet` when no visible sheet would be left.
   * - `sheet-range-blocked` when a sheet range (`Sheet1:Sheet3!A1`) starts or ends at the
   *   sheet. A range that only spans it keeps its text.
   * - `structure-blocked` for content as `renameSheet` refuses, a custom workbook view that
   *   shows the sheet, or a related part of the sheet other than hyperlinks and printer settings.
   *   When the sheet is not a worksheet, the refusal has the `about` kind `not-worksheet`. A
   *   formula or a defined name inside `mc:AlternateContent` that the delete would change refuses
   *   it too, and so does a defined name there whose local scope sheet changes position.
   * - `sheet-name-ambiguous`, with the other sheet as the subject, when a file gives the name of
   *   this sheet to another sheet too, in any case.
   * - `sheet-name-too-long`, with the full name as the subject, when a stored sheet name has more
   *   than 31 characters. A rename of that sheet can shorten it.
   * - `unknown-sheet` when no sheet is at `index`, also in view mode, because the index is
   *   checked before the mode.
   * - `view-mode` in view mode.
   * - `locked`, with the `workbook` rule, as `addSheet` states.
   */
  deleteSheet(index: number): EditResult;
  /**
   * Move the sheet at position `index` to position `to`. The shown sheet stays shown.
   * Sheet-scoped defined names and the active tab follow their sheets. Charts never refuse a
   * move. A move to the same position changes nothing. One undo step. Marks the cached results
   * stale when a formula or a defined name calls one of the functions that `addSheet` lists.
   * Refused with:
   *
   * - `sheet-range-blocked` when the move changes the sheets between the ends of a sheet range
   *   (`Sheet1:Sheet3!A1`) of a formula or a defined name, or which end comes first in tab
   *   order. Content that the editor does not rewrite and that holds a sheet range refuses too.
   * - `sheet-name-ambiguous` when another sheet has the same name in the same case, because the
   *   name cannot find this sheet.
   * - `sheet-name-too-long`, as `deleteSheet` states.
   * - `structure-blocked`, with the sheet name as the subject, when the sheet is not a
   *   worksheet (a chartsheet, a macro sheet, a dialog sheet, or a sheet whose part is
   *   missing), as `deleteSheet` refuses it. A worksheet moves past such a sheet. Also with the
   *   defined name as the subject, when the move changes the position of the sheet that the local
   *   scope of a defined name inside `mc:AlternateContent` names.
   * - `unknown-sheet` when no sheet is at `index`, also in view mode, as `deleteSheet` states.
   * - `bad-position` when `to` is outside the sheet list.
   * - `view-mode` in view mode.
   * - `locked`, with the `workbook` rule, as `addSheet` states.
   */
  moveSheet(index: number, to: number): EditResult;
  /**
   * Set the width of the columns `first` to `last` of the shown sheet (0-based, as `GridPoint`) to
   * `width` characters of the maximum digit width of the default font, with a custom width flag
   * (ECMA-376 Part 1, 18.3.1.13). The `x:col` elements that the range crosses split and keep their
   * other attributes. The refusals: `bad-size` for a width outside 0 to 255 (this editor's limit,
   * `MAX_COLUMN_WIDTH`; the schema states none); `bad-reference` for a range outside the grid;
   * `resize-anchor` when the resize would change a drawing, control, or object anchor that keeps
   * its size or position, or when the drawing part of the sheet cannot be read; `invalid-sheet`
   * when the sheet has a cell structure that the editor cannot change, or when an `x:cols` after
   * the first has an attribute that the merge of the column lists would drop; and `locked`, with the `size` rule, on a protected sheet that does
   * not allow formatting columns (`SheetProtection.allowFormatColumns`); there, a column heading shows no
   * resize edge. In view mode, it is refused with `view-mode`, and a heading edge starts no drag.
   * An open cell edit commits first. One undo step. A drag of the right edge of a column heading
   * calls it on pointer up. When a formula or a defined name calls `CELL` (`CELL("width")`) or
   * a macro function `GET.*` (`GET.CELL`), which can read a column width, the saved workbook asks
   * for a full calculation on load.
   */
  setColumnWidth(first: number, last: number, width: number): EditResult;
  /**
   * Set the height of the rows `first` to `last` of the shown sheet (0-based) to `height` points,
   * with a custom height flag (Part 1, 18.3.1.73). A row with no `x:row` gets an empty one. The
   * refusals: `bad-size` for a height outside 0 to 409 (this editor's limit, `MAX_ROW_HEIGHT`; the
   * schema states none); `bad-reference` for a range outside the grid; `too-many-rows` when it
   * would create more than `MAX_RESIZE_NEW_ROWS` rows; `resize-anchor` when the resize would change
   * an anchor that keeps its size or position, or when the drawing part of the sheet cannot be
   * read; `invalid-sheet` when the rows of the sheet are
   * out of order; and `locked`, with the `size` rule, on a protected sheet that does not allow
   * formatting rows (`SheetProtection.allowFormatRows`). In view mode, it is refused with
   * `view-mode`. When a formula or a defined name calls a macro function `GET.*` (`GET.CELL`),
   * which can read a row height, the saved workbook asks for a full calculation on load. As
   * `setColumnWidth` otherwise.
   */
  setRowHeight(first: number, last: number, height: number): EditResult;
  /**
   * Copy the selected range, as Mod+C does. The editor keeps what each cell holds (its value or
   * formula, and its cell format) and returns the content for the clipboard: `text`, the shown text
   * of the cells, tab-separated with rows joined by CR LF, for `text/plain`, and `token`, which
   * marks this copy, for the type `application/x-xlsx-editor`. A text that starts with `=`, `+`,
   * `-`, `@`, tab, or CR gets an apostrophe in `text`, except the shown text of a number or a date
   * that is one signed number with a digit after its sign and an optional currency symbol, such
   * as `-1,234.50`, `-$1,234.50`, or `-1.2E+03` (`-1+2` and `-$A$1` get it). A tab or a line
   * break inside the shown text becomes a space in `text` (approximation `copied-line-breaks`).
   * A date whose shown text does not read back to the same day, such as `2-Jan` or a two-digit
   * year outside 1930 to 2029, gives its edit text, such as `2035-06-01`.
   * Refused with `merged-range` when the range meets a merged range, `copy-too-large` for more
   * than `MAX_FORMAT_NEW_CELLS` cells, `copy-conditional-limit` when the range has more
   * conditional formatting than one copy can check, `copy-array-formula` when it meets an array
   * or data table range, also without its master cell (that range is the subject, as an edit
   * names it), and `copy-unreadable-formula` when the copy cannot get the text of a formula: an
   * empty formula, or a shared formula whose master cell is missing or cannot move to the cell.
   * `copy-wrapped-formula` refuses a cell whose formula is inside `mc:AlternateContent` (the cell
   * is the subject), so the paste never writes its cached value as a constant. The copy keeps any
   * other formula as its text, also one that the tokenizer cannot read; its paste is then refused
   * with `paste-unreadable-formula`. A copied or cut formula pasted on another sheet gives a name
   * that the source sheet scopes the source sheet prefix. A copy with such a name depends on its
   * source sheet for a paste on another sheet: such a paste after a rename or a delete of that
   * sheet is refused with `copy-expired`. A paste on the source sheet, renamed or not, keeps the
   * binding and is applied. A rich text cell copies as its plain text,
   * and a stored error value as text (approximation `pasted-values`). Hyperlinks, comments, data
   * validation, and conditional formatting stay with the source cells (approximation
   * `cut-cell-attachments`). An open cell edit commits first; while an input method composes, the
   * copy is refused with `composing`. Changes no cell, so it works in view mode and on a protected
   * sheet. The editor keeps one copy, and forgets it when another workbook opens. The live region
   * announces the copied range (`grid.copied`). A dashed outline marks the source range. A paste
   * that leaves the source as it was keeps the outline, and every other edit ends it. Escape on the
   * grid ends the outline and keeps the copy, so a later paste still writes the copied cells.
   */
  copy(): CopyResult;
  /**
   * Cut the selected range, as Mod+X does: a copy whose next paste also clears the values and cell
   * formats of the source, in the same transaction. A cut formula keeps the cells it refers to:
   * pasted on another sheet, each reference without a sheet name, and each name that the source
   * sheet scopes, gets the name of the source sheet. References to the cut cells do not move.
   * A dashed outline marks the source range until the cut ends. Escape on the grid forgets a
   * pending cut: the source stays, and a later paste of its token takes the `text` lane. When
   * the workbook changes between the cut and the paste, also by an undo or a redo, the paste is refused with `cut-expired`, the source stays, and every later paste of
   * the same cut is refused with `cut-expired` until a new copy or cut. Escape does not forget an
   * expired cut. After one paste, the next paste of the cut is a copy, with the rules of
   * `copy-expired` that `paste` states. Every refusal of a cut has the command `cut`, and its
   * message says that the cells were not cut; `CUT_REFUSAL_CODES` lists the codes with a message.
   * A refusal of the open cell edit that the cut commits first, such as `composing`, is not a
   * refusal of the cut: it has the command `commit` or no command. A cut is refused with `view-mode` in view mode. It runs
   * the checks of `copy` first, so a cut and a copy of one range give the same copy reason, such
   * as `copy-wrapped-formula` for a cell whose formula is inside `mc:AlternateContent`, also on a
   * protected sheet. The cut checks the clear of its source next. So it is also refused when the
   * paste could not clear the source, as `clearAll` states: on a protected sheet with `locked`
   * (the `cell` rule for a locked source cell that holds content, and the `format` rule when the
   * sheet does not allow formatting cells); with
   * `cell-alternate-content` when a source cell holds a second form of its value that the clear can
   * neither keep nor replace; with `shared-formula` when the clear would break a formula range; and
   * with `invalid-sheet` when the store cannot edit the sheet. It is refused with
   * `cut-wrapped-value`, with the sheet and cell as the subject, when a source cell holds its value
   * only in `mc:AlternateContent`, which the copy cannot carry. The paste clears only the source
   * cells that it does not write.
   * On a protected sheet, the clear keeps whether each source cell is locked or hidden.
   */
  cut(): CopyResult;
  /**
   * Paste clipboard data at the active cell, as Mod+V does, as one transaction and one undo step,
   * and select the pasted range. `token` equal to the last copy pastes that copy: values, formulas
   * moved by the offset from the source, and cell formats. Else `text` pastes as rows split at line
   * breaks and cells split at tab. Each cell is read as `enterValue` reads typed text, so a General
   * cell takes the number format of a typed form such as `12%`, and other cell formats are kept.
   * On a sheet whose protection does not allow formatting cells, or with a full styles part, the
   * paste stores the numbers and keeps the cell formats. Quotes have no meaning, so a quoted
   * value with a line break splits into rows. Else `html` refuses the paste with `paste-html`;
   * the HTML is never read. Rich text and stored error values of a copy paste as text
   * (approximation `pasted-values`). Refused with:
   *
   * - `grid-full` when the range would leave the grid.
   * - `paste-reference-outside` when a moved reference would leave the sheet.
   * - `range-too-large` for more than `MAX_FORMAT_NEW_CELLS` cells.
   * - `merged-range` when the range meets a merged range.
   * - `paste-unreadable-formula`, with the source cell as the subject, when a formula of the copy
   *   cannot be read to move it. Also when it fails the check of typed formula input (such as
   *   `1+` or `A1:S!B2`, also in a cut on the same sheet), or holds a sheet range that reads two
   *   ways.
   * - `unsupported-function`, with the function name, when a formula calls a function outside the
   *   predefined functions.
   * - `style-limit` when the clear of a cut source keeps the protection of its cells and that
   *   record would pass the record limit of the styles part, as `clearFormats` states.
   * - `cell-alternate-content`, with the sheet and cell as the subject, when a target cell holds
   *   a second form of its value that the paste can neither keep nor replace, such as a formula
   *   in `mc:AlternateContent`.
   * - `shared-formula`, with the blocking range or cell as the subject, when the paste would
   *   break a formula range, as `EDIT_REFUSAL_CODES` states: a cell inside an array or data table
   *   range, or a shared master whose children the editor cannot rewrite (`causeCell` names the
   *   cell).
   * - `cut-expired` for a cut that the workbook changed after.
   * - `copy-expired` when a sheet that the copy depends on, or a copied cell format, changed after
   *   the copy. A copy with formulas also expires at an insert or a delete of rows or columns on
   *   any sheet, and at a rename, a move, or a delete of a sheet that a formula names. A formula with a reference
   *   across a span of sheets also expires at an add, a rename, a move, or a delete of any sheet.
   *   An undo or a redo of such an edit does not bring the copy back.
   * - `paste-name-scope` when a copied or cut formula pasted on another sheet holds a name with no
   *   sheet prefix that the target sheet scopes and the source sheet does not. Also when it calls a
   *   function name that the source sheet scopes.
   * - `locked` on a protected sheet, with the command `paste`: with the `cell` rule when a
   *   target cell is locked, or when the paste would give a cell a cell format that changes
   *   whether it is locked or hidden, whatever the protection allows; and with the `format`
   *   rule when the paste would give a cell another cell format and the protection does not
   *   allow formatting cells.
   * - `view-mode` in view mode, and `composing` while an input method composes in the open cell
   *   editor.
   *
   * After `cut-expired` or `copy-expired`, every paste of the same `token` is refused with the same
   * code until a new copy or cut. Such a paste never takes the `text` lane. An open cell edit
   * commits first. When a paste changes a cell format and a formula or a defined name calls
   * `CELL` or a macro function `GET.*`, the saved workbook asks for a full calculation on load.
   */
  paste(data: PasteData): EditResult;
  /**
   * Undo the last edit, as the toolbar's Undo does. An open cell editor is not committed. When
   * it holds changed text (`snapshot().editChanged`), the edit is cancelled, the cell shows its
   * content again, and the history does not move. An open editor with unchanged text closes,
   * and the undo steps the history once. While an input method composes, the undo is refused
   * with `composing`. The history keys in an open cell editor never call this: they belong to
   * the input's own undo. An empty history is refused with `nothing-to-undo`, and `view` mode
   * with `view-mode`. After
   * the undo, the editor selects the selection that the edit started from, on its sheet, and
   * scrolls its active cell into view; an edit that kept no selection leaves it as it is.
   */
  undo(): EditResult;
  /**
   * Redo the last undone edit. An open cell edit with changed text (`snapshot().editChanged`)
   * refuses the redo with `edit-pending` before any write, as the toolbar disables it; an open edit
   * with no change closes first. While an input method composes, it is refused with `composing`. An
   * empty redo history is refused with `nothing-to-redo`. After the redo, the editor selects the
   * selection that the edit ended with, as `undo` does.
   */
  redo(): EditResult;
  /**
   * The workbook as bytes of the opened package's kind (`snapshot().workbookKind`: a workbook or a
   * template), with every edit, or `null` when none is open. An open cell editor is committed
   * first; when the store refuses that commit, `save` returns `null` and the snapshot holds the
   * refusal. While an input method composes, it is refused with `composing`.
   */
  save(): Uint8Array<ArrayBuffer> | null;
  /**
   * {@link XlsxEditorInstance.save} in short tasks, so the page stays responsive while a large
   * workbook saves. The workbook is captured at the call: an open cell editor commits then. A
   * refusal is in the snapshot when the call returns, but a workbook that cannot serialize is
   * refused when the failure occurs, with the code `save-failed`. An edit made while the save
   * runs belongs to the next save. A call during a running save of the same revision joins that
   * save, and gets its own copy of the bytes. Resolves to `null` when the save is refused.
   */
  saveInTasks(): Promise<SavedWorkbook | null>;
  /**
   * Unmount and release the editor. An open edit is committed, or its refusal is announced.
   * While an input method composes, the edit is dropped with `edit-dropped`, and the listeners
   * hear it before they go. After `destroy`, no listener is left to notify.
   */
  destroy(): void;
}
