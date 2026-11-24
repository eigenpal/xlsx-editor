// The clipboard of the editor (design D2 of the change `workbook-structure-editing`): copy,
// cut, and paste of a range, DOM-free. `clipboard-events.ts` connects it to the browser events.
//
// Follows DOCX Editor and PPTX Editor at `origin/main`:
// - docx-editor `packages/core/src/editor/clipboard-write.ts:7-9` (f3116761): the engine never
//   reads the clipboard. A paste takes the text of the paste event as an argument.
// - docx-editor `packages/core/src/editor/surface-clipboard-ops.ts:353-411` (f3116761): a paste
//   is one commit and one undo step.
// - docx-editor `packages/core/src/editor/docx-editor-exec.ts:549-580` (f3116761): a cut reads
//   its content before it deletes. Here the cut keeps a snapshot and clears the source in the
//   paste transaction.
// - pptx-editor `packages/core/src/editor/clipboard.ts:42,69-82` (79a5d22d): a private
//   clipboard type marks a copy of this editor. Difference: the type holds only a token of the
//   snapshot that the editor keeps, never content, so a paste never parses clipboard data; and
//   plain text equal to the last copy takes the plain lane.
//
// A copy keeps, for each cell of the range, its input (a value or formula text) and its cell
// format index, and gives the text for `text/plain` and the token. A paste routes: the token of
// the last copy to the internal lane; else `text/plain` to the plain lane; else `paste-html`
// when only `text/html` is there; else nothing changes. Every value of the clipboard is
// untrusted: the plain lane reads text as typed input, and HTML is never read.
//
// An internal paste uses the snapshot only while everything it depends on is unchanged: the
// store, the source cells of a cut (its revision), the sheets that a copied formula names, and
// the `x:xf` element of each copied cell format index (`clipboard-deps.ts`). An edit that can
// change what a copied formula refers to also ends the copy, and an undo does not bring it back.
// Otherwise the paste is refused with `cut-expired` or `copy-expired`. The editor then drops the
// snapshot and keeps only the token and its refusal, so the same token stays refused with that
// code until a new copy, cut, or reset; it never falls to the plain lane.
//
// The source range of a pending copy or cut shows a dashed outline over the grid. The outline is
// never part of the workbook, and `#markAt` is its only state, apart from the snapshot: the end
// of the outline never changes what a paste of the copy writes. It shows while the store is at
// the revision of the copy, of the cut, or of a paste of the copy. So any other edit, an undo,
// and a redo end it. A paste of a cut, a paste over the source, and Escape end it too. Escape on
// a copy ends only the outline. Escape on a cut also forgets the cut.

import type { GridPoint, GridRange } from '../contracts/index.ts';
import {
  attr,
  cellHasFormula,
  cellKey,
  child,
  columnFormats,
  formatCellRef,
  formatIndex,
  localNameKeys,
  moveCopiedFormula,
  qualifyFormula,
  readFormulaInput,
  sheetPrefixText,
  rowFormat,
  sharedFormulaText,
  sheetIndex,
  styleLists,
  xmlPartOf,
  MAX_COLS,
  MAX_FORMAT_NEW_CELLS,
  arrayFormulaIn,
  MAX_ROWS,
  type CellInput,
  type Element,
  type FormulaWriteProblem,
  type WorkbookDocumentStore,
  type WorkbookOp,
  type XlsxPackage,
  wrappedValueCell,
} from '../store/index.ts';
import {
  cellFormat,
  cellValue,
  conditionalFormat,
  dateToSerial,
  displayValue,
  formatCodeOf,
  LAYOUT_CHECK_BUDGET,
  type CellValue,
  type CheckBudget,
  type NumberLocale,
} from '../cascade/index.ts';
import { protectionChecks } from '../store/store/sheet-protection.ts';
import { clearOps } from '../automation/index.ts';
import {
  editingOf,
  type CellEditing,
  type EditTarget,
  type RefusedBy,
  type RunHistory,
} from './cell-editing.ts';
import { refusalOf, type EditRefusal, type EditResult } from './cell-input.ts';
import { isMultiRange, type Merges } from './selection.ts';
import { cellContent, selectionRef, typedInput } from './editor-model.ts';
import { copiedDateText } from './typed-forms.ts';
import { formatsFit, pastedFormatOps, type FormattedCell } from './entry-ops.ts';
import type { ConditionalCut } from './conditional-cut.ts';
import { copyDeps, depsHold, sameNode, usesLocalName, type CopyDeps } from './clipboard-deps.ts';

/** The private clipboard type of a copy: it holds only the token of the copy. @public */
export const CLIPBOARD_TYPE = 'application/x-xlsx-editor';

/** What a copy or a cut puts on the clipboard. @public */
export interface ClipboardContent {
  /** Tab-separated shown text, rows joined by CR LF, for `text/plain`. */
  readonly text: string;
  /** The token of the copy, for the type `CLIPBOARD_TYPE`. */
  readonly token: string;
}

/** The outcome of a copy or a cut. @public */
export type CopyResult =
  | { readonly ok: true; readonly content: ClipboardContent }
  | { readonly ok: false; readonly refusal: EditRefusal };

/** What a paste event carries. Every field is untrusted. @public */
export interface PasteData {
  /** The `text/plain` data. */
  readonly text?: string;
  /** The `CLIPBOARD_TYPE` data. */
  readonly token?: string;
  /**
   * Whether the clipboard holds `text/html`: the event's list of types holds it. The HTML itself
   * is never read.
   */
  readonly html?: boolean;
}

/** One cell of a copy: what it holds and its cell format index. */
interface CopiedCell {
  readonly input: CellInput;
  readonly format: number;
}

interface Snapshot {
  readonly token: string;
  /** The `text/plain` of the copy, for `tokenOfText`. */
  readonly text: string;
  /** The store of the copy. Another store (a workbook opened later) takes the plain lane. */
  readonly store: WorkbookDocumentStore;
  /** The source range, its sheet by name and part, at the revision of the copy. */
  readonly sheet: string;
  readonly partName: string;
  readonly origin: GridPoint;
  readonly rows: number;
  readonly cols: number;
  /** Row by row. */
  readonly cells: readonly CopiedCell[];
  /** What the copy depends on (`clipboard-deps.ts`). */
  readonly deps: CopyDeps;
  /**
   * The sheets that the copied formulas name, whether a formula spans sheets, and the store's
   * count of reference edits at the copy. Null when the copy holds no formula.
   */
  readonly references: {
    readonly sheets: readonly string[];
    readonly span: boolean;
    readonly at: number;
  } | null;
  /** A cut: the store revision of the cut, or null for a copy or a used cut. */
  cutAt: number | null;
}

/** An expired copy: only its token and its refusal stay. */
interface SpentCopy {
  readonly token: string;
  readonly store: WorkbookDocumentStore;
  readonly refusal: EditRefusal;
}

/** The source range of a snapshot as a reference with its sheet, the subject of a refusal. */
function sourceSubject(snap: Snapshot): string {
  const ref = selectionRef({
    sheetIndex: 0,
    active: snap.origin,
    anchor: snap.origin,
    range: {
      start: snap.origin,
      end: { row: snap.origin.row + snap.rows - 1, col: snap.origin.col + snap.cols - 1 },
    },
  });
  return `${snap.sheet}!${ref}`;
}

/** What the clipboard reads and calls on the editor. */
export interface ClipboardHost {
  readonly edits: CellEditing;
  target(): EditTarget | null;
  locale(): NumberLocale;
  merges(): Merges;
  select(anchor: GridPoint, focus: GridPoint): void;
  /** Put text in the live region of the grid. */
  announce(text: string): void;
  /** Paint the grid again: the outline of a pending copy or cut changed. */
  repaint(): void;
  t(key: string, params?: Readonly<Record<string, string | number>>): string;
  /** The conditional formatting budget flag: a sheet with the flag copies its cell formats. */
  readonly conditionalCut: ConditionalCut;
}

/** The first characters that make a text cell read as a formula in another program. */
const RISKY_START = new Set(['=', '+', '-', '@', '\t', '\r']);
/**
 * One signed number as a number format shows it: a sign, an optional currency symbol of a typed
 * form (`typed-forms.ts`), then digits with group, decimal, time, fraction, and percent
 * characters, an optional exponent, and an optional currency symbol at the end (`-1,234.50`,
 * `-$1,234.50`, `-5,50 €`, `-1:30`, `-1 1/2`, `-1.2E+03`, `-5%`). A second sign (`-1+2`) is not
 * a number. The pattern is linear: each character has one part that can match it.
 */
const SIGNED_NUMBER = /^[-+][$€£¥]?[0-9 .,:/\u00a0\u202f]*(?:[Ee][-+]?[0-9]+)?(?:[$€£¥]|%)?$/;

/**
 * Whether shown text is one signed number that starts with a digit after its sign and an
 * optional currency symbol, such as `-1,234.50`, `-$1,234.50`, `-5,50 €`, `-.5`, or `-1.2E+03`.
 * `-E5`, `-$A$1`, and `-$1:1` read as references in a formula, so they are not.
 */
export function signedNumberText(text: string): boolean {
  // After a leading currency symbol, `:` would read as a row range (`-$1:1`).
  const symbol = /^[-+][$€£¥]/.test(text);
  return (
    SIGNED_NUMBER.test(text) &&
    /^[-+][$€£¥]?[.,]?[0-9]/.test(text) &&
    !(symbol && text.includes(':'))
  );
}

/**
 * The text of one cell for `text/plain` (AGENTS.md, formula injection on export): text that
 * starts with `=`, `+`, `-`, `@`, tab, or CR gets an apostrophe, unless it is a shown number
 * with its sign. A tab or a line break inside the text becomes a space, so the text keeps its
 * cell when it is pasted.
 */
export function exportCellText(text: string, numeric: boolean): string {
  const first = text[0];
  const risky =
    first !== undefined &&
    RISKY_START.has(first) &&
    !(numeric && (first === '-' || first === '+') && signedNumberText(text));
  const out = risky ? `'${text}` : text;
  return out.replace(/[\t\r\n]/g, ' ');
}

/**
 * Plain text as rows of cells: rows split at CR LF, LF, or CR, and cells at tab. A line break at
 * the end does not start a row. Null when the text holds more than `limit` rows, or more than
 * `limit` line breaks and tabs: a scan counts them before any split, and stops at the limit,
 * so a large text never builds a large list.
 */
export function splitPlainText(text: string, limit = MAX_FORMAT_NEW_CELLS): string[][] | null {
  let breaks = 0;
  for (let i = 0; i < text.length; i += 1) {
    const ch = text.charCodeAt(i);
    if (ch === 13 /* CR */ && text.charCodeAt(i + 1) === 10) i += 1;
    if (ch === 9 || ch === 10 || ch === 13) breaks += 1;
    if (breaks > limit) return null;
  }
  const lines = text.split(/\r\n|\n|\r/);
  if (lines.length > 1 && lines[lines.length - 1] === '') lines.pop();
  if (lines.length > limit) return null;
  return lines.map((line) => line.split('\t'));
}

/** The input of a stored value, as a paste writes it again. */
function inputOf(value: CellValue): CellInput {
  switch (value.kind) {
    case 'number':
      return { type: 'number', value: value.value };
    case 'boolean':
      return { type: 'boolean', value: value.value };
    case 'date':
      return { type: 'date', iso: value.iso };
    // An error value pastes as its text: a cell input has no error value (approximation
    // `pasted-values`). Rich text reads as its plain text for the same reason.
    case 'error':
      return { type: 'string', value: value.code };
    case 'string':
      return { type: 'string', value: value.text };
    case 'empty':
      return { type: 'empty' };
  }
}

/** Whether a rectangle meets a merged range. */
function meetsMerge(merges: Merges, top: number, left: number, rows: number, cols: number) {
  const bottom = top + rows - 1;
  const right = left + cols - 1;
  return merges.some(
    (m) => m.firstRow <= bottom && m.lastRow >= top && m.firstCol <= right && m.lastCol >= left
  );
}

/** The refusal code of a copied formula that a paste cannot write. */
function pasteProblemCode(problem: FormulaWriteProblem): EditRefusal['code'] {
  switch (problem) {
    case 'out-of-grid':
      return 'paste-reference-outside';
    case 'name-rebinds':
    case 'scoped-function':
      return 'paste-name-scope';
    default:
      return 'paste-unreadable-formula';
  }
}

/** The clipboard of each editor, by its cell editing, for the menu bar's Paste (`clipboardOf`). */
const clipboards = new WeakMap<object, WorkbookClipboard>();

/** The clipboard of `editor`, when the engine made it. @internal */
export function clipboardOf(editor: object): WorkbookClipboard | undefined {
  const edits = editingOf(editor);
  return edits ? clipboards.get(edits) : undefined;
}

/** The copy, cut, and paste commands of one editor. */
export class WorkbookClipboard {
  readonly #host: ClipboardHost;
  /** Tokens of this editor start with a random part, so a copy of another editor never matches. */
  readonly #prefix = Math.random().toString(36).slice(2, 10);
  #count = 0;
  #snapshot: Snapshot | null = null;
  /** The store revision at which the outline of the source range shows, or null when it ended. */
  #markAt: number | null = null;
  #spent: SpentCopy | null = null;

  constructor(host: ClipboardHost) {
    this.#host = host;
    clipboards.set(host.edits, this);
  }

  /** Forget the last copy: the editor opened another workbook or was destroyed. */
  reset(): void {
    this.#forget();
    this.#spent = null;
  }

  /** Forget the snapshot and its outline. */
  #forget(): void {
    this.#snapshot = null;
    this.#markAt = null;
  }

  /** Drop the snapshot of an expired copy, and keep its token and its refusal. */
  #expire(snap: Snapshot, refusal: EditRefusal): EditRefusal {
    this.#forget();
    this.#spent = { token: snap.token, store: snap.store, refusal };
    return refusal;
  }

  /** Expire the snapshot when an edit ended a cut, or changed what its formulas refer to. */
  #sweep(): void {
    const snap = this.#snapshot;
    const target = this.#host.target();
    if (!snap || snap.store !== target?.store) return;
    const refs = snap.references;
    if (snap.cutAt !== null && snap.cutAt !== target.revision)
      this.#expire(snap, { code: 'cut-expired', subject: sourceSubject(snap) });
    else if (refs && target.store.referenceEdits(refs.sheets, refs.span) !== refs.at)
      this.#expire(snap, { code: 'copy-expired', subject: sourceSubject(snap) });
  }

  /** Whether the outline of `snap` shows: it waits at the current revision of its store. */
  #marked(snap: Snapshot): boolean {
    const target = this.#host.target();
    return (
      this.#markAt !== null && snap.store === target?.store && this.#markAt === target.revision
    );
  }

  /** The source range of the pending copy or cut on the sheet part `partName`, or null. */
  pendingRange(partName: string | undefined): GridRange | null {
    this.#sweep();
    const snap = this.#snapshot;
    if (!snap || snap.partName !== partName || !this.#marked(snap)) return null;
    const { origin, rows, cols } = snap;
    return { start: origin, end: { row: origin.row + rows - 1, col: origin.col + cols - 1 } };
  }

  /**
   * End a pending copy or cut on any sheet, as Escape on the grid does. A copy loses only its
   * outline: a later paste of its token still writes the copied cells, under the checks of
   * `paste`. A cut is forgotten: the source cells stay, and a later paste of its token takes the
   * plain-text lane.
   * False when none waits, so the key goes on. An expired copy or cut does not wait: every later
   * paste of its token is still refused with its code, as `paste` states.
   */
  endPending(): boolean {
    this.#sweep();
    const snap = this.#snapshot;
    if (!snap || !this.#marked(snap)) return false;
    const isCut = snap.cutAt !== null;
    if (isCut) this.#forget();
    else this.#markAt = null;
    this.#host.announce(this.#host.t(isCut ? 'grid.cutEnded' : 'grid.copyUnmarked'));
    this.#host.repaint();
    return true;
  }

  /** Copy the selection. An open cell edit commits first. */
  copy(): CopyResult {
    return this.#copy(false);
  }

  /** Cut the selection: a copy whose next paste clears the source in the same transaction. */
  cut(): CopyResult {
    // A cut changes the workbook at its paste: view mode refuses it, as it refuses the paste.
    if (this.#host.edits.viewing)
      return this.#refused({ code: 'view-mode', command: 'cut' }, 'cut');
    return this.#copy(true);
  }

  /**
   * A refused copy, shown through the refusal channel. It returns the refusal that the channel
   * shows, so the result and the snapshot agree (`loading` while a workbook loads).
   */
  #refused(
    refusal: EditRefusal,
    by?: RefusedBy
  ): { readonly ok: false; readonly refusal: EditRefusal } {
    const shown = this.#host.edits.refuse(refusal, by);
    return { ok: false, refusal: shown.ok ? refusal : shown.refusal };
  }

  #copy(cut: boolean): CopyResult {
    const open = this.#host.edits.commitOpenEdit();
    if (!open.ok) return { ok: false, refusal: open.refusal };
    // Every later refusal of a cut is marked as a cut, so its message names the cut.
    const refuse = (refusal: EditRefusal) =>
      cut ? this.#refused({ ...refusal, command: 'cut' }, 'cut') : this.#refused(refusal);
    const target = this.#host.target();
    const part = target && xmlPartOf(target.pkg, target.entry.partName);
    if (!target || !part) return refuse({ code: 'not-editable' });
    // A copy holds one range: a selection of several ranges is refused, not copied in part.
    if (isMultiRange(target.selection)) return refuse({ code: 'multiple-ranges' });
    const { start, end } = target.selection.range;
    const rows = end.row - start.row + 1;
    const cols = end.col - start.col + 1;
    // The snapshot walks the range, so its size is bounded first.
    if (rows * cols > MAX_FORMAT_NEW_CELLS) return refuse({ code: 'copy-too-large' });
    if (meetsMerge(this.#host.merges(), start.row, start.col, rows, cols))
      return refuse({ code: 'merged-range' });
    // A range that meets an array or data table range, also without its master cell, cannot
    // paste: the store keeps such a range whole. The range is the subject, as an edit names it.
    const array = arrayFormulaIn(part, start.row, start.col, end.row, end.col);
    if (array)
      return refuse({ code: 'copy-array-formula', subject: `${target.entry.name}!${array}` });
    const { pkg, index } = target;
    const sheet = sheetIndex(part);
    const colStyles = columnFormats(part.root, start.col, end.col);
    // An index past the cell formats shows format 0 (`cellFormat`, `inheritedFormat`): the copy
    // holds the format that the cell shows, so a paste never writes an index that the part
    // check refuses.
    const xfCount = Math.max(styleLists(pkg, index).cellXfs.length, 1);
    const shownFormat = (xf: number | undefined): number =>
      xf !== undefined && xf < xfCount ? xf : 0;
    const options = { locale: this.#host.locale(), date1904: index.date1904 };
    // The plain text is the text that the grid shows, under the true conditional formatting rules
    // of each cell. After a budget cut of the sheet, every cell gives the text of its cell format,
    // as the hint follows the cell format then. A copy has its own check budget, which charges
    // each row once in full, so the result depends only on the range; a copy past it is refused.
    const rules = !this.#host.conditionalCut.holds(target.store, target.entry.partName);
    const budget: CheckBudget = { left: LAYOUT_CHECK_BUDGET, everyRow: true };
    const cells: CopiedCell[] = [];
    const lines: string[] = [];
    for (let r = start.row; r <= end.row; r += 1) {
      const fromRow = rowFormat(sheet.rows.get(r));
      const texts: string[] = [];
      for (let c = start.col; c <= end.col; c += 1) {
        const entry = sheet.cells.get(cellKey(r, c));
        if (!entry) {
          cells.push({
            input: { type: 'empty' },
            format: shownFormat(fromRow ?? colStyles.get(c)),
          });
          texts.push('');
          continue;
        }
        const s = attr(entry.el, 's');
        const value = cellValue(pkg, index, entry.el, false).value;
        const f = child(entry.el, 'x:f');
        let input: CellInput;
        if (f) {
          const t = attr(f, 't');
          const subject = `${target.entry.name}!${formatCellRef({ row: r, col: c })}`;
          // An array or data table formula holds its whole range: it cannot paste as one cell.
          // Its range could not be read (`arrayFormulaIn`), so the cell is the subject.
          if (t === 'array' || t === 'dataTable')
            return refuse({ code: 'copy-array-formula', subject });
          const text = sharedFormulaText(part, entry);
          if (text === undefined) return refuse({ code: 'copy-unreadable-formula', subject });
          input = { type: 'formula', text };
        } else if (cellHasFormula(entry.el))
          // A formula inside `mc:AlternateContent` has no text that a paste can move: its copy
          // is refused, never pasted as its cached value.
          return refuse({
            code: 'copy-wrapped-formula',
            subject: `${target.entry.name}!${formatCellRef({ row: r, col: c })}`,
          });
        else input = inputOf(value);
        cells.push({ input, format: shownFormat(formatIndex(s)) });
        const base = cellFormat(pkg, index, s).format;
        const format =
          value.kind === 'empty' || !rules
            ? base
            : conditionalFormat(pkg, index, sheet.root, r, c, value, base, budget);
        // A shown date that reads back to another day copies its edit text (`copiedDateText`).
        const serial =
          value.kind === 'number'
            ? value.value
            : value.kind === 'date'
              ? dateToSerial(value.iso, index.date1904)
              : undefined;
        const text = displayValue(value, format, options).text;
        const shown =
          serial === undefined
            ? text
            : copiedDateText(
                text,
                serial,
                formatCodeOf(format, options.locale),
                options.locale,
                index.date1904
              );
        texts.push(exportCellText(shown, value.kind === 'number' || value.kind === 'date'));
      }
      lines.push(texts.join('\t'));
    }
    // Past the budget, some cells kept their cell formats: the copy is refused, never mixed.
    if (rules && budget.left <= 0) return refuse({ code: 'copy-conditional-limit' });
    // The paste of a cut clears its source as Clear all does: a cut that the clear would refuse
    // (on a protected sheet, a locked cell that holds content, or formats that the sheet does not
    // allow to clear) is refused now, so no paste is refused later for its source. The checks of
    // the copy ran first, so a cut and a copy of one range give the same copy reason.
    if (cut) {
      const box = [start.row, start.col, end.row, end.col] as const;
      const ops = clearOps(target.entry.name, selectionRef(target.selection), 'all');
      const refusal = protectionChecks();
      for (const op of ops) {
        const locked = refusal(target.pkg, op);
        if (locked) return refuse(refusalOf(locked));
      }
      // A cell whose only value is in `mc:AlternateContent` is refused too: the copy cannot
      // carry that value, so the clear of the paste would lose it (`wrappedValueCell`).
      const wrapped = wrappedValueCell(part, box);
      if (wrapped)
        return refuse({ code: 'cut-wrapped-value', subject: `${target.entry.name}!${wrapped}` });
      // Any other clear refusal, such as a cell that the clear can neither keep nor replace
      // (`clearBlocker`) or a shared formula range that it would break, is also given at the
      // cut, not at the paste.
      const cleared = target.store.dryRun(ops);
      if (!cleared.ok) return refuse(refusalOf(cleared.rejection));
    }
    this.#count += 1;
    const token = `${this.#prefix}-${this.#count}`;
    const formulas: string[] = [];
    for (const cell of cells) if (cell.input.type === 'formula') formulas.push(cell.input.text);
    const formats = new Set(cells.map((cell) => cell.format));
    // A cut depends on the source sheet. A copied formula with a name that the source sheet
    // scopes depends on it for a paste on another sheet, which reads the names that it scopes.
    const scoped = usesLocalName(formulas, localNameKeys(pkg, index, target.entry.index));
    const source = cut ? target.entry.name : undefined;
    const scopedSheet = !cut && scoped ? target.entry.name : undefined;
    const deps = copyDeps(pkg, index, formulas, formats, source, scopedSheet);
    const named = deps.sheets.map((dep) => dep.name);
    this.#spent = null;
    const text = lines.join('\r\n');
    this.#snapshot = {
      token,
      text,
      store: target.store,
      sheet: target.entry.name,
      partName: target.entry.partName!,
      origin: start,
      rows,
      cols,
      cells,
      deps,
      references:
        formulas.length > 0
          ? { sheets: named, span: deps.span, at: target.store.referenceEdits(named, deps.span) }
          : null,
      cutAt: cut ? target.revision : null,
    };
    this.#markAt = target.revision;
    this.#host.repaint();
    // A cut changes no cell until the paste: the live region says that the cells wait to move.
    // A copy says what it copied, in a short text. A successful copy or cut is the next
    // command, so an earlier refusal no longer shows.
    this.#host.edits.succeeded();
    const range = selectionRef(target.selection);
    this.#host.announce(this.#host.t(cut ? 'grid.cut' : 'grid.copied', { ref: range }));
    return { ok: true, content: { text, token } };
  }

  /**
   * The token of the pending copy or cut whose text is `text`, or undefined. A paste that reads
   * only the text of the system clipboard, such as the menu bar's Paste, gives it with the text,
   * so the paste takes the lane that the paste key takes with the private type: a cut moves its
   * cells, and a copy keeps its formulas and cell formats. Only a copy whose outline shows
   * matches (`#marked`): after Escape, an edit, or a paste that ends the outline, and for an
   * expired copy, the text takes the plain lane, as text from elsewhere does. The line breaks
   * (`\r\n`, `\r`, or `\n`) and one trailing line break do not count. @internal
   */
  tokenOfText(text: string): string | undefined {
    this.#sweep();
    const snap = this.#snapshot;
    if (!snap || !this.#marked(snap)) return undefined;
    const lines = (t: string) => t.replace(/\r\n?/g, '\n').replace(/\n$/, '');
    return lines(snap.text) === lines(text) ? snap.token : undefined;
  }

  /** Paste clipboard data at the active cell, as one transaction and one undo step. */
  paste(data: PasteData): EditResult {
    const edits = this.#host.edits;
    if (edits.viewing) return edits.refuse({ code: 'view-mode' });
    const open = edits.commitOpenEdit();
    if (!open.ok) return open;
    const target = this.#host.target();
    if (!target?.entry.partName) return edits.refuse({ code: 'not-editable' });
    // A paste writes from the active cell over one range: a selection of several ranges is
    // refused, as a copy of it is.
    if (isMultiRange(target.selection)) return edits.refuse({ code: 'multiple-ranges' });
    this.#sweep();
    // An expired copy stays refused with its code until a new copy, cut, or reset.
    const spent = this.#spent;
    if (spent && data.token === spent.token && spent.store === target.store)
      return edits.refuse(spent.refusal);
    const snapshot = this.#snapshot;
    if (snapshot && data.token && data.token === snapshot.token && snapshot.store === target.store)
      return this.#pasteInternal(target, snapshot);
    if (data.text) {
      const waiting = snapshot?.cutAt != null && snapshot.cutAt === target.revision;
      const r = this.#pastePlain(target, data.text);
      // A plain paste never moves the cells of a cut, also when it holds the text of the cut,
      // as a clipboard that drops the private type gives it. So it ends a pending cut and says
      // that the cells stay, as Escape does.
      if (r.ok && waiting) {
        this.#forget();
        this.#host.announce(this.#host.t('grid.cutEnded'));
        this.#host.repaint();
      }
      return r;
    }
    if (data.html) return edits.refuse({ code: 'paste-html' });
    return { ok: true };
  }

  /**
   * Why a paste of `rows` by `cols` cells at the active cell is refused, or null. `grid-full`
   * names the target range, clipped to the grid.
   */
  #limits(target: EditTarget, rows: number, cols: number): EditRefusal | null {
    const at = target.selection.active;
    if (rows * cols > MAX_FORMAT_NEW_CELLS) return { code: 'range-too-large' };
    if (at.row + rows > MAX_ROWS || at.col + cols > MAX_COLS) {
      const end = {
        row: Math.min(at.row + rows, MAX_ROWS) - 1,
        col: Math.min(at.col + cols, MAX_COLS) - 1,
      };
      const ref = selectionRef({
        sheetIndex: 0,
        active: at,
        anchor: at,
        range: { start: at, end },
      });
      const subject = `${target.entry.name}!${ref}`;
      return { code: 'grid-full', subject, about: { kind: 'cell-content', name: subject } };
    }
    if (meetsMerge(this.#host.merges(), at.row, at.col, rows, cols))
      return { code: 'merged-range' };
    return null;
  }

  #pasteInternal(target: EditTarget, snap: Snapshot): EditResult {
    const edits = this.#host.edits;
    // A copied formula can name a sheet that changed after the copy, and a copied cell format
    // index can name another format.
    if (!depsHold(target.pkg, target.index, snap.deps, target.entry.name))
      return edits.refuse(
        this.#expire(snap, { code: 'copy-expired', subject: sourceSubject(snap) })
      );
    const refused = this.#limits(target, snap.rows, snap.cols);
    if (refused) return edits.refuse(refused);
    const at = target.selection.active;
    const sheet = target.entry.name;
    const ops: WorkbookOp[] = [];
    const home = target.index.sheets.find((s) => s.partName === snap.partName);
    const source = snap.cutAt !== null ? home : undefined;
    const otherSheet = home?.partName !== target.entry.partName;
    // A cut clears the source cells that the paste does not write: on the same sheet, the
    // writes cover the overlap, so a clear never changes a cell that the paste then writes.
    if (source) {
      const box: Box = [snap.origin.row, snap.origin.col, snap.rows, snap.cols];
      const parts = otherSheet ? [box] : boxesOutside(box, [at.row, at.col, snap.rows, snap.cols]);
      for (const [row, col, rows, cols] of parts) {
        const ref = selectionRef({
          sheetIndex: 0,
          active: { row, col },
          anchor: { row, col },
          range: { start: { row, col }, end: { row: row + rows - 1, col: col + cols - 1 } },
        });
        for (const op of clearOps(source.name, ref, 'all')) ops.push(op);
      }
    }
    const locals =
      otherSheet && home ? localNameKeys(target.pkg, target.index, home.index) : new Set<string>();
    const targetLocals = otherSheet
      ? localNameKeys(target.pkg, target.index, target.entry.index)
      : new Set<string>();
    const prefix = home ? sheetPrefixText([home.name]) : '';
    const offset = { rows: at.row - snap.origin.row, cols: at.col - snap.origin.col };
    const formats: WorkbookOp[] = [];
    for (let i = 0; i < snap.cells.length; i += 1) {
      const cell = snap.cells[i]!;
      const r = Math.floor(i / snap.cols);
      const c = i % snap.cols;
      const ref = formatCellRef({ row: at.row + r, col: at.col + c });
      let value = cell.input;
      // A copied formula moves by the offset. A cut moves the cell, so its formula keeps its
      // references (design D2: references to cut cells do not move); on another sheet, each
      // reference without a sheet gets the source sheet, so it keeps its cells. On another
      // sheet, a copy and a cut apply one name rule: a name that the source sheet scopes gets
      // the source sheet, and a name that only the target sheet scopes refuses the paste.
      if (value.type === 'formula' && (!source || otherSheet)) {
        let moved = source
          ? qualifyFormula(value.text, prefix, locals, targetLocals)
          : moveCopiedFormula(value.text, offset);
        if (moved.ok && !source && otherSheet)
          moved = qualifyFormula(moved.text, prefix, locals, targetLocals, false);
        if (!moved.ok) {
          const from = formatCellRef({ row: snap.origin.row + r, col: snap.origin.col + c });
          const subject = `${snap.sheet}!${from}`;
          return edits.refuse({ code: pasteProblemCode(moved.problem), subject });
        }
        value = { type: 'formula', text: moved.text };
      }
      // The store checks a written formula as typed input (`readFormulaInput`). A copied or
      // cut formula that fails the check refuses as a paste, with the source cell as subject,
      // not with the code of typed input, whose text gives advice for typing. A call to a
      // function outside the predefined functions keeps its own code and names the function.
      const typed = value.type === 'formula' ? readFormulaInput(value.text) : null;
      if (typed && !typed.ok) {
        if (typed.problem === 'unsupported-function')
          return edits.refuse({ code: 'unsupported-function', subject: typed.subject });
        const from = formatCellRef({ row: snap.origin.row + r, col: snap.origin.col + c });
        return edits.refuse({ code: 'paste-unreadable-formula', subject: `${snap.sheet}!${from}` });
      }
      ops.push({ op: 'setCellValue', sheet, ref, value });
      formats.push({ op: 'setCellFormatIndex', sheet, ref, index: cell.format });
    }
    for (const op of formats) ops.push(op);
    const marked = this.#marked(snap);
    // The source cells under a paste of a copy on its own sheet, to tell whether the paste
    // changed one of them.
    const shared =
      marked && !source && !otherSheet
        ? boxesShared(
            [at.row, at.col, snap.rows, snap.cols],
            [snap.origin.row, snap.origin.col, snap.rows, snap.cols]
          )
        : null;
    const before = shared ? cellsIn(target.pkg, snap.partName, shared) : [];
    // A paste of a cut uses the cut, so a repaint during the run does not see an expired cut.
    const cutAt = snap.cutAt;
    snap.cutAt = null;
    const r = edits.run(ops, pastedHistory(sheet, at, snap.rows, snap.cols), 'paste');
    if (!r.ok) {
      snap.cutAt = cutAt;
      return r;
    }
    // A paste of a cut ends its outline. A copy stays pending after its paste, unless the paste
    // changed a source cell, also at the source itself: a rich text cell, an error, and a cached
    // formula value do not stay as they were. A paste over a source cell that stays as it was
    // keeps the outline.
    const after = this.#host.target();
    const now = after?.revision ?? null;
    const changed =
      shared !== null && (!after || !sameCells(before, cellsIn(after.pkg, snap.partName, shared)));
    this.#markAt = marked && !source && !changed ? now : null;
    this.#selectPasted(at, snap.rows, snap.cols);
    this.#host.repaint();
    return r;
  }

  #pastePlain(target: EditTarget, text: string): EditResult {
    const edits = this.#host.edits;
    const at = target.selection.active;
    const rows = splitPlainText(text);
    if (!rows) return edits.refuse({ code: 'range-too-large' });
    let cols = 0;
    for (const line of rows) cols = Math.max(cols, line.length);
    const refused = this.#limits(target, rows.length, cols);
    if (refused) return edits.refuse(refused);
    const { pkg, index, entry } = target;
    const sheet = entry.name;
    const ops: WorkbookOp[] = [];
    // The cells that take the number format of their typed form: their format ops follow the
    // value ops, one for each rectangle of one format (`pastedFormatOps`).
    const formatted: FormattedCell[] = [];
    for (let r = 0; r < rows.length; r += 1) {
      const line = rows[r]!;
      for (let c = 0; c < line.length; c += 1) {
        const cell = { row: at.row + r, col: at.col + c };
        const typed = line[c]!;
        // Text that matches the content writes nothing, as the cell editor does.
        if (typed === cellContent(pkg, index, entry.partName, cell, this.#host.locale())) continue;
        const read = typedInput(pkg, index, entry.partName, cell, typed, this.#host.locale());
        ops.push({ op: 'setCellValue', sheet, ref: formatCellRef(cell), value: read.value });
        if (read.numberFormat) formatted.push({ at: cell, numberFormat: read.numberFormat });
      }
    }
    // A full styles part takes the values only (`formatsFit`), as a protected sheet does.
    const formats = pastedFormatOps(pkg, sheet, formatted);
    if (formatsFit(pkg, formats, formatted.length, (o) => target.store.dryRun(o)))
      for (const op of formats) ops.push(op);
    const r =
      ops.length > 0
        ? edits.run(ops, pastedHistory(sheet, at, rows.length, cols), 'paste')
        : ({ ok: true } as const);
    if (r.ok) this.#selectPasted(at, rows.length, cols);
    return r;
  }

  #selectPasted(at: GridPoint, rows: number, cols: number): void {
    this.#host.select(at, pastedEnd(at, rows, cols));
  }
}

/** The last cell of a pasted range of `rows` by `cols` cells at `at`. */
const pastedEnd = (at: GridPoint, rows: number, cols: number): GridPoint => ({
  row: at.row + rows - 1,
  col: at.col + cols - 1,
});

/**
 * The history of a paste: a redo selects the pasted range, as the paste does, and an undo
 * selects the selection before the paste.
 */
function pastedHistory(sheet: string, at: GridPoint, rows: number, cols: number): RunHistory {
  const end = pastedEnd(at, rows, cols);
  return {
    sheet,
    after: () => ({
      sheet,
      range: [at.row, at.col, end.row, end.col],
      active: { row: at.row, col: at.col },
      anchor: { row: at.row, col: at.col },
      focus: end,
    }),
  };
}

/** A range as its 0-based top row, left column, row count, and column count. */
type Box = readonly [number, number, number, number];

/** The cells that two boxes share, as a box, or null when they share none. */
function boxesShared(a: Box, b: Box): Box | null {
  const top = Math.max(a[0], b[0]);
  const left = Math.max(a[1], b[1]);
  const rows = Math.min(a[0] + a[2], b[0] + b[2]) - top;
  const cols = Math.min(a[1] + a[3], b[1] + b[3]) - left;
  return rows > 0 && cols > 0 ? [top, left, rows, cols] : null;
}

/** The cell elements of `box` on the sheet part `partName`, row by row, or undefined. */
function cellsIn(pkg: XlsxPackage, partName: string, box: Box): (Element | undefined)[] {
  const part = xmlPartOf(pkg, partName);
  if (!part) return [];
  const cells = sheetIndex(part).cells;
  const out: (Element | undefined)[] = [];
  for (let r = box[0]; r < box[0] + box[2]; r += 1)
    for (let c = box[1]; c < box[1] + box[3]; c += 1) out.push(cells.get(cellKey(r, c))?.el);
  return out;
}

/** Whether two lists of cells hold the same cells. */
function sameCells(a: readonly (Element | undefined)[], b: readonly (Element | undefined)[]) {
  return a.length === b.length && a.every((el, i) => sameNode(el, b[i]));
}

/** The parts of `box` outside `cover`, as at most four boxes: above, below, left, right. */
function boxesOutside(box: Box, cover: Box): Box[] {
  const [top, left, rows, cols] = box;
  const bottom = top + rows;
  const right = left + cols;
  const cTop = Math.max(top, cover[0]);
  const cLeft = Math.max(left, cover[1]);
  const cBottom = Math.min(bottom, cover[0] + cover[2]);
  const cRight = Math.min(right, cover[1] + cover[3]);
  if (cTop >= cBottom || cLeft >= cRight) return [box];
  const out: Box[] = [];
  if (cTop > top) out.push([top, left, cTop - top, cols]);
  if (bottom > cBottom) out.push([cBottom, left, bottom - cBottom, cols]);
  if (cLeft > left) out.push([cTop, left, cBottom - cTop, cLeft - left]);
  if (right > cRight) out.push([cTop, cRight, cBottom - cTop, right - cRight]);
  return out;
}
