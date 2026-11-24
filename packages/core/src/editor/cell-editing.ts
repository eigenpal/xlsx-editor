// The editing commands of the workbook editor (tasks 7.2, 7.4, and 7.5): enter a value, clear
// values, formats, or both, apply a format patch, toggle bold, italic, and underline, insert and
// delete rows and columns, undo, redo, and the open cell editor. Each command is one store
// transaction through `transact`, so one command is one revision and one undo step, with the side
// effects of designs D3 and D5 in the same transaction. A refused command returns its typed reason,
// keeps it for the snapshot, and announces it; it never reports `ok`.
//
// Redo and save while the cell editor is open commit the open edit first. If that commit is
// refused, the command is refused too, with the commit's reason. Undo commits nothing: it
// cancels an open edit with changed text (`undo`). Closing the surface commits
// the open edit, or keeps its refusal visible: an edit is never dropped without a word.
//
// The refusal has one announcement: the live region of the grid surface. Chrome may show the
// refusal from the snapshot, but must not announce it again. The same refusal twice is
// announced twice. A new selection or another sheet clears the refusal.
//
// In `view` mode every command here is refused with `view-mode` before it reads the workbook.

import type { GridPoint } from '../contracts/index.ts';
import {
  formatCellRef,
  type TransactResult,
  type WorkbookDocumentStore,
  type StructuralOpName,
  type WorkbookOp,
} from '../store/index.ts';
import type { NumberLocale } from '../cascade/index.ts';
import { protectionChecks } from '../store/store/sheet-protection.ts';
import {
  MAX_COLS,
  MAX_ROWS,
  cellChildBlocker,
  cellKey,
  clearRangeBlocker,
  protectionRefusal,
  sheetIndex,
  xmlPartOf,
  type CellOverride,
  type FormatPatch,
  type HistorySelection,
} from '../store/index.ts';
import {
  clearOps,
  formatOps,
  isWholeSheet,
  mergeOps,
  structureOp,
  togglePatch,
  unmergeOp,
  type FormatToggle,
  type MergeOptions,
} from '../automation/index.ts';
import { rangeRect, rangeRects, type ViewState } from './view-map.ts';
import { selectCell, type Merges } from './selection.ts';
import type { StageBox } from './grid-surface.ts';
import { CellEditor, editMoveStep, type CellEditMode, type EditMove } from './cell-editor.ts';
import { refusalOf, refusalText, type EditRefusal, type EditResult } from './cell-input.ts';
import { cellContent, resolvedFormatAt, selectionRef, typedInput } from './editor-model.ts';
import { decimalStepOps } from './decimal-selection.ts';
import { entryOps, formatsFit, lineBreakWrapOp } from './entry-ops.ts';

export type { CellFont, EditingHost, EditTarget, RunHistory } from './editing-host.ts';
import type { EditingHost, EditTarget, RunHistory } from './editing-host.ts';

/**
 * The refusals that every press of a key repeats alike: view mode, a locked cell, a cell that
 * no commit can change (`cell-alternate-content`, `shared-formula`), and an undo or a redo with
 * nothing to step, and any command during a load. Each is announced once, until another
 * message replaces it.
 */
const ONCE = new Set<EditRefusal['code']>([
  'view-mode',
  'locked',
  'cell-alternate-content',
  'shared-formula',
  'cut-wrapped-value',
  'nothing-to-undo',
  'nothing-to-redo',
  'loading',
]);

/**
 * The commands of this module whose `cell-alternate-content` or `shared-formula` refusal names
 * them (`refusalText`). `commit` is the commit of an open cell edit, which leaves the cell editor
 * open. The cut sets its own command (`clipboard.ts`).
 */
const NAMED_COMMANDS = new Set<EditRefusal['command']>([
  'clear',
  'clearAll',
  'paste',
  'merge',
  'commit',
]);
/** The codes whose refusal of a clear, a paste, a merge, or a commit names that command. */
const NAMES_COMMAND = new Set<EditRefusal['code']>(['cell-alternate-content', 'shared-formula']);

/** What a refusal stopped: the open of a cell editor (`begin`, a key) or a command. */
export type RefusedBy = 'open' | NonNullable<EditRefusal['command']>;

/** Appended to an announcement that equals the current one, so it is announced again. */
export const REANNOUNCE = ' ';

/** Whether a point is a cell of the grid: 0-based integer coordinates inside the bounds. */
export const inGrid = (p: GridPoint): boolean =>
  Number.isInteger(p.row) &&
  Number.isInteger(p.col) &&
  p.row >= 0 &&
  p.row < MAX_ROWS &&
  p.col >= 0 &&
  p.col < MAX_COLS;

/** The cell editing of each editor, for the engine's own callers, such as the toolbar. */
const editings = new WeakMap<object, CellEditing>();

/** Keep `edits` as the cell editing of `editor`. */
export function registerEditing(editor: object, edits: CellEditing): void {
  editings.set(editor, edits);
}

/** The cell editing of `editor`, when the engine made it. */
export function editingOf(editor: object): CellEditing | undefined {
  return editings.get(editor);
}

/**
 * Translate the labels of the surface of `editor` again and paint it once, after its catalog
 * changed: the grid label, the accessible stale words, and the cell editor label. For the
 * adapters, which change the locale. Returns whether the engine made `editor`. @internal
 */
export function relabelEditor(editor: object): boolean {
  const edits = editings.get(editor);
  edits?.relabel();
  return !!edits;
}

export class CellEditing {
  #editor: CellEditor | null = null;
  /**
   * The cell that the open cell editor edits, and the number locale that the edit opened with:
   * the edit reads and compares its text in that locale, and a new locale (`setNumberLocale`)
   * applies to the next edit.
   */
  #editing: { sheet: number; cell: GridPoint; locale: NumberLocale } | null = null;
  /** The edit has typed input: a first character, or an input event in the editor. */
  #typed = false;
  /** An input method composes in the formula bar (`typeInBar`). */
  #barComposing = false;
  /** The cell's content when the editor opened. */
  #openedWith = '';
  #refusal: EditRefusal | null = null;
  /** What `#refusal` stopped (`RefusedBy`), or undefined for another edit. */
  #refusedBy: RefusedBy | undefined = undefined;
  #dropped = false;
  /** The stored `edit-dropped` refusal of the last drop, which a settled load shows again. */
  #dropRefusal: EditRefusal | null = null;
  readonly #host: EditingHost;

  constructor(host: EditingHost) {
    this.#host = host;
  }

  get editing(): boolean {
    return this.#editing !== null;
  }

  /**
   * Whether the open edit has typed input, or text that differs from the cell's content when
   * the edit opened. Typing `1` over a cell that holds `1` is a change.
   */
  get changed(): boolean {
    if (this.#editing === null || !this.#editor) return false;
    return this.#typed || this.#editor.value !== this.#openedWith;
  }

  /** An input method composes in the open editor. */
  get composing(): boolean {
    return this.#editing !== null && (this.#barComposing || !!this.#editor?.composing);
  }

  /** The text of the open edit, as typed, or null when no edit is open. */
  get draft(): string | null {
    return this.#editing !== null && this.#editor ? this.#editor.value : null;
  }

  /**
   * The open edit as a cell that a commit changes: whether the cell then holds content. Null when
   * no edit is open, and when the text equals the content at the open, which a commit does not
   * write (`#write`). An empty text and a lone apostrophe commit an empty cell.
   */
  get pending(): CellOverride | null {
    const draft = this.draft;
    if (draft === null || draft === this.#openedWith) return null;
    const holds = draft !== '' && draft !== "'";
    return { row: this.#editing!.cell.row, col: this.#editing!.cell.col, holds };
  }

  /** Whether the draft held content at the last publish of a key. */
  #draftHolds: boolean | null = null;

  /**
   * Publish a key of the open edit. A draft that turns empty, or turns not empty, changes the
   * cells that a merge would empty (`selectionDropsValuesAt`), so it publishes the whole snapshot.
   */
  #publishDraft(): void {
    const holds = this.pending?.holds ?? null;
    if (holds !== this.#draftHolds) this.#host.publish();
    else this.#host.publishDraft();
    this.#draftHolds = holds;
  }

  get refusal(): EditRefusal | null {
    return this.#refusal;
  }

  /** The stored `edit-dropped` refusal while the drop shows (`dropped`), else null. */
  get dropRefusal(): EditRefusal | null {
    return this.#dropped ? this.#dropRefusal : null;
  }

  /** The workbook and the shown sheet that an edit applies to, or null when none is ready. */
  get target(): EditTarget | null {
    return this.#host.target();
  }

  /** The merged ranges of the shown sheet. */
  get merges(): Merges {
    return this.#host.merges();
  }

  /**
   * An edit with changed text was dropped, and no command, edit, selection change, or other
   * load followed. A later refusal can replace `edit-dropped` in the snapshot; this keeps the
   * drop visible.
   */
  get dropped(): boolean {
    return this.#dropped;
  }

  /** The open editor's input, when it takes the keyboard. */
  get input(): HTMLTextAreaElement | null {
    return this.#editor?.isOpen ? this.#editor.input : null;
  }

  /**
   * Label the cell editor from the catalog now, and paint the grid with its labels once. The
   * live region empties, so an announcement in the old catalog is not read again.
   */
  relabel(): void {
    this.#host.silence();
    this.#editor?.input.setAttribute('aria-label', this.#host.t('editing.cellEditor'));
    this.#host.render();
  }

  /** Mount the cell editor in the stage of a surface. */
  attach(stage: HTMLElement): void {
    this.detach();
    this.#editor = new CellEditor(stage, this.#host.t('editing.cellEditor'), {
      commit: (_text, move, inRange) => this.commitAndMove(move, inRange),
      cancel: () => this.cancel(),
      // Each key publishes the draft alone, so `editText` follows the typed text without a
      // recompute of the rest of the snapshot.
      change: (typed) => {
        if (typed) this.#typed = true;
        // A key in the cell editor ends a composition that the formula bar reported: the bar
        // lost the keyboard, and can lose it with no composition end (an unmount).
        if (typed) this.#barComposing = false;
        this.#publishDraft();
      },
      reveal: () => {
        if (!this.#editing) return;
        this.#host.reveal(this.#editing.cell);
        this.#host.render();
      },
    });
  }

  /**
   * Unmount the cell editor. An open edit is committed; when the store refuses it, the refusal
   * stays in the snapshot and is announced, and the text is not kept.
   */
  detach(): void {
    this.endEdit();
    this.#editor?.dispose();
    this.#editor = null;
  }

  /**
   * Close the open edit: commit it, and when the store refuses it, keep the refusal in the
   * snapshot and announce it. The text is not kept. Called before the surface unmounts.
   */
  endEdit(): void {
    const editing = this.#editing;
    // Text that an input method still composes is not converted: it is dropped, not stored,
    // and the live region says so.
    if (editing && this.composing) {
      this.drop();
      this.#host.publish();
      return;
    }
    if (editing && this.#editor) this.#write(this.#editor.value, editing.cell, editing.locale);
    this.#close();
    if (editing) this.#host.publish();
  }

  /** Close the editor without a commit, and forget the refusal: a new workbook opened. */
  reset(): void {
    this.#close();
    this.#refusal = null;
    this.#dropped = false;
  }

  /**
   * The load settled, ready or failed: close the editor and forget the refusals of the load,
   * but keep `edit-dropped` (and `dropped`) when the load dropped an edit with changed text,
   * so the user still sees that the typed text was not stored. It is not announced again. It is
   * the object stored at the drop, also when a `loading` refusal replaced it during the load, so
   * the `rejection` event, which fires once for each refusal object, does not fire again.
   */
  settleLoad(): void {
    this.#close();
    // `#dropped` is set only by `drop`, which sets `#dropRefusal` too.
    this.#refusal = this.#dropped ? this.#dropRefusal : null;
  }

  /** A save wrote `bytes`: forget an older refusal, which no longer applies. */
  saved<T>(bytes: T): T {
    // A save is the next command, so the drop is no longer shown. `#dropped` is true only while
    // `#refusal` is set, so a save with no refusal has nothing to clear or publish.
    this.#dropped = false;
    if (!this.#refusal) return bytes;
    this.#refusal = null;
    this.#host.publish();
    return bytes;
  }

  /**
   * Drop the open edit for another workbook or none, or a surface that closes while an input
   * method composes, without a commit. An edit with changed text leaves the refusal
   * `edit-dropped` in the snapshot, which listeners see at once and the live region announces.
   * The status message shows it: under the loading text while a load runs, and in the status
   * strip, or under the empty or failure text, after the load settles (`settleLoad`). It stays
   * until the next command, edit, or selection change, or the next load.
   */
  drop(): void {
    const changed = this.changed;
    this.reset();
    if (!changed) return;
    this.#dropped = true;
    this.#refuse({ code: 'edit-dropped' });
    this.#dropRefusal = this.#refusal;
  }

  /**
   * Refuse a command with `loading` while a workbook loads: keep the refusal, announce it, and
   * publish, as for any engine refusal. The toolbar refuses its own slots through this, so its
   * result, the snapshot, and the live region agree.
   */
  refuseLoading(): EditResult {
    return this.#refuse({ code: 'loading' });
  }

  /**
   * Refuse a save: keep `reason` (the refusal of the open edit's commit), or `not-editable`
   * when no workbook is open, or `loading` while a workbook loads, as a refusal with
   * `command: 'save'`, which `refusalText` reads as a save that did not happen, and announce
   * it. Returns null, the bytes of no save.
   */
  refuseSave(reason: EditRefusal | null): null {
    this.#refuse({ ...(reason ?? { code: 'not-editable' }), command: 'save' });
    return null;
  }

  /** Forget the refusal: the selection or the sheet changed. The caller publishes. */
  clearRefusal(): void {
    this.#refusal = null;
    this.#dropped = false;
  }

  /**
   * A command that succeeded with no transaction, such as a copy or a cut: an earlier refusal
   * no longer shows, as after a transaction that applied.
   */
  succeeded(): void {
    if (!this.#refusal) return;
    this.#refusal = null;
    this.#dropped = false;
    this.#host.publish();
  }

  // ------------------------------------------------------------------ commands

  /**
   * Write `text` to the active cell or to `at`. An open edit of that cell commits first, so
   * its text is not lost; a refused commit refuses the write.
   */
  enterValue(text: string, at?: GridPoint): EditResult {
    const editing = this.#editing;
    if (editing && (!at || (at.row === editing.cell.row && at.col === editing.cell.col))) {
      const open = this.commitOpenEdit();
      if (!open.ok) return open;
    }
    return this.#write(text, at);
  }

  #write(
    text: string,
    at?: GridPoint,
    locale = this.#host.locale(),
    command?: 'commit'
  ): EditResult {
    if (this.#host.viewing()) return this.#refuse({ code: 'view-mode' });
    const target = this.#host.target();
    if (!target?.entry.partName) return this.#refuse({ code: 'not-editable' });
    const cell = at ?? target.selection.active;
    // A cell outside the grid, or with a coordinate that is not an integer, is refused.
    if (!inGrid(cell)) return this.#refuse({ code: 'bad-reference' });
    const { pkg, index } = target;
    const part = target.entry.partName;
    const current = cellContent(pkg, index, part, cell, locale);
    // Text that matches the content writes nothing: no revision and no undo step.
    if (text === current)
      return this.#settle(undefined, { ok: true, revision: target.store.revision });
    // A date cell (`t="d"`) shows its ISO 8601 text; ISO text typed in it stays a date.
    // A typed form such as `12%` also sets its number format on a General cell, in the same
    // transaction, so one undo restores the value and the format.
    const entry = typedInput(pkg, index, part, cell, text, locale);
    const ref = formatCellRef(cell);
    const { value, format } = entryOps(pkg, target.entry.name, ref, entry);
    // Text that the cell editor commits with a line break turns on wrap text, so each line
    // shows; a value written by a command keeps the cell format.
    const wrap =
      command === 'commit'
        ? lineBreakWrapOp(
            pkg,
            target.entry.name,
            ref,
            entry,
            resolvedFormatAt(pkg, index, part, cell).alignment.wrapText
          )
        : undefined;
    const formats = [format, wrap].filter((o): o is WorkbookOp => o !== undefined);
    // A full styles part takes the value only (`formatsFit`), as a protected sheet does.
    const fits =
      formats.length > 0 && formatsFit(pkg, formats, formats.length, (o) => target.store.dryRun(o));
    const ops = fits ? [value, ...formats] : [value];
    // A value written to another cell than the active one is not a command on the selection:
    // its history entry keeps no selection, so undo and redo leave the selection alone.
    const own =
      !at || (at.row === target.selection.active.row && at.col === target.selection.active.col);
    return this.#transact(target, ops, command, own);
  }

  /** Clear the values of the selection. An open edit is committed first, as for `save`. */
  clearValues(): EditResult {
    if (this.#host.viewing()) return this.#refuse({ code: 'view-mode' });
    const open = this.commitOpenEdit();
    if (!open.ok) return open;
    const target = this.#host.target();
    if (!target) return this.#refuse({ code: 'not-editable' });
    const ref = selectionRef(target.selection);
    return this.#transact(target, [{ op: 'clearValues', sheet: target.entry.name, ref }], 'clear');
  }

  /**
   * Clear the values, the formats, or both, of the selection, as one transaction. An open edit
   * is committed first; a refused commit refuses the command.
   */
  clear(kind: 'formats' | 'all'): EditResult {
    if (this.#host.viewing()) return this.#refuse({ code: 'view-mode' });
    const open = this.commitOpenEdit();
    if (!open.ok) return open;
    const target = this.#host.target();
    if (!target) return this.#refuse({ code: 'not-editable' });
    return this.#transact(
      target,
      clearOps(target.entry.name, selectionRef(target.selection), kind),
      kind === 'all' ? 'clearAll' : 'clear'
    );
  }

  /**
   * Apply a format patch to the selection: the active cell when it is the only cell. An open
   * edit is committed first; a refused commit refuses the command. A number format change
   * confirms through `EditingHost.confirm`, unless it commits typed text: the live region then
   * says the new text, also when only its spaces change.
   */
  applyFormat(patch: FormatPatch): EditResult {
    if (this.#host.viewing()) return this.#refuse({ code: 'view-mode' });
    const typed = this.changed;
    const open = this.commitOpenEdit();
    if (!open.ok) return open;
    const target = this.#host.target();
    if (!target) return this.#refuse({ code: 'not-editable' });
    const r = formatOps(target.entry.name, selectionRef(target.selection), patch);
    if (!r.ok) return this.#refuse(refusalOf(r.rejection));
    const done = this.#transact(target, r.ops, 'format');
    if (done.ok && patch.numberFormat !== undefined && !typed)
      this.#host.confirm?.(this.#host.t('grid.numberFormatChanged'));
    return done;
  }

  /**
   * Add (`step` 1) or remove (`step` -1) one decimal place in each cell of the selection, from the
   * cell's own number format (`decimalStepOps`, which states the `General` column rule), in one
   * transaction and one undo step. A cell that cannot step keeps its format. While an input method
   * composes, the step is refused with `composing`. An open edit with changed text refuses the step
   * with `edit-pending`; an open edit with no change closes first. A sheet that does not allow
   * formats refuses it with `locked`, before any cell is read. When no cell can step, the step is
   * refused with `decimals-unchanged`, whose `reason` is the disabled reason of the control. Past
   * `MAX_DECIMAL_STEP_CELLS` cells that exist, it is refused with `decimals-too-large`. A step that changes cells confirms through
   * `EditingHost.confirm`, so the live region says that the places changed when the active cell
   * keeps its words.
   */
  stepDecimals(step: 1 | -1): EditResult {
    if (this.#host.viewing()) return this.#refuse({ code: 'view-mode' });
    // The checks come in the order of the slot state: the input method, the changed text, and
    // the protection. The step reads the stored cells, which changed text would change.
    if (this.composing) return this.#refuse({ code: 'composing' });
    if (this.changed) return this.#refuse({ code: 'edit-pending' });
    const open = this.commitOpenEdit();
    if (!open.ok) return open;
    const target = this.#host.target();
    if (!target) return this.#refuse({ code: 'not-editable' });
    const { pkg, index, entry, selection } = target;
    // A sheet that does not allow formats refuses the step before it reads any cell, as the
    // slot state does, also when no cell could step.
    const ref = formatCellRef(selection.active);
    const lock = protectionRefusal(pkg, { op: 'applyFormat', sheet: entry.name, ref, patch: {} });
    if (lock) return this.#refuse({ ...refusalOf(lock), command: 'format' });
    const shown = this.#host.shownGeneral?.();
    const r = decimalStepOps(pkg, index, entry, selection, this.#host.locale(), step, shown);
    if (!r.ok) return this.#refuse(r.refusal);
    const done = this.#transact(target, r.ops, 'format');
    const said = step > 0 ? 'grid.decimalsIncreased' : 'grid.decimalsDecreased';
    if (done.ok) this.#host.confirm?.(this.#host.t(said));
    return done;
  }

  /**
   * Toggle bold, italic, or underline on the selection, from the format of the active cell:
   * the format that the snapshot shows (`activeFormat`).
   */
  toggleFormat(which: FormatToggle): EditResult {
    if (this.#host.viewing()) return this.#refuse({ code: 'view-mode' });
    const open = this.commitOpenEdit();
    if (!open.ok) return open;
    const target = this.#host.target();
    if (!target) return this.#refuse({ code: 'not-editable' });
    const at = target.selection.active;
    const format = resolvedFormatAt(target.pkg, target.index, target.entry.partName, at);
    return this.applyFormat(togglePatch(format, which));
  }

  /**
   * Insert or delete the rows or columns of the selection (task 7.5): the rows of a whole-row
   * selection or of a range, or the active cell's row (`structureOp`). The open cell editor
   * commits first; a refused commit refuses the command.
   */
  structure(name: StructuralOpName): EditResult {
    if (this.#host.viewing()) return this.#refuse({ code: 'view-mode' });
    const open = this.commitOpenEdit();
    if (!open.ok) return open;
    const target = this.#host.target();
    if (!target) return this.#refuse({ code: 'not-editable' });
    if (isWholeSheet(target.selection)) return this.#refuse({ code: 'whole-sheet' });
    return this.#transact(
      target,
      [structureOp(target.entry.name, target.selection, name)],
      'structure'
    );
  }

  /**
   * Merge the range of the selection into one cell, or unmerge the merged ranges that it meets
   * (`merge-commands.ts`). The open cell editor commits first; a refused commit refuses the
   * command. After a merge, a redo selects the merged cell with its top-left cell active.
   */
  merge(options: MergeOptions | 'unmerge'): EditResult {
    if (this.#host.viewing()) return this.#refuse({ code: 'view-mode' });
    const open = this.commitOpenEdit();
    if (!open.ok) return open;
    const target = this.#host.target();
    if (!target) return this.#refuse({ code: 'not-editable' });
    const name = target.entry.name;
    if (options === 'unmerge')
      return this.#transact(target, [unmergeOp(name, target.selection)], undefined);
    const after = (b: HistorySelection): HistorySelection => {
      const [top, left, bottom, right] = b.range;
      const at = { row: top, col: left };
      return { ...b, active: at, anchor: at, focus: { row: bottom, col: right } };
    };
    const ops = mergeOps(name, target.selection, options);
    // The merge rule of a protected sheet refuses first, also before the clear of `dropValues`.
    const locked = protectionRefusal(
      target.pkg,
      ops.find((op) => op.op === 'mergeCells')!
    );
    if (locked) return this.#refuse(refusalOf(locked));
    return this.#transact(target, ops, 'merge', true, { sheet: name, after });
  }

  /**
   * Apply ops that do not depend on the selection (the sheet commands, `sheet-commands.ts`, the
   * resize, and the paste) as one transaction and one undo step. An open edit is committed first;
   * a refused commit refuses the command. A refusal shows through the refusal channel, and a
   * `locked` refusal keeps `command`, so its message names the command. `check` refuses the
   * command after the view mode, the open edit, and the protection of each op, and before the
   * store applies the ops, so a command gives the same refusal in each state as the store does.
   */
  run(
    ops: readonly WorkbookOp[],
    history?: RunHistory,
    command?: EditRefusal['command'],
    check?: (target: EditTarget) => EditRefusal | null
  ): EditResult {
    if (this.#host.viewing()) return this.#refuse({ code: 'view-mode' });
    const open = this.commitOpenEdit();
    if (!open.ok) return open;
    const target = this.#host.target();
    if (!target) return this.#refuse({ code: 'not-editable' });
    if (check) {
      const protectionRefusalOf = protectionChecks();
      for (const op of ops) {
        const locked = protectionRefusalOf(target.pkg, op);
        if (locked) return this.#settle(command, { ok: false, rejection: locked });
      }
      const refusal = check(target);
      if (refusal) return this.#refuse(refusal);
    }
    return this.#transact(target, ops, command, true, history);
  }

  /** Whether the editor shows the workbook in view mode, where no command changes it. */
  get viewing(): boolean {
    return this.#host.viewing();
  }

  /** Refuse a picked file that did not open: an edit that the pick's load dropped stays shown. */
  refuseOpen(): void {
    const dropped = this.#dropped;
    this.#refuse({ code: 'open-failed' });
    this.#dropped = dropped;
    this.#host.publish();
  }

  /** Refuse a command before it reaches the store, through the refusal channel. */
  refuse(refusal: EditRefusal, by?: RefusedBy): EditResult {
    return this.#refuse(refusal, by);
  }

  /**
   * Undo, the command of the toolbar and the editor instance. The history keys in an open cell
   * editor never reach it: they belong to the input's own undo. An open edit commits nothing.
   * An edit with changed text is cancelled, the cell shows its content again, and the history
   * does not move. An edit with no changed text closes, and the undo steps the history once.
   * While an input method composes, undo is refused with `composing`.
   */
  undo(): EditResult {
    if (this.#host.viewing()) return this.#refuse({ code: 'view-mode' });
    // No history steps while no workbook is ready, also while another one loads.
    if (!this.#host.target()) return this.#refuse({ code: 'not-editable' });
    if (this.composing) return this.#refuse({ code: 'composing' });
    if (this.changed) {
      this.cancel();
      this.#host.say(this.#host.t('grid.typingUndone'));
      return { ok: true };
    }
    this.cancel();
    const store = this.#host.target()?.store;
    if (!store?.canUndo) return this.#refuse({ code: 'nothing-to-undo' });
    return this.#step(store, 'undo');
  }

  /**
   * Redo. An open edit with changed text refuses it with `edit-pending`, and an open edit with
   * no change closes first.
   */
  redo(): EditResult {
    if (this.#host.viewing()) return this.#refuse({ code: 'view-mode' });
    if (!this.#host.target()) return this.#refuse({ code: 'not-editable' });
    if (this.composing) return this.#refuse({ code: 'composing' });
    // A commit of changed text would clear the redo steps: the redo is refused before any write,
    // as the toolbar disables it (`editPending`). An edit with no change closes.
    if (this.changed) return this.#refuse({ code: 'edit-pending' });
    this.cancel();
    const store = this.#host.target()?.store;
    if (!store?.canRedo) return this.#refuse({ code: 'nothing-to-redo' });
    return this.#step(store, 'redo');
  }

  /**
   * Undo or redo, then hand the host the selection that the entry keeps: the selection before
   * its command for undo, after it for redo, or null when the command supplied none.
   */
  #step(store: WorkbookDocumentStore, step: 'undo' | 'redo'): EditResult {
    const undo = step === 'undo';
    const selection = undo ? store.selectionForUndo() : store.selectionForRedo();
    const sheet = undo ? store.sheetForUndo() : store.sheetForRedo();
    const label = undo ? store.labelForUndo() : store.labelForRedo();
    // The history keeps the name of a renamed sheet before the rename. A rename keeps the
    // position of its sheet, so the name at that position after a redo is the new name.
    const before = this.#host.target()?.index;
    const at = before?.sheets.findIndex((s) => s.name === sheet) ?? -1;
    const r = this.#settle(undefined, undo ? store.undo() : store.redo());
    if (!r.ok) return r;
    const renamed = !undo && label === 'renameSheet' && at >= 0;
    const changed = renamed ? (this.#host.target()?.index.sheets[at]?.name ?? sheet) : sheet;
    this.#host.restore(step, selection, sheet, label, changed);
    return r;
  }

  /** Commit the open edit, if any, and keep the active cell. */
  commitOpenEdit(): EditResult {
    return this.commitAndMove(null) ? { ok: true } : { ok: false, refusal: this.#refusal! };
  }

  /**
   * Open the cell editor over the active cell, with `text` or the cell's content. Without
   * `focus`, the keyboard stays where it is: the formula bar opens the edit so.
   */
  begin(
    text?: string,
    mode: CellEditMode = text === undefined ? 'edit' : 'enter',
    focus = true
  ): EditResult {
    if (this.#host.viewing()) return this.#refuse({ code: 'view-mode' }, 'open');
    const target = this.#host.target();
    // A sheet without a part can never commit: refuse before the editor opens.
    if (!this.#editor || !target?.entry.partName) return this.#refuse({ code: 'not-editable' });
    if (this.#editing) return { ok: true };
    const cell = target.selection.active;
    // A locked cell of a protected sheet opens no editor: the commit could never apply.
    const locked = protectionRefusal(target.pkg, {
      op: 'setCellValue',
      sheet: target.entry.name,
      ref: formatCellRef(cell),
      value: { type: 'empty' },
    });
    if (locked) return this.#refuse(refusalOf(locked), 'open');
    // So does a cell with content that no edit can keep or replace
    // (`cellChildBlocker`): every commit of it is refused.
    const part = xmlPartOf(target.pkg, target.entry.partName);
    const present = part && sheetIndex(part).cells.get(cellKey(cell.row, cell.col));
    const blocked = present && cellChildBlocker(target.entry.name, formatCellRef(cell), present.el);
    if (blocked) return this.#refuse(refusalOf(blocked), 'open');
    // So does a cell of a formula range that no write can change (`clearRangeBlocker`): a cell
    // of an array or data table range, or a shared master that cannot unshare. A child formula
    // that would move out of the grid is checked only at the commit.
    const box = [cell.row, cell.col, cell.row, cell.col] as const;
    const range = part && clearRangeBlocker(part, target.entry.name, box);
    if (range) return this.#refuse(refusalOf(range), 'open');
    const locale = this.#host.locale();
    this.#editing = { sheet: target.sheet, cell, locale };
    // The open publishes the whole snapshot: the first key compares with no draft.
    this.#draftHolds = null;
    const current = cellContent(target.pkg, target.index, target.entry.partName, cell, locale);
    const content = text ?? current;
    this.#openedWith = current;
    this.#typed = text !== undefined && mode === 'enter';
    this.#host.reveal(cell);
    this.#editor.open(content, mode, focus);
    // An earlier refusal is not a refusal of this edit, so it clears, as a new selection clears
    // it. A dropped edit stays announced until the next command.
    if (!this.#dropped) this.#refusal = null;
    this.#host.publish();
    this.#host.render();
    return { ok: true };
  }

  /** Commit the open editor. A refused commit keeps the editor open with its text. */
  commit(): EditResult {
    return this.commitOpenEdit();
  }

  /**
   * The formula bar holds `text`: it becomes the text of the open edit, which the cell editor
   * shows in the cell. With no open edit, it opens one over the active cell first, in `edit`
   * mode and without the focus, as `begin` refuses it. One edit has one text: the bar never
   * keeps a copy. `composing` is whether an input method composes in the bar; while it does,
   * each command that would commit the edit is refused with `composing`.
   */
  typeInBar(text: string, composing: boolean): EditResult {
    if (typeof text !== 'string') throw new TypeError('The edit text is not a string');
    if (!this.#editing) {
      const opened = this.begin(text, 'edit', false);
      if (!opened.ok) return opened;
    }
    this.#editor?.setText(text);
    this.#typed = true;
    this.#barComposing = composing;
    this.#publishDraft();
    this.#host.render();
    return { ok: true };
  }

  /**
   * Commit the open edit, then move the active cell one step in the direction `edit`, inside a
   * multi-cell selection, as Enter and Tab do in the cell editor. With no open edit, only move
   * it. The grid takes the keyboard when the commit applies; a refused commit keeps the edit
   * open with its text. A value that is not an `EditMove` throws a `TypeError` before anything
   * changes.
   */
  commitMoving(edit: EditMove): EditResult {
    const move = editMoveStep(edit);
    if (this.composing) return this.#refuse({ code: 'composing' });
    if (this.#editing) {
      return this.commitAndMove(move, true) ? { ok: true } : { ok: false, refusal: this.#refusal! };
    }
    const target = this.#host.target();
    if (!target) return this.#refuse({ code: 'not-editable' });
    // The move says the next cell first, so the focus that enters the grid says nothing more.
    this.#host.moveFrom(target.selection.active, move, true);
    this.#host.surface()?.grid.focus({ preventScroll: true });
    return { ok: true };
  }

  cancel(): void {
    if (!this.#editing) return;
    this.#close();
    // The refusal clears, and so does its announcement, so the live region holds no stale text.
    if (this.#refusal) this.#host.silence();
    this.#refusal = null;
    this.#dropped = false;
    this.#host.surface()?.grid.focus({ preventScroll: true });
    this.#host.publish();
    this.#host.schedule();
  }

  /** Commit the open editor and move the active cell; false when the store refused. */
  commitAndMove(move: readonly [number, number] | null, inRange = false): boolean {
    const editing = this.#editing;
    if (!editing || !this.#editor) return true;
    // The composed text belongs to the input method: no command commits it until it ends.
    if (this.composing) {
      this.#refuse({ code: 'composing' });
      return false;
    }
    const r = this.#write(this.#editor.value, editing.cell, editing.locale, 'commit');
    if (!r.ok) {
      // The keyboard stays where the user typed: the cell editor takes it back from the grid,
      // and a field outside the grid, such as the formula bar, keeps it.
      if (this.#keyboardInGrid()) this.#editor.input.focus({ preventScroll: true });
      return false;
    }
    this.#close();
    // The move says the next cell first, so the focus that enters the grid says nothing more.
    if (move) this.#host.moveFrom(editing.cell, move, inRange);
    this.#host.surface()?.grid.focus({ preventScroll: true });
    if (!move) {
      this.#host.publish();
      this.#host.schedule();
    }
    return true;
  }

  /**
   * Place the cell editor over the whole of its cell in the painted window of sheet `sheet`,
   * in the cell's font at the zoom. When no part of the cell is in the window, the editor stays
   * focused but is not seen.
   */
  place(sheet: number, mapView: ViewState, box: StageBox): void {
    const editing = this.#editing;
    if (!this.#editor || !editing || editing.sheet !== sheet) return;
    const range = selectCell(sheet, editing.cell, this.#host.merges()).range;
    const rect = rangeRect(mapView, range);
    const font = this.#host.cellFont(editing.cell);
    this.#editor.place(
      {
        left: box.headWidth + rect.x * box.scale,
        top: box.headHeight + rect.y * box.scale,
        width: rect.width * box.scale,
        height: rect.height * box.scale,
        fontSize: font.sizePt * box.scale,
        fontFamily: font.family,
      },
      rangeRects(mapView, range).length > 0
    );
  }

  /** Whether the keyboard is in the grid surface, its cell editor, or nowhere in particular. */
  #keyboardInGrid(): boolean {
    const doc = this.#editor?.input.ownerDocument;
    const active = doc?.activeElement;
    const grid = this.#host.surface()?.grid;
    return (
      !active || active === doc.body || active === this.#editor?.input || !!grid?.contains(active)
    );
  }

  #close(): void {
    this.#editing = null;
    this.#openedWith = '';
    this.#typed = false;
    this.#barComposing = false;
    this.#editor?.close();
  }

  // ------------------------------------------------------------------ results

  /**
   * Commit the ops of `command` (none for a cell edit), which a `locked` refusal names. With
   * `selected`, the history entry keeps the selection of the command as the selection before
   * it and after it, so undo and redo select it again. A move of the active cell after a
   * commit is not part of the command.
   */
  #transact(
    target: EditTarget,
    ops: readonly WorkbookOp[],
    command: EditRefusal['command'],
    selected = true,
    history?: RunHistory
  ): EditResult {
    const { range, active, anchor } = target.selection;
    const selection: HistorySelection = {
      sheet: target.entry.name,
      range: [range.start.row, range.start.col, range.end.row, range.end.col],
      active,
      anchor,
      focus: target.focus,
    };
    const after = history ? history.after(selection) : selection;
    // The ops name cells by reference at the target's revision: a later revision refuses them.
    return this.#settle(
      command,
      target.store.transact(ops, {
        label: ops[0]!.op,
        baseRevision: target.revision,
        // No cell command changes the selection: a value, a format, a clear, and an insert or a
        // delete of rows or columns all keep it, so the selection after equals the one before.
        // A sheet op gives its own (`RunHistory`), and a paste names its target sheet.
        sheet: history?.sheet ?? target.entry.name,
        ...(selected ? { selectionBefore: selection } : {}),
        ...(selected && after ? { selectionAfter: after } : {}),
      })
    );
  }

  /** One transaction as an edit result: a refusal is kept, shown, and announced. */
  #settle(command: EditRefusal['command'], r: TransactResult): EditResult {
    if (!r.ok) {
      const refusal = refusalOf(r.rejection);
      // A `locked` refusal names the command that it stopped, and so does a
      // `cell-alternate-content` or `shared-formula` refusal of a clear, a paste, a merge, or a
      // commit. A `locked` commit is a value edit: its rule names the message.
      const named =
        command &&
        ((refusal.code === 'locked' && command !== 'commit') ||
          (NAMES_COMMAND.has(refusal.code) && NAMED_COMMANDS.has(command)));
      return this.#refuse(named ? { ...refusal, command } : refusal, command);
    }
    if (this.#refusal) {
      this.#refusal = null;
      this.#dropped = false;
      this.#host.publish();
    }
    return { ok: true };
  }

  #refuse(given: EditRefusal, by?: RefusedBy): EditResult {
    // An object with no name in the file is named by a generic label.
    const refused =
      given.code === 'resize-anchor' && given.subject === undefined
        ? { ...given, subject: this.#host.t('editing.anchoredObject') }
        : given;
    // While a workbook loads, every command is refused with `loading`, whatever stopped it; a
    // save keeps `command`. The edit that the load dropped keeps `edit-dropped`.
    // A refused command after the load settled is the next command: the drop is no longer
    // shown. While a load runs, the drop stays shown with the `loading` refusal.
    if (!this.#host.loading() && refused.code !== 'edit-dropped') this.#dropped = false;
    // A new object for each refused command, also for a shared constant refusal, so the
    // `rejection` event (one for each refusal object) reports each one.
    const refusal: EditRefusal =
      this.#host.loading() && refused.code !== 'edit-dropped'
        ? { code: 'loading', ...(refused.command ? { command: refused.command } : {}) }
        : { ...refused };
    // A refusal that each key repeats alike (`ONCE`) is announced once, not again for each
    // key, until a selection or another refusal replaces it. Another message is another
    // refusal: a `locked` refusal by another rule, or with another subject, is announced. So
    // is a refusal of another command, or of a key after a command (`by`).
    const text = refusalText(refusal, this.#host.t, this.#host.locale());
    const repeat =
      ONCE.has(refusal.code) &&
      this.#refusal !== null &&
      this.#refusedBy === by &&
      this.#refusal.code === refusal.code &&
      refusalText(this.#refusal, this.#host.t, this.#host.locale()) === text;
    // Repeated presses during a load change nothing that shows, so they take no time from it.
    if (repeat && refusal.code === 'loading') return { ok: false, refusal };
    this.#refusal = refusal;
    this.#refusedBy = by;
    if (!repeat) this.#host.say(text);
    this.#host.publish();
    return { ok: false, refusal };
  }
}
