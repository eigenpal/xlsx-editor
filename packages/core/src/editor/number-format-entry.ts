// The custom number format entry, DOM-free: the check of a typed format code, its preview on
// the active cell's value, and the custom codes that the workbook holds (ECMA-376 Part 1,
// 18.8.30 and 18.8.31). The chrome's number format dialog reads these, and applies a code
// through the `number.custom` slot (`runToolbarCommand`), so the store writes the
// `x:numFmt` (an equal code reuses its id) in one transaction and one undo step.
//
// One parser decides: the check (`number-format-code.ts`) runs the store's own checks of a
// format patch, the checks that the apply runs, and the preview
// formats through the display path that the grid paints with (`displayValue`). A code that
// the check refuses is never applied and has no preview.
//
// The workbook's codes are file-derived text: the chrome renders them as text only, and this
// module caps their count.

import {
  cellKey,
  FIRST_CUSTOM_NUMFMT_ID,
  lowerNumberFormat,
  sheetIndex,
  styleLists,
  xmlPartOf,
} from '../store/index.ts';
import {
  builtinFormatCode,
  cellValue,
  displayValue,
  sameNumberLocale,
  sheetStale,
  formatValue,
  type CellValue,
  type FormattedValue,
  type NumberLocale,
} from '../cascade/index.ts';
import { editingOf } from './cell-editing.ts';
import { numberFormatCodeProblem, type NumberFormatCodeProblem } from './number-format-code.ts';
import { resolvedFormatAt } from './editor-model.ts';
import { selectCell } from './selection.ts';
import { toolbarCommandState } from './toolbar-commands.ts';
import type { XlsxEditorInstance } from './xlsx-editor.ts';

/**
 * The value that a preview formats when the active cell holds nothing: a number with a group,
 * decimals, and a sign, so the preview shows each part of a number format.
 *
 * @public
 */
export const NUMBER_FORMAT_SAMPLE = -1234.5678;

/**
 * A preview of a number format code: the text that the grid shows before it fits the value to
 * the column, and the `RRGGBB` color of the code's section. `sample` is true when the active
 * cell is empty, and the text formats `NUMBER_FORMAT_SAMPLE`. `stale` is true for a formula
 * result that an edit can have changed. `color` is always six hex digits from the engine's
 * fixed palette of named and indexed colors, never text from the file.
 *
 * @public
 */
export type NumberFormatPreview =
  | {
      readonly ok: true;
      readonly text: string;
      readonly color?: string;
      readonly sample: boolean;
      /** The value is a formula result that an edit can have changed (not recalculated). */
      readonly stale?: boolean;
    }
  | { readonly ok: false; readonly problem: NumberFormatCodeProblem };

/**
 * The preview of a code on the active cell's value, through the display path that the grid
 * paints with, in the editor's number locale and the workbook's date system. A code that the
 * store refuses has its problem and no text. With no workbook ready, the preview formats
 * `NUMBER_FORMAT_SAMPLE` in the 1900 date system. A rich text cell shows its stored text for
 * every code, because the grid paints rich text with no number format.
 *
 * @public
 */
export function previewNumberFormat(editor: XlsxEditorInstance, code: string): NumberFormatPreview {
  const problem = numberFormatCodeProblem(code);
  if (problem) return { ok: false, problem };
  // The apply writes a code that equals a built-in code as its id, and the grid shows an id
  // through the locale's code for it (`NumberLocale.builtins`): the preview does the same.
  const target = editingOf(editor)?.target;
  const defined = target ? styleLists(target.pkg, target.index).numFmts : new Map<number, string>();
  const lowered = lowerNumberFormat({ numberFormat: { code } }, defined);
  const nf = 'patch' in lowered ? lowered.patch.numberFormat : undefined;
  const id = nf && 'id' in nf ? nf.id : undefined;
  const locale = editor.snapshot().numberLocale;
  const done = (shown: FormattedValue, sample: boolean, stale = false): NumberFormatPreview => ({
    ok: true,
    text: shown.text,
    sample,
    ...(shown.color ? { color: shown.color } : {}),
    ...(stale ? { stale: true } : {}),
  });
  if (!target) {
    const shownCode = id === undefined ? code : (builtinFormatCode(id, locale) ?? code);
    return done(formatValue(NUMBER_FORMAT_SAMPLE, shownCode, { locale, date1904: false }), true);
  }
  const { pkg, index, entry, selection } = target;
  const at = selection.active;
  const part = xmlPartOf(pkg, entry.partName);
  const cell = part ? sheetIndex(part).cells.get(cellKey(at.row, at.col)) : undefined;
  // The stale state of the sheet, as the grid reads it: a stale result is not shown as current.
  const read = cell && part ? cellValue(pkg, index, cell.el, sheetStale(index, part.root)) : null;
  const stored: CellValue = read ? read.value : { kind: 'empty' };
  const stale = read?.stale ?? false;
  // The grid paints rich text runs as stored, with no number format: the preview does too.
  if (stored.kind === 'string' && stored.runs && stored.runs.length > 0)
    return { ok: true, text: stored.text, sample: false };
  const sample = stored.kind === 'empty';
  const value: CellValue = sample ? { kind: 'number', value: NUMBER_FORMAT_SAMPLE } : stored;
  // The cell's own format with the format that the apply writes.
  const base = resolvedFormatAt(pkg, index, entry.partName, at);
  const format =
    id === undefined
      ? { ...base, formatCode: code }
      : { ...base, numFmtId: id, formatCode: defined.get(id) };
  return done(displayValue(value, format, { locale, date1904: index.date1904 }), sample, stale);
}

/** True when the selection is the active cell alone, or the merged range that holds it. */
function selectsOneCell(editor: XlsxEditorInstance): boolean {
  const editing = editingOf(editor);
  const sel = editing?.target?.selection;
  if (!editing || !sel) return false;
  const own = selectCell(sel.sheetIndex, sel.active, editing.merges).range;
  const { start, end } = sel.range;
  return (
    own.start.row === start.row &&
    own.start.col === start.col &&
    own.end.row === end.row &&
    own.end.col === end.col
  );
}

/**
 * What a number format dialog opened with: the code to show first. Only
 * `numberFormatDialogStart` makes one; the engine keeps what it compares (the workbook, the
 * selection, and the revision) out of sight, keyed by the start object. Keep the object that
 * it returns: a copy is not a start, and its Apply closes.
 *
 * @public
 */
export interface NumberFormatDialogStart {
  /** The code that the dialog shows first: the `value` of the `number.custom` slot. */
  readonly initial: string;
}

/** What the engine compares for a start. A start that it did not make has no entry. */
interface StartFacts {
  readonly workbook: object;
  readonly selection: string;
  readonly revision: number;
  /** The number locale: `initial` is a code in it. */
  readonly locale: NumberLocale;
}

const startFacts = new WeakMap<NumberFormatDialogStart, StartFacts>();

/**
 * The start of a number format dialog: call it once, when the dialog opens. Null when the
 * `number.custom` slot is disabled (`toolbarCommandState`, the one source of enabled state)
 * or the commit of an open cell edit is refused: then do not open the dialog. The reason of a
 * disabled slot is its `disabledReason`; the reason of a refused commit is in the snapshot's
 * `refusal`. Otherwise it commits the open edit first, so the preview and the Apply read
 * the same cell, and `initial` is the slot's `value` after the commit. The commit is an edit:
 * call it in the event handler that opens the dialog, not during render.
 *
 * @public
 */
export function numberFormatDialogStart(
  editor: XlsxEditorInstance
): NumberFormatDialogStart | null {
  const editing = editingOf(editor);
  if (!editing?.target || !slotEnabled(editor)) return null;
  if (!editing.commitOpenEdit().ok) return null;
  const workbook = editing.target?.store ?? null;
  if (!workbook) return null;
  const snap = editor.snapshot();
  const initial = toolbarCommandState(snap, 'number.custom').value ?? '';
  const revision = editingOf(editor)?.target?.revision ?? -1;
  const start: NumberFormatDialogStart = Object.freeze({ initial });
  startFacts.set(start, {
    workbook,
    selection: selectionKey(editor),
    revision,
    locale: snap.numberLocale,
  });
  return start;
}

/** The shown sheet index and the selected range, so a start can tell when the selection moved. */
function selectionKey(editor: XlsxEditorInstance): string {
  const snap = editor.snapshot();
  return `${snap.activeSheet}!${snap.selectionRef ?? ''}`;
}

const slotEnabled = (editor: XlsxEditorInstance): boolean =>
  toolbarCommandState(editor.snapshot(), 'number.custom').enabled;

/**
 * True while a dialog's start still holds: the same workbook, selection, and number locale,
 * and the `number.custom` slot enabled. When it turns false, close the dialog.
 *
 * @public
 */
export function numberFormatDialogCurrent(
  editor: XlsxEditorInstance,
  start: NumberFormatDialogStart
): boolean {
  const facts = startFacts.get(start);
  const workbook = editingOf(editor)?.target?.store ?? null;
  return (
    facts !== undefined &&
    workbook !== null &&
    workbook === facts.workbook &&
    selectionKey(editor) === facts.selection &&
    sameNumberLocale(editor.snapshot().numberLocale, facts.locale) &&
    slotEnabled(editor)
  );
}

/**
 * What an Apply of a number format dialog does, from `numberFormatDialogApplyStep`.
 *
 * @public
 */
export type NumberFormatDialogApplyStep =
  | { readonly kind: 'closed' }
  | { readonly kind: 'unchanged' }
  | { readonly kind: 'apply' }
  | { readonly kind: 'problem'; readonly problem: NumberFormatCodeProblem };

/**
 * What an Apply of a number format dialog does with `code`, by `kind`:
 *
 * - `closed`: close the dialog and write nothing: `numberFormatDialogCurrent` is false.
 * - `unchanged`: write nothing and close. The code is the start's `initial`, the workbook
 *   has no edit since the start, and the selection is one cell (a merged cell is one
 *   cell). The cell keeps its format, also a format id that has no code in this editor and
 *   shows as `General`.
 * - `apply`: run the `number.custom` slot with the code. On a range, an unchanged code sets
 *   it on every cell.
 * - `problem`: keep the dialog open and show `problem`, a catalog key and its parameters.
 *
 * @example
 * ```ts
 * const step = numberFormatDialogApplyStep(editor, start, code);
 * if (step.kind === 'problem') showError(t(step.problem.key, step.problem.params));
 * ```
 *
 * @public
 */
export function numberFormatDialogApplyStep(
  editor: XlsxEditorInstance,
  start: NumberFormatDialogStart,
  code: string
): NumberFormatDialogApplyStep {
  if (!numberFormatDialogCurrent(editor, start)) return { kind: 'closed' };
  // An edit after the start (a host call, automation) can change the cell's format, so the
  // shortcut holds only at the start's revision.
  const revision = editingOf(editor)?.target?.revision;
  if (
    code === start.initial &&
    revision === startFacts.get(start)?.revision &&
    selectsOneCell(editor)
  )
    return { kind: 'unchanged' };
  const problem = numberFormatCodeProblem(code);
  return problem ? { kind: 'problem', problem } : { kind: 'apply' };
}

/** The most custom codes that `workbookNumberFormatCodes` lists. @public */
export const MAX_LISTED_NUMBER_FORMATS = 200;

/** The most `x:numFmt` records that `workbookNumberFormatCodes` checks, valid or not. */
const MAX_CHECKED_NUMBER_FORMATS = 2000;

/**
 * The custom number format codes of the open workbook (`x:numFmt` with an id of 164 or more),
 * by id, each once, and only the codes that the store applies. Empty when no workbook is
 * ready. It checks at most the first 2,000 records by id, so a file with many codes that the
 * store refuses can list fewer. The codes are file text: render them as text only.
 *
 * @public
 */
export function workbookNumberFormatCodes(editor: XlsxEditorInstance): readonly string[] {
  const target = editingOf(editor)?.target;
  if (!target) return [];
  const numFmts = styleLists(target.pkg, target.index).numFmts;
  const ids = [...numFmts.keys()]
    .filter((id) => id >= FIRST_CUSTOM_NUMFMT_ID)
    .sort((a, b) => a - b);
  const out: string[] = [];
  const seen = new Set<string>();
  for (const id of ids.slice(0, MAX_CHECKED_NUMBER_FORMATS)) {
    const code = numFmts.get(id)!;
    if (seen.has(code) || numberFormatCodeProblem(code)) continue;
    seen.add(code);
    out.push(code);
    if (out.length >= MAX_LISTED_NUMBER_FORMATS) break;
  }
  return out;
}
