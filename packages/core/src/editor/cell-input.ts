// What the cell editor commits, DOM-free: the text a user typed, read as a store input, and the
// typed refusals that the editor shows. The text reads as a formula when it starts with `=`, as
// a boolean when it is `TRUE` or `FALSE`, as empty when it is empty, as a number when it is a
// typed number form (`readCellEntry`), and as a string otherwise. A leading apostrophe keeps the
// rest as a string.

import { EN_US, parseDouble, type NumberLocale } from '../cascade/index.ts';
import {
  MAX_FORMAT_NEW_CELLS,
  type CellInput,
  type OpRejection,
  type StructureSubject,
} from '../store/index.ts';
import type { ChromeDisabledReason } from './toolbar-commands.ts';
import { DECIMAL_REASON_KEYS, MAX_DECIMAL_STEP_CELLS, groupDigits } from './decimal-reasons.ts';
import { MAX_SELECTION_RANGES } from './selection.ts';
import { readNumberForm, type CellEntryOptions, type NumberFormatRef } from './typed-forms.ts';

/**
 * What a structural refusal names, as a kind of content and a display name: the sheet name,
 * the defined name, or a reference such as `Data!B5`. Chrome shows it through
 * {@link refusalText} and {@link subjectText}, never the part path.
 *
 * @public
 */
export type RefusalSubject = StructureSubject;

export type { CellEntryOptions } from './typed-forms.ts';

/** Typed text read as a store input, and the number format that its form implies. */
export interface TypedEntry {
  /** The value to store. */
  readonly value: CellInput;
  /**
   * The number format of the typed form: a built-in format id of ECMA-376 Part 1, 18.8.30, or
   * a custom code: for a currency amount, an elapsed time with no seconds (`[h]:mm`), or a date
   * and a time that format 22 cannot show in the locale order, with its seconds, or on a 12-hour
   * clock. The editor applies it only to a cell with the General format (`typedInput`).
   */
  readonly numberFormat?: NumberFormatRef;
}

/**
 * Read the text of the cell editor as a store input: the value that a cell with the General
 * format stores. A typed form such as `12%` or `2026-11-12` reads as a number; the number
 * format that the editor also sets is not part of the result. `options` gives the date system
 * of the workbook (1900 when absent) and the year of a date with no year (the year of the host
 * clock when absent). Write cell values through the editor (`enterValue`), which applies the
 * number format rules of the cell.
 *
 * @public
 */
export function parseCellInput(
  text: string,
  locale: NumberLocale,
  options: CellEntryOptions = {}
): CellInput {
  return readCellEntry(text, locale, options).value;
}

/**
 * Read the text of the cell editor as a store input and the number format of its form. A
 * number form is a plain number, a percent (`12%`), a currency amount (`$5.50`, `5,50 €`), a
 * number with group separators (`1,234.5`), a negative in parentheses (`(100)`), a fraction with
 * a whole part (`1 1/2`), a date (`2026-11-12`, or the day and month order of the locale's short
 * date, built-in format 14), a time (`13:45`, `1:45 PM`), an elapsed time (`36:00`), or a date
 * and a time. The locale gives the decimal and group separators. When the locale's short date
 * has no day and month order, only ISO 8601 dates read as dates. Text of more than 64
 * characters, and a form that does not read whole, stays a string. The cell's own format rules
 * (the Text format, a non-General format) apply in `typedInput`, not here.
 */
export function readCellEntry(
  text: string,
  locale: NumberLocale,
  options: CellEntryOptions = {}
): TypedEntry {
  if (text.length === 0) return { value: { type: 'empty' } };
  // A lone `=` holds no formula: it is stored as the text `=`.
  if (text === '=') return { value: { type: 'string', value: text } };
  if (text.charCodeAt(0) === 61 /* = */) return { value: { type: 'formula', text: text.slice(1) } };
  // A leading `'` marks the rest as text. A lone `'` marks no text, so it clears the cell, as
  // an empty input does; it never stores an empty shared string item.
  if (text === "'") return { value: { type: 'empty' } };
  if (text.charCodeAt(0) === 39 /* ' */) return { value: { type: 'string', value: text.slice(1) } };
  const upper = text.toUpperCase();
  if (upper === 'TRUE' || upper === 'FALSE')
    return { value: { type: 'boolean', value: upper === 'TRUE' } };
  const plain = locale.decimal === '.' ? text : swapDecimal(text, locale.decimal);
  const n = plain === undefined ? undefined : parseDouble(plain);
  if (n !== undefined) return { value: { type: 'number', value: n } };
  const typed = readNumberForm(text, locale, options);
  if (!typed) return { value: { type: 'string', value: text } };
  const value: CellInput = { type: 'number', value: typed.value };
  return typed.format === undefined ? { value } : { value, numberFormat: typed.format };
}

/** The text with the locale decimal separator as `.`, or undefined when it holds a `.`. */
function swapDecimal(text: string, decimal: string): string | undefined {
  if (text.includes('.')) return undefined;
  return text.split(decimal).join('.');
}

/**
 * Why an edit was refused. Each code has a whole-sentence message in the shared catalog at
 * `editing.refused.<code>`; a test keeps the two lists equal. `edit-dropped` is the one code that
 * no command returns: the editor keeps it in `snapshot().refusal` when it drops an open edit with
 * changed text (another open, `close`, or a surface that closes while an input method composes).
 * `loading` refuses every command, a save included, while a workbook loads.
 *
 * `shared-formula` refuses an edit (a typed value or formula, or a paste), a clear, or the clear
 * of a merge with `dropValues` that would break a formula range. The editor keeps each formula
 * range whole. The cases:
 *
 * - The cell is the master of, or inside, an array or data table formula range. Its master can
 *   be inside `mc:AlternateContent`.
 * - The cell is a shared formula master whose text cannot be read.
 * - The cell is a shared formula master, and a child has its `x:f` inside `mc:AlternateContent`
 *   in its cell. The editor cannot rewrite that child. A paste over the master is refused, also
 *   when it covers that child.
 * - The edit would move the formula of a shared child out of the grid.
 *
 * An edit to any other shared formula master writes each child's own formula first. An edit or
 * a clear of a cell with a formula inside `mc:AlternateContent`, a formula range master or a
 * wrapped shared child too, is refused with `cell-alternate-content` instead.
 *
 * @public
 */
export const EDIT_REFUSAL_CODES = [
  'unknown-sheet',
  'bad-reference',
  'bad-value',
  'text-too-long',
  'bad-size',
  'merge-one-cell',
  'merge-overlap',
  'merge-drops-values',
  'bad-position',
  'invalid-sheet',
  'shared-formula',
  'unreadable-formula',
  'formula-syntax',
  'r1c1-formula',
  'unsupported-function',
  'unknown-formula-reference',
  'invalid-character',
  'unsupported-package',
  'cell-alternate-content',
  'bad-format',
  'unreadable-number-format',
  'range-too-large',
  'clear-too-large',
  'decimals-too-large',
  'decimals-unchanged',
  'style-limit',
  'invalid-part',
  'locked',
  'structure-blocked',
  'sheet-range-blocked',
  'grid-full',
  'stale-revision',
  'save-failed',
  'open-failed',
  'sheet-name-invalid',
  'sheet-name-too-long',
  'sheet-name-taken',
  'sheet-name-ambiguous',
  'last-sheet',
  'nothing-to-undo',
  'nothing-to-redo',
  'not-editable',
  'whole-sheet',
  'multiple-ranges',
  'too-many-ranges',
  'merged-range',
  'paste-html',
  'clipboard-read',
  'copy-too-large',
  'copy-conditional-limit',
  'copy-array-formula',
  'copy-unreadable-formula',
  'copy-wrapped-formula',
  'paste-unreadable-formula',
  'paste-name-scope',
  'cut-wrapped-value',
  'cut-expired',
  'copy-expired',
  'paste-reference-outside',
  'too-many-rows',
  'resize-anchor',
  'view-mode',
  'composing',
  'edit-pending',
  'edit-dropped',
  'loading',
  'go-to-reference',
  'go-to-name',
  'go-to-sheet',
] as const;

/**
 * Why an edit was refused: one of `EDIT_REFUSAL_CODES`.
 *
 * @public
 */
export type EditRefusalCode = (typeof EDIT_REFUSAL_CODES)[number];

/**
 * A refused edit: the code, and the fields that the code fills.
 *
 * @public
 */
export interface EditRefusal {
  readonly code: EditRefusalCode;
  /** The engine's detail for logs. It is never shown as user text, and its words can change. */
  readonly detail?: string;
  /**
   * What the refusal names, as the file or the typed text names it: an element and its part, a
   * part, a cell (`Data!B5`), a range (`B2:C3`), a sheet name, or the function that
   * `unsupported-function` refuses. A merge names a cell or a range of the shown sheet, with no
   * sheet name.
   * When the refusal has no `about`, `refusalText` shows it where the message has `{subject}`,
   * with bidirectional controls removed and cut to 64 characters; else `about` stands in its
   * place and this is for logs. A message whose refusal names nothing has a subject-free text. A
   * `resize-anchor` refusal of an object with no name shows a generic label.
   */
  readonly subject?: string;
  /**
   * The cell, with its sheet (`Data!A2`), whose formula the editor cannot rewrite, on a
   * `shared-formula` refusal of a shared formula master: a child inside `mc:AlternateContent`, a
   * child whose formula would move out of the grid, or a master whose text cannot be read. The
   * message names it beside `masterCell` (`editing.formulaCellRefused.<command>`, with `edit`
   * for a cell edit).
   */
  readonly causeCell?: string;
  /**
   * True when the formula of `causeCell` cannot be read: the message then says so
   * (`editing.formulaUnreadableRefused.<command>`). On a `resize-anchor` refusal, true when the
   * drawing part of the sheet cannot be read: the message then says so
   * (`editing.resizeUnreadDrawing`).
   */
  readonly causeUnreadable?: true;
  /**
   * The shared formula master, with its sheet (`Data!A1`), that the refused command would
   * change, beside `causeCell`. It is inside the edited, cleared, cut, or pasted cells, so the
   * message names it as the cell to leave out.
   */
  readonly masterCell?: string;
  /**
   * What a `structure-blocked` or `grid-full` refusal names, for user text. A
   * `sheet-range-blocked` refusal has it as `OpRejection.about` states.
   */
  readonly about?: RefusalSubject;
  /**
   * The command that the refusal stopped, when it is not a cell edit. `save` marks a save
   * (`save()`, the save control, or Control or Command and S) that wrote nothing, with the reason
   * as the code: `loading`, `not-editable`, `composing`, or the code that refused the commit of the
   * open cell edit (`SAVE_REFUSAL_CODES`). `format`, `clear` (of values or of formats), `clearAll`,
   * `structure`, `cut`, and `paste` mark a `locked` refusal of that kind of command, so its message
   * names the command. `clear`, `clearAll`, `paste`, and `merge` also mark a
   * `cell-alternate-content` or `shared-formula` refusal of that command, so its message names the
   * command. `merge` marks such a refusal of the clear of `dropValues`. `commit` marks such a
   * refusal of the commit of an open cell edit: the cell editor stays open, so its message tells
   * the user to cancel the cell edit. A save that commits first has `save` instead. A `locked`
   * refusal of a value edit, a column width or a row height, a merge or an unmerge, or a sheet op has no
   * command: its rule (`lock`: `cell`, `size`, `merge`, or `workbook`) names the message. `cut`
   * marks every refusal of a cut, whatever its code (`CUT_REFUSAL_CODES` names the codes with
   * their own cut message), but not a refusal of the open cell edit that the cut commits first:
   * that refusal has `commit` or no command, as any commit has.
   * When a load changes a refusal to `loading`, the refusal keeps its command. It is set only on
   * a refused save or cut, on a `locked` refusal of `format`, `clear`, `clearAll`, `structure`, or
   * `paste`, on a `cell-alternate-content` or `shared-formula` refusal of `clear`, `clearAll`,
   * `paste`, `merge`, or `commit`, and on a `loading` refusal that replaced one of those during a load; every
   * other refusal has no command.
   */
  readonly command?:
    'save' | 'format' | 'clear' | 'clearAll' | 'structure' | 'cut' | 'paste' | 'merge' | 'commit';
  /**
   * The protection rule of a `locked` refusal: a locked `cell` (or a new cell format that would
   * change the protection of a cell), `format` when the protected sheet does not allow
   * formatting cells, `structure` when it does not allow the insert or
   * the delete, `size` when it does not allow a column width or a row height, or `workbook`
   * when the workbook protection locks the sheet list (adding, renaming, deleting, and moving
   * sheets), or `merge` when the sheet is protected, because a protected sheet allows no merge
   * or unmerge. The message names the rule.
   */
  readonly lock?: 'cell' | 'format' | 'structure' | 'size' | 'workbook' | 'merge';
  /**
   * Why the command changed nothing, in the words of its disabled control: a catalog key and its
   * parameters. Set on `decimals-unchanged` and `decimals-too-large`, which name the cell or the
   * selection that stops the step. `refusalText` shows it, and `toolbarCommandState` gives it as
   * the disabled reason.
   */
  readonly reason?: ChromeDisabledReason;
}

/**
 * The outcome of an edit command. A refused write never reports `ok`.
 *
 * @public
 */
export type EditResult =
  { readonly ok: true } | { readonly ok: false; readonly refusal: EditRefusal };

/** The editor refusal of a store rejection. */
export function refusalOf(rejection: OpRejection): EditRefusal {
  return {
    code: rejection.code,
    detail: rejection.detail,
    ...(rejection.lock !== undefined ? { lock: rejection.lock } : {}),
    ...(rejection.subject !== undefined ? { subject: rejection.subject } : {}),
    ...(rejection.causeCell !== undefined ? { causeCell: rejection.causeCell } : {}),
    ...(rejection.causeUnreadable ? { causeUnreadable: true as const } : {}),
    ...(rejection.masterCell !== undefined ? { masterCell: rejection.masterCell } : {}),
    ...(rejection.about !== undefined ? { about: rejection.about } : {}),
  };
}

/**
 * A translation function: a key of the shared catalog and its parameters, to display text. The
 * text helpers (`refusalText`, `subjectText`, `disabledReasonText`) and the editor take it.
 *
 * @public
 */
export type EditorTranslate = {
  // Method syntax makes the parameters bivariant, so a catalog's typed translator, whose key
  // is a union of catalog keys (the adapters' `TFunction`), passes where a string key is
  // read. The engine passes only catalog keys.
  translate(key: string, params?: Readonly<Record<string, string | number>>): string;
}['translate'];

/** The longest display name that a refusal shows, in characters. */
const MAX_SUBJECT_NAME = 64;

/**
 * A name from the file as a message shows it: without control characters (C0, U+007F, and C1)
 * and bidirectional format controls, which could reorder the message around it, and cut to 64
 * characters, the last one an ellipsis. Use it for each sheet, cell, range, or part name in a
 * message.
 *
 * @public
 */
export function displayName(name: string): string {
  const plain = name.replace(
    /[\u0000-\u001f\u007f-\u009f\u061c\u200e\u200f\u202a-\u202e\u2066-\u2069]/g,
    ''
  );
  const chars = [...plain];
  return chars.length > MAX_SUBJECT_NAME
    ? `${chars.slice(0, MAX_SUBJECT_NAME - 1).join('')}…`
    : plain;
}

/**
 * The user words of a refusal subject: a noun of the catalog's `structureSubject` keys, such as
 * "a drawing" or "the defined name Rates". The name loses bidirectional controls and is cut to
 * 64 characters.
 *
 * @public
 */
export function subjectText(about: RefusalSubject, t: EditorTranslate): string {
  return t(`structureSubject.${about.kind}`, { name: displayName(about.name ?? '') });
}

/**
 * The codes that a save can return: a load, no open workbook, a composition, and the codes
 * that refuse the commit of the open cell edit. Each has a whole-sentence message at
 * `editing.saveRefused.<code>`; a save with another code shows `editing.saveRefused.other`.
 *
 * @public
 */
export const SAVE_REFUSAL_CODES = [
  'loading',
  'not-editable',
  'composing',
  'view-mode',
  'unknown-sheet',
  'bad-reference',
  'bad-value',
  'text-too-long',
  'invalid-sheet',
  'shared-formula',
  'unreadable-formula',
  'formula-syntax',
  'r1c1-formula',
  'unsupported-function',
  'unknown-formula-reference',
  'invalid-character',
  'unsupported-package',
  'cell-alternate-content',
  'style-limit',
  'invalid-part',
  'locked',
  'stale-revision',
  'save-failed',
] as const satisfies readonly EditRefusalCode[];

const SAVE_CODES: ReadonlySet<EditRefusalCode> = new Set(SAVE_REFUSAL_CODES);

/**
 * The codes of a cut refusal that have their own message at `editing.cutRefused.<code>`: a load,
 * view mode, no open workbook, the copy codes (`copy-wrapped-formula` for a formula that only
 * `mc:AlternateContent` holds, as the copy gives it), `cut-wrapped-value` for a value that the
 * copy cannot carry, and `cell-alternate-content`, `shared-formula`, or `invalid-sheet` when the
 * store cannot clear the cut cells. A `locked` cut has an `editing.lockedRefused` message. The
 * dry run of the clear can also give another store code, such as `style-limit` or
 * `invalid-part`: that cut shows `editing.cutRefused.other`. A refusal of the open cell edit
 * that the cut commits first is not a cut refusal.
 *
 * @public
 */
export const CUT_REFUSAL_CODES = [
  'loading',
  'view-mode',
  'not-editable',
  'merged-range',
  'copy-too-large',
  'copy-conditional-limit',
  'copy-array-formula',
  'copy-unreadable-formula',
  'copy-wrapped-formula',
  'cut-wrapped-value',
  'cell-alternate-content',
  'shared-formula',
  'invalid-sheet',
  'multiple-ranges',
] as const satisfies readonly EditRefusalCode[];

const CUT_CODES: ReadonlySet<EditRefusalCode> = new Set(CUT_REFUSAL_CODES);

/** The codes whose message names a `{subject}`: each has a subject-free message too. */
const NAMES_SUBJECT: ReadonlySet<EditRefusalCode> = new Set([
  'unsupported-function',
  'cell-alternate-content',
  'merge-drops-values',
  'merge-overlap',
  'structure-blocked',
  'grid-full',
  'sheet-range-blocked',
  'sheet-name-too-long',
  'sheet-name-taken',
  'sheet-name-ambiguous',
  'shared-formula',
  'copy-array-formula',
  'copy-unreadable-formula',
  'copy-wrapped-formula',
  'cut-wrapped-value',
  'paste-unreadable-formula',
  'paste-name-scope',
  'paste-reference-outside',
  'resize-anchor',
  'go-to-name',
  'go-to-sheet',
]);

export { groupDigits } from './decimal-reasons.ts';

/**
 * The catalog key of a `cell-alternate-content` refusal of a clear, a paste, a merge, or a
 * commit, which names the command, or undefined for a value edit.
 */
function alternateKey(command: EditRefusal['command']): string | undefined {
  switch (command) {
    case 'clear':
      return 'editing.alternateRefused.clear';
    case 'clearAll':
      return 'editing.alternateRefused.clearAll';
    case 'paste':
      return 'editing.alternateRefused.paste';
    case 'merge':
      return 'editing.alternateRefused.merge';
    case 'commit':
      return 'editing.alternateRefused.commit';
    default:
      return undefined;
  }
}

/**
 * The catalog key of a `shared-formula` refusal of a clear, a paste, a merge, or a commit, which
 * names that command:
 * `editing.formulaRangeRefused.<command>`, or `editing.formulaRangeRefusedUnnamed.<command>`
 * when the refusal names no range or cell.
 */
function formulaRangeKey(command: EditRefusal['command'], unnamed: boolean): string | undefined {
  switch (command) {
    case 'clear':
      return unnamed
        ? 'editing.formulaRangeRefusedUnnamed.clear'
        : 'editing.formulaRangeRefused.clear';
    case 'clearAll':
      return unnamed
        ? 'editing.formulaRangeRefusedUnnamed.clearAll'
        : 'editing.formulaRangeRefused.clearAll';
    case 'paste':
      return unnamed
        ? 'editing.formulaRangeRefusedUnnamed.paste'
        : 'editing.formulaRangeRefused.paste';
    case 'merge':
      return unnamed
        ? 'editing.formulaRangeRefusedUnnamed.merge'
        : 'editing.formulaRangeRefused.merge';
    case 'commit':
      return unnamed
        ? 'editing.formulaRangeRefusedUnnamed.commit'
        : 'editing.formulaRangeRefused.commit';
    default:
      return undefined;
  }
}

/**
 * The key in `editing.formulaCellRefused` and `editing.formulaUnreadableRefused` of a command: a
 * cell edit has `edit`. No format or structure command gives a `causeCell`; a refusal of one that
 * did would show the cell edit text.
 */
function formulaCellKey(command: EditRefusal['command']): string {
  switch (command) {
    case 'save':
    case 'cut':
    case 'clear':
    case 'clearAll':
    case 'paste':
    case 'merge':
    case 'commit':
      return command;
    default:
      return 'edit';
  }
}

/**
 * The catalog key of a `locked` refusal, by its protection rule, then by its command: a
 * format lock never shows the message of a locked cell.
 */
function lockedKey(refusal: EditRefusal): string {
  const { command } = refusal;
  switch (refusal.lock) {
    case 'structure':
      return 'editing.lockedRefused.structure';
    case 'size':
      return 'editing.lockedRefused.size';
    case 'workbook':
      return 'editing.lockedRefused.workbook';
    case 'merge':
      return 'editing.lockedRefused.merge';
    case 'format':
      return command === 'cut'
        ? 'editing.lockedRefused.cutFormats'
        : command === 'paste'
          ? 'editing.lockedRefused.pasteFormats'
          : command === 'clear'
            ? 'editing.lockedRefused.clearFormats'
            : command === 'clearAll'
              ? 'editing.lockedRefused.clearAll'
              : 'editing.lockedRefused.format';
    default:
      return command === 'cut'
        ? 'editing.lockedRefused.cut'
        : command === 'paste'
          ? 'editing.lockedRefused.paste'
          : command === 'clear' || command === 'clearAll'
            ? 'editing.lockedRefused.clear'
            : command === 'structure'
              ? 'editing.lockedRefused.structure'
              : command === 'format'
                ? 'editing.lockedRefused.format'
                : 'editing.refused.locked';
  }
}

/**
 * The user text of a refusal: one whole-sentence message of the catalog, never composed from parts
 * and never cut. It is `editing.refused.<code>` for a command, `editing.saveRefused.<code>` for a
 * save, `editing.cutRefused.<code>` for a cut, and an `editing.lockedRefused` key for a `locked`
 * refusal with a protection rule (`lock`: `cell`, `format`, `structure`, `size`, or `workbook`)
 * or a command, by the rule that refused it. A `cell-alternate-content` or `shared-formula`
 * refusal of a clear, a paste, or a merge shows an `editing.alternateRefused` or
 * `editing.formulaRangeRefused` key of that command. With no subject, a `shared-formula` refusal
 * of one of them shows `editing.formulaRangeRefusedUnnamed`, and a `cell-alternate-content`
 * refusal of a merge or a commit shows `editing.alternateRefusedUnnamed`. A `shared-formula` refusal with a
 * `causeCell` and a `masterCell` shows `editing.formulaCellRefused.<command>` (`edit` for a
 * cell edit), or `editing.formulaUnreadableRefused.<command>` with `causeUnreadable`. It names
 * the master to leave out, and the cell whose formula blocks it. A save or a cut with a code
 * outside its list (`SAVE_REFUSAL_CODES`, `CUT_REFUSAL_CODES`) shows `editing.saveRefused.other`
 * or `editing.cutRefused.other`. `{subject}` is the words of `about` when the refusal has one, else
 * the name as the file names it (a cell such as `Data!B5`); a refusal that names nothing shows
 * the subject-free message (`editing.refusedUnnamed.<code>`, `editing.saveRefusedUnnamed.<code>`,
 * `editing.cutRefusedUnnamed.<code>`), so no placeholder shows. A `structure-blocked` refusal
 * whose `about` kind is `not-worksheet` shows `editing.notWorksheetRefused`. A
 * `structure-blocked` or `sheet-range-blocked` refusal whose `about` kind is `wrapped-formula`
 * shows `editing.wrappedFormulaRefused`, and a `resize-anchor`
 * refusal with `causeUnreadable` shows `editing.resizeUnreadDrawing`. `{limit}` is grouped
 * with the separator of `locale`, US English by default. Chrome and the live region show refusals
 * through this one function.
 *
 * @public
 */
export function refusalText(
  refusal: EditRefusal,
  t: EditorTranslate,
  locale: NumberLocale = EN_US
): string {
  // A raw subject is file or user text: it gets the same cleanup as a name in `about`.
  const subject = refusal.about
    ? subjectText(refusal.about, t)
    : refusal.subject === undefined
      ? undefined
      : displayName(refusal.subject);
  const { code, command } = refusal;
  // `range-too-large`, `clear-too-large`, `decimals-too-large`, and `copy-too-large` name the
  // limit of one command, in cells.
  const params: Record<string, string | number> = {};
  if (code === 'decimals-too-large') params.limit = groupDigits(MAX_DECIMAL_STEP_CELLS, locale);
  else if (code === 'range-too-large' || code === 'clear-too-large' || code === 'copy-too-large')
    params.limit = groupDigits(MAX_FORMAT_NEW_CELLS, locale);
  else if (code === 'too-many-ranges') params.limit = groupDigits(MAX_SELECTION_RANGES, locale);
  if (subject !== undefined) params.subject = subject;
  const p = Object.keys(params).length ? params : undefined;
  const unnamed = subject === undefined && NAMES_SUBJECT.has(code);
  // A shared formula master refused for the formula of one cell names that cell too.
  const { causeCell, masterCell } = refusal;
  if (code === 'shared-formula' && causeCell !== undefined && masterCell !== undefined) {
    const family = refusal.causeUnreadable ? 'formulaUnreadableRefused' : 'formulaCellRefused';
    return t(`editing.${family}.${formulaCellKey(command)}`, {
      causeCell: displayName(causeCell),
      masterCell: displayName(masterCell),
    });
  }
  if (command === 'save') {
    if (!SAVE_CODES.has(code)) return t('editing.saveRefused.other');
    return unnamed ? t(`editing.saveRefusedUnnamed.${code}`) : t(`editing.saveRefused.${code}`, p);
  }
  if (code === 'locked' && (command !== undefined || refusal.lock !== undefined))
    return t(lockedKey(refusal));
  // A refused cut names the cut, not a copy or an edit.
  if (command === 'cut') {
    if (!CUT_CODES.has(code)) return t('editing.cutRefused.other');
    return unnamed ? t(`editing.cutRefusedUnnamed.${code}`) : t(`editing.cutRefused.${code}`, p);
  }
  // A decimal step that no cell can take says why, in the words of the disabled control.
  const reason = refusal.reason;
  if (code === 'decimals-unchanged' && reason && DECIMAL_REASON_KEYS.has(reason.key))
    return t('editing.decimalsUnchanged', { reason: t(reason.key, reason.params) });
  const alternate = code === 'cell-alternate-content' ? alternateKey(command) : undefined;
  if (alternate !== undefined && subject !== undefined) return t(alternate, p);
  if (code === 'cell-alternate-content' && command === 'merge')
    return t('editing.alternateRefusedUnnamed.merge');
  if (code === 'cell-alternate-content' && command === 'commit')
    return t('editing.alternateRefusedUnnamed.commit');
  const range = code === 'shared-formula' ? formulaRangeKey(command, unnamed) : undefined;
  if (range !== undefined) return t(range, p);
  // A resize of a sheet whose drawing cannot be read checks no anchor: its message says so,
  // not that an object keeps its size.
  if (code === 'resize-anchor' && refusal.causeUnreadable) return t('editing.resizeUnreadDrawing');
  // A sheet that is not a worksheet refers to nothing that the edit changes: its message says
  // that the editor moves and deletes only worksheets.
  if (code === 'structure-blocked' && refusal.about?.kind === 'not-worksheet') {
    const name = refusal.about.name ?? refusal.subject;
    return name === undefined
      ? t('editing.notWorksheetRefused.unnamed')
      : t('editing.notWorksheetRefused.named', { subject: displayName(name) });
  }
  // A formula inside `mc:AlternateContent`: its message says why the editor cannot update it.
  if (
    (code === 'structure-blocked' || code === 'sheet-range-blocked') &&
    refusal.about?.kind === 'wrapped-formula'
  ) {
    const name = refusal.about.name ?? refusal.subject;
    return name === undefined
      ? t('editing.wrappedFormulaRefused.unnamed')
      : t('editing.wrappedFormulaRefused.named', { subject: displayName(name) });
  }
  return unnamed ? t(`editing.refusedUnnamed.${code}`) : t(`editing.refused.${code}`, p);
}
