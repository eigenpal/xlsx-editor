// The toolbar's one command table, DOM-free: the state of each chrome slot and the command it
// runs. After docx-editor/packages/core/src/editor/toolbar-commands.ts at a0825f1f (EigenPal):
// `toolbarCommandState` is the only source of a control's enabled state, its pressed state,
// its value, and the engine's reason when it is disabled, and both adapters call it.
// Differences: the state is a pure function of the editor snapshot (the active cell's format
// and the refusals that the snapshot carries), so it reads the selection only and never scans
// a sheet; the disabled reason is a key of the shared catalog with its parameters, never
// English text; and `runToolbarCommand` runs the slot on the workbook editor.
//
// A command and its state read the same source. A toggle (bold, an alignment, wrap) reads the
// active cell's format from the snapshot (`activeFormat`), and the editor's `toggleFormat`
// reads the same resolved format, so the pressed state is always what the press undoes.

import type { ChromeSlotId } from './chrome-controls.ts';
import { chromeControlForSlot } from './chrome-controls.ts';
import {
  displayName,
  subjectText,
  type EditorTranslate,
  type EditRefusal,
  type EditResult,
  type RefusalSubject,
} from './cell-input.ts';
import type { ActiveCellFormat, EditorSnapshot } from './editor-model.ts';
import type { SavedWorkbook } from './editor-instance.ts';
import type { XlsxEditorInstance } from './xlsx-editor.ts';
import { editingOf } from './cell-editing.ts';
import { sharedSaveOf } from './editor-save.ts';
import { numberFormatCodeProblem } from './number-format-code.ts';
import { MENU_ONLY_SLOTS, menuSlotState, runMenuSlot } from './menu-commands.ts';
import type { PasteData } from './clipboard.ts';
import {
  FIRST_CUSTOM_NUMFMT_ID,
  styleLists,
  STRUCTURE_PERMIT,
  type FormatPatch,
  type StructuralOpName,
} from '../store/index.ts';
import { builtinFormatCode, EN_US, type NumberLocale } from '../cascade/index.ts';
import {
  builtinOrCode,
  isNumberFormatChoice,
  NUMBER_FORMAT_CHOICES,
  numberFormatChoiceCode,
  numberFormatValue,
} from './number-format-commands.ts';

export { CUSTOM_NUMBER_FORMAT_PREFIX, numberFormatValue } from './number-format-commands.ts';

/**
 * Why a control is disabled: a key of the shared catalog (the English catalog of
 * `@xlsx-editor.dev/i18n`), its parameters, and the subject of a structure refusal. The adapters
 * translate it with `disabledReasonText`; the engine never returns English text.
 *
 * @public
 */
export interface ChromeDisabledReason {
  readonly key: DisabledReasonKey;
  readonly params?: Readonly<Record<string, string>>;
  /** What refuses a structural command; its words are the `{subject}` of the key. */
  readonly about?: RefusalSubject;
}

/**
 * The user text of a disabled reason: its key with its parameters, and `{subject}` as the words
 * of `about` (`subjectText`).
 *
 * @public
 */
export function disabledReasonText(reason: ChromeDisabledReason, t: EditorTranslate): string {
  const params = reason.about
    ? { ...reason.params, subject: subjectText(reason.about, t) }
    : reason.params;
  return t(reason.key, params);
}

/**
 * Every catalog key that `toolbarCommandState`, `sheetCommandState`, and `mergeCommandState` can
 * return as a disabled reason. A test keeps each one in the catalog.
 *
 * @public
 */
export const DISABLED_REASON_KEYS = [
  'disabledReason.editorNotReady',
  'disabledReason.loading',
  'disabledReason.notWired',
  'disabledReason.nothingToUndo',
  'disabledReason.nothingToRedo',
  'disabledReason.editPending',
  'disabledReason.rangeTooLarge',
  'disabledReason.clearTooLarge',
  'disabledReason.structureBlocked',
  'disabledReason.structureBlockedUnnamed',
  'disabledReason.structureUnavailable',
  'disabledReason.wholeSheet',
  'disabledReason.noSave',
  'disabledReason.viewMode',
  'disabledReason.locked',
  'disabledReason.cellAlternateContent',
  'disabledReason.cellAlternateContentClearAll',
  'disabledReason.formulaRange',
  'disabledReason.formulaRangeClearAll',
  'disabledReason.formulaCell',
  'disabledReason.formulaCellClearAll',
  'disabledReason.formulaUnreadable',
  'disabledReason.formulaUnreadableClearAll',
  'disabledReason.formatLocked',
  'disabledReason.stylesUnread',
  'disabledReason.stylesUnreadProtected',
  'disabledReason.structureLocked',
  'disabledReason.workbookLocked',
  'disabledReason.lastSheet',
  'disabledReason.notWorksheet',
  'disabledReason.sheetEditBlocked',
  'disabledReason.composing',
  'disabledReason.decimalsNoNumberSelection',
  'disabledReason.decimalsNoNumber',
  'disabledReason.decimalsNoShownPlaces',
  'disabledReason.decimalsGeneralSection',
  'disabledReason.decimalsGeneralSectionSelection',
  'disabledReason.decimalsDate',
  'disabledReason.decimalsDateSelection',
  'disabledReason.decimalsText',
  'disabledReason.decimalsTextSelection',
  'disabledReason.decimalsFraction',
  'disabledReason.decimalsFractionSelection',
  'disabledReason.decimalsNoDecimals',
  'disabledReason.decimalsNoDecimalsSelection',
  'disabledReason.decimalsNoDecimalsCellSelection',
  'disabledReason.decimalsMaxDecimals',
  'disabledReason.decimalsMaxDecimalsSelection',
  'disabledReason.decimalsMaxDecimalsCellSelection',
  'disabledReason.decimalsTooLong',
  'disabledReason.decimalsTooLongSelection',
  'disabledReason.decimalsUnreadable',
  'disabledReason.decimalsUnreadableSelection',
  'disabledReason.decimalsNoCode',
  'disabledReason.decimalsNoCodeSelection',
  'disabledReason.decimalsMissingCode',
  'disabledReason.decimalsMissingCodeSelection',
  'disabledReason.decimalsTooLarge',
  'disabledReason.mergeLocked',
  'disabledReason.mergeOneCell',
  'disabledReason.mergeOverlap',
  'disabledReason.mergeDropsValues',
  'disabledReason.noMergedRange',
  'disabledReason.noOpen',
  'disabledReason.noClipboard',
  'disabledReason.zoomMax',
  'disabledReason.zoomMin',
  'disabledReason.zoomActual',
] as const;

/**
 * Why an enabled control can seem to do nothing: a key of the shared catalog. The adapters
 * translate it, as they translate a {@link ChromeDisabledReason}.
 *
 * @public
 */
export interface ChromeHint {
  readonly key: ToolbarHintKey;
}

/**
 * Every catalog key that `toolbarCommandState` can return as a hint of an enabled control. A
 * test keeps each one in the catalog.
 *
 * @public
 */
export const TOOLBAR_HINT_KEYS = ['hint.conditionalFormat', 'hint.decimalsLargeSelection'] as const;

/** A hint key. @public */
export type ToolbarHintKey = (typeof TOOLBAR_HINT_KEYS)[number];

/** A disabled-reason key. @public */
export type DisabledReasonKey = (typeof DISABLED_REASON_KEYS)[number];

/**
 * The state of one control.
 *
 * @public
 */
export interface ToolbarCommandState {
  readonly id: ChromeSlotId;
  readonly enabled: boolean;
  /** The engine's reason when disabled, as a catalog key. Shown as the tooltip. */
  readonly disabledReason: ChromeDisabledReason | null;
  /** The command is applied at the active cell: bold is on, the cell is left aligned. */
  readonly active: boolean;
  /**
   * What a value control shows: the font name, the size in points, the color as `RRGGBB` (or
   * `none` for no fill), the number format choice (a `value` of `numberFormatChoices`, or
   * `CUSTOM_NUMBER_FORMAT_PREFIX` and the format code of any other format), or for
   * `number.custom` the format code of the active cell, which its dialog starts from
   * (`General` for a format id with no code here, such as a locale-dependent built-in, so a
   * run of it sets id 0). Absent for a control that has no value.
   */
  readonly value?: string;
  // DOCX Editor differs: its engine states a reason in English, and its catalog maps the sentence
  // to a key (`i18n/src/disabled-reasons.ts` at `f3116761`). DOCX Editor has no hint field.
  /**
   * Why an enabled control can seem to do nothing: `hint.conditionalFormat` when a true
   * conditional formatting rule shows a state of this control at the active cell that differs
   * from the cell format, which the control reads and sets (`EditorSnapshot.conditionalFormat`).
   * Or why it is enabled while the other format controls are not:
   * `hint.decimalsLargeSelection` on a decimal step when the selection is too large to format
   * (`formatRefusal`), because the step changes only the selected cells that hold numbers.
   * Absent otherwise, and absent on a disabled control. The toolbar shows it after the label in
   * the tooltip, and as the accessible description. The engine names a catalog key and the
   * adapter translates it, as this engine does for `disabledReason`.
   */
  readonly hint?: ChromeHint;
}

/**
 * What the host gives the toolbar beyond the snapshot.
 *
 * @public
 */
export interface ToolbarHost {
  /**
   * Receives the workbook bytes of `file.save` when they are ready, with what they hold. Without
   * it, `file.save` is disabled.
   */
  readonly save?: (bytes: Uint8Array<ArrayBuffer>, saved: SavedWorkbook) => void;
  /**
   * Called once when a press of `file.save` starts a delivery, with the promise of its result,
   * so a host can show that the save runs. A press that joins the running save of an earlier
   * press and a refused save do not call it. A press that joins a save from `saveInTasks` calls
   * it. The promise resolves to `null` when the save is refused later.
   */
  readonly saveStart?: (saving: Promise<SavedWorkbook | null>) => void;
  /**
   * Whether the host opens another workbook now, while the editor still shows the old one: for
   * example during a download or a read before `load`. While it returns true, `file.save` is
   * disabled with `disabledReason.loading`, and `runToolbarCommand` refuses the save with
   * `loading` (`command: 'save'`) before any save work starts, so a key press does not delay the
   * open. The command state and each run read it at their call.
   */
  readonly opening?: () => boolean;
  /**
   * Opens another workbook for `file.open`: for example a file picker that passes the bytes to
   * `load`. Without it, `file.open` is disabled. The engine never fetches a workbook.
   */
  readonly open?: () => void;
  /**
   * Reads the text of the system clipboard for `edit.paste`. Without it, `edit.paste` is
   * disabled. The menu bar gives one when the host window has a clipboard that a page can read.
   */
  readonly readClipboard?: () => Promise<PasteData>;
}

/**
 * The font families of the font list. A family name is what the workbook stores; the list
 * also shows the active cell's own family when it is not here. Each has a metric-compatible
 * substitute face in `loadSubstituteFonts`.
 *
 * @public
 */
export const FONT_FAMILY_CHOICES: readonly string[] = [
  'Arial',
  'Calibri',
  'Cambria',
  'Century Gothic',
  'Courier New',
  'Times New Roman',
];

/** The font sizes of the size list, in points. @public */
export const FONT_SIZE_CHOICES: readonly number[] = [
  8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72,
];

/**
 * The colors of the font color and fill color palettes, as `RRGGBB`.
 *
 * @public
 */
export const COLOR_CHOICES = [
  { value: '000000', labelKey: 'colorPicker.colors.black' },
  { value: 'FFFFFF', labelKey: 'colorPicker.colors.white' },
  { value: '808080', labelKey: 'colorPicker.colors.darkGray' },
  { value: 'C0C0C0', labelKey: 'colorPicker.colors.lightGray' },
  { value: 'C00000', labelKey: 'colorPicker.colors.darkRed' },
  { value: 'FF0000', labelKey: 'colorPicker.colors.red' },
  { value: 'FFC000', labelKey: 'colorPicker.colors.orange' },
  { value: 'FFFF00', labelKey: 'colorPicker.colors.yellow' },
  { value: '92D050', labelKey: 'colorPicker.colors.lightGreen' },
  { value: '00B050', labelKey: 'colorPicker.colors.green' },
  { value: '00B0F0', labelKey: 'colorPicker.colors.lightBlue' },
  { value: '0070C0', labelKey: 'colorPicker.colors.blue' },
  { value: '002060', labelKey: 'colorPicker.colors.darkBlue' },
  { value: '7030A0', labelKey: 'colorPicker.colors.purple' },
] as const;

/** The value of the fill color palette that removes the fill. @public */
export const NO_FILL = 'none';

const HEX6 = /^[0-9A-Fa-f]{6}$/;

/** The slots that a format patch drives, and so the format range limit refuses. */
const FORMAT_SLOTS = new Set<ChromeSlotId>([
  'font.family',
  'font.size',
  'font.bold',
  'font.italic',
  'font.underline',
  'font.color',
  'fill.color',
  'borders.outline',
  'borders.all',
  'borders.none',
  'alignment.left',
  'alignment.center',
  'alignment.right',
  'alignment.top',
  'alignment.middle',
  'alignment.bottom',
  'alignment.wrap',
  'number.format',
  'number.custom',
  'number.increaseDecimals',
  'number.decreaseDecimals',
]);

/** The decimal step slots, and the step of each. */
const DECIMAL_STEPS: ReadonlyMap<ChromeSlotId, 1 | -1> = new Map<ChromeSlotId, 1 | -1>([
  ['number.increaseDecimals', 1],
  ['number.decreaseDecimals', -1],
]);

const STRUCTURE_SLOTS = new Set<ChromeSlotId>([
  'cells.insertRows',
  'cells.deleteRows',
  'cells.insertColumns',
  'cells.deleteColumns',
]);

/**
 * Whether the protection of the shown sheet locks a slot: the format commands and clearing
 * formats unless it allows formatting, each structural command unless it allows that one
 * (`STRUCTURE_PERMIT`, the store's map), and clearing contents when the selection holds a
 * locked cell with content (`selectionContentLocked`, the store's rule).
 */
function lockedBy(snapshot: EditorSnapshot, id: ChromeSlotId): DisabledReasonKey | null {
  const p = snapshot.protection;
  if (!p) return null;
  // The same protection rule as the store's refusal (`protectionRefusal`), so the tooltip
  // names what the command would be refused for: formatting, a locked cell, or the structure.
  const format = !p.allowFormatCells ? 'disabledReason.formatLocked' : null;
  const cell = snapshot.selectionContentLocked ? 'disabledReason.locked' : null;
  if (FORMAT_SLOTS.has(id) || id === 'clear.formats') return format;
  if (id === 'clear.all') return cell ?? format;
  if (id === 'clear.contents') return cell;
  if (!STRUCTURE_SLOTS.has(id)) return null;
  return p[STRUCTURE_PERMIT[id.slice('cells.'.length) as StructuralOpName]]
    ? null
    : 'disabledReason.structureLocked';
}

const HORIZONTAL = {
  'alignment.left': 'left',
  'alignment.center': 'center',
  'alignment.right': 'right',
} as const;
const VERTICAL = {
  'alignment.top': 'top',
  'alignment.middle': 'center',
  'alignment.bottom': 'bottom',
} as const;

const enabled = (id: ChromeSlotId, active = false, value?: string): ToolbarCommandState => ({
  id,
  enabled: true,
  disabledReason: null,
  active,
  ...(value !== undefined ? { value } : {}),
});

const disabled = (
  id: ChromeSlotId,
  reason: ChromeDisabledReason,
  value?: string
): ToolbarCommandState => ({
  id,
  enabled: false,
  disabledReason: reason,
  active: false,
  ...(value !== undefined ? { value } : {}),
});

/** The disabled reason of a structure refusal. */
function structureReason(refusal: EditRefusal): ChromeDisabledReason {
  if (refusal.code === 'structure-blocked')
    return refusal.about
      ? { key: 'disabledReason.structureBlocked', about: refusal.about }
      : { key: 'disabledReason.structureBlockedUnnamed' };
  if (refusal.code === 'whole-sheet') return { key: 'disabledReason.wholeSheet' };
  return { key: 'disabledReason.structureUnavailable' };
}

/** What a value slot shows for the active cell's format. */
function valueOf(
  id: ChromeSlotId,
  f: ActiveCellFormat | null,
  locale: NumberLocale
): string | undefined {
  if (!f) return undefined;
  switch (id) {
    case 'font.family':
      return f.fontName;
    case 'font.size':
      return String(f.fontSize);
    case 'font.color':
      return f.fontColor;
    case 'fill.color':
      return f.fillColor ?? NO_FILL;
    case 'number.format':
      return numberFormatValue(f.numberFormat, locale);
    case 'number.custom':
      return f.numberFormat.code;
    default:
      return undefined;
  }
}

/** Whether a toggle slot is applied at the active cell. */
function activeOf(id: ChromeSlotId, f: ActiveCellFormat | null): boolean {
  if (!f) return false;
  switch (id) {
    case 'font.bold':
      return f.bold;
    case 'font.italic':
      return f.italic;
    case 'font.underline':
      return f.underline;
    case 'alignment.left':
    case 'alignment.center':
    case 'alignment.right':
      return f.horizontal === HORIZONTAL[id];
    case 'alignment.top':
    case 'alignment.middle':
    case 'alignment.bottom':
      return f.vertical === VERTICAL[id];
    case 'alignment.wrap':
      return f.wrapText;
    default:
      return false;
  }
}

/**
 * The state of one control, from the editor snapshot alone: enabled, pressed, its value, and
 * the reason when it is disabled.
 *
 * @public
 */
export function toolbarCommandState(
  snapshot: EditorSnapshot,
  id: ChromeSlotId,
  host: ToolbarHost = {}
): ToolbarCommandState {
  const control = chromeControlForSlot(id);
  if (!control) return disabled(id, { key: 'disabledReason.notWired' });
  if (MENU_ONLY_SLOTS.has(id)) return menuSlotState(snapshot, id, host);
  if (snapshot.status === 'loading') return disabled(id, { key: 'disabledReason.loading' });
  if (snapshot.status !== 'ready' || !snapshot.selection)
    return disabled(id, { key: 'disabledReason.editorNotReady' });
  const f = snapshot.activeFormat;
  const value = valueOf(id, f, snapshot.numberLocale);
  // In view mode every slot but save edits the workbook, so each is disabled with the mode.
  if (snapshot.mode === 'view' && id !== 'file.save')
    return disabled(id, { key: 'disabledReason.viewMode' }, value);
  // Every slot commits the open edit or undoes it, and the composed text belongs to the input
  // method until it ends (`CellEditing.commitAndMove`).
  if (snapshot.editing && snapshot.composing)
    return disabled(id, { key: 'disabledReason.composing' }, value);
  // Undo of an open edit with changed text cancels the edit (`CellEditing.undo`), so undo is
  // available; with no changed text it closes the edit and steps the history. Redo is refused
  // with a changed edit (`edit-pending`): its commit would clear the redo steps.
  const pending = snapshot.editing && snapshot.editChanged;
  if (id === 'history.undo')
    return snapshot.canUndo || pending
      ? enabled(id)
      : disabled(id, { key: 'disabledReason.nothingToUndo' });
  if (id === 'history.redo')
    return pending
      ? disabled(id, { key: 'disabledReason.editPending' })
      : snapshot.canRedo
        ? enabled(id)
        : disabled(id, { key: 'disabledReason.nothingToRedo' });
  if (id === 'file.save') {
    if (!host.save) return disabled(id, { key: 'disabledReason.noSave' });
    return host.opening?.() ? disabled(id, { key: 'disabledReason.loading' }) : enabled(id);
  }
  // A decimal step reads the stored cells, which changed text does not show yet. The step
  // refuses it with `edit-pending` before any other check, so the state gives it first too.
  if (pending && DECIMAL_STEPS.has(id)) return disabled(id, { key: 'disabledReason.editPending' });
  const locked = lockedBy(snapshot, id);
  if (locked) return disabled(id, { key: locked }, value);
  // While the styles part stays as bytes, a format command is refused (`unsupported-package`).
  // On a sheet that is not protected, Clear formats and Clear all give the default format 0, so
  // they stay enabled. On a protected sheet, the store refuses them, because the protection of
  // each cell is in the unread part (`clearFormatsRange`).
  if (snapshot.partsPastLimit.some((p) => p.kind === 'styles')) {
    if (FORMAT_SLOTS.has(id)) return disabled(id, { key: 'disabledReason.stylesUnread' }, value);
    if (snapshot.protection && (id === 'clear.formats' || id === 'clear.all'))
      return disabled(id, { key: 'disabledReason.stylesUnreadProtected' }, value);
  }
  // A clear of the values refuses a cell with content that it can neither keep nor replace.
  const blocked = snapshot.selectionContentBlocked;
  if ((id === 'clear.contents' || id === 'clear.all') && blocked !== null) {
    const params = { subject: displayName(blocked) };
    // Only Clear all points to Clear formats, and only while Clear formats is enabled.
    const formats = id === 'clear.all' && toolbarCommandState(snapshot, 'clear.formats').enabled;
    const key = formats
      ? 'disabledReason.cellAlternateContentClearAll'
      : 'disabledReason.cellAlternateContent';
    return disabled(id, { key, params }, value);
  }
  // A clear of the values that would break a formula range is refused with `shared-formula`.
  const range = snapshot.selectionFormulaRangeBlocked;
  if ((id === 'clear.contents' || id === 'clear.all') && range !== null) {
    // A shared master blocked by the formula of one cell is named as the cell to leave out,
    // beside that cell.
    const causeCell = snapshot.selectionFormulaRangeCauseCell;
    const masterCell = snapshot.selectionFormulaRangeMasterCell;
    const named = causeCell !== null && masterCell !== null;
    const params: Record<string, string> = named
      ? { causeCell: displayName(causeCell), masterCell: displayName(masterCell) }
      : { subject: displayName(range) };
    // Only Clear all points to Clear formats, and only while Clear formats is enabled.
    const formats = id === 'clear.all' && toolbarCommandState(snapshot, 'clear.formats').enabled;
    const base = !named
      ? 'formulaRange'
      : snapshot.selectionFormulaRangeCauseUnreadable
        ? 'formulaUnreadable'
        : 'formulaCell';
    const key = `disabledReason.${base}${formats ? 'ClearAll' : ''}` as const;
    return disabled(id, { key, params }, value);
  }
  if (STRUCTURE_SLOTS.has(id) && snapshot.structureRefusal)
    return disabled(id, structureReason(snapshot.structureRefusal));
  const step = DECIMAL_STEPS.get(id);
  if (step !== undefined) {
    // Explicit codes step per cell; the `General` numbers of a column share one step. No cell
    // is created but the active one, so the limit of `formatRefusal` does not apply: the
    // selection state has its own limit.
    if (!f) return disabled(id, { key: 'disabledReason.editorNotReady' });
    const refusal = snapshot.decimalSteps[step > 0 ? 'increase' : 'decrease'];
    if (refusal)
      return disabled(id, refusal.reason ?? { key: 'disabledReason.decimalsNoNumberSelection' });
    // A true rule that shows another number format: the step reads and sets the cell format.
    const shown = snapshot.conditionalFormat;
    if (shown !== null && shown.numberFormat.code !== f.numberFormat.code)
      return { ...enabled(id), hint: { key: 'hint.conditionalFormat' } };
    // The other format controls are disabled with `rangeTooLarge`: the hint says why the step
    // is not.
    return snapshot.formatRefusal
      ? { ...enabled(id), hint: { key: 'hint.decimalsLargeSelection' } }
      : enabled(id);
  }
  if (FORMAT_SLOTS.has(id) && snapshot.formatRefusal)
    return disabled(id, { key: 'disabledReason.rangeTooLarge' }, value);
  if ((id === 'clear.formats' || id === 'clear.all') && snapshot.clearRefusal)
    return disabled(id, { key: 'disabledReason.clearTooLarge' }, value);
  const active = control.toggle ? activeOf(id, f) : false;
  const state = enabled(id, active, value);
  // A true rule shows a state that differs from the cell format that this control sets.
  const shown = snapshot.conditionalFormat;
  const differs =
    shown !== null &&
    FORMAT_SLOTS.has(id) &&
    (valueOf(id, shown, snapshot.numberLocale) !== value ||
      (control.toggle && activeOf(id, shown) !== active));
  return differs ? { ...state, hint: { key: 'hint.conditionalFormat' } } : state;
}

/**
 * The state of every slot of a list, from one snapshot, in the order of `ids`.
 *
 * @public
 */
export function toolbarCommandStates(
  snapshot: EditorSnapshot,
  ids: readonly ChromeSlotId[],
  host: ToolbarHost = {}
): readonly ToolbarCommandState[] {
  return ids.map((id) => toolbarCommandState(snapshot, id, host));
}

const refused = (code: EditRefusal['code'], detail: string): EditResult => ({
  ok: false,
  refusal: { code, detail },
});

/**
 * The format patch of a value slot and a picked value, or null when the value is not one. A
 * number format entry is its code in the number locale (`numberFormatChoiceCode`).
 */
export function valuePatch(
  id: ChromeSlotId,
  value: string,
  locale: NumberLocale = EN_US
): FormatPatch | null {
  switch (id) {
    case 'font.family': {
      const name = value.trim();
      return name.length > 0 ? { fontName: name } : null;
    }
    case 'font.size': {
      const size = Number(value);
      return value.trim() !== '' && Number.isFinite(size) ? { fontSize: size } : null;
    }
    case 'font.color':
      return HEX6.test(value) ? { fontColor: value.toUpperCase() } : null;
    case 'fill.color':
      if (value === NO_FILL) return { fill: { type: 'none' } };
      return HEX6.test(value) ? { fill: { type: 'solid', color: value.toUpperCase() } } : null;
    case 'number.format': {
      // The code of the entry in the locale; `runToolbarCommand` writes the built-in id of the
      // entry instead when the workbook gives that id no other code (`builtinOrCode`).
      const code = isNumberFormatChoice(value) ? numberFormatChoiceCode(value, locale) : undefined;
      return code === undefined ? null : { numberFormat: { code } };
    }
    case 'number.custom':
      // A format code, as typed: the store's own checks decide (`numberFormatCodeProblem`).
      return numberFormatCodeProblem(value) ? null : { numberFormat: { code: value } };
    default:
      return null;
  }
}

/** The alignment patch of an alignment slot: a press on the applied alignment resets it. */
function alignmentPatch(id: ChromeSlotId, f: ActiveCellFormat | null): FormatPatch | null {
  const on = activeOf(id, f);
  switch (id) {
    case 'alignment.left':
    case 'alignment.center':
    case 'alignment.right':
      return { horizontal: on ? 'general' : HORIZONTAL[id] };
    case 'alignment.top':
    case 'alignment.middle':
    case 'alignment.bottom':
      // Bottom is the vertical alignment of a cell that states none: `x:alignment/@vertical`
      // defaults to `bottom` (`CT_CellAlignment`, sml.xsd).
      return on
        ? id === 'alignment.bottom'
          ? null
          : { vertical: 'bottom' }
        : { vertical: VERTICAL[id] };
    case 'alignment.wrap':
      return { wrapText: !on };
    default:
      return null;
  }
}

/** The newest host that pressed Save for each running save, which gets the one delivery. */
const saveDeliveries = new WeakMap<Promise<SavedWorkbook | null>, ToolbarHost>();

/** The presses of toolbar saves of each editor, and the newest press that delivered its bytes. */
const saveOrder = new WeakMap<object, { pressed: number; delivered: number }>();

/**
 * Run a slot on the editor: the command of a `command` slot, the picked `value` of a `value` slot,
 * or the save of `file.save`, which hands the bytes to `host.save` when they are ready, after this
 * call returns: `ok` then means that the save started or was already running. One undo step for
 * each edit. A refused edit returns a refusal, never `ok`: the engine's refusal, or one of the
 * toolbar itself, for an unknown slot (`not-editable`), a value that is not a value of the slot
 * (`bad-format`), or a save with no `host.save` (`not-editable`). While `host.opening` returns
 * true, a save is refused with `loading` before any save work starts. While a workbook loads, every
 * slot is refused with `loading` through the engine, which keeps and announces the refusal, also
 * where the toolbar refuses on its own; a save to `host.save` has `command: 'save'`.
 *
 * @public
 */
export function runToolbarCommand(
  editor: XlsxEditorInstance,
  id: ChromeSlotId,
  value?: string,
  host: ToolbarHost = {}
): EditResult {
  if (editor.snapshot().status !== 'loading') return runSlot(editor, id, value, host);
  // While a workbook loads, no slot runs: a toggle such as an alignment would read a format
  // of the old workbook. A save of the host is a save, refused as a save key's is; every other
  // slot is refused by the engine (`refuseLoading`), which keeps and announces the refusal, so
  // the result, the snapshot, and the live region agree.
  if (id === 'file.save' && host.save) {
    editor.save();
    return { ok: false, refusal: { code: 'loading', command: 'save' } };
  }
  return editingOf(editor)?.refuseLoading() ?? { ok: false, refusal: { code: 'loading' } };
}

/**
 * The built-in id that a value of a number format slot names, or undefined: the id of a list
 * entry, or for `number.custom` the active cell's built-in id when the value is the locale's
 * code of that id (from `NumberLocale.builtins`), so a value sent back through `execute` keeps
 * the locale-dependent id.
 */
function builtinIdOf(
  snapshot: EditorSnapshot,
  id: ChromeSlotId,
  value: string
): number | undefined {
  if (id === 'number.format')
    return NUMBER_FORMAT_CHOICES.find((c) => c.value === value)?.id ?? undefined;
  if (id !== 'number.custom') return undefined;
  const active = snapshot.activeFormat?.numberFormat.id;
  if (active === undefined || active === 0 || active >= FIRST_CUSTOM_NUMFMT_ID) return undefined;
  return builtinFormatCode(active, snapshot.numberLocale) === value ? active : undefined;
}

/**
 * The patch of a value slot. A number format slot value that names a built-in id writes that
 * id, unless the workbook gives the id another code: then it writes the code, which the cell
 * shows (`builtinOrCode`). One rule for the list entries and for `number.custom`.
 */
function slotPatch(
  editor: XlsxEditorInstance,
  id: ChromeSlotId,
  value: string
): FormatPatch | null {
  const snapshot = editor.snapshot();
  const patch = valuePatch(id, value, snapshot.numberLocale);
  if (patch?.numberFormat?.code === undefined) return patch;
  const builtin = builtinIdOf(snapshot, id, value);
  if (builtin === undefined) return patch;
  const target = editingOf(editor)?.target;
  const defined = target ? styleLists(target.pkg, target.index).numFmts : new Map<number, string>();
  return { numberFormat: builtinOrCode(builtin, patch.numberFormat.code, defined) };
}

/**
 * Whether a pick of a number format list entry keeps the format of the selection: one cell or
 * one merged range that shows the entry already, with no changed text in an open edit. The
 * pick then writes nothing, adds no undo step, and says nothing, as the number format dialog
 * does for its first code (`numberFormatDialogApplyStep`). A cell format that does not apply
 * its number format (`applyNumberFormat` absent) shows the same entry, so it gets no new
 * cell format either.
 */
function pickKeeps(editor: XlsxEditorInstance, value: string): boolean {
  const snapshot = editor.snapshot();
  const format = snapshot.activeFormat?.numberFormat;
  // A disabled list (view mode, a locked sheet, a composing input method) goes on to
  // `applyFormat`, so the pick gives the refusal that the state names, never `ok`.
  return (
    toolbarCommandState(snapshot, 'number.format').enabled &&
    format !== undefined &&
    snapshot.selectionSingle &&
    !editingOf(editor)?.changed &&
    numberFormatValue(format, snapshot.numberLocale) === value
  );
}

function runSlot(
  editor: XlsxEditorInstance,
  id: ChromeSlotId,
  value: string | undefined,
  host: ToolbarHost
): EditResult {
  const control = chromeControlForSlot(id);
  if (!control) return refused('not-editable', `no chrome slot ${id}`);
  const menu = runMenuSlot(editor, id, host);
  if (menu) return menu;
  if (control.kind === 'value') {
    // A kept pick still closes an open edit with no change, as every other slot does.
    if (id === 'number.format' && value !== undefined && pickKeeps(editor, value))
      return editingOf(editor)?.commitOpenEdit() ?? { ok: true };
    const patch = value === undefined ? null : slotPatch(editor, id, value);
    if (!patch) return refused('bad-format', `${id}: not a value of the slot`);
    return editor.applyFormat(patch);
  }
  switch (id) {
    case 'history.undo':
      return editor.undo();
    case 'history.redo':
      return editor.redo();
    case 'font.bold':
      return editor.toggleFormat('bold');
    case 'font.italic':
      return editor.toggleFormat('italic');
    case 'font.underline':
      return editor.toggleFormat('underline');
    case 'borders.outline':
      return editor.applyFormat({ borders: { edges: 'outline' } });
    case 'borders.all':
      return editor.applyFormat({ borders: { edges: 'all' } });
    case 'borders.none':
      return editor.applyFormat({ borders: { edges: 'none' } });
    case 'cells.insertRows':
      return editor.insertRows();
    case 'cells.deleteRows':
      return editor.deleteRows();
    case 'cells.insertColumns':
      return editor.insertColumns();
    case 'cells.deleteColumns':
      return editor.deleteColumns();
    case 'number.increaseDecimals':
    case 'number.decreaseDecimals':
      return stepDecimalsOf(editor, DECIMAL_STEPS.get(id)!);
    case 'clear.contents':
      return editor.clearValues();
    case 'clear.formats':
      return editor.clearFormats();
    case 'clear.all':
      return editor.clearAll();
    case 'file.save': {
      if (!host.save) return refused('not-editable', 'the host takes no saved bytes');
      // The host opens another workbook: no save work starts, so the press cannot delay the
      // open. The editor keeps and announces the refusal, as for a save during a load.
      if (host.opening?.()) {
        editingOf(editor)?.refuseSave({ code: 'loading' });
        return { ok: false, refusal: { code: 'loading', command: 'save' } };
      }
      const saving = editor.saveInTasks();
      // The editor keeps a refused save in the snapshot (`command: 'save'`) at the call, and
      // announces it: no workbook is ready, or an open cell edit does not commit.
      const refusal = editor.snapshot().refusal;
      if (refusal?.command === 'save') return { ok: false, refusal };
      if (editor.snapshot().status !== 'ready')
        return refused('not-editable', 'no workbook is open');
      // A press during a save of the same workbook and revision joins it: one delivery.
      const shared = sharedSaveOf(saving);
      const joined = saveDeliveries.has(shared);
      saveDeliveries.set(shared, host);
      if (joined) return { ok: true };
      // Saves of different revisions can end out of order. Older bytes never follow newer bytes.
      const order = saveOrder.get(editor) ?? { pressed: 0, delivered: 0 };
      saveOrder.set(editor, order);
      const press = ++order.pressed;
      void saving.then((saved) => {
        if (!saved || press < order.delivered) return;
        order.delivered = press;
        saveDeliveries.get(shared)?.save?.(saved.bytes, saved);
      });
      host.saveStart?.(saving);
      return { ok: true };
    }
    default: {
      if (!id.startsWith('alignment.')) return refused('not-editable', `${id} runs no command`);
      // The press on bottom alignment when the cell is bottom aligned changes nothing.
      const patch = alignmentPatch(id, editor.snapshot().activeFormat);
      return patch ? editor.applyFormat(patch) : { ok: true };
    }
  }
}

/**
 * Add or remove one decimal place in each cell of the selection, from the cell's own number format
 * (`CellEditing.stepDecimals`), in one undo step. On a selection of more cells, every `General`
 * number of a column takes the step of the first `General` number of that column (the active cell
 * first, then row order) when it succeeds; else each `General` number steps from its own places. A
 * `General` number in exponent form steps alone. A cell that cannot step keeps its format. When no
 * cell can step, the command is refused with `decimals-unchanged`, and its `reason` is the
 * disabled reason of the control. Past `MAX_DECIMAL_STEP_CELLS` cells that exist inside
 * the selection, it is refused with `decimals-too-large`. An open edit with changed text refuses it
 * with `edit-pending`. The slot state names each as a disabled reason first.
 */
function stepDecimalsOf(editor: XlsxEditorInstance, step: 1 | -1): EditResult {
  return editingOf(editor)?.stepDecimals(step) ?? refused('not-editable', 'no workbook is open');
}
