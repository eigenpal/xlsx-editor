// The keys of a chrome field that edits the open cell edit, such as the formula bar: the
// platform-neutral rules that each adapter reads, so the React and Vue bars act the same.
//
// - Enter, Tab, and Shift with them commit and move, as in the cell editor (`editMoveOf`).
// - A key that an input method composes belongs to it (`composesKey`).
// - The history keys step the workbook history only when the field holds no typing of its own:
//   no open cell edit and no typed draft (`fieldHistoryKey`). Else they belong to the field's
//   own undo, as in the cell editor. PPTX Editor's `panelHistoryKey` keeps every history key of
//   a text area in the field; the formula bar is a text area with no typing of its own while no
//   edit is open, so its keys step the history then, as on the grid.
// - A refusal of a typed reference marks the reference field, and a refusal of the edit text
//   marks the field that shows the text (`refusesReference`, `refusesEditText`).

import type { EditMove } from './cell-editor.ts';
import { historyShortcutOf, type HistoryShortcutKeys } from './history-shortcut.ts';
import type { EditRefusalCode } from './cell-input.ts';

/**
 * Modifier and key state of one keydown, as both DOM and framework key events carry it.
 *
 * @public
 */
export interface FieldKeys extends HistoryShortcutKeys {
  readonly altKey: boolean;
}

/**
 * The move of Enter (`down`), Shift+Enter (`up`), Tab (`right`), or Shift+Tab (`left`) in a
 * field that edits the open cell edit, for `commitEdit(move)`, or null for another key. As in
 * the cell editor, Alt+Enter is a line break, and Control or Command with Enter commits.
 *
 * @public
 */
export function editMoveOf(keys: FieldKeys): EditMove | null {
  if (keys.key === 'Enter') return keys.altKey ? null : keys.shiftKey ? 'up' : 'down';
  if (keys.key === 'Tab') return keys.shiftKey ? 'left' : 'right';
  return null;
}

/**
 * Whether a keydown belongs to an input method that composes text: the event says so
 * (`isComposing`, key code 229), or the field saw a composition start and no end (`composing`).
 *
 * @public
 */
export function composesKey(
  event: { readonly isComposing?: boolean; readonly keyCode?: number },
  composing: boolean
): boolean {
  return composing || event.isComposing === true || event.keyCode === 229;
}

/**
 * The history step that a keydown in a field asks of the workbook, or null: the history keys
 * (`historyShortcutOf`) when the field holds no typing of its own (`typing` false: no open cell
 * edit, and no typed draft). With typing, the keys belong to the field.
 *
 * @public
 */
export function fieldHistoryKey(keys: FieldKeys, typing: boolean): 'undo' | 'redo' | null {
  return typing ? null : historyShortcutOf(keys);
}

/** The refusals of the name box (`selectReference`), which refuse the typed reference. */
const GO_TO_CODES: ReadonlySet<EditRefusalCode> = new Set([
  'go-to-reference',
  'go-to-name',
  'go-to-sheet',
]);

/**
 * Whether the refusal `code` refuses the reference that a field such as the name box gave to
 * `selectReference`. Another refusal of `selectReference`, such as a refused commit of the open
 * edit, refuses the edit and not the reference.
 *
 * @public
 */
export function refusesReference(code: EditRefusalCode): boolean {
  return GO_TO_CODES.has(code);
}

/**
 * Whether the refusal `code`, while a cell edit is open, can refuse the text of that edit, so
 * that a field that shows the text, such as the formula bar, marks it invalid. A refusal of a
 * typed reference, a composing input method, a pending edit, or a dropped edit does not. The
 * code alone cannot tell a refusal of the open edit from a refusal of a call that leaves the
 * edit open, such as `enterValue(text, at)` for another cell: a field that marks the text
 * marks it only for a refusal that comes while the edit is open.
 *
 * @public
 */
export function refusesEditText(code: EditRefusalCode): boolean {
  return (
    !GO_TO_CODES.has(code) &&
    code !== 'composing' &&
    code !== 'edit-pending' &&
    code !== 'edit-dropped'
  );
}
