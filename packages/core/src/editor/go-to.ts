// The name box: read a typed reference and select it (`selectReference`).
//
// The text is read by the bounded formula tokenizer (design D4), never by a regular expression,
// so a long or hostile text costs time in proportion to its capped length. The text must be one
// reference: a cell (`B2`, `$B$2`), a range (`B2:C5`), whole columns (`C:E`), or whole rows
// (`3:5`), with an optional sheet name (`Sheet2!A1`, `'My sheet'!A1:B2`), or a defined name
// whose text is one such reference with a sheet name. Letters can be in any case. Anything else
// is refused with a typed reason, never ignored: a defined name that names no one range on a
// visible worksheet (`go-to-name`), a sheet that is not a visible worksheet (`go-to-sheet`),
// and any other text (`go-to-reference`). A name that the shown sheet scopes
// (`x:definedName/@localSheetId`, `CT_DefinedName`, sml.xsd) wins over a workbook name, as a
// formula on that sheet reads it.

import type { GridPoint, GridSelection } from '../contracts/index.ts';
import { selectBand } from './selection.ts';
import {
  MAX_COLS,
  MAX_ROWS,
  attr,
  child,
  children,
  intAttr,
  sheetNameKey,
  textOf,
  tokenizeFormula,
  xmlPartOf,
  type FormulaToken,
  type RefEnd,
  type WorkbookIndex,
  type XlsxPackage,
} from '../store/index.ts';
import type { EditRefusal, EditResult } from './cell-input.ts';
import type { CellEditing } from './cell-editing.ts';
import type { EditorSheet } from './editor-model.ts';
import type { XlsxEditorInstance } from './editor-instance.ts';

/** Where a name box reference goes: the sheet position, and the corners of the range. */
export interface GoToTarget {
  readonly sheet: number;
  readonly anchor: GridPoint;
  readonly focus: GridPoint;
  /**
   * Whole columns or whole rows: the selection is that band, as a heading press selects it,
   * and does not grow through merged cells.
   */
  readonly band?: 'columns' | 'rows';
}

/**
 * The longest text that `selectReference` reads, in UTF-16 code units: a longer text is refused
 * with `go-to-reference`, and the name box takes no more.
 *
 * @public
 */
export const MAX_GO_TO_LENGTH = 255;

/** The corner of a reference end; a whole-column end starts or ends at a row edge. */
function corner(end: RefEnd, last: boolean): GridPoint {
  return {
    row: end.row?.index ?? (last ? MAX_ROWS - 1 : 0),
    col: end.col?.index ?? (last ? MAX_COLS - 1 : 0),
  };
}

/** The one token of `text`, or undefined when it is not one token. */
function oneToken(text: string): FormulaToken | undefined {
  if (text.length === 0 || text.length > MAX_GO_TO_LENGTH || text.startsWith('=')) return;
  const read = tokenizeFormula(text);
  return read.ok && read.tokens.length === 1 ? read.tokens[0] : undefined;
}

/**
 * The text of the defined name `name` as the sheet at `scope` reads it: the name that this
 * sheet scopes (`@localSheetId`), else, when `local` is false, the workbook name. `exists` is
 * whether the workbook has a name with this key in any scope.
 */
function definedNameText(
  pkg: XlsxPackage,
  index: WorkbookIndex,
  name: string,
  scope: number,
  local: boolean
): { readonly text?: string; readonly exists: boolean } {
  const key = sheetNameKey(name);
  let workbook: string | undefined;
  let exists = false;
  const root = xmlPartOf(pkg, index.mainPart)?.root;
  for (const el of children(child(root, 'x:definedNames'), 'x:definedName')) {
    if (sheetNameKey(attr(el, 'name') ?? '') !== key) continue;
    exists = true;
    const at = intAttr(el, 'localSheetId');
    if (at === scope) return { text: textOf(el).trim(), exists };
    if (at === undefined && !local) workbook ??= textOf(el).trim();
  }
  return workbook === undefined ? { exists } : { text: workbook, exists };
}

const UNREADABLE: EditRefusal = { code: 'go-to-reference', detail: 'not one reference' };

/** The position of the first visible worksheet with the name `name`, compared without case. */
function visibleSheet(sheets: readonly EditorSheet[], name: string): number {
  const key = sheetNameKey(name);
  return sheets.findIndex((s) => s.visible && s.worksheet && sheetNameKey(s.name) === key);
}

/**
 * Read the name box text `text` in the workbook of `pkg` and `index`, with `sheets`, of which
 * `active` shows. Returns the target, or the refusal that says why the text selects nothing.
 */
export function readGoTo(
  text: string,
  pkg: XlsxPackage,
  index: WorkbookIndex,
  sheets: readonly EditorSheet[],
  active: number
): GoToTarget | EditRefusal {
  // The cap holds before the trim, so the cost of a read follows the capped length.
  const token = text.length > MAX_GO_TO_LENGTH ? undefined : oneToken(text.trim());
  if (token?.kind === 'name') {
    // A defined name goes to its range when its text is one reference with a sheet name. A
    // sheet prefix (`Data!Rates`) reads the name that the sheet scopes, and only that one.
    const prefix = token.sheet;
    if (prefix && (prefix.book !== undefined || prefix.sheets.length !== 1)) return UNREADABLE;
    const bare = prefix ? token.text.slice(prefix.raw.length) : token.text;
    const scope = prefix ? visibleSheet(sheets, prefix.sheets[0]!) : active;
    if (prefix && scope < 0) return { code: 'go-to-sheet', subject: prefix.sheets[0]! };
    const named = definedNameText(pkg, index, bare, scope, prefix !== undefined);
    if (!named.exists) return UNREADABLE;
    const ref = named.text === undefined ? undefined : oneToken(named.text);
    const to = ref?.kind === 'reference' && ref.sheet ? targetOf(ref, sheets, active) : null;
    return to && !('code' in to) ? to : { code: 'go-to-name', subject: token.text };
  }
  return token?.kind === 'reference' ? targetOf(token, sheets, active) : UNREADABLE;
}

/** The target of a reference token, or the refusal of its sheet. */
function targetOf(
  token: Extract<FormulaToken, { kind: 'reference' }>,
  sheets: readonly EditorSheet[],
  active: number
): GoToTarget | EditRefusal {
  const prefix = token.sheet;
  if (prefix && (prefix.book !== undefined || prefix.sheets.length !== 1)) return UNREADABLE;
  let sheet = active;
  if (prefix) {
    const name = prefix.sheets[0]!;
    sheet = visibleSheet(sheets, name);
    if (sheet < 0) return { code: 'go-to-sheet', subject: name };
  }
  const a = corner(token.from, false);
  const b = corner(token.to ?? token.from, true);
  const band = token.shape === 'columns' || token.shape === 'rows' ? token.shape : undefined;
  return {
    sheet,
    anchor: { row: Math.min(a.row, b.row), col: Math.min(a.col, b.col) },
    focus: { row: Math.max(a.row, b.row), col: Math.max(a.col, b.col) },
    ...(band ? { band } : {}),
  };
}

/**
 * Select the range that `text` names (`selectReference`), on the workbook that `shown` holds,
 * or refuse it through the refusal channel of `edits`: with `not-editable` when no workbook is
 * ready (`loading` while one loads, as `edits` maps it). A sheet change or a selection that
 * `editor` refuses, such as a refused commit of an open edit, returns that refusal.
 */
export function goTo(
  text: string,
  edits: CellEditing,
  shown: {
    readonly pkg: XlsxPackage;
    readonly index: WorkbookIndex;
    readonly sheets: readonly EditorSheet[];
    readonly active: number;
  } | null,
  editor: Pick<XlsxEditorInstance, 'setActiveSheet'> & {
    select(anchor: GridPoint, focus: GridPoint, reveal: GridPoint): boolean;
    setSelection(selection: GridSelection, focus: GridPoint, reveal?: GridPoint): void;
  }
): EditResult {
  if (!shown) return edits.refuse({ code: 'not-editable' });
  // The composed text of an open edit belongs to the input method: nothing is read until it ends.
  if (edits.composing) return edits.refuse({ code: 'composing' });
  const to = readGoTo(text, shown.pkg, shown.index, shown.sheets, shown.active);
  if ('code' in to) return edits.refuse(to);
  // The view scrolls to the active cell, the top-left cell of the range, not to its far corner.
  const shows = editor.setActiveSheet(to.sheet);
  if (shows && to.band) {
    // Whole columns or rows select the band, as a heading press does. The open edit, which
    // `setActiveSheet` committed on another sheet, commits here on the same sheet.
    const committed = edits.commitOpenEdit();
    if (!committed.ok) return committed;
    const [from, last] =
      to.band === 'columns' ? [to.anchor.col, to.focus.col] : [to.anchor.row, to.focus.row];
    editor.setSelection(selectBand(to.sheet, to.band, from, last), to.focus, to.anchor);
    return { ok: true };
  }
  if (shows && editor.select(to.anchor, to.focus, to.anchor)) return { ok: true };
  return edits.refusal
    ? { ok: false, refusal: edits.refusal }
    : edits.refuse({ code: 'not-editable' });
}
