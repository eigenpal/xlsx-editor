// The sheet commands of the editor (design D1 and D4 of the change `workbook-structure-editing`):
// add, rename, delete, and move a sheet. Each command is one store transaction and one undo
// step, through `CellEditing.run`, so an open cell edit commits first and a refusal shows
// through the refusal channel with its code.
//
// The history entry of a sheet op names the sheet that it changes and the selection that a
// redo restores (`RunHistory`): a redo of an add shows the added sheet, a redo of a rename
// selects again on the new name, and a redo of a delete of the shown sheet keeps the sheet
// that shows after it. An undo or a redo of an add or a delete names the change and the shown
// sheet (`sheetStepAnnouncement`).
//
// The editor keeps its view state (scroll, zoom, selection) by sheet position. A sheet op, or
// its undo, changes the positions, so `followSheets` moves each view to the new position of
// its sheet, found by its part name, and keeps the shown sheet when it still exists. When the
// shown sheet is gone, the nearest visible worksheet shows.

import type { GridSelection } from '../contracts/index.ts';
import { sheetByName, sheetNameKey, sheetsByKey, type WorkbookIndex } from '../store/index.ts';
import type { CellEditing } from './cell-editing.ts';
import type { EditRefusal, EditResult } from './cell-input.ts';
import { editorSheets, type EditorSheet } from './editor-model.ts';

/** What the sheet commands read from the editor. */
export interface SheetCommandHost {
  readonly edits: CellEditing;
  /** The workbook index of the current revision, when a workbook is open. */
  index(): WorkbookIndex | null;
  /** The position of the shown sheet. */
  active(): number;
  t(key: string, params?: Readonly<Record<string, string | number>>): string;
  /** Show the sheet at a position. */
  show(index: number): boolean;
  /** Set the text of the live region of the grid. */
  announce(text: string): void;
}

/** The sheet commands of one editor. */
export class SheetCommands {
  readonly #host: SheetCommandHost;

  constructor(host: SheetCommandHost) {
    this.#host = host;
  }

  /**
   * Show the next sheet that can show, in `direction` (Control and Page Down or Page Up), past
   * hidden sheets and chartsheets.
   */
  cycle(direction: 1 | -1): void {
    const count = this.#host.index()?.sheets.length ?? 0;
    const active = this.#host.active();
    for (let i = 1; i <= count; i += 1)
      if (this.#host.show((active + direction * i + count) % count)) return;
  }

  /**
   * The default name of a new sheet: `sheetTabs.newSheetName` with the smallest number from 1
   * that no sheet name has, compared without case.
   */
  newSheetName(index: WorkbookIndex): string {
    const taken = new Set(index.sheets.map((s) => sheetNameKey(s.name)));
    for (let n = 1; n <= index.sheets.length + 1; n += 1) {
      const name = this.#host.t('sheetTabs.newSheetName', { n });
      if (!taken.has(sheetNameKey(name))) return name;
    }
    return this.#host.t('sheetTabs.newSheetName', { n: index.sheets.length + 1 });
  }

  /**
   * Add a sheet after the shown sheet, and show it. The live region announces the added sheet
   * (`grid.sheetAdded`), as a delete announces its sheet.
   */
  add(name?: string): EditResult {
    const index = this.#host.index();
    if (!index) return this.#host.edits.refuse({ code: 'not-editable' });
    const at = this.#host.active() + 1;
    const added = name ?? this.newSheetName(index);
    // A redo shows the added sheet at its first cell, as the add does.
    const origin = { row: 0, col: 0 };
    const r = this.#host.edits.run([{ op: 'addSheet', name: added, index: at }], {
      sheet: added,
      after: () => ({
        sheet: added,
        range: [0, 0, 0, 0],
        active: origin,
        anchor: origin,
        focus: origin,
      }),
    });
    if (r.ok) {
      this.#host.show(at);
      this.#host.announce(this.#host.t('grid.sheetAdded', { name: added }));
    }
    return r;
  }

  /** Rename a sheet. A redo selects again on the new name when the sheet shows. */
  rename(position: number, name: string): EditResult {
    const sheet = this.#sheetAt(position);
    if (typeof sheet !== 'string') return this.#host.edits.refuse(sheet);
    return this.#host.edits.run([{ op: 'renameSheet', sheet, name }], {
      sheet,
      after: (before) => (before.sheet === sheet ? { ...before, sheet: name } : before),
    });
  }

  /**
   * Delete a sheet. `followSheets` shows a neighbor when the shown sheet goes. The live region
   * announces the delete and that undo restores the sheet (`grid.sheetDeleted`).
   */
  delete(position: number): EditResult {
    const sheet = this.#sheetAt(position);
    if (typeof sheet !== 'string') return this.#host.edits.refuse(sheet);
    // A redo of a delete of the shown sheet keeps the sheet that shows after it.
    const r = this.#host.edits.run([{ op: 'deleteSheet', sheet }], {
      sheet,
      after: (before) => (before.sheet === sheet ? null : before),
    });
    if (r.ok) this.#host.announce(this.#host.t('grid.sheetDeleted', { name: sheet }));
    return r;
  }

  /**
   * Move a sheet. The shown sheet stays shown at its new position (`followSheets`). The live
   * region announces the moved sheet (`grid.sheetMoved`) when its position changed. The history
   * names the moved sheet, so an undo and a redo announce it. A sheet that is not a worksheet
   * (a chartsheet, a macro sheet, a dialog sheet, or a sheet whose part is missing) is refused
   * with `structure-blocked` and the `about` kind `not-worksheet`, as a delete refuses it: the
   * history names only worksheets, and the tabs disable the move (`disabledReason.notWorksheet`).
   * The view mode, an open edit, and the workbook protection refuse first, as for a delete.
   */
  move(position: number, to: number): EditResult {
    const sheet = this.#sheetAt(position);
    if (typeof sheet !== 'string') return this.#host.edits.refuse(sheet);
    const r = this.#host.edits.run(
      [{ op: 'moveSheet', sheet, index: to }],
      { sheet, after: (before) => before },
      undefined,
      (target) => {
        const entry = target.index.sheets[position];
        return entry && (entry.kind !== 'worksheet' || !entry.partName)
          ? {
              code: 'structure-blocked',
              subject: sheet,
              about: { kind: 'not-worksheet', name: sheet },
              detail: `${sheet}: the sheet is not a worksheet that the editor can move`,
            }
          : null;
      }
    );
    // A move to the same position changes nothing, so it announces nothing.
    if (r.ok && to !== position)
      this.#host.announce(this.#host.t('grid.sheetMoved', { name: sheet }));
    return r;
  }

  /**
   * The name that the ops resolve to the sheet at a tab position, or the refusal of the
   * command. The tab names its sheet by identity: when the name finds another sheet, or none
   * (two sheets of a file with the same name), the command is refused with
   * `sheet-name-ambiguous` and the other sheet as the subject, so it never acts on another
   * sheet. No sheet at the position is `unknown-sheet`.
   */
  #sheetAt(position: number): string | EditRefusal {
    const index = this.#host.index();
    const entry = Number.isInteger(position) ? index?.sheets[position] : undefined;
    if (!index || !entry) return { code: 'unknown-sheet' };
    if (sheetByName(index, entry.name)?.index === entry.index) return entry.name;
    const twin = sheetsByKey(index, entry.name).find((s) => s.index !== entry.index);
    return { code: 'sheet-name-ambiguous', subject: twin?.name ?? entry.name };
  }
}

/** The key that finds a sheet across revisions: its part name, or its name for a gap. */
const sheetKey = (s: WorkbookIndex['sheets'][number]): string => s.partName ?? `#${s.name}`;

/** The sheet list, the views, and the shown sheet after a change of the sheet list. */
export function followSheets<V extends { selection: GridSelection }>(
  before: WorkbookIndex,
  after: WorkbookIndex,
  views: Map<number, V>,
  active: number
): { sheets: EditorSheet[]; views: Map<number, V>; active: number } {
  const sheets = editorSheets(after);
  const same =
    before.sheets.length === after.sheets.length &&
    before.sheets.every((s, i) => sheetKey(s) === sheetKey(after.sheets[i]!));
  if (same) return { sheets, views, active };
  const position = new Map<string, number>();
  after.sheets.forEach((s, i) => position.set(sheetKey(s), i));
  const moved = new Map<number, V>();
  for (const [old, view] of views) {
    const entry = before.sheets[old];
    const p = entry ? position.get(sheetKey(entry)) : undefined;
    if (p === undefined) continue;
    view.selection = { ...view.selection, sheetIndex: p };
    moved.set(p, view);
  }
  const shown = before.sheets[active];
  let next = shown ? position.get(sheetKey(shown)) : undefined;
  if (next === undefined) {
    // The shown sheet is gone: show the nearest sheet that can show.
    const ok = (s: EditorSheet | undefined) => !!s && s.visible && s.worksheet;
    const from = Math.min(Math.max(active, 0), sheets.length - 1);
    for (let d = 0; d < sheets.length && next === undefined; d += 1) {
      if (ok(sheets[from + d])) next = from + d;
      else if (ok(sheets[from - d])) next = from - d;
    }
  }
  return { sheets, views: moved, active: next ?? -1 };
}

/** The catalog keys of the undo and redo announcements of a sheet op. */
const SHEET_STEPS: Readonly<Record<string, Readonly<Record<'undo' | 'redo', string>>>> = {
  addSheet: { undo: 'grid.sheetAddUndone', redo: 'grid.sheetAddRedone' },
  deleteSheet: { undo: 'grid.sheetDeleteUndone', redo: 'grid.sheetDeleteRedone' },
  renameSheet: { undo: 'grid.sheetRenameUndone', redo: 'grid.sheetRenameRedone' },
  moveSheet: { undo: 'grid.sheetMoveUndone', redo: 'grid.sheetMoveRedone' },
};

/**
 * The announcement of an undo or a redo of a sheet op: the change, with the name of the sheet
 * after the step (the added, deleted, renamed, or moved sheet), then the shown sheet and its
 * selection. Undefined for another history entry, which `stepAnnouncement` names.
 */
export function sheetStepAnnouncement(
  t: (key: string, params?: Readonly<Record<string, string>>) => string,
  step: 'undo' | 'redo',
  label: string | null,
  name: string | null,
  shown: { readonly sheet: string; readonly selection: string } | undefined
): string | undefined {
  const keys = label !== null && Object.hasOwn(SHEET_STEPS, label) ? SHEET_STEPS[label] : null;
  if (!keys || name === null || !shown) return undefined;
  return t(keys[step]!, { name, sheet: shown.sheet, selection: shown.selection });
}
