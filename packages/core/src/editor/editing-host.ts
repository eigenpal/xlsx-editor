// The types that the editing commands share with the editor that hosts them: the edit target,
// the host interface, and the history of a run. Split from `cell-editing.ts`, which re-exports
// them, so that each file stays within its line limit.

import type { GridPoint, GridSelection } from '../contracts/index.ts';
import type {
  HistorySelection,
  SheetEntry,
  WorkbookDocumentStore,
  WorkbookIndex,
  XlsxPackage,
} from '../store/index.ts';
import type { NumberLocale } from '../cascade/index.ts';
import type { Merges } from './selection.ts';
import type { EditorTranslate } from './cell-input.ts';
import type { ShownGeneralText } from './decimal-selection.ts';

/** The workbook and the shown sheet that an edit applies to. */
export interface EditTarget {
  readonly store: WorkbookDocumentStore;
  readonly pkg: XlsxPackage;
  readonly index: WorkbookIndex;
  /** The position of the shown sheet. */
  readonly sheet: number;
  readonly entry: SheetEntry;
  readonly selection: GridSelection;
  /** The moving corner of the selection, which Shift+Arrow extends. */
  readonly focus: GridPoint;
  /** The store revision that `pkg` and `selection` were read at. */
  readonly revision: number;
}

/** The font of a cell's text: a CSS family list, and the size in points. */
export interface CellFont {
  readonly family: string;
  readonly sizePt: number;
}

/** What the editing commands read and call on the editor. */
export interface EditingHost {
  /** The edit target, or null when no workbook is ready to edit. */
  target(): EditTarget | null;
  /** Whether the editor is in `view` mode, where every edit command is refused. */
  viewing(): boolean;
  /** Whether a workbook loads, while every command is refused with `loading`. */
  loading(): boolean;
  /** The number locale now: it can change with `setNumberLocale`. */
  locale(): NumberLocale;
  t: EditorTranslate;
  /** The grid of the mounted surface, or null. */
  surface(): { readonly grid: HTMLElement } | null;
  /** Say `text` in the live region (`LiveAnnouncer.say`). The same text again is said again. */
  say(text: string): void;
  /** Empty the live region, so its words are not read again (`LiveAnnouncer.silence`). */
  silence(): void;
  merges(): Merges;
  /** The font that the painted grid draws a cell's text with. */
  cellFont(cell: GridPoint): CellFont;
  /** The text that the grid shows for a `General` number, as the snapshot's state reads it. */
  shownGeneral?(): ShownGeneralText | undefined;
  /**
   * Say `text` at the next paint when the change of the last command leaves the words of the
   * active cell as they were (`LiveAnnouncer.confirm`).
   */
  confirm?(text: string): void;
  publish(): void;
  /** Publish the open edit's text alone: a key in the cell editor changes nothing else. */
  publishDraft(): void;
  schedule(): void;
  render(): void;
  reveal(cell: GridPoint): void;
  /**
   * Make the cell one step from `cell` active. With `inRange`, inside a multi-cell selection the
   * active cell moves in it and the selection stays (`cycleInRange`).
   */
  moveFrom(cell: GridPoint, move: readonly [number, number], inRange: boolean): void;
  /**
   * After a committed undo or redo: select `selection` on its sheet when there is one, and
   * announce the step. `sheet` and `label` are the sheet and the label of the history entry
   * (`addSheet`), so a sheet op can name its change. `changed` is the name of that sheet after
   * the step, which differs from `sheet` after the redo of a rename.
   */
  restore(
    step: 'undo' | 'redo',
    selection: HistorySelection | null,
    sheet: string | null,
    label: string | null,
    changed: string | null
  ): void;
}

/**
 * The history of a command that `CellEditing.run` applies, when it is not the default (the
 * shown sheet, and the selection before as the selection after). A sheet op names the sheet
 * that it changes, and a paste its target sheet, with the selection that a redo restores.
 */
export interface RunHistory {
  /**
   * The sheet that the command changes: the added, renamed, deleted, or moved sheet, or the
   * target sheet of a paste.
   */
  readonly sheet: string;
  /**
   * The selection after the command, from the selection before it. Null keeps none, so a
   * redo keeps the sheet that the editor shows.
   */
  readonly after: (before: HistorySelection) => HistorySelection | null;
}
