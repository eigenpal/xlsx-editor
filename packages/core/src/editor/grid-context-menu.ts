// The grid context menu, DOM-free in its data: the request that a secondary press, Shift+F10, or
// the Menu key makes on the grid, and the default rows of the menu. After the context menu of
// docx-editor/packages/react/src/editor/contextmenu/DocxEditorContextMenu.tsx at f5cc27dfb
// (EigenPal): the menu opens at the pointer, or at the caret for a key; with no handler the
// browser keeps its own menu; and its rows are chrome slots, so a row and its toolbar control
// share one label, one icon, one command, and one enabled state (`toolbarCommandState`). The
// keys and the request shape follow pptx-editor/packages/react/src/ContextMenu.tsx at 9d7d60b7.
//
// Differences from DOCX Editor:
// - The engine makes the request, as an editor event (`EditorEvents.contextMenu`), so the React
//   and Vue parts read one request and the engine selects the pressed cells first.
// - The rows are those of a workbook: the clipboard, the insert and delete of rows and columns,
//   the clear commands, and the merge commands. The rows follow what the press hit: a press on
//   a row heading shows the row commands, and a press on a column heading the column commands.
// - A key opens the menu at the active cell, not at a text caret.

import type { GridRange } from '../contracts/index.ts';
import type { ChromeMenuEntry, ChromeMenuItemEntry } from './chrome-menus.ts';
import type { ChromeSlotId } from './chrome-controls.ts';
import type { GridPointer } from './grid-pointer.ts';
import type { SurfaceElements, listener } from './grid-surface.ts';
import { MAX_COLS, MAX_ROWS } from '../store/index.ts';
import { isRightToLeft } from './chrome-direction.ts';

/**
 * What a grid context menu acts on: cells, whole rows, whole columns, or the whole sheet. It is
 * what the press hit; a key reads the selection.
 *
 * @public
 */
export type GridContextMenuTargetKind = 'cells' | 'rows' | 'columns' | 'all';

/**
 * A request for the grid context menu (`EditorEvents.contextMenu`).
 *
 * @public
 */
export interface GridContextMenuRequest {
  /** Where the menu opens, in window pixels: the pointer, or the active cell for a key. */
  readonly x: number;
  readonly y: number;
  /** `'pointer'` for a secondary press, `'keyboard'` for Shift+F10 or the Menu key. */
  readonly trigger: 'pointer' | 'keyboard';
  /** What the menu acts on. */
  readonly kind: GridContextMenuTargetKind;
  /** The sheet of the selection, 0-based. */
  readonly sheetIndex: number;
  /** The selected range after the press, 0-based and inclusive. */
  readonly range: GridRange;
}

const item = (slot: ChromeSlotId, shortcut?: string): ChromeMenuItemEntry =>
  shortcut === undefined ? { kind: 'item', slot } : { kind: 'item', slot, shortcut };
const SEPARATOR = { kind: 'separator' } as const;

const ROW_COMMANDS: readonly ChromeMenuEntry[] = [
  item('cells.insertRows'),
  item('cells.deleteRows'),
];
const COLUMN_COMMANDS: readonly ChromeMenuEntry[] = [
  item('cells.insertColumns'),
  item('cells.deleteColumns'),
];

/**
 * The default rows of the grid context menu for what the menu acts on, as menu entries of the
 * chrome registry. Each row is a chrome slot.
 *
 * @public
 */
export function gridContextMenuEntries(
  kind: GridContextMenuTargetKind
): readonly ChromeMenuEntry[] {
  const structure =
    kind === 'rows'
      ? ROW_COMMANDS
      : kind === 'columns'
        ? COLUMN_COMMANDS
        : kind === 'cells'
          ? [
              item('cells.insertRows'),
              item('cells.insertColumns'),
              item('cells.deleteRows'),
              item('cells.deleteColumns'),
            ]
          : [];
  return [
    item('edit.cut', 'Mod+X'),
    item('edit.copy', 'Mod+C'),
    item('edit.paste', 'Mod+V'),
    SEPARATOR,
    ...structure,
    SEPARATOR,
    item('clear.contents'),
    item('clear.formats'),
    ...(kind === 'cells' ? [SEPARATOR, item('merge.merge'), item('merge.unmerge')] : []),
    SEPARATOR,
    item('number.custom'),
  ];
}

/** What `bindGridContextMenu` reads and calls. */
export interface GridContextMenuInput {
  readonly pointer: Pick<GridPointer, 'contextPress'>;
  /** The cell editor's input, when one is open: it keeps the browser's menu for its text. */
  input(): HTMLElement | null;
  /** The selection of the shown sheet, for a key, or null when no sheet shows. */
  selection(): { readonly sheetIndex: number; readonly range: GridRange } | null;
  /** The handlers of `EditorEvents.contextMenu`. */
  readonly menus: {
    hasContextMenu(): boolean;
    requestContextMenu(request: GridContextMenuRequest): void;
  };
}

/** What a selection is: whole rows, whole columns, the sheet, or cells. */
function kindOf(
  range: GridRange,
  last: { readonly row: number; readonly col: number }
): GridContextMenuTargetKind {
  const rows = range.start.col === 0 && range.end.col >= last.col;
  const cols = range.start.row === 0 && range.end.row >= last.row;
  if (rows && cols) return 'all';
  return rows ? 'rows' : cols ? 'columns' : 'cells';
}

/** The last row and column of the grid, 0-based. */
const GRID_LAST = { row: MAX_ROWS - 1, col: MAX_COLS - 1 } as const;

/**
 * Listen for the grid context menu on a mounted surface: a secondary press on the stage, and
 * Shift+F10 or the Menu key on the grid. Nothing happens while no handler listens, so the browser
 * shows its own menu. A key opens the menu at the start corner of the active cell, under it.
 * The listeners add their cleanups through `on`.
 */
export function bindGridContextMenu(
  on: ReturnType<typeof listener>,
  els: SurfaceElements,
  s: GridContextMenuInput
): void {
  // A key that opened the menu: the browser can fire `contextmenu` for the same key after it.
  // A secondary press always starts with `pointerdown`, which ends the wait.
  let keyOpened = false;
  on(els.stage, 'pointerdown', () => (keyOpened = false), { capture: true });
  on(els.stage, 'contextmenu', (e) => {
    if (!s.menus.hasContextMenu()) return;
    if (keyOpened) {
      keyOpened = false;
      e.preventDefault();
      return;
    }
    const input = s.input();
    if (input && input.contains(e.target as Node)) return;
    const pressed = s.pointer.contextPress(e);
    if (!pressed) return;
    e.preventDefault();
    const hit = pressed.hit;
    s.menus.requestContextMenu({
      x: e.clientX,
      y: e.clientY,
      trigger: 'pointer',
      kind: hit.kind === 'corner' ? 'all' : hit.kind,
      sheetIndex: pressed.selection.sheetIndex,
      range: pressed.selection.range,
    });
  });
  on(els.scroller, 'keydown', (e) => {
    const chord =
      (e.key === 'F10' && e.shiftKey && !e.ctrlKey && !e.metaKey && !e.altKey) ||
      e.key === 'ContextMenu';
    if (!chord || e.target !== els.scroller || !s.menus.hasContextMenu()) return;
    const selection = s.selection();
    if (!selection) return;
    e.preventDefault();
    e.stopPropagation();
    keyOpened = true;
    const cell = els.overlay.querySelector<HTMLElement>('.xlsx-active-cell');
    const box = (cell ?? els.scroller).getBoundingClientRect();
    s.menus.requestContextMenu({
      x: isRightToLeft(els.scroller) ? box.right : box.left,
      y: box.bottom,
      trigger: 'keyboard',
      kind: kindOf(selection.range, GRID_LAST),
      sheetIndex: selection.sheetIndex,
      range: selection.range,
    });
  });
}
