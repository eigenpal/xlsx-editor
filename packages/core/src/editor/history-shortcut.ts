// Taken from docx-editor/packages/core/src/editor/history-shortcut.ts at ec393d3c (EigenPal).
// Intended differences: the doc comments only; the code is unchanged. The grid reads the
// history keys through it to step the store history. The cell editor reads them only to keep
// them: in an open cell editor the keys belong to the input's own undo and redo.

/** The keys that undo and redo: the grid steps the history, and the cell editor keeps them. */

/** Modifier and key state of one keydown, as both DOM and framework key events carry it. */
export interface HistoryShortcutKeys {
  readonly key: string;
  readonly metaKey: boolean;
  readonly ctrlKey: boolean;
  readonly shiftKey: boolean;
}

/**
 * Which history step a keydown asks for, or `null` for any other key.
 *
 * Ctrl+Y is the Windows redo; Ctrl/Cmd+Shift+Z is the macOS one. Both are accepted on every
 * platform, also Meta+Y, which the grid takes as redo.
 */
export function historyShortcutOf(event: HistoryShortcutKeys): 'undo' | 'redo' | null {
  if (!(event.metaKey || event.ctrlKey)) return null;
  const key = event.key.toLowerCase();
  if (key === 'y') return 'redo';
  if (key === 'z') return event.shiftKey ? 'redo' : 'undo';
  return null;
}
