/** @jsxImportSource react */
// One chrome slot as the toolbar sees it: the public `useEditorCommand` state and command, and
// the slot's control and label. The adapter keeps no command state of its own: every part of
// the toolbar reads a slot through this hook, so the bar, the menus, the overflow panel, and a
// host's own control show the same answer, from the one source `toolbarCommandState`.

import { useCallback, useContext, type ReactNode } from 'react';
import {
  chromeControlForSlot,
  editorRowsSlice,
  slotWords,
  type ChromeControl,
  type ChromeSlotId,
} from '@xlsx-editor.dev/core/editor';
import { useEditorState, useTranslation, useXlsxEditor } from '../context.ts';
import {
  ToolbarHostContext,
  translateKey,
  useEditorCommand,
  type EditorCommandState,
} from '../useEditorCommand.ts';

export { ToolbarHostContext, translateKey };

/** The state, words, and command of one slot. */
export interface ToolbarSlot extends EditorCommandState {
  readonly slot: ChromeSlotId;
  readonly control: ChromeControl;
  /** The accessible name. */
  readonly label: string;
  /** The disabled reason or the hint: the accessible description, or null. */
  readonly description: string | null;
  /**
   * The tooltip: the label, and the disabled reason after it when there is one; else the label
   * with the hint on a second line.
   */
  readonly title: string;
  /** Give the keyboard back to the grid, and keep the selection. */
  focusGrid(): void;
}

/** One slot of the chrome registry. */
export function useToolbarSlot(slot: ChromeSlotId): ToolbarSlot {
  const editor = useXlsxEditor();
  const { t } = useTranslation();
  const command = useEditorCommand(slot);
  const control = chromeControlForSlot(slot)!;
  return {
    ...command,
    ...slotWords(control.labelKey, command, t),
    slot,
    control,
    focusGrid: () => editor?.focus(),
  };
}

/** The state of a control that opens rows: enabled, and the reason when disabled. */
export interface RowsState {
  readonly isEnabled: boolean;
  readonly disabledReason: string | null;
}

const sameRows = (
  a: ReturnType<typeof editorRowsSlice>,
  b: ReturnType<typeof editorRowsSlice>
): boolean => a.enabled === b.enabled && a.disabledReason === b.disabledReason;

/**
 * The state of a control that only opens the rows `slots`, such as a group menu or a submenu:
 * disabled with the reason of its first row when every row is disabled (`toolbarRowsState`).
 * `key` names the list, so a new array of the same slots reads no new state.
 */
export function useRowsState(slots: readonly ChromeSlotId[]): RowsState {
  const host = useContext(ToolbarHostContext);
  const { t } = useTranslation();
  const key = slots.join('\n');
  const select = useCallback(
    (snapshot: Parameters<typeof editorRowsSlice>[0]) =>
      editorRowsSlice(snapshot, key ? (key.split('\n') as ChromeSlotId[]) : [], host, t),
    [key, host, t]
  );
  const slice = useEditorState(select, sameRows);
  return { isEnabled: slice.enabled, disabledReason: slice.disabledReason };
}

/**
 * Registry icon paths as inline SVG (Material Symbols view box), hidden from assistive
 * technology.
 */
export function ChromeIcon({
  paths,
  className,
}: {
  readonly paths: readonly string[] | null;
  readonly className?: string;
}): ReactNode {
  if (!paths) return null;
  return (
    <svg className={className} viewBox="0 -960 960 960" aria-hidden="true" focusable="false">
      {paths.map((d) => (
        <path key={d} d={d} />
      ))}
    </svg>
  );
}
