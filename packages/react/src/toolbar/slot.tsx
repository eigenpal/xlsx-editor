/** @jsxImportSource react */
// One chrome slot as the toolbar sees it: the public `useEditorCommand` state and command, and
// the slot's control and label. The adapter keeps no command state of its own: every part of
// the toolbar reads a slot through this hook, so the bar, the menus, the overflow panel, and a
// host's own control show the same answer, from the one source `toolbarCommandState`.

import type { ReactNode } from 'react';
import {
  chromeControlForSlot,
  slotWords,
  type ChromeControl,
  type ChromeSlotId,
} from '@xlsx-editor.dev/core/editor';
import { useTranslation, useXlsxEditor } from '../context.ts';
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
