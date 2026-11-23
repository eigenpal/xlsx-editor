// One chrome slot as a hook, after docx-editor/packages/react/src/editor/useEditorCommand.ts
// and PPTX Editor's slot-only form (`packages/react/src/context.ts`).
//
// The state comes from core's `toolbarCommandState` over the editor snapshot, and `execute`
// runs core's `runToolbarCommand`. The toolbar reads every slot through this hook, so a host
// control and the toolbar show the same answer.
//
// Differences from DOCX Editor: the hook takes a slot only, as PPTX Editor's does, because
// every command of the workbook editor has a slot. `execute` takes the picked value of a value
// slot as a string. The state is a pure function of the snapshot, so the selector reads the
// snapshot and does not ask the editor again. The state also has `hint`, the translated hint of
// an enabled command; DOCX Editor's hook has no hint.

import { createContext, useCallback, useContext, useMemo } from 'react';
import {
  editorCommandSlice,
  runToolbarCommand,
  sameEditorCommandSlice,
  type ChromeSlotId,
  type EditorSnapshot,
  type ToolbarHost,
} from '@xlsx-editor.dev/core/editor';
import type { TFunction } from '@xlsx-editor.dev/i18n';
import { useEditorState, useTranslation, useXlsxEditor } from './context.ts';

/** What the enclosing toolbar gives its slots: the receiver of saved bytes, and its open state. */
export const ToolbarHostContext = createContext<ToolbarHost>({});

/** A catalog key that the engine names, translated. */
export function translateKey(
  t: TFunction,
  key: string,
  params?: Readonly<Record<string, string | number>>
): string {
  return t(key as Parameters<TFunction>[0], params ? { ...params } : undefined);
}

/**
 * The live state of one chrome slot, and its action.
 *
 * @public
 */
export interface EditorCommandState {
  /**
   * Run the slot's command. Pass the picked value of a value slot, such as `'Arial'` for
   * `font.family` or `'FF0000'` for `font.color`.
   *
   * @returns `true` when the editor ran the command; `false` when it refused it, or before
   * the editor mounts.
   */
  readonly execute: (value?: string) => boolean;
  /** Whether the command is applied at the active cell, such as bold on a bold cell. */
  readonly isActive: boolean;
  /** Whether the editor runs the command now. */
  readonly isEnabled: boolean;
  /** The engine's reason when the command is disabled, translated. Null when it is enabled. */
  readonly disabledReason: string | null;
  /**
   * The engine's hint for an enabled command, translated, or null. The text of
   * `hint.conditionalFormat` means that a true conditional formatting rule shows another state
   * of this control at the active cell; the command still reads and sets the cell format.
   */
  readonly hint: string | null;
  /**
   * What a value slot shows: the font name, the size in points, the color as `RRGGBB` (or
   * `none` for no fill), the number format choice, or the format code of the active cell for
   * `number.custom`. Null for a slot that has no value.
   */
  readonly value: string | null;
}

/**
 * Bind a chrome slot, such as `font.bold` or `history.undo`, to the editor. The component
 * re-renders only when the slot's own state changes, and the result keeps its identity while
 * its fields are equal.
 *
 * `file.save` hands the bytes to the `onSave` of the enclosing toolbar. Outside a toolbar it is
 * disabled: call `save()` on the editor ref instead.
 *
 * @public
 */
export function useEditorCommand(slot: ChromeSlotId): EditorCommandState {
  const editor = useXlsxEditor();
  const host = useContext(ToolbarHostContext);
  const { t } = useTranslation();
  const select = useCallback(
    (snapshot: EditorSnapshot) => editorCommandSlice(snapshot, slot, host, t),
    [slot, host, t]
  );
  const slice = useEditorState(select, sameEditorCommandSlice);
  const execute = useCallback(
    (value?: string): boolean => (editor ? runToolbarCommand(editor, slot, value, host).ok : false),
    [editor, slot, host]
  );
  return useMemo(
    () => ({
      execute,
      isActive: slice.active,
      isEnabled: slice.enabled,
      disabledReason: slice.disabledReason,
      hint: slice.hint,
      value: slice.value,
    }),
    [execute, slice]
  );
}
