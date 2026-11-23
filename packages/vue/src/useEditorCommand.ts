// One chrome slot as a composable: the Vue twin of the React adapter's `useEditorCommand.ts`,
// after PPTX Editor's Vue `useEditorCommand` (pptx-editor/packages/vue/src/composables.ts at
// d4502552), which follows DOCX Editor's Vue adapter.
//
// The state comes from core's `toolbarCommandState` over the editor snapshot
// (`editorCommandSlice`), and `execute` runs core's `runToolbarCommand`, as in React. The
// toolbar reads every slot through this composable, so a host control and the toolbar show the
// same answer.
//
// Differences from React: the state fields are computed refs, and the slot may be a ref or a
// getter, as PPTX Editor's and DOCX Editor's Vue composables take it. The enclosing toolbar
// provides its host as a ref (`toolbarHostKey`), so a change of its `onSave` prop changes the
// state at once.

import {
  computed,
  inject,
  shallowRef,
  toValue,
  type ComputedRef,
  type InjectionKey,
  type MaybeRefOrGetter,
  type ShallowRef,
} from 'vue';
import {
  editorCommandSlice,
  runToolbarCommand,
  sameEditorCommandSlice,
  type ChromeSlotId,
  type EditorSnapshot,
  type ToolbarHost,
} from '@xlsx-editor.dev/core/editor';
import type { TFunction } from '@xlsx-editor.dev/i18n';
import { useEditorState, useTranslator, useXlsxEditor } from './context.ts';

/** What the enclosing toolbar gives its slots: the receiver of saved bytes, and its open state. */
export const toolbarHostKey: InjectionKey<Readonly<ShallowRef<ToolbarHost>>> =
  Symbol('xlsxToolbarHost');

/** A catalog key that the engine names, translated. */
export function translateKey(
  t: TFunction,
  key: string,
  params?: Readonly<Record<string, string | number>>
): string {
  return t(key as Parameters<TFunction>[0], params ? { ...params } : undefined);
}

const noHost = shallowRef<ToolbarHost>({});

/**
 * The live state of one chrome slot, and its action: the twin of React's
 * `EditorCommandState`, with refs.
 *
 * @public
 */
export interface UseEditorCommandReturn {
  /**
   * Run the slot's command. Pass the picked value of a value slot, such as `'Arial'` for
   * `font.family` or `'FF0000'` for `font.color`.
   *
   * @returns `true` when the editor ran the command; `false` when it refused it, or before
   * the editor mounts.
   */
  readonly execute: (value?: string) => boolean;
  /** Whether the command is applied at the active cell, such as bold on a bold cell. */
  readonly isActive: ComputedRef<boolean>;
  /** Whether the editor runs the command now. */
  readonly isEnabled: ComputedRef<boolean>;
  /** The engine's reason when the command is disabled, translated. Null when it is enabled. */
  readonly disabledReason: ComputedRef<string | null>;
  /**
   * The engine's hint for an enabled command, translated, or null. The text of
   * `hint.conditionalFormat` means that a true conditional formatting rule shows another state
   * of this control at the active cell; the command still reads and sets the cell format.
   */
  readonly hint: ComputedRef<string | null>;
  /**
   * What a value slot shows: the font name, the size in points, the color as `RRGGBB` (or
   * `none` for no fill), the number format choice, or the format code of the active cell for
   * `number.custom`. Null for a slot that has no value.
   */
  readonly value: ComputedRef<string | null>;
}

/**
 * Bind a chrome slot, such as `font.bold` or `history.undo`, to the editor. The refs change
 * only when the slot's own state changes. The slot may be a ref or a getter: the state follows
 * it.
 *
 * `file.save` hands the bytes to the `onSave` of the enclosing toolbar. Outside a toolbar it is
 * disabled: call `save()` on the editor ref instead.
 *
 * @public
 */
export function useEditorCommand(slot: MaybeRefOrGetter<ChromeSlotId>): UseEditorCommandReturn {
  const editor = useXlsxEditor();
  const host = inject(toolbarHostKey, noHost);
  const t = useTranslator();
  const slice = useEditorState(
    (snapshot: EditorSnapshot) => editorCommandSlice(snapshot, toValue(slot), host.value, t.value),
    sameEditorCommandSlice,
    // The slot, the toolbar's host, and the root's translator change the slice without an
    // editor change.
    () => [toValue(slot), host.value, t.value]
  );
  return {
    execute: (value?: string): boolean =>
      editor.value ? runToolbarCommand(editor.value, toValue(slot), value, host.value).ok : false,
    isActive: computed(() => slice.value.active),
    isEnabled: computed(() => slice.value.enabled),
    disabledReason: computed(() => slice.value.disabledReason),
    hint: computed(() => slice.value.hint),
    value: computed(() => slice.value.value),
  };
}
