// Typed editor events as a composable: the Vue twin of the React adapter's
// `customization-hooks.ts`. Adapted from DOCX Editor packages/vue/src/editor/useEditorEvent.ts
// at e9a35d0c4: the composable subscribes with the editor's `on` method, and subscribes again
// when the editor instance changes.
//
// Intended differences from DOCX Editor:
// - The file is named after PPTX Editor's (packages/vue/src/customization-hooks.ts at
//   4d976427), as the React twin is. It holds only `useEditorEvent`.
// - The event map is this editor's `EditorEvents`.
// - It returns the named type `UseEditorEventReturn`, as every composable of this adapter does.
// - It watches with the default flush, not `post`, so a host that changes the editor sees its
//   events from the change on; and it does nothing outside an effect scope, where no cleanup
//   could end the subscription.
// - `setup` runs once, so the handler that it receives is the one called. DOCX Editor's handler
//   object holds the same one handler.

import { getCurrentScope, watch } from 'vue';
import type { EditorEvents } from '@xlsx-editor.dev/core/editor';
import { useXlsxEditor } from './context.ts';

/** What {@link useEditorEvent} returns: nothing; the subscription ends with the scope. @public */
export type UseEditorEventReturn = void;

/**
 * Subscribe to an editor event (`'change'`, `'selectionChange'`, `'rejection'`,
 * `'openFailure'`) for the lifetime of the component, the twin of React's `useEditorEvent`.
 * Call it in `setup`. It does nothing until `XlsxEditorRoot` creates the editor, and subscribes
 * again when the instance changes.
 *
 * `'change'` fires only when the workbook content changes. The root's `change` emit fires after
 * every state change, also a selection or a zoom change.
 *
 * @example
 * ```ts
 * const { t } = useTranslation();
 * useEditorEvent('rejection', (refusal, snapshot) =>
 *   toast(refusalText(refusal, t, snapshot.numberLocale))
 * );
 * ```
 *
 * @public
 */
export function useEditorEvent<E extends keyof EditorEvents>(
  event: E,
  handler: EditorEvents[E]
): UseEditorEventReturn {
  const editor = useXlsxEditor();
  if (!getCurrentScope()) return;
  watch(
    editor,
    (instance, _prev, onCleanup) => {
      if (instance) onCleanup(instance.on(event, handler));
    },
    { immediate: true }
  );
}
