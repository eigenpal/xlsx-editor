// Typed editor events as a hook. Adapted from DOCX Editor
// packages/react/src/editor/useEditorEvent.ts at e9a35d0c4: the hook subscribes with the
// editor's `on` method, a new handler on each render does not subscribe again (only a new
// editor or a new event name does), and the latest handler is the one called. The Vue twin is in
// `customization-hooks.ts` of the Vue adapter.
//
// Intended differences from DOCX Editor:
// - The file is named after PPTX Editor's (packages/react/src/customization-hooks.ts at
//   4d976427), the module for the hooks that build host chrome. It holds only `useEditorEvent`.
// - The event map is this editor's `EditorEvents`.
// - The latest handler is kept in an effect, not during render, so a render that React drops
//   does not leave its handler behind.

import { useEffect, useRef } from 'react';
import type { EditorEvents } from '@xlsx-editor.dev/core/editor';
import { useXlsxEditor } from './context.ts';

/**
 * Subscribe to an editor event (`'change'`, `'selectionChange'`, `'rejection'`,
 * `'openFailure'`) for the lifetime of the component. It does nothing until the nearest
 * `XlsxEditor.Root` creates the editor, and subscribes again when the instance changes. A new
 * handler on each render does not subscribe again; the latest one is called.
 *
 * `'change'` fires only when the workbook content changes. The root's `onChange` prop fires
 * after every state change, also a selection or a zoom change.
 *
 * @example
 * ```tsx
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
): void {
  const editor = useXlsxEditor();
  const handlerRef = useRef(handler);
  useEffect(() => {
    handlerRef.current = handler;
  });
  useEffect(() => {
    if (!editor) return undefined;
    const forward = ((...args: unknown[]) =>
      (handlerRef.current as (...a: unknown[]) => void)(...args)) as EditorEvents[E];
    return editor.on(event, forward);
  }, [editor, event]);
}
