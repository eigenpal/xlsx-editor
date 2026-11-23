// Provider context and hooks, after PPTX Editor's `context.ts`: the editor instance and the
// translator that `XlsxEditor.Root` publishes, and a snapshot hook over the editor's
// subscription (`useSyncExternalStore`), so chrome re-renders only when the editor changes.

import {
  createContext,
  useCallback,
  useContext,
  useEffect,
  useMemo,
  useSyncExternalStore,
} from 'react';
import {
  createT,
  deepMerge,
  en,
  type LocaleStrings,
  type TFunction,
  type Translations,
} from '@xlsx-editor.dev/i18n';
import {
  emptySnapshotOf,
  type EditorMode,
  type EditorSnapshot,
  type XlsxEditorInstance,
} from '@xlsx-editor.dev/core/editor';

/** The editor published by `XlsxEditor.Root`. @public */
export const XlsxEditorContext = createContext<XlsxEditorInstance | null>(null);

/** A translator over a locale catalog: missing or `null` entries fall back to English. @public */
export function translatorFor(locale: Translations | undefined): TFunction {
  const merged = deepMerge(
    en as Record<string, unknown>,
    locale as Record<string, unknown> | undefined
  );
  return createT(merged as LocaleStrings, typeof locale?._lang === 'string' ? locale._lang : 'en');
}

/** The translator published by `XlsxEditor.Root` (English by default). @public */
export const TranslationContext = createContext<TFunction>(translatorFor(undefined));

/** The editor, or null before it mounts. @public */
export function useXlsxEditor(): XlsxEditorInstance | null {
  return useContext(XlsxEditorContext);
}

/** The translator of the enclosing root. @public */
export function useTranslation(): { readonly t: TFunction } {
  return { t: useContext(TranslationContext) };
}

/**
 * The `mode` prop of the enclosing `XlsxEditor.Root`, for the snapshot that hooks read before
 * the editor mounts. Internal: not exported from the package.
 */
export const RootModeContext = createContext<EditorMode>('edit');

/**
 * Whether a change of the `mode` prop to `view` waits for the open cell edit, which the store
 * refuses: `StatusMessage` shows a notice while it waits. Internal: not exported.
 */
export const ModeWaitsContext = createContext(false);

/**
 * Whether a change of the root's `mode` prop to view mode waits for the open cell edit, as
 * `XlsxEditor.StatusMessage` shows it. For a layout that shows the wait in its own chrome.
 *
 * @public
 */
export function useViewModeWaits(): boolean {
  return useContext(ModeWaitsContext);
}

/**
 * How many grids (`XlsxEditor.Content`) and `XlsxEditor.StatusMessage` elements the enclosing
 * root holds, so the root can warn in development when a layout shows the grid without the only
 * place that shows refusals. Internal.
 */
export interface LayoutParts {
  grids: number;
  statusMessages: number;
}

/** The layout parts of the enclosing root. Internal. */
export const LayoutPartsContext = createContext<LayoutParts | null>(null);

/** Count one layout part while it is mounted. Internal. */
export function useLayoutPart(part: keyof LayoutParts): void {
  const parts = useContext(LayoutPartsContext);
  useEffect(() => {
    if (!parts) return;
    parts[part] += 1;
    return () => {
      parts[part] -= 1;
    };
  }, [parts, part]);
}

/**
 * A slice of the editor's state, after DOCX Editor's `useEditorState`. The component re-renders
 * only when `selector`'s result changes by `isEqual` (default `Object.is`). Before the editor
 * mounts, the selector reads an empty snapshot with the root's `mode`, never `null`.
 *
 * @public
 */
export function useEditorState<T>(
  selector: (snapshot: EditorSnapshot) => T,
  isEqual: (a: T, b: T) => boolean = Object.is
): T {
  const editor = useXlsxEditor();
  const empty = emptySnapshotOf(useContext(RootModeContext));
  // Notifications reach React at once, not deferred and merged as in DOCX Editor's
  // `useEditorState`. Its two reasons do not hold here, as measured in `adapter.test.tsx`
  // ("reads no stale snapshot and survives key repeat"): `publish()` builds the whole new
  // snapshot before it notifies, so a read inside a notification is never stale; and 500
  // synchronous key presses in one batch stay under React's update-depth limit, because each
  // key publishes once and an equal slice skips the render.
  const subscribe = useCallback(
    (onChange: () => void) => (editor ? editor.subscribe(onChange) : () => undefined),
    [editor]
  );
  // The last slice, keyed on the editor, the selector, and `isEqual`, as DOCX Editor keys it: a
  // new selector starts a new memo, so it never returns the slice of the old one. An equal
  // slice keeps its reference, so React skips the render. Pass a stable selector (module level
  // or `useCallback`); an inline one still reads correctly, it only starts a new memo each render.
  const read = useMemo(() => {
    let last: { snapshot: EditorSnapshot; slice: T } | null = null;
    return (): T => {
      const snapshot = editor ? editor.snapshot() : empty;
      if (last && last.snapshot === snapshot) return last.slice;
      const slice = selector(snapshot);
      const kept = last && isEqual(last.slice, slice) ? last.slice : slice;
      last = { snapshot, slice: kept };
      return kept;
    };
  }, [editor, empty, selector, isEqual]);
  return useSyncExternalStore(subscribe, read, read);
}

const whole = (snapshot: EditorSnapshot): EditorSnapshot => snapshot;

/**
 * The editor's whole state, re-read after each change. Before it mounts, an empty snapshot
 * with the root's `mode`.
 *
 * @public
 */
export function useEditorSnapshot(): EditorSnapshot {
  return useEditorState(whole);
}
