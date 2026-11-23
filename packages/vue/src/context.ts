// Injection keys and composables: the Vue twin of the React adapter's `context.ts`, after PPTX
// Editor's Vue `composables.ts` (pptx-editor/packages/vue/src/composables.ts at d4502552), which
// follows DOCX Editor's Vue adapter. The root holds ONE editor subscription and publishes a
// change tick, and every composable reads the tick, so a toolbar of many controls costs one
// subscription, not one each. Every composable annotates a named `Use<Name>Return` type, so
// core's internal types do not leak into the API snapshot (AGENTS.md "Public API").
//
// Differences from React: a composable returns refs, not plain values, as PPTX Editor's do.
// `useEditorState` reads the snapshot again on the root's tick, and keeps the slice's identity
// while `isEqual` holds, as React's memo keeps it. The tick watch flushes synchronously (PPTX
// Editor's runs before the render): a host that reads a composable right after an editor call,
// as a React host reads `useEditorState` after `act`, sees the new state with no `nextTick`.

import {
  getCurrentScope,
  inject,
  onMounted,
  onScopeDispose,
  provide,
  shallowRef,
  watch,
  type InjectionKey,
  type ShallowRef,
} from 'vue';
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

/** The editor published by `XlsxEditorRoot`, the twin of React's `XlsxEditorContext`. @public */
export const xlsxEditorKey: InjectionKey<ShallowRef<XlsxEditorInstance | null>> =
  Symbol('xlsxEditor');

/** The translator published by `XlsxEditorRoot`, the twin of React's `TranslationContext`. @public */
export const xlsxTranslationKey: InjectionKey<ShallowRef<TFunction>> = Symbol('xlsxTranslation');

/** A counter that the root bumps on every editor change: one subscription for all composables. */
const xlsxEditorTickKey: InjectionKey<ShallowRef<number>> = Symbol('xlsxEditorTick');

/** The `mode` prop of the root, for the snapshot that composables read before mount. Internal. */
export const rootModeKey: InjectionKey<() => EditorMode> = Symbol('xlsxRootMode');

/** Whether a change of the `mode` prop waits for the open cell edit. Internal. */
export const modeWaitsKey: InjectionKey<ShallowRef<boolean>> = Symbol('xlsxModeWaits');

/** A translator over a locale catalog: missing or `null` entries fall back to English. @public */
export function translatorFor(locale: Translations | undefined): TFunction {
  const merged = deepMerge(
    en as Record<string, unknown>,
    locale as Record<string, unknown> | undefined
  );
  return createT(merged as LocaleStrings, typeof locale?._lang === 'string' ? locale._lang : 'en');
}

const nullEditor = shallowRef<XlsxEditorInstance | null>(null);
const english = shallowRef<TFunction>(translatorFor(undefined));
const noWait = shallowRef(false);
const editMode = (): EditorMode => 'edit';

/**
 * Publish `editor` and one shared change tick to the components below. `XlsxEditorRoot` calls
 * it in its `setup`. Internal.
 */
export function provideXlsxEditor(editor: ShallowRef<XlsxEditorInstance | null>): void {
  provide(xlsxEditorKey, editor);
  const tick = shallowRef(0);
  provide(xlsxEditorTickKey, tick);
  let off: (() => void) | null = null;
  watch(
    editor,
    (next) => {
      off?.();
      off = next ? next.subscribe(() => (tick.value += 1)) : null;
      tick.value += 1;
    },
    { immediate: true }
  );
  onScopeDispose(() => {
    off?.();
    off = null;
  });
}

/** What `useXlsxEditor` returns: the editor, or null before it mounts. @public */
export type UseXlsxEditorReturn = Readonly<ShallowRef<XlsxEditorInstance | null>>;

/** The editor, or null before it mounts. @public */
export function useXlsxEditor(): UseXlsxEditorReturn {
  return inject(xlsxEditorKey, nullEditor);
}

/** What `useTranslation` returns. @public */
export interface UseTranslationReturn {
  /** Translate a key. A render that calls it renders again when the root's `locale` changes. */
  readonly t: TFunction;
}

/** The translator ref of the enclosing root, for a watcher that follows a locale change. Internal. */
export function useTranslator(): Readonly<ShallowRef<TFunction>> {
  return inject(xlsxTranslationKey, english);
}

/** The translator of the enclosing root. @public */
export function useTranslation(): UseTranslationReturn {
  const base = inject(xlsxTranslationKey, english);
  return { t: ((key, params) => base.value(key, params)) as TFunction };
}

/**
 * Whether the parent listens to `event` on a component, with `@event`, or with `@event.once`
 * when `once` is true. Vue passes `.once` as `on<Event>Once`. Internal.
 */
export function hasListener(
  props: Readonly<Record<string, unknown>> | null,
  event: string,
  once = true
): boolean {
  const name = `on${event[0]!.toUpperCase()}${event.slice(1)}`;
  return Boolean(props?.[name] ?? (once ? props?.[`${name}Once`] : undefined));
}

/** What `useViewModeWaits` returns. @public */
export type UseViewModeWaitsReturn = Readonly<ShallowRef<boolean>>;

/**
 * Whether a change of the root's `mode` prop to view mode waits for the open cell edit, as
 * `XlsxEditorStatusMessage` shows it. For a layout that shows the wait in its own chrome.
 *
 * @public
 */
export function useViewModeWaits(): UseViewModeWaitsReturn {
  return inject(modeWaitsKey, noWait);
}

/**
 * How many grids (`XlsxEditorContent`) and `StatusMessage` elements the enclosing root holds, so
 * the root can warn in development when a layout shows the grid without the only place that
 * shows refusals. Internal.
 */
export interface LayoutParts {
  grids: number;
  statusMessages: number;
}

/** The layout parts of the enclosing root. Internal. */
export const layoutPartsKey: InjectionKey<LayoutParts> = Symbol('xlsxLayoutParts');

/** Count one layout part while it is mounted. Internal. */
export function useLayoutPart(part: keyof LayoutParts): void {
  const parts = inject(layoutPartsKey, null);
  if (!parts) return;
  // A part that unmounts before it mounts, as an async setup can, was never counted.
  let counted = false;
  onMounted(() => {
    parts[part] += 1;
    counted = true;
  });
  onScopeDispose(() => {
    if (counted) parts[part] -= 1;
  });
}

/** What `useEditorState` returns: the selected slice, updated when it changes. @public */
export type UseEditorStateReturn<T> = Readonly<ShallowRef<T>>;

/**
 * A slice of the editor's state, the twin of React's `useEditorState`. The ref changes only
 * when `selector`'s result changes by `isEqual` (default `Object.is`). Before the editor mounts,
 * the selector reads an empty snapshot with the root's `mode`, never `null`. `deps` names a
 * source whose change reads the slice again, such as a slot that the selector reads.
 *
 * @public
 */
export function useEditorState<T>(
  selector: (snapshot: EditorSnapshot) => T,
  isEqual: (a: T, b: T) => boolean = Object.is,
  deps?: () => unknown
): UseEditorStateReturn<T> {
  const editorRef = useXlsxEditor();
  const tick = inject(xlsxEditorTickKey, null);
  const mode = inject(rootModeKey, editMode);
  // The snapshot of the last read: a tick with the same snapshot skips the selector, as React's
  // memo does, so a selector that builds a new value each call changes the ref only with the
  // state.
  let seen: EditorSnapshot | null = null;
  const current = (): EditorSnapshot => editorRef.value?.snapshot() ?? emptySnapshotOf(mode());
  const read = (): T => {
    seen = current();
    return selector(seen);
  };
  const out = shallowRef(read()) as ShallowRef<T>;
  const update = (): void => {
    const next = read();
    if (!isEqual(out.value, next)) out.value = next;
  };
  const onChange = (): void => {
    if (current() !== seen) update();
  };
  if (!getCurrentScope()) return out;
  if (deps) watch(deps, update);
  // The root's `mode` prop changes the empty snapshot before mount.
  watch(mode, onChange);
  if (tick) {
    watch(editorRef, update);
    watch(tick, onChange, { flush: 'sync' });
    return out;
  }
  // No root above (a host that provides only the editor): subscribe directly.
  let off: (() => void) | null = null;
  watch(
    editorRef,
    (editor) => {
      off?.();
      off = editor ? editor.subscribe(onChange) : null;
      update();
    },
    { immediate: true }
  );
  onScopeDispose(() => {
    off?.();
    off = null;
  });
  return out;
}

/** What `useEditorSnapshot` returns. @public */
export type UseEditorSnapshotReturn = Readonly<ShallowRef<EditorSnapshot>>;

const whole = (snapshot: EditorSnapshot): EditorSnapshot => snapshot;

/**
 * The editor's whole state, read again after each change. Before it mounts, an empty snapshot
 * with the root's `mode`.
 *
 * @public
 */
export function useEditorSnapshot(): UseEditorSnapshotReturn {
  return useEditorState(whole);
}
