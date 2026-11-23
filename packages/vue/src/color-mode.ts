// The chrome color mode that an `XlsxEditorRoot` declares (`colorMode`): the Vue twin of the
// React adapter's `color-mode.ts`. Copied from pptx-editor/packages/vue/src/color-mode.ts at
// d4502552 (EigenPal). Differences: the scope is `.xlsx-editor`, the core import is
// `@xlsx-editor.dev/core/editor`, and the keys are `xlsxChromeDark` and `xlsxScopedByAncestor`.
// The internal functions say `Internal.` in prose, as the internal helpers of `context.ts` do,
// and this header replaces the source's first paragraph. The ancestor mark
// (`provideScopedByAncestor`, `useScopedByAncestor`) is from the same file at 929ca860.
//
// `useScopeClassName` is after docx-editor/packages/vue/src/editor/scope-context.ts at 27720ad52
// (EigenPal), with the same name and values. Differences:
// - The scope class is `xlsx-editor`, this editor's name for `docx-editor`.
// - It returns a computed ref (`UseScopeClassNameReturn`), as
//   pptx-editor/packages/vue/src/customization-hooks.ts at 929ca860 does, not a string read
//   once: a part keeps the scope of the current `colorMode` without a second call, which the
//   React hook does by its re-render.
// - Only `XlsxEditorViewport` provides the ancestor mark, as in PPTX Editor. DOCX Editor also
//   sets it in its packaged host and in parts that carry their own scope. The dialog frame
//   (`dialog-frame.tsx`) also adds `.xlsx-editor` to itself when it is outside a scope, but it
//   renders no host content, so it sets no mark.
// - The composable is in this file, beside the theme it reads.
//
// Core resolves the mode (`resolveDarkChrome`, `systemPrefersDark`, `subscribeSystemDark`); this
// file only binds it to Vue. `XlsxEditorViewport` puts `.dark` beside `.xlsx-editor`.

import {
  computed,
  inject,
  onMounted,
  onScopeDispose,
  provide,
  shallowRef,
  watch,
  type ComputedRef,
  type InjectionKey,
} from 'vue';
import {
  resolveDarkChrome,
  subscribeSystemDark,
  systemPrefersDark,
  type ColorMode,
  type ResolvedColorMode,
} from '@xlsx-editor.dev/core/editor';

const chromeDarkKey: InjectionKey<ComputedRef<boolean>> = Symbol('xlsxChromeDark');

/** True inside `XlsxEditorViewport`: an ancestor element already carries `.xlsx-editor`. */
const scopedByAncestorKey: InjectionKey<boolean> = Symbol('xlsxScopedByAncestor');

/** Mark the components below as inside the viewport's scope. Internal. */
export function provideScopedByAncestor(): void {
  provide(scopedByAncestorKey, true);
}

/** Whether an ancestor element already carries `.xlsx-editor`. Internal. */
export function useScopedByAncestor(): boolean {
  return inject(scopedByAncestorKey, false);
}

const lightChrome = computed(() => false);

/** Resolve and provide the chrome color mode for one root. Internal. */
export function provideChromeDark(mode: () => ColorMode | undefined): void {
  const systemDark = shallowRef(false);
  let off: (() => void) | null = null;
  const follow = (next: ColorMode | undefined): void => {
    off?.();
    off = null;
    if (next !== 'system') return;
    systemDark.value = systemPrefersDark();
    off = subscribeSystemDark(() => (systemDark.value = systemPrefersDark()));
  };
  // The operating system is read after mount, as DOCX Editor's Vue adapter does: a server
  // render and the hydrating client then start from the same light chrome.
  onMounted(() => follow(mode()));
  watch(mode, follow);
  onScopeDispose(() => off?.());
  provide(
    chromeDarkKey,
    computed(() => resolveDarkChrome(mode(), systemDark.value))
  );
}

/** Whether the enclosing root resolved a dark chrome. Internal. */
export function useChromeDark(): ComputedRef<boolean> {
  return inject(chromeDarkKey, lightChrome);
}

/** What `useEditorColorMode` returns. @public */
export interface UseEditorColorModeReturn {
  /** The resolved theme: `'dark'` or `'light'`. Updates when `colorMode` or the OS setting changes. */
  readonly colorMode: ComputedRef<ResolvedColorMode>;
}

/**
 * The theme the enclosing `XlsxEditorRoot` resolved from its `colorMode`, as React's
 * `useEditorColorMode` and DOCX Editor's. Outside a root, or without `colorMode`, it is
 * `'light'`. Call it in `setup`. @public
 */
export function useEditorColorMode(): UseEditorColorModeReturn {
  const dark = useChromeDark();
  return { colorMode: computed(() => (dark.value ? 'dark' : 'light')) };
}

/** What {@link useScopeClassName} returns. @public */
export type UseScopeClassNameReturn = ComputedRef<'' | 'xlsx-editor ' | 'xlsx-editor dark '>;

/**
 * The scope classes a part you build yourself adds to its own root element, so the editor's
 * `--doc-*` tokens and chrome styles apply to it, as React's `useScopeClassName`. Inside
 * `XlsxEditorViewport` it is `''`. Outside it, it is `'xlsx-editor '`, or `'xlsx-editor dark '`
 * when the root's `colorMode` resolved dark, and `'xlsx-editor '` outside a root. It is a ref,
 * so it follows a change of `colorMode` and a `'system'` color mode. The value ends with a
 * space, so it concatenates cleanly. Call it in `setup`. @public
 */
export function useScopeClassName(): UseScopeClassNameReturn {
  const scoped = useScopedByAncestor();
  const dark = useChromeDark();
  return computed(() => (scoped ? '' : dark.value ? 'xlsx-editor dark ' : 'xlsx-editor '));
}
