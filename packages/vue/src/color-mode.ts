// The chrome color mode that an `XlsxEditorRoot` declares (`colorMode`): the Vue twin of the
// React adapter's `color-mode.ts`. Copied from pptx-editor/packages/vue/src/color-mode.ts at
// d4502552 (EigenPal). Differences: the scope is `.xlsx-editor`, the core import is
// `@xlsx-editor.dev/core/editor`, the key is `xlsxChromeDark`, and the ancestor scope mark is
// left out, because the one dialog of this adapter reads the DOM for its scope, as React's does.
// The two internal functions say `Internal.` in prose, as the internal helpers of `context.ts`
// do, and this header replaces the source's first paragraph.
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
