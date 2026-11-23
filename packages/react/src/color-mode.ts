// The React glue of the chrome color theme. Core resolves it (`resolveDarkChrome`,
// `systemPrefersDark`, `subscribeSystemDark` in `@xlsx-editor.dev/core/editor`), as PPTX
// Editor's core does; this file only binds it to React. The root resolves `colorMode` once;
// `XlsxEditor.Viewport` puts `.dark` beside `.xlsx-editor`, so the dark token block of the core
// stylesheet applies to every part inside it. `useEditorColorMode` tells a part the host builds
// which theme the editor resolved, as DOCX Editor's hook of that name does.

import { createContext, useContext, useSyncExternalStore } from 'react';
import {
  resolveDarkChrome,
  subscribeSystemDark,
  systemPrefersDark,
  type ColorMode,
  type ResolvedColorMode,
} from '@xlsx-editor.dev/core/editor';

const light = () => false;
const nothing = () => () => {};

/** Whether `mode` resolves to a dark chrome now; `'system'` follows the OS setting. */
export function useResolvedDark(mode: ColorMode | undefined): boolean {
  const following = mode === 'system';
  const systemDark = useSyncExternalStore(
    following ? subscribeSystemDark : nothing,
    following ? systemPrefersDark : light,
    light
  );
  return resolveDarkChrome(mode, systemDark);
}

/** Whether the enclosing root resolved a dark chrome. Internal. */
export const ChromeDarkContext = createContext(false);

/** Whether the enclosing root resolved a dark chrome. Internal. */
export function useChromeDark(): boolean {
  return useContext(ChromeDarkContext);
}

/**
 * What {@link useEditorColorMode} returns.
 *
 * @public
 */
export interface UseEditorColorModeReturn {
  /**
   * The resolved theme: `'dark'` or `'light'`. Re-renders when `colorMode` or the OS setting
   * changes.
   */
  readonly colorMode: ResolvedColorMode;
}

// One object per theme, so the answer keeps its identity until the theme changes.
const LIGHT: UseEditorColorModeReturn = Object.freeze({ colorMode: 'light' });
const DARK: UseEditorColorModeReturn = Object.freeze({ colorMode: 'dark' });

/**
 * The theme the enclosing `XlsxEditor.Root` resolved from its `colorMode`, so parts you build
 * yourself can match the editor chrome, as DOCX Editor's hook of this name does. `'system'`
 * resolves to the operating system's current setting and re-renders when it changes. Outside a
 * `Root`, or without `colorMode`, the answer is `'light'`.
 *
 * @example
 * ```tsx
 * function MyPanel() {
 *   const { colorMode } = useEditorColorMode();
 *   return <aside className={colorMode === 'dark' ? 'panel panel--dark' : 'panel'} />;
 * }
 * ```
 *
 * @public
 */
export function useEditorColorMode(): UseEditorColorModeReturn {
  return useContext(ChromeDarkContext) ? DARK : LIGHT;
}
