// The React glue of the chrome color theme. Core resolves it (`resolveDarkChrome`,
// `systemPrefersDark`, `subscribeSystemDark` in `@xlsx-editor.dev/core/editor`), as PPTX
// Editor's core does; this file only binds it to React. The root resolves `colorMode` once;
// `XlsxEditor.Viewport` puts `.dark` beside `.xlsx-editor`, so the dark token block of the core
// stylesheet applies to every part inside it. `useEditorColorMode` tells a part the host builds
// which theme the editor resolved, as DOCX Editor's hook of that name does.
//
// `ScopedByAncestorContext` and `useScopeClassName` are copied from
// docx-editor/packages/react/src/editor/scope-context.ts at 27720ad52 (EigenPal), with the same
// name, signature, and return shape. Differences:
// - The scope class is `xlsx-editor`, this editor's name for `docx-editor`.
// - Only `XlsxEditor.Viewport` provides `ScopedByAncestorContext`, as
//   pptx-editor/packages/react/src/color-mode.ts at 9d7d60b7 does. DOCX Editor also sets the
//   mark in its packaged host and in parts that carry their own scope. Here the viewport carries
//   `.xlsx-editor` in the packaged `XlsxEditor` and in a composed layout alike. The dialog frame
//   (`dialog-frame.tsx`) also adds `.xlsx-editor` to itself when it is outside a scope, but it
//   renders no host content, so it sets no mark.
// - The hook and the context are in this file, beside the theme they read, because this adapter
//   has no `editor/` directory. `ChromeDarkContext` is this file's own.

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

/** True inside `XlsxEditor.Viewport`: an ancestor already carries `.xlsx-editor`. Internal. */
export const ScopedByAncestorContext = createContext(false);

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

/**
 * The scope classes that a part you build yourself adds to its own root element, so the editor's
 * `--doc-*` tokens and chrome styles apply to it, as DOCX Editor's `useScopeClassName`. Inside
 * `XlsxEditor.Viewport` it is `''`: the viewport already carries the scope. Outside it, it is
 * `'xlsx-editor '`, or `'xlsx-editor dark '` when the root's `colorMode` resolved dark, and it
 * re-renders when that changes. Outside a root, it is `'xlsx-editor '`. A part inside the
 * viewport that portals out of it (to `document.body`) adds `xlsx-editor` itself, with
 * `useEditorColorMode` for the theme.
 *
 * The value ends with a space, so it concatenates cleanly, and `''` collapses away.
 *
 * @example
 * ```tsx
 * const scope = useScopeClassName();
 * <aside className={`${scope}my-panel`} />
 * ```
 *
 * @public
 */
export function useScopeClassName(): '' | 'xlsx-editor ' | 'xlsx-editor dark ' {
  const scoped = useContext(ScopedByAncestorContext);
  const dark = useContext(ChromeDarkContext);
  if (scoped) return '';
  return dark ? 'xlsx-editor dark ' : 'xlsx-editor ';
}
