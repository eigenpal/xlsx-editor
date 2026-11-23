// The color theme of the chrome, resolved without a framework, after
// pptx-editor/packages/core/src/editor/adapter-support.ts at 1943d2c9 (EigenPal), which keeps the
// resolution in core so the React and Vue adapters cannot drift. Each adapter keeps only its
// glue: it reads the system setting through these functions and puts `.dark` beside
// `.xlsx-editor`. The workbook canvas is never themed.

/** The color theme of the chrome: `'system'` follows the OS setting. DOCX Editor's type. @public */
export type ColorMode = 'light' | 'dark' | 'system';

/** The theme a root resolved from `colorMode`, never `'system'`. DOCX Editor's type. @public */
export type ResolvedColorMode = 'light' | 'dark';

/** Whether the chrome is dark for `mode`, given the system setting. @internal */
export function resolveDarkChrome(mode: ColorMode | undefined, systemDark: boolean): boolean {
  return mode === 'dark' || (mode === 'system' && systemDark);
}

const DARK_QUERY = '(prefers-color-scheme: dark)';

/** Whether the operating system prefers a dark scheme now. False without `matchMedia`. @internal */
export function systemPrefersDark(): boolean {
  return typeof matchMedia === 'function' && matchMedia(DARK_QUERY).matches;
}

/** Be told when the system setting changes. Returns the unsubscribe function. @internal */
export function subscribeSystemDark(listener: () => void): () => void {
  if (typeof matchMedia !== 'function') return () => undefined;
  const query = matchMedia(DARK_QUERY);
  query.addEventListener('change', listener);
  return () => query.removeEventListener('change', listener);
}
