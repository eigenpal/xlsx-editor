// The light and dark theme of the demo shell: the start theme, its memory, the names, the icon
// paths, and the arrow-key step of the toggle. The React toggle (`ThemeToggle.tsx`) and its Vue
// twin (examples/vue/src/ThemeToggle.vue) both use this module, so the two demos keep one rule.
// It holds the logic of pptx-editor/examples/vite/src/ThemeToggle.tsx at 5b6ead87 (EigenPal),
// with the differences that `ThemeToggle.tsx` lists.

export type Theme = 'light' | 'dark';
/** The themes, in the order of the toggle. */
export const MODES = ['light', 'dark'] as const;
/** The name of each theme, for the toggle. */
export const NAMES: Record<Theme, string> = { light: 'Light', dark: 'Dark' };

/** The rays of the sun icon (Light). */
export const SUN =
  'M12 2v2M12 20v2M4.9 4.9l1.4 1.4M17.7 17.7l1.4 1.4M2 12h2M20 12h2M4.9 19.1l1.4-1.4M17.7 6.3l1.4-1.4';
/** The moon icon (Dark). */
export const MOON = 'M12 3a6 6 0 0 0 9 9 9 9 0 1 1-9-9Z';

/** Where the demo keeps the theme that the user picked. */
const THEME_KEY = 'xlsx-editor-demo-theme';

const isTheme = (v: unknown): v is Theme => v === 'light' || v === 'dark';

/** The theme picked before, or null. Storage can be off (a private window): then null. */
function storedTheme(): Theme | null {
  try {
    const v = globalThis.localStorage?.getItem(THEME_KEY);
    return isTheme(v) ? v : null;
  } catch {
    return null;
  }
}

/**
 * The theme the demo starts in: the URL's (`?theme=dark` or `?theme=light`), else the theme
 * picked before, else light. Dark is only ever picked by hand. A copy of `initialTheme` in
 * pptx-editor/examples/shared/demo-files.ts at 5b6ead87.
 */
export function initialTheme(search: string): Theme {
  const asked = new URLSearchParams(search).get('theme');
  return isTheme(asked) ? asked : (storedTheme() ?? 'light');
}

/** Keeps the picked theme for the next visit. A copy of `rememberTheme` in the same file. */
export function rememberTheme(theme: Theme): void {
  try {
    globalThis.localStorage?.setItem(THEME_KEY, theme);
  } catch {
    // Storage is off: the next visit starts light again.
  }
}

/** The next theme for an arrow key in the radio group, or null for another key. */
export function step(value: Theme, key: string): Theme | null {
  const by =
    key === 'ArrowRight' || key === 'ArrowDown'
      ? 1
      : key === 'ArrowLeft' || key === 'ArrowUp'
        ? -1
        : 0;
  if (!by) return null;
  return MODES[(MODES.indexOf(value) + by + MODES.length) % MODES.length]!;
}
