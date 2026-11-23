/** @jsxImportSource react */
// The light/dark toggle in the top bar. A copy of pptx-editor/examples/vite/src/ThemeToggle.tsx
// at 5b6ead87 (EigenPal), with the shape of the DOCX Editor demos' toggle. Differences: the
// names are the demo's own English text (AGENTS.md, "i18n"), the arrow-key step is inline,
// because the editor core has no `radioGroupStep`, and the start theme and its memory
// (`initialTheme`, `rememberTheme`, from PPTX Editor's `demo-files.ts`) live here, because this
// demo has no shared demo files. `Theme` is exported for the App, and the storage key is
// `xlsx-editor-demo-theme`, so the two demos on one origin keep their own pick.
// The JSX import source pragma on the first line is added, as every React file here has it.

export type Theme = 'light' | 'dark';
const MODES = ['light', 'dark'] as const;
const NAMES: Record<Theme, string> = { light: 'Light', dark: 'Dark' };

const SUN =
  'M12 2v2M12 20v2M4.9 4.9l1.4 1.4M17.7 17.7l1.4 1.4M2 12h2M20 12h2M4.9 19.1l1.4-1.4M17.7 6.3l1.4-1.4';
const MOON = 'M12 3a6 6 0 0 0 9 9 9 9 0 1 1-9-9Z';

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
function step(value: Theme, key: string): Theme | null {
  const by =
    key === 'ArrowRight' || key === 'ArrowDown'
      ? 1
      : key === 'ArrowLeft' || key === 'ArrowUp'
        ? -1
        : 0;
  if (!by) return null;
  return MODES[(MODES.indexOf(value) + by + MODES.length) % MODES.length]!;
}

export function ThemeToggle({ value, onChange }: { value: Theme; onChange: (t: Theme) => void }) {
  return (
    <div
      className="demo-theme-toggle"
      role="radiogroup"
      aria-label="Appearance"
      // One Tab stop: the checked option; the arrow keys move the choice, as a radio group does.
      onKeyDown={(e) => {
        const next = step(value, e.key);
        if (!next) return;
        e.preventDefault();
        onChange(next);
        e.currentTarget.querySelector<HTMLElement>(`[data-value="${next}"]`)?.focus();
      }}
    >
      {MODES.map((mode) => (
        <button
          key={mode}
          type="button"
          role="radio"
          className="demo-theme-toggle__option"
          data-value={mode}
          aria-checked={value === mode}
          tabIndex={value === mode ? 0 : -1}
          aria-label={NAMES[mode]}
          title={NAMES[mode]}
          onMouseDown={(e) => e.preventDefault()}
          onClick={() => onChange(mode)}
        >
          <svg viewBox="0 0 24 24" aria-hidden="true">
            {mode === 'light' ? <circle cx="12" cy="12" r="4" /> : null}
            <path d={mode === 'light' ? SUN : MOON} />
          </svg>
        </button>
      ))}
    </div>
  );
}
