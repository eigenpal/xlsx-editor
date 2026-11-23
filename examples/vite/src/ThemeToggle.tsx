/** @jsxImportSource react */
// The light/dark toggle in the top bar. A copy of pptx-editor/examples/vite/src/ThemeToggle.tsx
// at 5b6ead87 (EigenPal), with the shape of the DOCX Editor demos' toggle. Differences: the
// names are the demo's own English text (AGENTS.md, "i18n"), the arrow-key step is inline,
// because the editor core has no `radioGroupStep`, and the start theme and its memory
// (`initialTheme`, `rememberTheme`, from PPTX Editor's `demo-files.ts`) live here, because this
// demo has no shared demo files. `Theme` is exported for the App, and the storage key is
// `xlsx-editor-demo-theme`, so the demos of each editor keep their own pick. The logic is in
// `theme.ts`, which the Vue twin (examples/vue/src/ThemeToggle.vue) shares, so the React and
// Vue demos of one origin keep one pick.
// The JSX import source pragma on the first line is added, as every React file here has it.

import { MODES, MOON, NAMES, step, SUN, type Theme } from './theme.ts';

export { initialTheme, rememberTheme, type Theme } from './theme.ts';

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
