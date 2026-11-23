/** @jsxImportSource react */
// The React | Vue switcher of the top bar. The Vue twin is examples/vue/src/AdapterSwitcher.vue;
// both take the demos and their addresses from `adapters.ts`, and the styles from `demo.css`.
//
// A navigation, not tabs: the other demo is a link that keeps the page's query and sets `theme`
// to the theme on screen, and the current demo is marked `aria-current="page"` and is not a
// link, so the switcher is one Tab stop.
//
// A copy of pptx-editor/examples/vite/src/AdapterSwitcher.tsx at 412d750f (EigenPal).
// Differences, each on purpose:
// - The name of the navigation is the literal "Framework": the demo's own text is English and
//   stays out of `packages/i18n` (AGENTS.md, "i18n"), so there is no `useTranslation`.
// - The list of demos, the show flag, and `adapterHref` are in `adapters.ts`, shared with the
//   Vue twin; the development ports are 5180 and 5181.
// - It takes the theme on screen (`theme`), which the link sets in its query, and each address
//   goes through `sanitizeHref` (AGENTS.md, "Security"). The demo never changes its own query,
//   so the link has no click handler that reads the query again: a new theme renders a new
//   address.
// - The JSX import source pragma on the first line is added, as every React file here has it.

import { ADAPTERS, adapterHref, linkHref, sanitizeHref, type AdapterId } from './adapters.ts';
import type { Theme } from './theme.ts';

export { SHOW_ADAPTER_SWITCHER } from './adapters.ts';

/** The switcher: `current` is this demo, and `theme` is the theme on screen. */
export function AdapterSwitcher({ current, theme }: { current: AdapterId; theme: Theme }) {
  return (
    <nav className="demo-adapter-switcher" aria-label="Framework">
      {ADAPTERS.map((a) =>
        a.id === current ? (
          <span key={a.id} className="demo-adapter-switcher__option" aria-current="page">
            {a.label}
          </span>
        ) : (
          <a
            key={a.id}
            href={linkHref(sanitizeHref(adapterHref(a.href, location.search, theme)), a.href)}
            className="demo-adapter-switcher__option"
          >
            {a.label}
          </a>
        )
      )}
    </nav>
  );
}
