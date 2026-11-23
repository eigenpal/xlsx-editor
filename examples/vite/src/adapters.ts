// The two demos of the React | Vue switcher, and the address of each. The hosted build (`bun run
// build:demos`) serves the React demo at `/react/` and the Vue demo at `/vue/`; in development
// the two demos run on their own ports, 5180 (`bun run demo`) and 5181 (`bun run dev:vue`).
// `AdapterSwitcher.tsx` and `examples/vue/src/AdapterSwitcher.vue` render the switcher from
// this module.
//
// After pptx-editor/examples/vite/src/AdapterSwitcher.tsx at 412d750f and `adapterHref` in
// pptx-editor/examples/shared/demo-files.ts at 39dc46ed (EigenPal). Differences:
// - The ports are 5180 and 5181. The list and the show flag are in this module, so that the
//   React and Vue switchers share them. This demo has no shared demo files, so `adapterHref`
//   is here.
// - `adapterHref` sets `theme` to the theme on screen. The toggle keeps a pick only in the
//   storage of its origin, and the two development servers are two origins, so the query
//   carries the pick.
// - A switcher passes each address through `sanitizeHref` (AGENTS.md, "Security") and
//   `linkHref`, as every bound `href` in this repository does.
// - It imports `sanitizeHref` from the ooxml-core source, because the `core` package does not
//   export it.

import type { HrefProjection } from '../../../packages/ooxml-core/src/sinks.ts';
import type { Theme } from './theme.ts';

export { sanitizeHref } from '../../../packages/ooxml-core/src/sinks.ts';

const DEV = import.meta.env.DEV;

/** Shown in development and in the hosted build, which sets `VITE_ENABLE_FRAMEWORK_SWITCHER`. */
export const SHOW_ADAPTER_SWITCHER: boolean =
  !!DEV || import.meta.env.VITE_ENABLE_FRAMEWORK_SWITCHER === 'true';

/** The development and hosted addresses of each demo. */
export const ADAPTER_BASES = {
  react: { dev: 'http://localhost:5180/', build: '/react/' },
  vue: { dev: 'http://localhost:5181/', build: '/vue/' },
} as const;

/** The demo of each adapter, in the order of the switcher. */
export const ADAPTERS = [
  { id: 'react', label: 'React', href: ADAPTER_BASES.react[DEV ? 'dev' : 'build'] },
  { id: 'vue', label: 'Vue', href: ADAPTER_BASES.vue[DEV ? 'dev' : 'build'] },
] as const;

/** The id of an adapter demo. */
export type AdapterId = (typeof ADAPTERS)[number]['id'];

/**
 * The other demo's address with this page's query, and with `theme` set to the theme on screen,
 * so the switch keeps the query and the theme.
 */
export function adapterHref(base: string, search: string, theme: Theme): string {
  const query = new URLSearchParams(search);
  query.set('theme', theme);
  return `${base}?${query.toString()}`;
}

/** The address of a sanitized projection, or `fallback` when the projection is inert. */
export const linkHref = (projection: HrefProjection, fallback: string): string =>
  projection.ok ? projection.href : fallback;
