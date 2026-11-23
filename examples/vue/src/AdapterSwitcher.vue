<script setup lang="ts">
// The React | Vue switcher: the Vue twin of examples/vite/src/AdapterSwitcher.tsx. Both take the
// demos and their addresses from examples/vite/src/adapters.ts, and the styles from
// examples/vite/src/demo.css. A navigation: the other demo is a link that keeps the page's
// query and sets `theme` to the theme on screen, and the current demo is marked
// `aria-current="page"`.
//
// A copy of pptx-editor/examples/vue/src/AdapterSwitcher.vue at 412d750f (EigenPal).
// Differences: the name of the navigation is the literal "Framework", because the demo's text
// is English and stays out of `packages/i18n`, and the list of demos is in `adapters.ts`, shared
// with the React switcher; the development ports are 5180 and 5181. It takes the theme on screen
// (`theme`), and each address goes through `sanitizeHref` (AGENTS.md, "Security"); a new theme
// renders a new address, so the link has no click handler.
import {
  ADAPTERS,
  adapterHref,
  linkHref,
  sanitizeHref,
  type AdapterId,
} from '../../vite/src/adapters.ts';
import type { Theme } from '../../vite/src/theme.ts';

defineProps<{ current: AdapterId; theme: Theme }>();

// The demo never changes its own query, so the query of the first render stays right.
const search = location.search;
</script>

<template>
  <nav class="demo-adapter-switcher" aria-label="Framework">
    <template v-for="a in ADAPTERS" :key="a.id">
      <span v-if="current === a.id" class="demo-adapter-switcher__option" aria-current="page">{{
        a.label
      }}</span>
      <a
        v-else
        :href="linkHref(sanitizeHref(adapterHref(a.href, search, theme)), a.href)"
        class="demo-adapter-switcher__option"
        >{{ a.label }}</a
      >
    </template>
  </nav>
</template>
