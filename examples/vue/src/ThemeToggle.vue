<script setup lang="ts">
// The light/dark toggle in the top bar: the Vue twin of examples/vite/src/ThemeToggle.tsx. Both
// take the names, the icons, and the arrow-key step from examples/vite/src/theme.ts, and the
// `demo-theme-toggle` styles from examples/vite/src/demo.css.
//
// After pptx-editor/examples/vue/src/ThemeToggle.vue at 17425725 (EigenPal). Differences: the
// names are the demo's own English text (AGENTS.md, "i18n"), so there is no `useTranslation`,
// and the arrow-key step is `step` in `theme.ts`, because the editor core has no
// `radioGroupStep`.
import { MODES, MOON, NAMES, step, SUN, type Theme } from '../../vite/src/theme.ts';

const value = defineModel<Theme>({ required: true });

// One Tab stop: the checked option; the arrow keys move the choice, as a radio group does.
function onKeydown(e: KeyboardEvent): void {
  const next = step(value.value, e.key);
  if (!next) return;
  e.preventDefault();
  value.value = next;
  (e.currentTarget as HTMLElement).querySelector<HTMLElement>(`[data-value="${next}"]`)?.focus();
}
</script>

<template>
  <div class="demo-theme-toggle" role="radiogroup" aria-label="Appearance" @keydown="onKeydown">
    <button
      v-for="mode in MODES"
      :key="mode"
      type="button"
      role="radio"
      class="demo-theme-toggle__option"
      :data-value="mode"
      :aria-checked="value === mode"
      :tabindex="value === mode ? 0 : -1"
      :aria-label="NAMES[mode]"
      :title="NAMES[mode]"
      @mousedown.prevent
      @click="value = mode"
    >
      <svg viewBox="0 0 24 24" aria-hidden="true">
        <circle v-if="mode === 'light'" cx="12" cy="12" r="4" />
        <path :d="mode === 'light' ? SUN : MOON" />
      </svg>
    </button>
  </div>
</template>
