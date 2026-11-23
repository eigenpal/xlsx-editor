<script setup lang="ts">
// A shell icon: the Vue twin of `Icon` in examples/vite/src/icons.tsx, with the same path data
// (examples/vite/src/icon-paths.ts). Decorative, so the control around it carries the name.
import { PATHS, type IconName } from '../../vite/src/icon-paths.ts';

defineProps<{ name: IconName }>();
</script>

<template>
  <svg viewBox="0 -960 960 960" aria-hidden="true" focusable="false">
    <path :d="PATHS[name]" />
  </svg>
</template>
