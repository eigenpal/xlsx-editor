// The Vue twin of `examples/vite`: `bun run dev:vue` at the repository root starts it.
// It builds against the workspace sources, so `bun install && bun run dev:vue` works from a clean
// clone with no package build. It serves the public files of the React demo (`publicDir`), so
// both demos open the same sample workbook. The HarfBuzz runtime and the font package find their
// files with `new URL(file, import.meta.url)`, which dependency pre-bundling would break, so they
// are excluded from it.
//
// After pptx-editor/examples/vue/vite.config.ts at d4502552 (EigenPal). Differences: the scope is
// `@xlsx-editor.dev`, with no Pro or PDF export, and port 5181 is next to the React demo's 5180.

import { fileURLToPath } from 'node:url';
import { defineConfig } from 'vite';
import vue from '@vitejs/plugin-vue';

const pkg = (path: string): string =>
  fileURLToPath(new URL(`../../packages/${path}`, import.meta.url));

export default defineConfig({
  root: fileURLToPath(new URL('.', import.meta.url)),
  plugins: [vue()],
  // The hosted build serves each demo under its own path (`bun run build:demos`).
  base: process.env.VITE_BASE_PATH ?? '/',
  publicDir: fileURLToPath(new URL('../vite/public', import.meta.url)),
  resolve: {
    alias: [
      {
        find: /^@xlsx-editor\.dev\/core\/styles\/editor\.css$/,
        replacement: pkg('core/src/styles/editor.css'),
      },
      { find: /^@xlsx-editor\.dev\/core\/(\w+)$/, replacement: pkg('core/src/$1/index.ts') },
      { find: /^@xlsx-editor\.dev\/core$/, replacement: pkg('core/src/index.ts') },
      {
        find: /^@xlsx-editor\.dev\/vue\/styles\.css$/,
        replacement: pkg('vue/src/styles/editor.css'),
      },
      { find: /^@xlsx-editor\.dev\/vue$/, replacement: pkg('vue/src/index.ts') },
      { find: /^@xlsx-editor\.dev\/i18n\/([\w-]+)$/, replacement: pkg('i18n/src/$1.ts') },
      { find: /^@xlsx-editor\.dev\/i18n$/, replacement: pkg('i18n/src/index.ts') },
    ],
  },
  optimizeDeps: { exclude: ['harfbuzzjs', '@docx-editor.dev/fonts'] },
  // The adapter's complex parts are Vue TSX (`/** @jsxImportSource vue */`); its tsconfig keeps
  // `jsx: preserve` for the type checker, so the demo compiles their JSX with Vue's runtime.
  oxc: { jsx: { runtime: 'automatic', importSource: 'vue' } },
  server: { port: 5181, strictPort: true, fs: { allow: ['../..'] } },
  preview: { port: 5181 },
});
