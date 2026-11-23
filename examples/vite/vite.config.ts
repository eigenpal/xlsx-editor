// The demo: a Vite app over the workspace sources. `bun run demo` at the repository root starts
// it. The HarfBuzz runtime and the font package find their files with
// `new URL(file, import.meta.url)`, which dependency pre-bundling would break, so they are
// excluded from it. The server may read the repository root, for the committed fixtures.
//
// The packages export their built `dist/`. The aliases point the demo at the sources, so
// `bun install && bun run demo` works from a clean clone with no package build, as in
// pptx-editor/examples/vite/vite.config.ts at e22f3f3b.

import { fileURLToPath } from 'node:url';
import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

const repo = fileURLToPath(new URL('../..', import.meta.url));
const pkg = (path: string): string =>
  fileURLToPath(new URL(`../../packages/${path}`, import.meta.url));

export default defineConfig({
  root: fileURLToPath(new URL('.', import.meta.url)),
  plugins: [react()],
  // The hosted build serves each demo under its own path (`bun run build:demos`).
  base: process.env.VITE_BASE_PATH ?? '/',
  resolve: {
    alias: [
      {
        find: /^@xlsx-editor\.dev\/core\/styles\/editor\.css$/,
        replacement: pkg('core/src/styles/editor.css'),
      },
      { find: /^@xlsx-editor\.dev\/core\/(\w+)$/, replacement: pkg('core/src/$1/index.ts') },
      { find: /^@xlsx-editor\.dev\/core$/, replacement: pkg('core/src/index.ts') },
      {
        find: /^@xlsx-editor\.dev\/react\/styles\.css$/,
        replacement: pkg('react/src/styles/editor.css'),
      },
      { find: /^@xlsx-editor\.dev\/react$/, replacement: pkg('react/src/index.ts') },
      { find: /^@xlsx-editor\.dev\/i18n\/([\w-]+)$/, replacement: pkg('i18n/src/$1.ts') },
      { find: /^@xlsx-editor\.dev\/i18n$/, replacement: pkg('i18n/src/index.ts') },
    ],
  },
  optimizeDeps: { exclude: ['harfbuzzjs', '@docx-editor.dev/fonts'] },
  server: { port: 5180, fs: { allow: [repo] } },
  preview: { port: 5180 },
});
