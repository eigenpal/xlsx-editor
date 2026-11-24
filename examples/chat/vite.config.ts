// The chat example: a chat app with the workbook editor in a panel beside the conversation.
// `bun run dev:chat` at the repository root starts it on port 5182. It builds against the
// workspace sources, as the demo does (`examples/vite/vite.config.ts`), and it reuses the demo's
// sample workbook and icons from `examples/vite/public`. Adapted from
// pptx-editor/examples/chat/vite.config.ts at e0ac8351 (EigenPal).
//
// Differences from PPTX Editor:
// - The aliases are the ones of this repository's demo, with the React stylesheet entry.
// - No font package: the editor fetches its packaged substitute faces, and the chat uses the
//   system UI face, so the example adds no font download of its own.

import { fileURLToPath } from 'node:url';
import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

const repo = fileURLToPath(new URL('../..', import.meta.url));
const pkg = (path: string): string =>
  fileURLToPath(new URL(`../../packages/${path}`, import.meta.url));

export default defineConfig({
  root: fileURLToPath(new URL('.', import.meta.url)),
  plugins: [react()],
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
        find: /^@xlsx-editor\.dev\/react\/styles\.css$/,
        replacement: pkg('react/src/styles/editor.css'),
      },
      { find: /^@xlsx-editor\.dev\/react$/, replacement: pkg('react/src/index.ts') },
      { find: /^@xlsx-editor\.dev\/i18n\/([\w-]+)$/, replacement: pkg('i18n/src/$1.ts') },
      { find: /^@xlsx-editor\.dev\/i18n$/, replacement: pkg('i18n/src/index.ts') },
    ],
  },
  optimizeDeps: { exclude: ['harfbuzzjs', '@docx-editor.dev/fonts'] },
  build: { target: 'es2022' },
  server: { port: 5182, strictPort: true, fs: { allow: [repo] } },
  preview: { port: 5182 },
});
