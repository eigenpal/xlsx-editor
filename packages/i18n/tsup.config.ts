// After pptx-editor/packages/i18n/tsup.config.ts at e22f3f3b (EigenPal), with the locale entries of
// docx-editor/packages/i18n/tsup.config.ts at 708f28b8: one entry for each shipped locale
// (`@xlsx-editor.dev/i18n/<code>`), read from `locale-files.mjs`, the same rule that
// `scripts/validate-i18n.mjs` uses for the `exports` map. Differences: ESM only, as in PPTX
// Editor, and TypeScript 7 builds the declarations through `withDeclarations`, as in DOCX Editor.
import { defineConfig } from 'tsup';
import { withDeclarations } from '../../scripts/build-declarations.mjs';
import { readLocaleCodes } from './locale-files.mjs';

const localeCodes = readLocaleCodes(import.meta.dirname);

export default defineConfig(
  withDeclarations(import.meta.url, {
    entry: {
      index: 'src/index.ts',
      ...Object.fromEntries(localeCodes.map((code) => [code, `src/${code}.ts`])),
    },
    format: ['esm'],
    clean: true,
    metafile: true,
    sourcemap: false,
    // Keep non-ASCII text raw instead of `\uHHHH` escapes, as in DOCX Editor.
    esbuildOptions(options) {
      options.charset = 'utf8';
    },
  })
);
