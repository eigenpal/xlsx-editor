// Copied from pptx-editor/packages/vue/tsup.config.ts at d4502552 (EigenPal). Differences: the
// external scope is `@xlsx-editor.dev`, and TypeScript 7 builds the declarations through
// `withDeclarations`, as for the React adapter here (`packages/react/tsup.config.ts`), in place
// of `dts: true`.
import { defineConfig } from 'tsup';
import { withDeclarations } from '../../scripts/build-declarations.mjs';

// `core` stays external and a peer, as in the React adapter: the engine must resolve to ONE
// copy (shaper, caches). Simple parts use `h()`; complex parts are `.tsx` with Vue's JSX
// runtime, as in DOCX Editor's Vue adapter.
export default defineConfig(
  withDeclarations(import.meta.url, {
    entry: { index: 'src/index.ts' },
    format: ['esm'],
    clean: true,
    metafile: true,
    sourcemap: false,
    external: ['vue', /^@xlsx-editor\.dev\//],
    esbuildOptions(options) {
      options.jsx = 'automatic';
      options.jsxImportSource = 'vue';
    },
  })
);
