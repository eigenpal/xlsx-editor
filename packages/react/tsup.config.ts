// Copied from pptx-editor/packages/react/tsup.config.ts at e22f3f3b (EigenPal). Differences: the
// external scope is `@xlsx-editor.dev`, and TypeScript 7 builds the declarations through
// `withDeclarations`, as in DOCX Editor (docx-editor/packages/react/tsup.config.ts at 4603fc01).
import { defineConfig } from 'tsup';
import { withDeclarations } from '../../scripts/build-declarations.mjs';

// `core` stays external and a peer: the engine must resolve to ONE copy (shaper, caches).
export default defineConfig(
  withDeclarations(import.meta.url, {
    entry: { index: 'src/index.ts' },
    format: ['esm'],
    clean: true,
    metafile: true,
    sourcemap: false,
    external: ['react', 'react-dom', /^@xlsx-editor\.dev\//],
  })
);
