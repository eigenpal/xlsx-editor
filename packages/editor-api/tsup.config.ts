// After PPTX Editor's `packages/editor-api/tsup.config.ts` at d4502552. Differences: the
// external scope is `@xlsx-editor.dev`, and TypeScript 7 builds the declarations through
// `withDeclarations`, as for the other packages of this repository.
import { readFileSync } from 'node:fs';
import { defineConfig } from 'tsup';
import { withDeclarations } from '../../scripts/build-declarations.mjs';

// The license header opens every JS and .d.ts file in dist as a legal comment (`/*!`), which
// minifiers keep.
const banner = `/*!\n${readFileSync(new URL('./license-header.txt', import.meta.url), 'utf8').trim()}\n*/`;

// Minified and without source maps, so the built files do not carry the source. `core` stays
// external and a peer, as in the adapters: the engine must resolve to ONE copy.
export default defineConfig(
  withDeclarations(import.meta.url, {
    entry: { index: 'src/index.ts' },
    format: ['esm'],
    clean: true,
    metafile: true,
    sourcemap: false,
    minify: true,
    banner: { js: banner },
    declarations: { banner },
    external: [/^@xlsx-editor\.dev\//],
  })
);
