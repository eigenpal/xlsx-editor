// Copied from pptx-editor/scripts/vendored-notices.mjs at d4502552 (EigenPal). Differences: the
// entries name the XLSX Editor packages and the data their source files hold. The Ecma section
// lists the tables that core bundles: the preset shape geometry and the preset color names that
// it inlines from `@eigenpal/ooxml-core`, its predefined function names, its built-in number
// formats, and its default indexed color palette. DOCX Editor lists the preset color names under
// the same notice (docx-editor/packages/core/licenses/ECMA-376.md at 35a29d6d). The react entry
// names the Material Symbols path data of `src/icons.tsx`. There is no text warp table and no
// MicroType Express decoder. The Ecma notice read refuses a NOTICE with no COPYRIGHT NOTICE
// heading, as docx-editor/scripts/ecma-legal-comment.mjs does. DOCX Editor ships the Ecma and
// Material Symbols notices as files in `packages/core/licenses/`; this file follows PPTX Editor
// and puts them in THIRD_PARTY_NOTICES.md, as AGENTS.md asks for a second format. The comment
// below names the scan scope of `notices-sources.test.ts`.
//
// Notices for third-party data copied into source files. npm packages are found from the
// bundler metafile; these are not, so they are listed here by hand. `notices-sources.test.ts`
// fails when a source note in the `src` of a package in `scripts/lib-packages.mjs` names a
// third-party origin that is not listed.
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { HARFBUZZ_NOTICE } from './harfbuzz-notices.mjs';

// The Ecma copyright notice must travel with the generated data (reference/ecma-376/NOTICE).
const ecmaFile = readFileSync(resolve(import.meta.dirname, '../reference/ecma-376/NOTICE'), 'utf8');
const ecmaStart = ecmaFile.indexOf('COPYRIGHT NOTICE');
if (ecmaStart < 0) throw new Error('reference/ecma-376/NOTICE has no COPYRIGHT NOTICE heading.');
const ECMA_NOTICE = ecmaFile.slice(ecmaStart).trim();
const MATERIAL_SYMBOLS = (path) =>
  `## Material Symbols\n\n\`${path}\` holds SVG path data from Material Symbols by Google, licensed under the Apache License, Version 2.0 (https://github.com/google/material-design-icons). The license text is in this package's \`LICENSE\` file.\n`;
// Data copied into source files (not npm packages, so the metafile cannot see it).
export const VENDORED = {
  '@xlsx-editor.dev/core': [
    MATERIAL_SYMBOLS('src/editor/icon-paths.ts'),
    '## ECMA-376 tables\n\nThese tables come from ECMA-376 Part 1, 5th edition (2016), published by Ecma International (https://ecma-international.org/publications-and-standards/standards/ecma-376/). They are used under the Ecma copyright notice below.\n\n' +
      '- The preset shape geometry table, which this package inlines from `@eigenpal/ooxml-core` (`drawingml-preset-geometry.generated.ts`) at build time. It is generated from `presetShapeDefinitions.xml`.\n' +
      '- The preset color names of `ST_PresetColorVal` (20.1.10.48), which this package inlines from `@eigenpal/ooxml-core` (`drawingml-preset-colors.ts`) at build time.\n' +
      '- The predefined function names of 18.17.7, in `src/store/formula/predefined-functions.generated.ts`.\n' +
      '- The built-in number format codes of 18.8.30, in `src/store/number-format/builtins.ts`.\n' +
      '- The default indexed color palette of 18.8.27, in `src/cascade/color.ts` and `src/store/number-format/parse.ts`.\n\n```\n' +
      ECMA_NOTICE +
      '\n```\n',
    HARFBUZZ_NOTICE,
  ],
  '@xlsx-editor.dev/react': [MATERIAL_SYMBOLS('src/icons.tsx')],
  '@xlsx-editor.dev/vue': [MATERIAL_SYMBOLS('src/icons.tsx')],
};
