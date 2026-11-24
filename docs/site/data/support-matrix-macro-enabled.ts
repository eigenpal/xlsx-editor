/**
 * The row of the support matrix for macro-enabled workbooks and templates. `support-matrix.ts`
 * lists it in its order; the rules of that file hold here.
 */

import type { SupportFeature } from './support-matrix.ts';

const CORE = 'packages/core/src';

export const MACRO_ENABLED_FEATURES: readonly SupportFeature[] = [
  {
    id: 'save.macro-enabled',
    name: 'Macro-enabled workbooks and templates',
    category: 'save',
    status: 'supported',
    notes:
      'A macro-enabled workbook (`.xlsm`) or template (`.xltm`) opens, edits, and saves. The macro part stays as bytes that are never run, and a save keeps it and the content type of the main part. When the workbook holds macros, row and column inserts and deletes, and sheet adds, renames, deletes, and moves, are refused. The macros can refer to the cells and sheets that these edits change, or read a sheet by its position. The editor never rewrites macro code. `snapshot().macroEnabled` is true, so a host names the saved file with the macro-enabled extension. An add-in package is refused. The fixture `macro-enabled` is a macro-enabled workbook with no macro part. Its feature edit keeps the content type of the main part.',
    evidence: {
      fixtures: ['macro-enabled'],
      saveVerified: ['macro-enabled'],
      tests: [
        `${CORE}/editor/__tests__/macro-enabled.test.ts`,
        'packages/editor-api/src/__tests__/macro-enabled.test.ts',
        'packages/react/src/__tests__/sheet-tabs.test.tsx',
        'packages/vue/test/sheet-tabs.test.ts',
        `scripts/fidelity/feature-edits.test.ts`,
      ],
    },
  },
];
