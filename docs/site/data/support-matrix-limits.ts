/**
 * The rows of the support matrix for the limits of a workbook read. `support-matrix.ts` lists
 * them in its order; the rules of that file hold here. The workbook limits page
 * (`docs/site/content/workbook-limits.mdx`) gives the values.
 */

import type { SupportFeature } from './support-matrix.ts';

const CORE = 'packages/core/src';
const LOAD_STEPS_TEST = `${CORE}/editor/__tests__/load-steps.test.ts`;
const LIMIT_MESSAGES_TEST = 'packages/react/src/__tests__/limit-messages.test.tsx';

export const LIMIT_FEATURES: readonly SupportFeature[] = [
  {
    id: 'viewer.package-limits',
    name: 'Package limits',
    category: 'viewer',
    status: 'supported',
    notes:
      'A file past the limit of ZIP entries, of uncompressed content, or of the compression ratio does not open. The ratio limit applies only above 8 MiB of content. The editor reads no part of such a file. The workbook limits page gives the values.',
    docsLink: '/docs/workbook-limits',
    evidence: {
      tests: [
        'packages/ooxml-core/src/__tests__/zip.test.ts',
        `${CORE}/store/__tests__/package-read.test.ts`,
        LOAD_STEPS_TEST,
        'examples/vite/src/open-check.test.ts',
      ],
    },
  },
  {
    id: 'viewer.xml-part-limits',
    name: 'XML part limits',
    category: 'viewer',
    status: 'supported',
    notes:
      'Each XML part has a limit of size, of elements, and of nested levels. The content types part, the package relationships, the workbook part, and each worksheet refuse the file past a limit. So do the relationship parts of the workbook and of each worksheet. The styles part, the theme part, each drawing, and each other relationship part stay in the file as bytes past a limit, with a notice. The workbook limits page gives the values.',
    docsLink: '/docs/workbook-limits',
    evidence: {
      tests: [
        'packages/ooxml-core/src/__tests__/xml-reader.test.ts',
        `${CORE}/store/__tests__/package-read.test.ts`,
        LOAD_STEPS_TEST,
        'examples/vite/src/open-check.test.ts',
        `${CORE}/editor/__tests__/parts-past-limit.test.ts`,
      ],
    },
  },
  {
    id: 'viewer.shared-string-limit',
    name: 'Shared string table past a limit',
    category: 'viewer',
    status: 'supported',
    notes:
      'A shared string table past an XML part limit refuses the file, also when no relationship uses it. Without the table, each cell that refers to a shared string would show no text. A shared string table that is not valid XML stays in the file as bytes.',
    evidence: { tests: [LOAD_STEPS_TEST] },
  },
  {
    id: 'viewer.failure-limit',
    name: 'The limit that a refused file passes',
    category: 'viewer',
    status: 'supported',
    notes:
      'When a file passes a limit, the snapshot turns `failed` with the code `package`. `failure.limit` names the limit, such as `zip.maxRatio` or `xml.maxDepth`, with the names that DOCX Editor uses. A file that is not valid XML names no limit. For a part that does not read, `failure.part` names the part, and `failure.detail` names it for logs.',
    evidence: { tests: [LOAD_STEPS_TEST] },
  },
  {
    id: 'viewer.too-large-message',
    name: 'The message of a file past a limit',
    category: 'viewer',
    status: 'supported',
    notes:
      'When a refused file names a limit, the editor shows that the file is too large or too complex to open, as DOCX Editor does. The message names the part, for a limit of one part, and the limit with its value. A refused file with no limit keeps the message of a file that is not a workbook.',
    evidence: { tests: [LIMIT_MESSAGES_TEST, LOAD_STEPS_TEST] },
  },
  {
    id: 'viewer.part-past-limit-notice',
    name: 'The notice of a part past a limit',
    category: 'viewer',
    status: 'supported',
    notes:
      'A styles part, a theme part, a drawing, or another relationship part past an XML part limit stays in the file as bytes. The workbook opens without it, and a save keeps its bytes. While a drawing stays as bytes, the editor refuses to insert, delete, or resize rows and columns of its sheet. While the relationship part of a drawing stays as bytes, a row or column insert or delete on its sheet is refused. `snapshot().partsPastLimit` lists each such part with its kind and its limit. The editor shows a notice that names the part and what the workbook shows without it. While the styles part stays as bytes, a paste, Clear formats, and Clear all give the default format on a sheet that is not protected. Every other format change is refused, and the toolbar disables its controls with the reason. On a protected sheet, the toolbar also disables Clear formats and Clear all. An unused styles or theme part past a limit gets the kind `other` and locks no command.',
    evidence: {
      tests: [
        `${CORE}/editor/__tests__/parts-past-limit.test.ts`,
        `${CORE}/editor/__tests__/unread-styles-editing.test.ts`,
        `${CORE}/store/__tests__/drawing-kept-as-bytes.test.ts`,
        LIMIT_MESSAGES_TEST,
      ],
    },
  },
];
