// Reads the code samples that a README or a React guide marks for `check:consumer-install`. A
// marked sample is the fenced block that follows a `<!-- sample: name -->` comment, so the check
// builds the text that a reader copies.

/**
 * @param {string} markdown The README text.
 * @param {string} name The name in the `<!-- sample: name -->` comment.
 * @returns {string} The code of the fenced block after the comment.
 */
export function extractSample(markdown, name) {
  const lines = markdown.split(/\r?\n/);
  const markers = lines.flatMap((line, index) =>
    line.trim() === `<!-- sample: ${name} -->` ? [index] : []
  );
  if (markers.length !== 1) {
    throw new Error(`README sample "${name}": expected one marker, found ${markers.length}`);
  }
  let open = markers[0] + 1;
  // A `<!-- prettier-ignore -->` comment may sit between the marker and the block.
  while (open < lines.length && ['', '<!-- prettier-ignore -->'].includes(lines[open].trim())) {
    open += 1;
  }
  if (!lines[open]?.startsWith('```')) {
    throw new Error(`README sample "${name}": no fenced block follows the marker`);
  }
  const close = lines.findIndex((line, index) => index > open && line.startsWith('```'));
  if (close === -1) throw new Error(`README sample "${name}": the fenced block does not end`);
  return `${lines.slice(open + 1, close).join('\n')}\n`;
}

const editorRef = `import type { RefObject } from 'react';
import type { XlsxEditorRef } from '@xlsx-editor.dev/react';

declare const ref: RefObject<XlsxEditorRef | null>;
`;
const editorElement = `import { XlsxEditor } from '@xlsx-editor.dev/react';
`;

/**
 * The marked samples of the React guides in `packages/react/docs/`, which `check:consumer-install`
 * type-checks. A sample that is a fragment uses names that the guide text defines, so its
 * `prelude` declares them; the check writes the prelude before the sample as written.
 *
 * @type {readonly { guide: string, name: string, prelude: string }[]}
 */
export const GUIDE_SAMPLES = [
  { guide: 'editing-cells.md', name: 'guide-cell-field', prelude: '' },
  { guide: 'formatting-and-structure.md', name: 'guide-apply-format', prelude: editorRef },
  { guide: 'formatting-and-structure.md', name: 'guide-merge-button', prelude: '' },
  { guide: 'sheets-clipboard-and-sizes.md', name: 'guide-sheet-calls', prelude: editorRef },
  {
    guide: 'toolbar-and-keyboard.md',
    name: 'guide-on-save',
    prelude: `${editorElement}
declare const bytes: Uint8Array;
declare function upload(bytes: Uint8Array): void;
`,
  },
  { guide: 'toolbar-and-keyboard.md', name: 'guide-hidden-slots', prelude: editorElement },
  {
    guide: 'toolbar-and-keyboard.md',
    name: 'guide-toolbar-parts',
    prelude: `${editorElement}
import type { ReactElement } from 'react';

declare function SaveToCloudButton(): ReactElement;
declare function ShareButton(): ReactElement;
`,
  },
  { guide: 'toolbar-and-keyboard.md', name: 'guide-toolbar-own', prelude: editorElement },
  { guide: 'toolbar-and-keyboard.md', name: 'guide-number-format-dialog', prelude: '' },
  { guide: 'toolbar-and-keyboard.md', name: 'guide-bold-button', prelude: '' },
  {
    guide: 'view-mode-and-protection.md',
    name: 'guide-view-mode',
    prelude: `${editorElement}
declare const bytes: Uint8Array;
declare const canEdit: boolean;
`,
  },
];
