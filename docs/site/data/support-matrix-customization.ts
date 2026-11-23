/**
 * The rows of the support matrix for the customization of the chrome by host code: the theme
 * scope and the editor events. `support-matrix.ts` lists them in its order; the rules of that
 * file hold here.
 */

import type { SupportFeature } from './support-matrix.ts';

const CORE = 'packages/core/src';

export const CUSTOMIZATION_FEATURES: readonly SupportFeature[] = [
  {
    id: 'viewer.chrome-theme',
    name: 'Chrome color mode and theme scope',
    category: 'viewer',
    status: 'supported',
    notes:
      'The `colorMode` prop sets a light, dark, or system theme for the chrome. The grid keeps the colors of the workbook. `useScopeClassName()` gives the scope classes for a part that you build outside the viewport, such as a popup in `document.body`. That part then reads the same `--doc-*` tokens and follows the color mode.',
    docsLink: '/docs/customization/appearance',
    evidence: {
      tests: ['packages/react/src/__tests__/adapter.test.tsx', 'packages/vue/test/adapter.test.ts'],
    },
  },
  {
    id: 'viewer.editor-events',
    name: 'Editor events for host code',
    category: 'viewer',
    status: 'supported',
    notes:
      'A host subscribes to four events with `useEditorEvent` or `editor.on`. `change` fires for an edit, an undo, a redo, a settled load, and a close. `selectionChange` fires when the sheet or the selection changes. `rejection` fires once for each refused command. `openFailure` fires when a workbook does not open. A zoom change fires no event.',
    docsLink: '/docs/customization/events',
    evidence: {
      tests: [
        `${CORE}/editor/__tests__/editor-events.test.ts`,
        'packages/react/src/__tests__/editor-event.test.tsx',
        'packages/vue/test/editor-event.test.ts',
      ],
    },
  },
];
