/**
 * The rows of the support matrix for the customization of the chrome by host code: the theme
 * scope, the editor events, and the language of the chrome. `support-matrix.ts` lists them in
 * its order; the rules of that file hold here.
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
  {
    id: 'viewer.chrome-language',
    name: 'Chrome language, label overrides, and right-to-left chrome',
    category: 'viewer',
    status: 'supported',
    notes:
      'The root `locale` prop takes a partial catalog. A missing or `null` entry falls back to the catalog of an outer `LocaleProvider`, then to English. A change keeps the open workbook. `LocaleProvider` overrides the labels of a region, or sets the catalog of the app around the root. The toolbar, the formula bar, the sheet tabs, the zoom control, and the status message each take a `t` prop. A `t` receives a key and the values of its placeholders. The toolbar `t` also reaches its number format dialog. `useChromeTranslate(overrides)` builds such a `t`, and fills the placeholders and plural branches of an override. A right-to-left catalog or the `dir` of `LocaleProvider` mirrors the chrome parts and their arrow keys. It also mirrors the toolbar popups and the sheet tab menu. The grid keeps a left-to-right layout, also on a right-to-left host page. The cell editor and the formula bar fields set their direction from their text.',
    docsLink: '/docs/customization/language',
    evidence: {
      tests: [
        `${CORE}/editor/__tests__/chrome-translate.test.ts`,
        'packages/i18n/src/direction.test.ts',
        'packages/react/src/__tests__/chrome-translate.test.tsx',
        'packages/vue/test/chrome-translate.test.ts',
      ],
    },
  },
];
