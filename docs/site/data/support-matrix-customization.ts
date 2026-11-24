/**
 * The rows of the support matrix for the customization of the chrome by host code: the theme
 * scope, the editor events, the language of the chrome, the menu bar, and the grid context menu. `support-matrix.ts` lists
 * them in its order; the rules of that file hold here.
 */

import type { SupportFeature } from './support-matrix.ts';

const CORE = 'packages/core/src';

export const CUSTOMIZATION_FEATURES: readonly SupportFeature[] = [
  {
    id: 'viewer.sheet-preview',
    name: 'Sheet preview for host surfaces',
    category: 'viewer',
    status: 'supported',
    notes:
      'The `SheetPreview` part and `editor.renderPreview` paint the top-left window of a worksheet, with no headings and no selection. The preview paints again after an edit, a sheet change, or a size change. It shows the stored values and evaluates no formula. A chart sheet gives no preview.',
    docsLink: '/docs/customization/layout',
    evidence: {
      tests: [
        `${CORE}/editor/__tests__/sheet-preview.test.ts`,
        'packages/react/src/__tests__/sheet-preview.test.tsx',
        'packages/vue/test/sheet-preview.test.ts',
      ],
    },
  },
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
      'A host subscribes to five events with `useEditorEvent` or `editor.on`. `change` fires for an edit, an undo, a redo, a settled load, and a close. `selectionChange` fires when the sheet or the selection changes. `rejection` fires once for each refused command. `openFailure` fires when a workbook does not open. `contextMenu` fires when the user asks for the grid context menu. A zoom change fires no event.',
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
  {
    id: 'viewer.menu-bar',
    name: 'Menu bar with host rows and menus',
    category: 'viewer',
    status: 'supported',
    notes:
      'File, Edit, View, Insert, and Format above the toolbar. Each packaged row is a chrome slot with the state of `toolbarCommandState`. A host hides or replaces a row by slot, adds rows, groups, submenus, and menus, and sets `preset={false}`. The bar follows the WAI-ARIA menubar pattern and moves the menus that do not fit into a "More menus" menu. Paste needs a clipboard that the browser lets a page read. Merge cells is disabled when cells other than the top-left cell hold values.',
    docsLink: '/docs/customization/menus',
    evidence: {
      tests: [
        `${CORE}/editor/__tests__/menu-bar.test.ts`,
        'packages/react/src/__tests__/menu.test.tsx',
        'packages/vue/test/menu.test.ts',
      ],
    },
  },
  {
    id: 'viewer.context-menu',
    name: 'Grid context menu',
    category: 'viewer',
    status: 'supported',
    notes:
      "A secondary press on the grid opens a context menu at the pointer, and Shift+F10 or the Menu key opens it at the active cell. A press outside the selection selects the cell, row, column, or sheet under the pointer first. The rows follow what the press hit: the clipboard, insert and delete of rows or columns, clear, merge for cells, and the number format dialog. Each row is a chrome slot with the label, enabled state, and command of the toolbar. A secondary press in an open cell edit keeps the browser menu. Hosts hide, replace, and add rows, or remove the menu. The menu writes no change of its own: each row runs the edit of its slot, which that edit's row covers.",
    docsLink: '/docs/customization/menus',
    evidence: {
      tests: [
        `${CORE}/editor/__tests__/grid-context-menu.test.ts`,
        'packages/react/src/__tests__/context-menu.test.tsx',
        'packages/vue/test/context-menu.test.ts',
      ],
    },
  },
];
