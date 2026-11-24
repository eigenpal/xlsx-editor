/**
 * The rows of the support matrix for the customization parts of the toolbar. `support-matrix.ts`
 * lists them after the editing toolbar row; the rules of that file hold here. The toolbar page
 * (`docs/site/content/customization/toolbar.mdx`) shows the parts.
 */

import type { SupportFeature } from './support-matrix.ts';

const CORE = 'packages/core/src';

export const TOOLBAR_FEATURES: readonly SupportFeature[] = [
  {
    id: 'viewer.toolbar-parts',
    name: 'Named toolbar parts',
    category: 'viewer',
    status: 'supported',
    notes:
      'Each toolbar slot has a named part, such as `Toolbar.Bold`, that takes `hidden` and `className`. A button part also takes `icon` and `asChild`. A slot changed in a menu group shows after the menu button. A value control ignores `icon` and `asChild`, with a development warning.',
    docsLink: '/docs/customization/toolbar',
    evidence: {
      tests: [
        `${CORE}/editor/__tests__/toolbar-arrange.test.ts`,
        `${CORE}/editor/__tests__/toolbar-fit.test.ts`,
        `packages/react/src/__tests__/toolbar-panel.test.tsx`,
        `packages/vue/test/toolbar-panel.test.ts`,
        `packages/react/src/__tests__/toolbar-customize.test.tsx`,
        `packages/vue/test/toolbar-customize.test.ts`,
      ],
    },
  },
  {
    id: 'viewer.toolbar-actions',
    name: 'Toolbar actions and separators',
    category: 'viewer',
    status: 'supported',
    notes:
      '`Toolbar.Action` is a host button with the toolbar style and the mousedown guard. A disabled action stays focusable and gives its reason. In a collapsed group, it is a row of the overflow panel. Outside the toolbar, inside the viewport, it keeps the button states. `Icon` draws a chrome icon or host path data for it. `Toolbar.Separator` draws a rule.',
    docsLink: '/docs/customization/toolbar',
    evidence: {
      tests: [
        `packages/react/src/__tests__/toolbar-customize.test.tsx`,
        `packages/vue/test/toolbar-customize.test.ts`,
        `packages/react/src/__tests__/icon.test.tsx`,
        `packages/vue/test/icon.test.ts`,
      ],
    },
  },
  {
    id: 'viewer.toolbar-overflow',
    name: 'Toolbar collapse priority',
    category: 'viewer',
    status: 'supported',
    notes:
      'Groups leave the bar in priority order, one whole group at a time. Before that, the font family and number format fields move to the overflow panel, and come back when they fit. A host group without a `priority` leaves first. A `pinned` group stays at every width. `overflowContent` sets what a group or a slot shows in the overflow panel.',
    docsLink: '/docs/customization/toolbar',
    evidence: {
      tests: [
        `${CORE}/editor/__tests__/toolbar-arrange.test.ts`,
        `packages/react/src/__tests__/toolbar-customize.test.tsx`,
        `packages/vue/test/toolbar-customize.test.ts`,
      ],
    },
  },
];
