/**
 * The rows of the support matrix for frozen panes. `support-matrix.ts` lists them first in the
 * viewer rows; the rules of that file hold here.
 */

import type { SupportFeature } from './support-matrix.ts';

const CORE = 'packages/core/src';

export const PANE_FEATURES: readonly SupportFeature[] = [
  {
    id: 'viewer.frozen-panes',
    name: 'Frozen panes',
    category: 'viewer',
    status: 'partial',
    notes:
      'Frozen rows and columns (`x:pane/@state` set to `frozen` or `frozenSplit`, with `xSplit` and `ySplit`) stay in place while the rest of the sheet scrolls. A line marks the edge of each frozen band. The band starts at `x:sheetView/@topLeftCell`, and the split counts the frozen rows and columns from that cell (approximation "Frozen band origin"). The rows and columns before the band do not show while the pane is frozen, and the keys skip them. An inserted or deleted row or column inside a frozen band grows or shrinks the band, and an edit before the band moves it. A band that reaches 0 goes, with its pane. A frozen sheet opens scrolled to `x:pane/@topLeftCell`, with the active cell of the selection of `x:pane/@activePane`. A printed page shows no frozen panes, so no reference PDF covers them.',
    evidence: {
      approximations: ['frozen-band-origin'],
      tests: [
        `${CORE}/layout/__tests__/panes.test.ts`,
        `${CORE}/editor/__tests__/sheet-opening.test.ts`,
        `${CORE}/editor/__tests__/frozen-opening-scroll.test.ts`,
        `${CORE}/editor/__tests__/frozen-band.test.ts`,
        `${CORE}/editor/__tests__/heading-clip.test.ts`,
        `${CORE}/output/__tests__/paint-sheet.test.ts`,
        `${CORE}/layout/__tests__/sheet-layout.test.ts`,
        `${CORE}/store/__tests__/structure-cases.test.ts`,
        `${CORE}/layout/__tests__/structure-layout.test.ts`,
      ],
    },
  },
];
