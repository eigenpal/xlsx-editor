// The frozen panes of the hang tests (`cases.ts`). A frozen band starts at the cell of
// `x:sheetView/@topLeftCell` (ECMA-376 Part 1, 18.3.1.87), and `x:pane/@xSplit` and `@ySplit`
// (18.3.1.66) count its columns and rows. Splits past the end of the grid make the widest panes
// that layout allows (`WIDE_WINDOW` in `engine-run.ts`).

import { buildXlsx, type FixtureSheet } from '../fixtures/builder.ts';
import type { HangCase } from './cases.ts';

const HUGE_SPLITS = '<pane xSplit="4294967295" ySplit="4294967295" state="frozen"/>';

/** The cases of frozen panes. `first` is the first sheet, with the text `first` in A1. */
export const frozenPaneCases = (first: FixtureSheet): readonly HangCase[] => [
  {
    id: 'frozen-split-huge',
    covers: 'a frozen pane whose splits name more rows and columns than the grid holds',
    expect: 'open',
    bytes: () => buildXlsx({ sheets: [{ ...first, viewChildren: HUGE_SPLITS }] }),
  },
  {
    id: 'frozen-band-mid-grid',
    covers: 'a frozen band that starts in the middle of the grid, with splits past its end',
    expect: 'open',
    bytes: () =>
      buildXlsx({
        sheets: [{ ...first, viewAttrs: 'topLeftCell="XA524288"', viewChildren: HUGE_SPLITS }],
      }),
  },
];
