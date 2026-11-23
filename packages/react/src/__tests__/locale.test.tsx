/** @jsxImportSource react */
// A change of the `locale` prop relabels the mounted surface once: the grid label, the accessible
// stale words, and the cell editor label, which the engine wrote with the catalog of the first
// render.

import { afterEach, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { cleanup, render, waitFor } from '@testing-library/react';
import { buildXlsx, c, row } from '../../../../scripts/fixtures/builder.ts';
import {
  XlsxEditor,
  type FontSource,
  type Translations,
  type XlsxEditorInstance,
} from '../index.ts';

const fontsDir = join(
  dirname(
    Bun.resolveSync('@docx-editor.dev/fonts/package.json', join(import.meta.dir, '../../../core'))
  ),
  'assets'
);
const FONTS: FontSource[] = [
  {
    family: 'Carlito',
    bold: false,
    italic: false,
    bytes: new Uint8Array(readFileSync(join(fontsDir, 'Carlito-Regular.ttf'))),
    origin: 'substitute',
  },
];

afterEach(cleanup);

// `fullCalcOnLoad` marks the cached formula results as stale: B1 has one, and C1 has none.
const workbook = buildXlsx({
  calcPr: 'fullCalcOnLoad="1"',
  sheets: [
    {
      name: 'First',
      rows: row(1, [c('A1', 2), c('B1', 4, { f: 'A1*2' }), c('C1', undefined, { f: 'A1+1' })]),
    },
  ],
});

test('a locale change relabels the grid, its stale words, and the cell editor', async () => {
  const ready: { editor: XlsxEditorInstance | null } = { editor: null };
  const view = render(
    <div style={{ height: 400 }}>
      <XlsxEditor document={workbook} fonts={FONTS} onReady={(e) => (ready.editor = e)} />
    </div>
  );
  const grid = () => view.container.querySelector('[role="grid"]')!;
  const input = () => view.container.querySelector('textarea.xlsx-cell-editor')!;
  await waitFor(() => expect(grid()?.getAttribute('aria-label')).toBe('Worksheet First'), {
    timeout: 5000,
  });
  const cell = (ref: string) => view.container.querySelector(`[data-cell="${ref}"]`)!;
  // happy-dom lays out nothing: give the grid a size, so the cells get accessible rows.
  Object.defineProperty(grid(), 'clientWidth', { get: () => 800 });
  Object.defineProperty(grid(), 'clientHeight', { get: () => 400 });
  // A selection paints again, and leaves an announcement in the old catalog, which the
  // relabel empties.
  const live = () => view.container.querySelector('.xlsx-live')!;
  ready.editor!.select({ row: 0, col: 1 });
  expect(live().textContent).not.toBe('');
  await waitFor(() => expect(cell('B1')?.textContent).toBe('4, not recalculated'));
  expect(cell('C1').textContent).toBe('No result, not recalculated');
  const before = input().getAttribute('aria-label');
  const locale: Translations = {
    grid: {
      label: 'Blatt {sheet}',
      staleResult: '{text}, nicht neu berechnet',
      staleNoResult: 'Kein Ergebnis, nicht neu berechnet',
    },
    editing: { cellEditor: 'Zelleditor' },
  };
  view.rerender(
    <div style={{ height: 400 }}>
      <XlsxEditor document={workbook} fonts={FONTS} locale={locale} />
    </div>
  );
  await waitFor(() => expect(grid().getAttribute('aria-label')).toBe('Blatt First'));
  expect(before).not.toBe('Zelleditor');
  expect(input().getAttribute('aria-label')).toBe('Zelleditor');
  expect(cell('B1').textContent).toBe('4, nicht neu berechnet');
  expect(cell('C1').textContent).toBe('Kein Ergebnis, nicht neu berechnet');
  expect(live().textContent).toBe('');
});
