// A change of the `locale` prop relabels the mounted surface once: the grid label, the accessible
// stale words, and the cell editor label, which the engine wrote with the catalog of the first
// render.
// The Vue twin of the React adapter's `locale.test.tsx`: Vue renders in the next microtask, so a
// step waits for `nextTick` before it reads the DOM.

import { afterEach, expect, test } from 'bun:test';
import { waitFor } from '@testing-library/react';
import { h, nextTick, reactive } from 'vue';
import { buildXlsx, c, row } from '../../../scripts/fixtures/builder.ts';
import { XlsxEditor, type Translations, type XlsxEditorInstance } from '../src/index.ts';
import { FONTS, mount, unmountAll } from './mount.ts';

afterEach(unmountAll);

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
  const state = reactive<{ locale?: Translations }>({});
  const view = mount(() =>
    h('div', { style: { height: '400px' } }, [
      h(XlsxEditor, {
        document: workbook,
        fonts: FONTS,
        ...(state.locale ? { locale: state.locale } : {}),
        onReady: (e: XlsxEditorInstance) => (ready.editor = e),
      }),
    ])
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
  state.locale = locale;
  await nextTick();
  await waitFor(() => expect(grid().getAttribute('aria-label')).toBe('Blatt First'));
  expect(before).not.toBe('Zelleditor');
  expect(input().getAttribute('aria-label')).toBe('Zelleditor');
  expect(cell('B1').textContent).toBe('4, nicht neu berechnet');
  expect(cell('C1').textContent).toBe('Kein Ergebnis, nicht neu berechnet');
  expect(live().textContent).toBe('');
});
