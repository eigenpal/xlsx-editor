/** @jsxImportSource react */
// Two editors on one page do not interfere: a shortcut in one editor runs only in that editor,
// and the ids that tie a control to its label or reason are unique on the page. After the
// embedding checks of pptx-editor at 3bf89a88 (EigenPal), where a shortcut answers only for its
// own editor. The scoped save key and the ids came before this test: it guards them. The Vue
// twin is `packages/vue/test/two-editors.test.ts`.

import { afterEach, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { cleanup, fireEvent, render, waitFor, within } from '@testing-library/react';
import { buildXlsx, c, row } from '../../../../scripts/fixtures/builder.ts';
import { XlsxEditor, type FontSource } from '../index.ts';

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

const workbook = () => buildXlsx({ sheets: [{ name: 'S', rows: row(1, [c('A1', 1)]) }] });

/** Two editors on one page, each with its own save collector. */
async function openTwo() {
  const saved = { a: 0, b: 0 };
  const view = render(
    <>
      <div style={{ height: 400 }} data-test-editor="a">
        <XlsxEditor document={workbook()} fonts={FONTS} onSave={() => void saved.a++} />
      </div>
      <div style={{ height: 400 }} data-test-editor="b">
        <XlsxEditor document={workbook()} fonts={FONTS} onSave={() => void saved.b++} />
      </div>
    </>
  );
  await waitFor(() => expect(view.getAllByRole('tab').length).toBe(2), { timeout: 5000 });
  const grid = (name: string) =>
    view.container.querySelector(`[data-test-editor="${name}"] .xlsx-scroller`) as HTMLElement;
  const list = (name: string, label: RegExp) =>
    within(view.container.querySelector(`[data-test-editor="${name}"]`) as HTMLElement).getByRole(
      'button',
      { name: label }
    );
  return { view, saved, grid, list };
}

describe('two editors on one page', () => {
  test('Control and S in one editor saves only that editor', async () => {
    const { saved, grid } = await openTwo();
    fireEvent.keyDown(grid('b'), { key: 's', ctrlKey: true });
    await waitFor(() => expect(saved.b).toBe(1), { timeout: 5000 });
    expect(saved.a).toBe(0);
    fireEvent.keyDown(grid('a'), { key: 's', metaKey: true });
    await waitFor(() => expect(saved.a).toBe(1), { timeout: 5000 });
    expect(saved.b).toBe(1);
  });

  test('the ids on the page are unique, also with a list open in each editor', async () => {
    const { list } = await openTwo();
    // The number format list in one editor, the font list in the other.
    fireEvent.click(list('a', /^Number format/));
    fireEvent.click(list('b', /^Select font family/));
    expect(document.querySelectorAll('[data-popup]').length).toBe(2);
    const ids = [...document.querySelectorAll('[id]')].map((el) => el.id);
    expect(ids.length).toBeGreaterThan(0);
    expect(new Set(ids).size).toBe(ids.length);
  });
});
