// Two editors on one page do not interfere: a shortcut in one editor runs only in that editor,
// and the ids that tie a control to its label or reason are unique on the page. The Vue twin of
// `packages/react/src/__tests__/two-editors.test.tsx`, a guard of behavior that came before it.

import { afterEach, describe, expect, test } from 'bun:test';
import { fireEvent, waitFor, within } from '@testing-library/react';
import { h, nextTick } from 'vue';
import { buildXlsx, c, row } from '../../../scripts/fixtures/builder.ts';
import { XlsxEditor } from '../src/index.ts';
import { FONTS, mount, tabsShow, unmountAll } from './mount.ts';

afterEach(unmountAll);

const workbook = () => buildXlsx({ sheets: [{ name: 'S', rows: row(1, [c('A1', 1)]) }] });

/** Two editors on one page, each with its own save collector. */
async function openTwo() {
  const saved = { a: 0, b: 0 };
  const editor = (name: 'a' | 'b') =>
    h('div', { style: { height: '400px' }, 'data-test-editor': name }, [
      h(XlsxEditor, { document: workbook(), fonts: FONTS, onSave: () => void saved[name]++ }),
    ]);
  const view = mount(() => h('div', [editor('a'), editor('b')]));
  await tabsShow(view, 2);
  const grid = (name: string) =>
    view.container.querySelector(`[data-test-editor="${name}"] .xlsx-scroller`) as HTMLElement;
  const list = (name: string, label: RegExp) =>
    within(view.container.querySelector(`[data-test-editor="${name}"]`) as HTMLElement).getByRole(
      'button',
      { name: label }
    );
  return { saved, grid, list };
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
    await nextTick();
    expect(document.querySelectorAll('[data-popup]').length).toBe(2);
    const ids = [...document.querySelectorAll('[id]')].map((el) => el.id);
    expect(ids.length).toBeGreaterThan(0);
    expect(new Set(ids).size).toBe(ids.length);
  });
});
