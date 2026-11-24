/** @jsxImportSource react */
// The demo at start, in happy-dom (`demo-harness.tsx`): the sample opens with its headings, the
// top bar names it, and every sheet shows; F6 and Shift+F6 move among the regions of the page;
// view only refuses edits; the theme toggle switches and keeps its pick; and a file that hides
// its headings shows none, and its save keeps that flag. These were steps of the browser check
// (`scripts/verify-demo.ts`), which keeps the paint of each sheet.

import { afterEach, describe, expect, test } from 'bun:test';
import { act, cleanup, fireEvent } from '@testing-library/react';
import { buildXlsx, c, row } from '../../../../scripts/fixtures/builder.ts';
import { SHEET_GUIDE } from '../../../../scripts/sample-workbook/showcase-sheets.ts';
import { mountDemo } from './demo-harness.tsx';
import { holdsText } from './demo-race-helpers.ts';

afterEach(() => {
  cleanup();
  localStorage.clear();
});

describe('the demo at start', () => {
  test('the sample opens with its headings, the top bar names it, and every sheet shows', async () => {
    const demo = await mountDemo();
    expect(demo.workbookName()).toBe('Sample workbook.xlsx');
    expect(demo.headingsShown()).toBe(true);
    expect(demo.view.getAllByRole('tab').map((t) => t.textContent)).toEqual(
      SHEET_GUIDE.map(([n]) => n)
    );
    for (const [name] of SHEET_GUIDE) await demo.tab(name);
  });

  test('F6 and Shift+F6 move the focus among the top bar, the menu bar, the toolbar, the grid, and the footer', async () => {
    const demo = await mountDemo();
    const inGrid = () => {
      const at = document.activeElement;
      return !!at?.closest('.xlsx-editor') && !at.closest('[role="toolbar"], .xlsx-footer');
    };
    const focusIn = (selector: string) => () => !!document.activeElement?.closest(selector);
    const inOpen = () => document.activeElement === demo.openInput;
    act(() => demo.grid.focus());
    for (const [shift, where, check] of [
      [false, 'the footer', focusIn('.xlsx-footer')],
      [false, 'Open a file', inOpen],
      [false, 'the menu bar', focusIn('[role="menubar"]')],
      [false, 'the toolbar', focusIn('[role="toolbar"]')],
      [true, 'the menu bar', focusIn('[role="menubar"]')],
      [true, 'Open a file', inOpen],
      [true, 'the footer', focusIn('.xlsx-footer')],
      [true, 'the grid', inGrid],
    ] as const) {
      demo.key('F6', { shiftKey: shift });
      expect([where, check()]).toEqual([where, true]);
    }
  });

  test('view only refuses a typed value and Bold, and edit mode takes Bold again', async () => {
    const demo = await mountDemo();
    const bold = () => demo.toolbarButton('Bold');
    const disabled = () =>
      (bold() as HTMLButtonElement).disabled || bold().getAttribute('aria-disabled') === 'true';
    act(() => void fireEvent.click(demo.viewOnlyBox));
    expect(disabled()).toBe(true);
    const before = demo.content.value;
    act(() => demo.grid.focus());
    demo.key('9', {}, demo.grid);
    demo.key('Enter', {}, demo.grid);
    demo.key('ArrowUp', {}, demo.grid);
    expect(demo.content.value).toBe(before);
    expect(demo.view.container.querySelector('.xlsx-cell-editor:not([hidden])')).toBeNull();
    act(() => void fireEvent.click(demo.viewOnlyBox));
    expect(disabled()).toBe(false);
  });

  test('the theme toggle switches by click and by arrow key, and the pick stays over a reload', async () => {
    let demo = await mountDemo();
    const dark = () => document.querySelector('.demo-root')!.classList.contains('dark');
    const radio = (name: string) => demo.view.getByRole('radio', { name }) as HTMLElement;
    expect(dark()).toBe(false);
    act(() => void fireEvent.click(radio('Dark')));
    expect(dark()).toBe(true);
    act(() => radio('Dark').focus());
    demo.key('ArrowLeft', {}, radio('Dark'));
    expect(dark()).toBe(false);
    expect(document.activeElement).toBe(radio('Light'));
    // A pick is kept for the next visit: dark after a reload, then light again.
    act(() => void fireEvent.click(radio('Dark')));
    cleanup();
    demo = await mountDemo();
    expect(dark()).toBe(true);
    act(() => void fireEvent.click(radio('Light')));
    expect(dark()).toBe(false);
  });

  test('a file that hides its headings shows none, and its save keeps that flag', async () => {
    const demo = await mountDemo();
    // `x:sheetView/@showRowColHeaders="0"` (CT_SheetView, ECMA-376 Part 1, 18.3.1.87).
    const hidden = buildXlsx({
      sheets: [{ name: 'Hidden', headings: false, rows: row(1, [c('A1', 'x')]) }],
    });
    demo.pickFile('hidden.xlsx', hidden);
    await demo.showsSheet('Hidden');
    expect(demo.headingsShown()).toBe(false);
    expect(holdsText(await demo.save(), 'showRowColHeaders="0"')).toBe(true);
  });
});
