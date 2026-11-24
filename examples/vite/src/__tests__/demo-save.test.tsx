/** @jsxImportSource react */
// The saves of the demo, in happy-dom (`demo-harness.tsx`): the top bar Save, the toolbar Save,
// and Control+S each write one copy that keeps the edit and every sheet; the number format
// dialog, the number format list, the decimal step, and typed entry write codes that a save
// keeps. These were steps of the browser check (`scripts/verify-demo.ts`).

import { afterEach, describe, expect, test } from 'bun:test';
import { act, cleanup, fireEvent, waitFor } from '@testing-library/react';
import { EN_US } from '../../../../packages/core/src/cascade/index.ts';
import { numberFormatChoiceCode } from '../../../../packages/core/src/editor/index.ts';
import { attr, child, textOf, workbookIndex } from '../../../../packages/core/src/store/index.ts';
import { SAMPLE_BOLD_EDIT } from '../../../../scripts/sample-workbook/edits.ts';
import { SHEET_GUIDE } from '../../../../scripts/sample-workbook/showcase-sheets.ts';
import { formatAt, mountDemo, readSaved, settle, WAIT_MS, type Demo } from './demo-harness.tsx';

afterEach(cleanup);

const { sheet, ref } = SAMPLE_BOLD_EDIT;

/** Reads a saved copy, and checks that it holds every sheet of the sample. */
function savedSample(bytes: Uint8Array) {
  const pkg = readSaved(bytes);
  expect(workbookIndex(pkg)?.sheets.map((s) => s.name)).toEqual(SHEET_GUIDE.map(([n]) => n));
  return pkg;
}

const boldIn = (bytes: Uint8Array) => formatAt(savedSample(bytes), sheet, ref).bold;

/** Mounts the demo on the sheet of the Bold edit. */
async function onEditSheet(): Promise<Demo> {
  const demo = await mountDemo();
  await demo.tab(sheet);
  return demo;
}

describe('a save of the demo', () => {
  test('the top bar Save, the toolbar Save, and Control+S each write one copy with the edit', async () => {
    const demo = await onEditSheet();
    demo.bold();
    expect(boldIn(await demo.save('header'))).toBe(true);
    // The message of the save clears on the next edit. Two Bold clicks leave the cell bold.
    await demo.statusSays('Saved ');
    expect(demo.leaveAsks()).toBe(false);
    demo.bold();
    await waitFor(() => expect(demo.notice()).not.toContain('Saved '), { timeout: WAIT_MS });
    expect(demo.leaveAsks()).toBe(true);
    demo.bold();
    expect(boldIn(await demo.save('toolbar'))).toBe(true);
    expect(boldIn(await demo.save('keys'))).toBe(true);
    // The keys save from the top bar too.
    act(() => demo.openInput.focus());
    const before = demo.saved.length;
    demo.saveKey(demo.openInput);
    await waitFor(() => expect(demo.saved.length).toBe(before + 1), { timeout: WAIT_MS });
    expect(boldIn(demo.saved.at(-1)!)).toBe(true);
    // A held key saves once: its repeats stay in the page.
    act(() => demo.grid.focus());
    const held = demo.saved.length;
    for (let i = 0; i < 4; i++) demo.saveKey(demo.grid, { repeat: i > 0 });
    await settle(300);
    await waitFor(() => expect(demo.busy()).toBe(''), { timeout: WAIT_MS });
    expect(demo.saved.length - held).toBe(1);
  });

  test('the number format dialog refuses a code that cannot be read, and applies a typed code that a save keeps', async () => {
    const demo = await onEditSheet();
    const code = '#,##0.000 "kg"';
    act(() => void fireEvent.click(demo.view.getByRole('button', { name: /^Number format/ })));
    act(
      () => void fireEvent.click(demo.view.getByRole('button', { name: 'More number formats…' }))
    );
    const dialog = demo.view.getByRole('dialog', { name: 'Number format' });
    const field = demo.view.getByRole('textbox', { name: 'Format code' }) as HTMLInputElement;
    act(() => void fireEvent.change(field, { target: { value: '"open' } }));
    act(() => void fireEvent.click(demo.view.getByRole('button', { name: 'Apply' })));
    expect(dialog.querySelector('[role=alert]')?.textContent).toContain('cannot be read');
    expect(field.getAttribute('aria-invalid')).toBe('true');
    act(() => void fireEvent.change(field, { target: { value: code } }));
    // The active cell holds text, which a number section shows as it is.
    const output = dialog.querySelector('output')!;
    expect(output.textContent).not.toBe('');
    expect(output.getAttribute('data-problem')).toBeNull();
    act(() => void fireEvent.click(demo.view.getByRole('button', { name: 'Apply' })));
    expect(demo.view.queryByRole('dialog')).toBeNull();
    const list = demo.view.getByRole('button', { name: /^Number format/ });
    expect(list.getAttribute('aria-label')).toContain(code);
    expect(formatAt(savedSample(await demo.save('toolbar')), sheet, ref).code).toBe(code);
  });

  test('the currency entry and the increase decimals command apply, and a save keeps each code', async () => {
    const demo = await onEditSheet();
    // A cell of the sample that holds a number: a decimal step needs one.
    demo.goTo('C5');
    const list = () => demo.view.getByRole('button', { name: /^Number format/ });
    act(() => void fireEvent.click(list()));
    act(() => void fireEvent.click(demo.view.getByRole('option', { name: 'Currency' })));
    expect(list().textContent).toContain('Currency');
    const currency = numberFormatChoiceCode('currency', EN_US)!;
    expect(formatAt(savedSample(await demo.save('toolbar')), sheet, 'C5').code).toBe(currency);
    demo.closeMessage();
    act(() => void fireEvent.click(demo.toolbarButton('Increase decimal places')));
    // A code that differs from an entry only in its decimal places shows the name of that entry.
    expect(list().textContent).toContain('Currency');
    const stepped = formatAt(savedSample(await demo.save('toolbar')), sheet, 'C5').code;
    expect(stepped).toBe('"$"#,##0.000');
  });

  test('a typed percent and a typed date store numbers with formats 9 and 14, shown formatted', async () => {
    const demo = await onEditSheet();
    act(() => demo.grid.focus());
    // Control+Home goes to A1; K5 is below the two merged title rows and the empty row 3.
    demo.key('Home', { ctrlKey: true }, demo.grid);
    for (let i = 0; i < 4; i++) demo.key('ArrowDown', {}, demo.grid);
    for (let i = 0; i < 10; i++) demo.key('ArrowRight', {}, demo.grid);
    expect(demo.nameBox.value).toBe('K5');
    demo.typeInGrid('12%');
    demo.typeInGrid('2026-11-12');
    await settle(50);
    const cell = (r: string) => document.querySelector(`[data-cell="${r}"]`)?.textContent;
    expect(cell('K5')).toBe('12%');
    // Built-in format 14 in the demo's number locale (US English) is `mm-dd-yy`.
    expect(cell('K6')).toBe('11-12-26');
    const pkg = savedSample(await demo.save('header'));
    const at = (r: string) => {
      const { cell: el, numFmtId } = formatAt(pkg, sheet, r);
      return { type: attr(el?.el, 't') ?? 'n', value: textOf(child(el?.el, 'x:v')), numFmtId };
    };
    expect(at('K5')).toEqual({ type: 'n', value: '0.12', numFmtId: 9 });
    expect(at('K6')).toEqual({ type: 'n', value: '46338', numFmtId: 14 });
  });
});
