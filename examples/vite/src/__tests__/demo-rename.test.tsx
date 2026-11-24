/** @jsxImportSource react */
// The workbook name field of the demo top bar (`WorkbookName.tsx`), in happy-dom
// (`demo-harness.tsx`): a typed name and Enter rename the workbook, and Save suggests that name;
// Escape puts back the name before the edit; an empty name keeps the name before it; a move of
// the focus away keeps the typed name; File > Rename workbook focuses the field; and the next
// open starts with the name of its file.

import { afterEach, describe, expect, test } from 'bun:test';
import { act, cleanup, fireEvent } from '@testing-library/react';
import { mountDemo, type Demo } from './demo-harness.tsx';

afterEach(cleanup);

/** The workbook name field of the top bar. */
const nameField = () =>
  document.querySelector<HTMLInputElement>('.demo-topbar input.demo-file__name')!;

/** Focuses the name field, as a click does, and types `text` over the name. */
function typeName(text: string): HTMLInputElement {
  const field = nameField();
  act(() => field.focus());
  act(() => void fireEvent.change(field, { target: { value: text } }));
  return field;
}

/** Presses `key` in the name field. */
const press = (field: HTMLInputElement, key: string) =>
  act(() => void fireEvent.keyDown(field, { key }));

/** The suggested name of the next save. */
async function suggestedOnSave(demo: Demo): Promise<string | undefined> {
  await demo.save();
  return demo.services.suggested.at(-1);
}

describe('the workbook name in the top bar', () => {
  test('is a named text field, and Enter keeps a typed name that Save suggests', async () => {
    const demo = await mountDemo();
    const field = nameField();
    expect([field.tagName, field.getAttribute('aria-label')]).toEqual(['INPUT', 'Workbook name']);
    expect(demo.workbookName()).toBe('Sample workbook.xlsx');
    expect(await suggestedOnSave(demo)).toBe('sample-edited.xlsx');

    press(typeName('  Budget 2026.xlsx '), 'Enter');
    expect(document.activeElement).not.toBe(field);
    // The typed extension goes, because the extension shows apart.
    expect(demo.workbookName()).toBe('Budget 2026.xlsx');
    expect(await suggestedOnSave(demo)).toBe('Budget 2026.xlsx');
  });

  test('Escape puts back the name before the edit, and an empty name keeps it', async () => {
    const demo = await mountDemo();
    press(typeName('Plan'), 'Enter');
    press(typeName('Draft'), 'Escape');
    expect(document.activeElement).not.toBe(nameField());
    expect(demo.workbookName()).toBe('Plan.xlsx');
    press(typeName('   '), 'Enter');
    expect(demo.workbookName()).toBe('Plan.xlsx');
    expect(await suggestedOnSave(demo)).toBe('Plan.xlsx');
  });

  test('a move of the focus away keeps the typed name', async () => {
    const demo = await mountDemo();
    const field = typeName('Moved away');
    act(() => field.blur());
    expect(demo.workbookName()).toBe('Moved away.xlsx');
  });

  test('File > Rename workbook focuses the field and selects the name', async () => {
    const demo = await mountDemo();
    const bar = document.querySelector<HTMLElement>('.demo-topbar [role="menubar"]')!;
    act(() => {
      fireEvent.click(bar.querySelector<HTMLElement>('[data-menu="file"] > [role="menuitem"]')!);
    });
    const row = [...document.querySelectorAll<HTMLElement>('[role="menuitem"]')].find(
      (el) => el.textContent?.trim() === 'Rename workbook'
    );
    expect(row?.getAttribute('aria-disabled')).not.toBe('true');
    act(() => void fireEvent.click(row!));
    const field = nameField();
    expect(document.activeElement).toBe(field);
    expect([field.selectionStart, field.selectionEnd]).toEqual([0, field.value.length]);
    press(typeName('From the menu'), 'Enter');
    expect(demo.workbookName()).toBe('From the menu.xlsx');
  });

  test('the next open starts with the name of its file', async () => {
    const demo = await mountDemo();
    press(typeName('Renamed'), 'Enter');
    demo.pickFixture('basic-values');
    await demo.showsSheet('Values');
    expect(demo.workbookName()).toBe('basic-values.xlsx');
    expect(await suggestedOnSave(demo)).toBe('basic-values-edited.xlsx');
  });
});
