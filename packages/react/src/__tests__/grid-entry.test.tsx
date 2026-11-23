/** @jsxImportSource react */
// The live region when the chrome gives the grid the focus: F6 says the active cell, and a step
// that says its own words, such as Enter in the formula bar or a press on a sheet tab, says them
// once.

import { afterEach, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { act, cleanup, fireEvent, render, waitFor } from '@testing-library/react';
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

const workbook = buildXlsx({
  sheets: [
    { name: 'First', rows: row(1, [c('A1', 'one')]) + row(2, [c('A2', 'next')]) },
    { name: 'Second', rows: row(1, [c('A1', 'two')]) },
  ],
});

/** Record each text that `live` takes, also two writes of one step. */
function record(live: HTMLElement): string[] {
  let proto: object | null = live;
  let desc: PropertyDescriptor | undefined;
  while (proto && !desc) {
    desc = Object.getOwnPropertyDescriptor(proto, 'textContent');
    proto = Object.getPrototypeOf(proto);
  }
  const said: string[] = [];
  Object.defineProperty(live, 'textContent', {
    configurable: true,
    get: () => desc!.get!.call(live),
    set: (text: string) => {
      said.push(text);
      desc!.set!.call(live, text);
    },
  });
  return said;
}

/** Open the workbook, and record each text that the live region takes after the open. */
async function open() {
  const view = render(
    <div style={{ height: 400 }}>
      <XlsxEditor document={workbook} fonts={FONTS} />
    </div>
  );
  await waitFor(() => expect(view.getAllByRole('tab').length).toBe(2), { timeout: 5000 });
  const said = record(view.container.querySelector('.xlsx-live') as HTMLElement);
  /** The texts said since the last call, after a timer ends the step, as a task ends it. */
  const heard = async () => {
    await act(() => new Promise((done) => setTimeout(done, 0)));
    return said.splice(0).filter((text) => text !== '');
  };
  const grid = view.container.querySelector('.xlsx-scroller') as HTMLElement;
  const nameBox = view.getByLabelText('Cell reference') as HTMLInputElement;
  const content = view.getByLabelText('Cell content') as HTMLTextAreaElement;
  return { view, heard, grid, nameBox, content };
}

describe('the chrome gives the grid the focus', () => {
  test('F6 from the formula bar says the active cell', async () => {
    const { heard, grid, nameBox } = await open();
    act(() => nameBox.focus());
    await heard();
    fireEvent.keyDown(nameBox, { key: 'F6' });
    expect(document.activeElement).toBe(grid);
    expect(await heard()).toEqual(['A1, one']);
  });

  test('Enter in the formula bar says the next cell once', async () => {
    const { heard, grid, content } = await open();
    act(() => content.focus());
    await heard();
    fireEvent.keyDown(content, { key: 'Enter' });
    expect(document.activeElement).toBe(grid);
    expect(await heard()).toEqual(['A2, next']);
  });

  test('Enter after an edit that changes nothing says the next cell once', async () => {
    const { heard, grid, content } = await open();
    act(() => content.focus());
    act(() => void fireEvent.change(content, { target: { value: 'onex' } }));
    act(() => void fireEvent.change(content, { target: { value: 'one' } }));
    await heard();
    fireEvent.keyDown(content, { key: 'Enter' });
    expect(document.activeElement).toBe(grid);
    expect(await heard()).toEqual(['A2, next']);
  });

  test('Escape after a refused commit says the active cell', async () => {
    const { heard, grid, content } = await open();
    act(() => content.focus());
    act(() => void fireEvent.change(content, { target: { value: '=1+' } }));
    fireEvent.keyDown(content, { key: 'Enter' });
    expect(await heard()).toHaveLength(1);
    // Escape empties the region with the refusal, so the grid says the cell that it enters.
    fireEvent.keyDown(content, { key: 'Escape' });
    expect(document.activeElement).toBe(grid);
    expect(await heard()).toEqual(['A1, one']);
  });

  test('a refused rename that loses the focus to the grid says only its reason', async () => {
    const { view, heard, grid } = await open();
    fireEvent.doubleClick(view.getByRole('tab', { name: 'First' }));
    const field = view.getByLabelText('Sheet name') as HTMLInputElement;
    act(() => void fireEvent.change(field, { target: { value: 'Second' } }));
    await heard();
    act(() => grid.focus());
    const said = await heard();
    expect(said).toHaveLength(1);
    expect(said[0]).toContain('Second');
    expect(said[0]).not.toContain('A1');
  });

  test('a press on a tab that ends a rename says only the sheet that it chooses', async () => {
    const { view, heard, grid } = await open();
    const second = view.getByRole('tab', { name: 'Second' });
    fireEvent.click(second);
    fireEvent.doubleClick(second);
    const field = view.getByLabelText('Sheet name') as HTMLInputElement;
    act(() => void fireEvent.change(field, { target: { value: 'Renamed' } }));
    await heard();
    const first = view.getByRole('tab', { name: 'First' });
    fireEvent.mouseDown(first, { detail: 1 });
    // The editor paints between the press and the click.
    expect(await heard()).toEqual([]);
    fireEvent.mouseUp(first, { detail: 1 });
    fireEvent.click(first, { detail: 1 });
    expect(document.activeElement).toBe(grid);
    expect(await heard()).toEqual(['Sheet First, A1, one']);
  });

  test('a press on a tab that ends a rename and has no click gives the grid the focus', async () => {
    const { view, heard, grid } = await open();
    const second = view.getByRole('tab', { name: 'Second' });
    fireEvent.click(second);
    fireEvent.doubleClick(second);
    const field = view.getByLabelText('Sheet name') as HTMLInputElement;
    act(() => void fireEvent.change(field, { target: { value: 'Renamed' } }));
    await heard();
    // A drag off the tab ends the press with no click.
    fireEvent.mouseDown(view.getByRole('tab', { name: 'First' }), { detail: 1 });
    fireEvent.mouseUp(document.body, { detail: 1 });
    expect(await heard()).toEqual(['A1, two']);
    expect(document.activeElement).toBe(grid);
    expect(view.getByRole('tab', { name: 'Renamed' })).toBeTruthy();
    // A later release on the page takes no focus.
    act(() => grid.blur());
    fireEvent.mouseUp(document.body, { detail: 1 });
    expect(await heard()).toEqual([]);
    expect(document.activeElement).toBe(document.body);
  });

  test('a press on a tab with no rename and no click takes no focus', async () => {
    const { view, heard } = await open();
    await heard();
    fireEvent.mouseDown(view.getByRole('tab', { name: 'Second' }), { detail: 1 });
    fireEvent.mouseUp(document.body, { detail: 1 });
    expect(await heard()).toEqual([]);
    expect(document.activeElement).toBe(document.body);
  });

  test('a right press on a tab that ends a rename keeps its menu after the release', async () => {
    const { view, heard } = await open();
    fireEvent.doubleClick(view.getByRole('tab', { name: 'First' }));
    const field = view.getByLabelText('Sheet name') as HTMLInputElement;
    act(() => void fireEvent.change(field, { target: { value: 'Renamed' } }));
    await heard();
    const second = view.getByRole('tab', { name: 'Second' });
    fireEvent.mouseDown(second, { button: 2, detail: 1 });
    fireEvent.contextMenu(second, { clientX: 40 });
    fireEvent.mouseUp(second, { button: 2, detail: 1 });
    // The menu takes the focus, so the release gives the grid no focus.
    expect(await heard()).toEqual(['Sheet Second, A1, two']);
    expect(view.getByRole('menu')).toBeTruthy();
  });

  test('a press on a sheet tab says the sheet once', async () => {
    const { view, heard, grid, nameBox } = await open();
    act(() => nameBox.focus());
    await heard();
    fireEvent.click(view.getByRole('tab', { name: 'Second' }));
    expect(document.activeElement).toBe(grid);
    expect(await heard()).toEqual(['Sheet Second, A1, two']);
  });
});
