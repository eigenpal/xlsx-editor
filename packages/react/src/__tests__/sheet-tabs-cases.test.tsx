/** @jsxImportSource react */
// More sheet tab cases in happy-dom: a refused rename on blur, a rename while the sheet list
// changes, the rename field over a scrolled tab list, focus and dismissal of the tab menu, the
// shown tab scrolled into view, and Delete on the last visible sheet.

import { afterEach, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { act, cleanup, fireEvent, render, waitFor } from '@testing-library/react';
import { buildXlsx, row, c } from '../../../../scripts/fixtures/builder.ts';
import { XlsxEditor, type FontSource, type XlsxEditorInstance } from '../index.ts';

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

const three = buildXlsx({
  sheets: [
    { name: 'A', rows: row(1, [c('A1', 1)]) },
    { name: 'B', rows: '' },
    { name: 'C', rows: '' },
  ],
});

async function open(bytes = three, tabs = 3) {
  let viewer: XlsxEditorInstance | null = null;
  const view = render(
    <div style={{ height: 400 }}>
      <button type="button">outside</button>
      <XlsxEditor document={bytes} fonts={FONTS} onReady={(v) => (viewer = v)} />
    </div>
  );
  await waitFor(() => expect(view.getAllByRole('tab').length).toBe(tabs), { timeout: 5000 });
  return { view, viewer: viewer! };
}

type View = Awaited<ReturnType<typeof open>>['view'];
const names = (view: View) => view.getAllByRole('tab').map((t) => t.textContent);
const sheetNames = (viewer: XlsxEditorInstance) =>
  viewer
    .snapshot()
    .sheets.map((s) => s.name)
    .join();

describe('rename', () => {
  test('a refused blur ends the rename and keeps the old name; a refused Enter keeps the field', async () => {
    const { view, viewer } = await open();
    fireEvent.doubleClick(view.getAllByRole('tab')[0]!);
    let field = view.getByLabelText('Sheet name') as HTMLInputElement;
    fireEvent.change(field, { target: { value: 'b' } });
    fireEvent.blur(field);
    expect(view.queryByLabelText('Sheet name')).toBeNull();
    expect(names(view)).toEqual(['A', 'B', 'C']);
    await waitFor(() =>
      expect(view.container.querySelector('.xlsx-refusal')?.textContent).toContain(
        'has the name B.'
      )
    );
    fireEvent.doubleClick(view.getAllByRole('tab')[0]!);
    field = view.getByLabelText('Sheet name') as HTMLInputElement;
    fireEvent.change(field, { target: { value: 'b' } });
    fireEvent.keyDown(field, { key: 'Enter' });
    expect(view.getByLabelText('Sheet name')).toBe(field);
    expect(sheetNames(viewer)).toBe('A,B,C');
  });

  test('a change of the sheet list ends the rename, so no other sheet is renamed', async () => {
    // A hidden sheet after C takes the position of C when A goes.
    const bytes = buildXlsx({
      sheets: [
        { name: 'A', rows: '' },
        { name: 'B', rows: '' },
        { name: 'C', rows: '' },
        { name: 'H', rows: '', state: 'hidden' },
      ],
    });
    const { view, viewer } = await open(bytes, 3);
    fireEvent.doubleClick(view.getAllByRole('tab')[2]!);
    const field = view.getByLabelText('Sheet name') as HTMLInputElement;
    fireEvent.change(field, { target: { value: 'Renamed' } });
    act(() => {
      expect(viewer.deleteSheet(0).ok).toBe(true);
    });
    await waitFor(() => expect(view.queryByLabelText('Sheet name')).toBeNull());
    expect(sheetNames(viewer)).toBe('B,C,H');
    // A new rename of C with Enter renames C.
    fireEvent.doubleClick(view.getAllByRole('tab')[1]!);
    const next = view.getByLabelText('Sheet name') as HTMLInputElement;
    expect(next.value).toBe('C');
    fireEvent.change(next, { target: { value: 'Renamed' } });
    fireEvent.keyDown(next, { key: 'Enter' });
    expect(sheetNames(viewer)).toBe('B,Renamed,H');
  });

  test('the field follows its tab when the tab list scrolls', async () => {
    const { view } = await open();
    const tab = view.getAllByRole('tab')[2]!;
    let left = 120;
    tab.getBoundingClientRect = () =>
      ({ left, width: 60, top: 0, right: left + 60, bottom: 20, height: 20 }) as DOMRect;
    fireEvent.doubleClick(tab);
    const field = view.getByLabelText('Sheet name') as HTMLInputElement;
    expect(field.style.left).toBe('120px');
    left = 70;
    fireEvent.scroll(view.getByRole('tablist'));
    expect(view.getByLabelText('Sheet name').style.left).toBe('70px');
  });
});

describe('the tab menu', () => {
  test('Tab, Escape, and a pointer Delete return focus to the element that had it', async () => {
    const { view, viewer } = await open();
    const outside = view.getByRole('button', { name: 'outside' });
    const doc = outside.ownerDocument;
    for (const key of ['Tab', 'Escape']) {
      outside.focus();
      fireEvent.contextMenu(view.getAllByRole('tab')[1]!);
      expect(doc.activeElement?.textContent).toBe('Rename');
      fireEvent.keyDown(doc.activeElement!, { key });
      expect(view.queryByRole('menu')).toBeNull();
      expect(doc.activeElement).toBe(outside);
    }
    outside.focus();
    fireEvent.contextMenu(view.getAllByRole('tab')[1]!);
    fireEvent.click(view.getByRole('menuitem', { name: 'Delete' }));
    expect(sheetNames(viewer)).toBe('A,C');
    await waitFor(() => expect(doc.activeElement).toBe(outside));
  });

  test('a keyboard open returns focus to its tab, and a move focuses the moved tab', async () => {
    const { view } = await open();
    const doc = view.container.ownerDocument;
    fireEvent.keyDown(view.getAllByRole('tab')[0]!, { key: 'F10', shiftKey: true });
    fireEvent.keyDown(doc.activeElement!, { key: 'Tab' });
    await waitFor(() => expect(doc.activeElement).toBe(view.getAllByRole('tab')[0]!));
    fireEvent.keyDown(view.getAllByRole('tab')[0]!, { key: 'ContextMenu' });
    fireEvent.click(view.getByRole('menuitem', { name: 'Move right' }));
    expect(names(view)).toEqual(['B', 'A', 'C']);
    await waitFor(() => expect(doc.activeElement?.textContent).toBe('A'));
  });

  test('a window blur, a resize, a page scroll, and an outside wheel close the menu', async () => {
    const { view } = await open();
    const win = view.container.ownerDocument.defaultView!;
    const doc = view.container.ownerDocument;
    for (const fire of [
      () => win.dispatchEvent(new win.Event('blur')),
      () => win.dispatchEvent(new win.Event('resize')),
      () => doc.dispatchEvent(new win.Event('scroll')),
      () => doc.documentElement.dispatchEvent(new win.Event('scroll', { bubbles: true })),
      () => view.getByText('outside').dispatchEvent(new win.Event('wheel', { bubbles: true })),
    ]) {
      fireEvent.contextMenu(view.getAllByRole('tab')[1]!);
      expect(view.getByRole('menu')).toBeDefined();
      act(fire);
      expect(view.queryByRole('menu')).toBeNull();
    }
  });

  test('a scroll of the grid or the tab list and a wheel in the menu keep the menu', async () => {
    const { view } = await open();
    const win = view.container.ownerDocument.defaultView!;
    fireEvent.contextMenu(view.getAllByRole('tab')[1]!);
    const scroller = view.container.querySelector('.xlsx-scroller')!;
    act(() => {
      scroller.dispatchEvent(new win.Event('scroll'));
      view.getByRole('tablist').dispatchEvent(new win.Event('scroll'));
      view.getByRole('menu').dispatchEvent(new win.Event('wheel', { bubbles: true }));
    });
    expect(view.getByRole('menu')).toBeDefined();
  });

  test('Delete is disabled with its reason when one visible sheet is left', async () => {
    const bytes = buildXlsx({
      sheets: [
        { name: 'Only', rows: '' },
        { name: 'Hidden', rows: '', state: 'hidden' },
      ],
    });
    const { view, viewer } = await open(bytes, 1);
    fireEvent.contextMenu(view.getAllByRole('tab')[0]!);
    const del = view.getByRole('menuitem', { name: 'Delete' });
    expect(del.getAttribute('aria-disabled')).toBe('true');
    const reason = 'The sheet cannot be deleted: a workbook must keep one visible sheet.';
    expect(del.getAttribute('title')).toBe(`Delete: ${reason}`);
    const described = view.container.ownerDocument.getElementById(
      del.getAttribute('aria-describedby')!
    );
    expect(described?.textContent).toBe(reason);
    fireEvent.click(del);
    expect(sheetNames(viewer)).toBe('Only,Hidden');
    // With two visible sheets, Delete is enabled.
    cleanup();
    const two = await open(three, 3);
    fireEvent.contextMenu(two.view.getAllByRole('tab')[0]!);
    expect(two.view.getByRole('menuitem', { name: 'Delete' }).getAttribute('aria-disabled')).toBe(
      null
    );
  });
});

describe('the shown tab', () => {
  test('scrolls into view in the tab list when the shown sheet changes', async () => {
    const { view, viewer } = await open();
    const list = view.getByRole('tablist');
    list.style.overflowX = 'auto';
    Object.defineProperty(list, 'scrollWidth', { get: () => 500 });
    Object.defineProperty(list, 'clientWidth', { get: () => 100 });
    const scrolled: (string | null)[] = [];
    let measured: string | null = null;
    const proto = Object.getPrototypeOf(view.getAllByRole('tab')[0]!) as {
      getBoundingClientRect: () => DOMRect;
    };
    const before = proto.getBoundingClientRect;
    proto.getBoundingClientRect = function (this: HTMLElement) {
      if (this.getAttribute('role') !== 'tab') return before.call(this);
      measured = this.textContent;
      return { left: 300, top: 0, width: 60, height: 20, right: 360, bottom: 20 } as DOMRect;
    };
    list.scrollBy = (() => void scrolled.push(measured)) as typeof list.scrollBy;
    try {
      fireEvent.click(view.getByRole('button', { name: 'Add a sheet' }));
      await waitFor(() => expect(scrolled.at(-1)).toBe('Sheet1'));
      act(() => {
        viewer.setActiveSheet(3);
      });
      await waitFor(() => expect(scrolled.at(-1)).toBe('C'));
    } finally {
      proto.getBoundingClientRect = before;
    }
  });
});
