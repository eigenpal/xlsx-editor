/** @jsxImportSource react */
// Sheet tab cases of design D4 of the change `workbook-structure-editing` in happy-dom: a right
// press on a tab after a refused rename, a toolbar command with an open rename, the arrow keys
// past a chartsheet tab, focus after a click on a tab and after a menu action, and the width and
// length of the rename field.

import { afterEach, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { cleanup, fireEvent, render, waitFor } from '@testing-library/react';
import { buildXlsx, row, c, type FixtureWorkbook } from '../../../../scripts/fixtures/builder.ts';
import {
  parseXmlPart,
  readXlsx,
  serializeXmlPart,
  workbookIndex,
  writeXlsx,
  xmlPartOf,
} from '../../../core/src/store/index.ts';
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

const book: FixtureWorkbook = {
  sheets: [
    { name: 'A', rows: row(1, [c('A1', 1)]) },
    { name: 'B', rows: '' },
    { name: 'C', rows: '' },
  ],
};

/** `A`, a chartsheet `Chart1`, `B`, and `C`. */
function withChartsheet(): Uint8Array {
  const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const read = readXlsx(
    buildXlsx({
      ...book,
      extraParts: [
        [
          'xl/chartsheets/sheet9.xml',
          'application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml',
          new TextEncoder().encode(
            '<chartsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView workbookViewId="0"/></sheetViews></chartsheet>'
          ),
          `${REL}/chartsheet`,
        ],
      ],
    })
  );
  if (!read.ok) throw new Error('read');
  const main = workbookIndex(read.pkg)!.mainPart;
  const text = serializeXmlPart(xmlPartOf(read.pkg, main)!).replace(
    '<sheet name="B"',
    '<sheet name="Chart1" sheetId="9" r:id="rIdX1"/><sheet name="B"'
  );
  const parsed = parseXmlPart(main, text);
  if (!parsed.ok) throw new Error(parsed.reason);
  const parts = new Map(read.pkg.parts);
  parts.set(main, { kind: 'xml', partName: main, xml: parsed.part });
  return writeXlsx({ parts });
}

async function open(bytes: Uint8Array = buildXlsx(book), onSave?: (b: Uint8Array) => void) {
  let viewer: XlsxEditorInstance | null = null;
  const view = render(
    <div style={{ height: 400 }}>
      <XlsxEditor
        document={bytes}
        fonts={FONTS}
        onReady={(v) => (viewer = v)}
        {...(onSave ? { onSave } : {})}
      />
    </div>
  );
  await waitFor(() => expect(view.getAllByRole('tab').length).toBeGreaterThan(2), {
    timeout: 5000,
  });
  return { view, viewer: viewer! };
}

type View = Awaited<ReturnType<typeof open>>['view'];
const names = (view: View) => view.getAllByRole('tab').map((t) => t.textContent);
const active = (view: View) => view.container.ownerDocument.activeElement;
const onGrid = (view: View) => active(view)?.classList.contains('xlsx-scroller') ?? false;

/** Start a rename of the first tab and type `text` in its field. */
function typeName(view: View, text: string): HTMLInputElement {
  fireEvent.doubleClick(view.getAllByRole('tab')[0]!);
  const field = view.getByLabelText('Sheet name') as HTMLInputElement;
  fireEvent.change(field, { target: { value: text } });
  return field;
}

/** A full press on an element: mousedown, then click. */
function press(el: HTMLElement) {
  fireEvent.mouseDown(el, { detail: 1 });
  fireEvent.click(el, { detail: 1 });
}

describe('a right press on another tab after a refused rename', () => {
  test('opens no menu, and the reason stays', async () => {
    const { view, viewer } = await open();
    typeName(view, 'b');
    const tab = view.getAllByRole('tab')[2]!;
    fireEvent.mouseDown(tab, { button: 2, detail: 1 });
    fireEvent.contextMenu(tab, { clientX: 40 });
    expect(view.queryByRole('menu')).toBeNull();
    expect(viewer.snapshot().activeSheet).toBe(0);
    expect(viewer.snapshot().refusal?.code).toBe('sheet-name-taken');
  });
});

describe('a toolbar command during a rename', () => {
  test('Save writes the typed name first', async () => {
    const saved: Uint8Array[] = [];
    const { view } = await open(buildXlsx(book), (b) => saved.push(b));
    typeName(view, 'Renamed');
    press(view.getByRole('button', { name: 'Save' }));
    await waitFor(() => expect(saved.length).toBe(1));
    const read = readXlsx(saved[0]!);
    if (!read.ok) throw new Error('read');
    expect(workbookIndex(read.pkg)!.sheets.map((s) => s.name)).toEqual(['Renamed', 'B', 'C']);
    expect(view.queryByLabelText('Sheet name')).toBeNull();
    expect(onGrid(view)).toBe(true);
  });

  test('Undo commits the rename first, then undoes it', async () => {
    const { view, viewer } = await open();
    expect(viewer.enterValue('5', { row: 1, col: 0 }).ok).toBe(true);
    typeName(view, 'Renamed');
    press(view.getByRole('button', { name: 'Undo' }));
    await waitFor(() => expect(names(view)).toEqual(['A', 'B', 'C']));
    expect(viewer.snapshot().canRedo).toBe(true);
  });

  test('a press on the disabled Undo commits the rename, and its click undoes nothing', async () => {
    const { view, viewer } = await open();
    const undo = view.getByRole('button', { name: 'Undo' });
    expect(undo.getAttribute('aria-disabled')).toBe('true');
    typeName(view, 'Renamed');
    press(undo);
    expect(names(view)).toEqual(['Renamed', 'B', 'C']);
    expect(viewer.snapshot().canUndo).toBe(true);
    expect(viewer.snapshot().canRedo).toBe(false);
    // One click only: the next click on the now enabled Undo, with no press (a key), runs it.
    fireEvent.click(view.getByRole('button', { name: 'Undo' }), { detail: 0 });
    expect(names(view)).toEqual(['A', 'B', 'C']);
  });

  test('a refused name runs no command, and its reason stays', async () => {
    const saved: Uint8Array[] = [];
    const { view, viewer } = await open(buildXlsx(book), (b) => saved.push(b));
    expect(viewer.enterValue('5', { row: 1, col: 0 }).ok).toBe(true);
    for (const name of ['Undo', 'Save']) {
      typeName(view, 'b');
      press(view.getByRole('button', { name }));
      expect(view.queryByLabelText('Sheet name')).toBeNull();
      expect(viewer.snapshot().refusal?.code).toBe('sheet-name-taken');
    }
    expect(saved.length).toBe(0);
    expect(viewer.snapshot().canUndo).toBe(true);
    expect(names(view)).toEqual(['A', 'B', 'C']);
    // The next press runs its command.
    press(view.getByRole('button', { name: 'Save' }));
    await waitFor(() => expect(saved.length).toBe(1));
  });
});

describe('a right press on a chartsheet tab', () => {
  test('opens no menu', async () => {
    const { view } = await open(withChartsheet());
    const chart = view.getAllByRole('tab')[1]!;
    expect(chart.textContent).toBe('Chart1');
    fireEvent.contextMenu(chart, { clientX: 40 });
    expect(view.queryByRole('menu')).toBeNull();
    // A worksheet tab still opens it.
    fireEvent.contextMenu(view.getAllByRole('tab')[0]!, { clientX: 40 });
    expect(view.getByRole('menu')).toBeDefined();
  });
});

describe('the arrow keys and a chartsheet tab', () => {
  test('skip the disabled chartsheet tab', async () => {
    const { view, viewer } = await open(withChartsheet());
    expect(names(view)).toEqual(['A', 'Chart1', 'B', 'C']);
    const tabs = view.getAllByRole('tab');
    tabs[0]!.focus();
    fireEvent.keyDown(tabs[0]!, { key: 'ArrowRight' });
    expect(viewer.snapshot().activeSheet).toBe(2);
    expect(active(view)).toBe(view.getAllByRole('tab')[2]!);
    fireEvent.keyDown(active(view)!, { key: 'ArrowLeft' });
    expect(viewer.snapshot().activeSheet).toBe(0);
    fireEvent.keyDown(active(view)!, { key: 'End' });
    expect(viewer.snapshot().activeSheet).toBe(3);
    fireEvent.keyDown(active(view)!, { key: 'Home' });
    expect(viewer.snapshot().activeSheet).toBe(0);
    expect(active(view)).toBe(view.getAllByRole('tab')[0]!);
  });
});

describe('focus after a click on another tab', () => {
  test('stays in the tab list when it had focus there', async () => {
    const { view } = await open();
    view.getAllByRole('tab')[0]!.focus();
    press(view.getAllByRole('tab')[2]!);
    expect(active(view)).toBe(view.getAllByRole('tab')[2]!);
  });

  test('goes to the grid otherwise', async () => {
    const { view } = await open();
    (active(view) as HTMLElement | null)?.blur();
    press(view.getAllByRole('tab')[1]!);
    expect(onGrid(view)).toBe(true);
  });
});

describe('a right press on another tab while the tab list has focus', () => {
  test('gives focus back to the pressed tab when the menu closes', async () => {
    const { view, viewer } = await open();
    view.getAllByRole('tab')[0]!.focus();
    fireEvent.contextMenu(view.getAllByRole('tab')[2]!, { clientX: 40 });
    expect(viewer.snapshot().activeSheet).toBe(2);
    fireEvent.keyDown(active(view)!, { key: 'Escape' });
    expect(view.queryByRole('menu')).toBeNull();
    await waitFor(() => expect(active(view)).toBe(view.getAllByRole('tab')[2]!));
  });
});

describe('a menu action when the page had focus at the open', () => {
  test('gives focus to the grid', async () => {
    const { view, viewer } = await open();
    (active(view) as HTMLElement | null)?.blur();
    expect(active(view)).toBe(view.container.ownerDocument.body);
    fireEvent.contextMenu(view.getAllByRole('tab')[0]!, { clientX: 10 });
    fireEvent.click(view.getByRole('menuitem', { name: 'Move right' }));
    expect(names(view)).toEqual(['B', 'A', 'C']);
    expect(viewer.snapshot().activeSheet).toBe(1);
    expect(onGrid(view)).toBe(true);
  });
});

describe('the rename field', () => {
  test('ends before the add button', async () => {
    const { view } = await open();
    const add = view.getByRole('button', { name: 'Add a sheet' });
    const rect = (left: number, width: number) =>
      ({
        left,
        right: left + width,
        width,
        top: 0,
        bottom: 20,
        height: 20,
        x: left,
        y: 0,
      }) as DOMRect;
    add.getBoundingClientRect = () => rect(150, 28);
    view.getAllByRole('tab')[0]!.getBoundingClientRect = () => rect(10, 40);
    const field = typeName(view, 'A long name that would cover the add button');
    expect(field.style.left).toBe('10px');
    // The bar starts at 0: the field may grow from 10 to 2 pixels before the button at 150.
    expect(field.style.maxWidth).toBe('138px');
  });

  test('takes a long name, and the store refuses it with its reason', async () => {
    const { view, viewer } = await open();
    const field = typeName(view, 'x'.repeat(40));
    expect(field.hasAttribute('maxlength')).toBe(false);
    expect(field.value).toBe('x'.repeat(40));
    fireEvent.keyDown(field, { key: 'Enter' });
    expect(viewer.snapshot().refusal?.code).toBe('sheet-name-invalid');
    expect(names(view)).toEqual(['A', 'B', 'C']);
  });
});

describe('a chartsheet tab and the empty bar', () => {
  test('the tab is aria-disabled, out of keyboard focus, with no menu popup', async () => {
    const { view } = await open(withChartsheet());
    const chart = view.getAllByRole('tab')[1]!;
    expect(chart.hasAttribute('disabled')).toBe(false);
    expect(chart.getAttribute('aria-disabled')).toBe('true');
    expect(chart.hasAttribute('aria-haspopup')).toBe(false);
    expect(chart.tabIndex).toBe(-1);
    expect(view.getAllByRole('tab')[0]!.getAttribute('aria-haspopup')).toBe('menu');
  });

  test('a press keeps the grid focus, and the click chooses nothing and renames nothing', async () => {
    const { view, viewer } = await open(withChartsheet());
    viewer.focus();
    const chart = view.getAllByRole('tab')[1]!;
    expect(fireEvent.mouseDown(chart, { detail: 1 })).toBe(false);
    fireEvent.click(chart, { detail: 1 });
    fireEvent.doubleClick(chart);
    expect(viewer.snapshot().activeSheet).toBe(0);
    expect(view.queryByLabelText('Sheet name')).toBeNull();
    expect(onGrid(view)).toBe(true);
  });

  test('the chartsheet tab says why it is disabled, and the sheets take no rename or delete', async () => {
    const { view, viewer } = await open(withChartsheet());
    const chart = view.getAllByRole('tab')[1]!;
    const REASON =
      'This sheet is not a worksheet, or its part is missing. The editor cannot show, rename, delete, or move it.';
    expect(chart.getAttribute('title')).toBe(`Chart1: ${REASON}`);
    const why = chart.ownerDocument.getElementById(chart.getAttribute('aria-describedby') ?? '');
    expect(why?.textContent).toBe(REASON);
    // The chartsheet refuses a rename or a delete of any sheet: the menu says so, and F2 or a
    // double-click on a worksheet tab starts no rename field. The tab says why in its title
    // and its description.
    expect(viewer.snapshot().sheetEditRefusal?.code).toBe('structure-blocked');
    const tab = view.getAllByRole('tab')[0]!;
    fireEvent.keyDown(tab, { key: 'F2' });
    fireEvent.doubleClick(tab);
    expect(view.queryByLabelText('Sheet name')).toBeNull();
    expect(tab.getAttribute('title')).toStartWith(`${tab.textContent}: Sheets cannot be renamed`);
    const tabWhy = tab.ownerDocument.getElementById(tab.getAttribute('aria-describedby') ?? '');
    expect(tabWhy?.textContent).toStartWith('Sheets cannot be renamed or deleted:');
    expect(tabWhy?.closest('[role="tablist"]')).toBeNull();
    fireEvent.contextMenu(tab, { clientX: 10 });
    // The sheet menu's rows, not the menu bar's triggers.
    const [rename, del, , right] = view
      .getAllByRole('menuitem')
      .filter((m) => !m.closest('[role="menubar"]'));
    for (const row of [rename!, del!]) {
      expect(row.getAttribute('aria-disabled')).toBe('true');
      expect(row.getAttribute('title')).toStartWith(
        `${row.textContent}: Sheets cannot be renamed or deleted:`
      );
    }
    expect(right!.hasAttribute('aria-disabled')).toBe(false);
  });

  test('a press on the empty part of the bar keeps the grid focus', async () => {
    const { view } = await open();
    const bar = view.container.querySelector('.xlsx-sheet-bar') as HTMLElement;
    expect(fireEvent.mouseDown(bar, { detail: 1 })).toBe(false);
  });
});

describe('focus after a press on the add button', () => {
  test('goes to the new tab when the tab list had focus', async () => {
    const { view, viewer } = await open();
    view.getAllByRole('tab')[0]!.focus();
    press(view.getByRole('button', { name: 'Add a sheet' }));
    await waitFor(() => expect(names(view)).toHaveLength(4));
    const shown = viewer.snapshot().activeSheet;
    await waitFor(() => expect(active(view)?.getAttribute('data-sheet-index')).toBe(String(shown)));
  });

  test('stays on the grid otherwise', async () => {
    const { view, viewer } = await open();
    viewer.focus();
    press(view.getByRole('button', { name: 'Add a sheet' }));
    await waitFor(() => expect(names(view)).toHaveLength(4));
    expect(onGrid(view)).toBe(true);
  });
});

describe('the undo keys on the add button', () => {
  test('undo the added sheet, and redo adds it again', async () => {
    const { view, viewer } = await open();
    const add = view.getByRole('button', { name: 'Add a sheet' });
    fireEvent.click(add);
    await waitFor(() => expect(names(view)).toHaveLength(4));
    fireEvent.keyDown(add, { key: 'z', ctrlKey: true });
    await waitFor(() => expect(names(view)).toEqual(['A', 'B', 'C']));
    fireEvent.keyDown(add, { key: 'y', ctrlKey: true });
    await waitFor(() => expect(names(view)).toHaveLength(4));
    expect(viewer.snapshot().sheets).toHaveLength(4);
  });
});

describe('a rename and a refused cell edit', () => {
  test('no rename starts while a refused cell edit is open; its reason stays', async () => {
    const { view, viewer } = await open();
    expect(viewer.beginEdit('=1+').ok).toBe(true);
    expect(viewer.commitEdit().ok).toBe(false);
    const code = viewer.snapshot().refusal?.code;
    expect(code).toBe('formula-syntax');
    fireEvent.doubleClick(view.getAllByRole('tab')[0]!);
    expect(view.queryByLabelText('Sheet name')).toBeNull();
    expect(viewer.snapshot().refusal?.code).toBe(code);
  });

  test('a right press on another tab opens no menu; the edit and its reason stay', async () => {
    const { view, viewer } = await open();
    expect(viewer.beginEdit('=1+').ok).toBe(true);
    expect(viewer.commitEdit().ok).toBe(false);
    const tab = view.getAllByRole('tab')[2]!;
    fireEvent.mouseDown(tab, { button: 2, detail: 1 });
    fireEvent.contextMenu(tab, { clientX: 40 });
    expect(view.queryByRole('menu')).toBeNull();
    expect(viewer.snapshot().activeSheet).toBe(0);
    expect(viewer.snapshot().refusal?.code).toBe('formula-syntax');
  });

  test('Escape clears only a reason that the rename made', async () => {
    const { view, viewer } = await open();
    expect(viewer.moveSheet(0, 99).ok).toBe(false);
    expect(viewer.snapshot().refusal?.code).toBe('bad-position');
    const field = typeName(view, 'Other');
    fireEvent.keyDown(field, { key: 'Escape' });
    expect(viewer.snapshot().refusal?.code).toBe('bad-position');
    const again = typeName(view, 'b');
    fireEvent.keyDown(again, { key: 'Enter' });
    expect(viewer.snapshot().refusal?.code).toBe('sheet-name-taken');
    fireEvent.keyDown(again, { key: 'Escape' });
    expect(viewer.snapshot().refusal).toBeNull();
  });
});

describe('the rename field over a scrolled tab list', () => {
  test('stays inside the tab list when its tab scrolls out', async () => {
    const { view } = await open();
    const rect = (left: number, width: number) =>
      ({
        left,
        right: left + width,
        width,
        top: 0,
        bottom: 20,
        height: 20,
        x: left,
        y: 0,
      }) as DOMRect;
    view.getByRole('tablist').getBoundingClientRect = () => rect(20, 200);
    const tab = view.getAllByRole('tab')[0]!;
    let left = 40;
    tab.getBoundingClientRect = () => rect(left, 60);
    fireEvent.doubleClick(tab);
    expect((view.getByLabelText('Sheet name') as HTMLInputElement).style.left).toBe('40px');
    left = -100;
    fireEvent.scroll(view.getByRole('tablist'));
    expect((view.getByLabelText('Sheet name') as HTMLInputElement).style.left).toBe('20px');
    left = 500;
    fireEvent.scroll(view.getByRole('tablist'));
    // The list ends at 220: the field of width 60 starts at 160.
    expect((view.getByLabelText('Sheet name') as HTMLInputElement).style.left).toBe('160px');
  });
});

describe('an undo key on a tab after a move from the key menu', () => {
  test('leaves the focus on the selected tab', async () => {
    const { view } = await open();
    const first = view.getAllByRole('tab')[0]!;
    first.focus();
    fireEvent.keyDown(first, { key: 'F10', shiftKey: true });
    fireEvent.click(view.getByRole('menuitem', { name: 'Move right' }));
    await waitFor(() => expect(names(view)).toEqual(['B', 'A', 'C']));
    const selected = () =>
      view.getAllByRole('tab').find((t) => t.getAttribute('aria-selected') === 'true');
    expect(active(view)).toBe(selected()!);
    // The undo moves A back to the first position: the focus follows its tab.
    fireEvent.keyDown(active(view)!, { key: 'z', ctrlKey: true });
    await waitFor(() => expect(names(view)).toEqual(['A', 'B', 'C']));
    await waitFor(() => expect(active(view)).toBe(selected()!));
    expect(selected()!.textContent).toBe('A');
  });
});

describe('the tab of a newly shown sheet', () => {
  test('scrolls into view in the tab list only, never the host page', async () => {
    const { view, viewer } = await open();
    const list = view.getByRole('tablist');
    list.style.overflowX = 'auto';
    Object.defineProperty(list, 'scrollWidth', { get: () => 500 });
    Object.defineProperty(list, 'clientWidth', { get: () => 100 });
    const third = view.getAllByRole('tab')[2]!;
    third.getBoundingClientRect = () =>
      ({ left: 300, top: 0, width: 60, height: 20, right: 360, bottom: 20 }) as DOMRect;
    const moves: { left?: number }[] = [];
    list.scrollBy = ((o: { left?: number }) => void moves.push(o)) as typeof list.scrollBy;
    let pageScrolls = 0;
    const proto = Object.getPrototypeOf(third) as { scrollIntoView?: unknown };
    const before = proto.scrollIntoView;
    proto.scrollIntoView = () => void (pageScrolls += 1);
    try {
      expect(viewer.setActiveSheet(2)).toBe(true);
      await waitFor(() => expect(moves.length).toBe(1));
      // The least distance: the tab's right edge to the right edge of the list.
      expect(moves[0]!.left).toBe(260);
      expect(pageScrolls).toBe(0);
    } finally {
      proto.scrollIntoView = before;
    }
  });
});
