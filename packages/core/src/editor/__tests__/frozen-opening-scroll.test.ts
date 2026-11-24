// A frozen sheet opens at the scroll position that the file stores: the engine takes
// `x:pane/@topLeftCell` (`CT_Pane`, ECMA-376 Part 1, 18.3.1.66) through `sheetOpening`, the
// opening view, and the first layout. The engine is driven directly, in happy-dom.

import { afterEach, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { Window } from 'happy-dom';
import { buildXlsx, c, row } from '../../../../../scripts/fixtures/builder.ts';
import { createXlsxEditor, type XlsxEditorInstance } from '../index.ts';
import { flushEditor } from './flush-editor.ts';
import type { FontSource } from '../../layout/index.ts';

const fontsDir = join(
  dirname(Bun.resolveSync('@docx-editor.dev/fonts/package.json', import.meta.dir)),
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

const editors: XlsxEditorInstance[] = [];
afterEach(() => {
  for (const editor of editors.splice(0)) editor.destroy();
});

/** The first shown column and row headings of a sheet with the pane `pane`. */
async function openingHeadings(pane: string) {
  const win = new Window();
  const doc = win.document as unknown as Document;
  const host = doc.createElement('div');
  doc.body.append(host);
  const editor = createXlsxEditor({ fonts: FONTS });
  editors.push(editor);
  editor.attach(host);
  await editor.load(
    buildXlsx({
      sheets: [{ name: 'S', rows: row(1, [c('A1', 1)]), headings: true, viewChildren: pane }],
    })
  );
  const scroller = host.querySelector<HTMLElement>('.xlsx-scroller')!;
  Object.defineProperty(scroller, 'clientWidth', { get: () => 800 });
  Object.defineProperty(scroller, 'clientHeight', { get: () => 400 });
  flushEditor(editor);
  flushEditor(editor);
  const heads = (sel: string, key: 'col' | 'row') =>
    [...host.querySelectorAll<HTMLElement>(`${sel} .xlsx-head`)]
      .map((h) => Number(h.dataset[key]))
      .sort((a, b) => a - b)
      .slice(0, 3);
  return {
    cols: heads('.xlsx-column-heads', 'col'),
    rows: heads('.xlsx-row-heads', 'row'),
    active: editor.snapshot().selectionRef,
  };
}

test('frozen columns open with the stored top-left column next to them', async () => {
  const { cols, rows } = await openingHeadings(
    '<pane xSplit="2" topLeftCell="E1" activePane="topRight" state="frozen"/>'
  );
  expect(cols).toEqual([0, 1, 4]);
  expect(rows).toEqual([0, 1, 2]);
});

test('frozen rows and columns open with the stored top-left cell next to them', async () => {
  const { cols, rows, active } = await openingHeadings(
    '<pane xSplit="1" ySplit="1" topLeftCell="D9" activePane="bottomRight" state="frozen"/>' +
      '<selection pane="topRight" activeCell="C1" sqref="C1"/>' +
      '<selection pane="bottomRight" activeCell="E10" sqref="E10"/>'
  );
  expect(cols).toEqual([0, 3, 4]);
  expect(rows).toEqual([0, 8, 9]);
  // The active cell is the one of the active pane.
  expect(active).toBe('E10');
});
