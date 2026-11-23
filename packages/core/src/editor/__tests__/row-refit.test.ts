// A typed value fits its row again, in a happy-dom window of its own: a row with a fitted height
// (`@ht` without `@customHeight`) grows to the lines of a longer wrapped text, and a row with a
// height set by hand (`customHeight="1"`) keeps its height (ECMA-376 Part 1, 18.3.1.73).

import { afterEach, describe, expect, test } from 'bun:test';
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

let viewers: XlsxEditorInstance[] = [];
afterEach(() => {
  for (const v of viewers) v.destroy();
  viewers = [];
});

const WRAP =
  '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment wrapText="1"/></xf>';

async function mount() {
  const win = new Window();
  const doc = win.document as unknown as Document;
  const host = doc.createElement('div');
  doc.body.appendChild(host);
  const viewer = createXlsxEditor({ fonts: FONTS });
  viewers.push(viewer);
  viewer.attach(host);
  const scroller = host.querySelector('.xlsx-scroller') as HTMLElement;
  Object.defineProperty(scroller, 'clientWidth', { get: () => 800 });
  Object.defineProperty(scroller, 'clientHeight', { get: () => 400 });
  await viewer.load(
    buildXlsx({
      styles: { cellXfs: [WRAP] },
      sheets: [
        {
          name: 'S',
          headings: true,
          rows:
            row(1, [c('A1', 'one', { s: 1 })], 'ht="15"') +
            row(2, [c('A2', 'two', { s: 1 })], 'ht="15" customHeight="1"'),
        },
      ],
    })
  );
  flushEditor(viewer);
  return { viewer, host };
}

/** The height in pixels of the heading of a 0-based row. */
const rowHeight = (host: HTMLElement, index: number) => {
  const el = host.querySelector(`.xlsx-row-heads [data-row="${index}"]`) as HTMLElement;
  return Number.parseFloat(el.style.height);
};

const LONG = 'a long text that wraps over several lines in this column';

describe('a typed value in a row with a fitted height', () => {
  test('grows the row to its wrapped lines', async () => {
    const { viewer, host } = await mount();
    const before = rowHeight(host, 0);
    viewer.select({ row: 0, col: 0 });
    expect(viewer.beginEdit(LONG).ok).toBe(true);
    expect(viewer.commitEdit().ok).toBe(true);
    flushEditor(viewer);
    expect(rowHeight(host, 0)).toBeGreaterThan(before);
  });

  test('keeps a height set by hand', async () => {
    const { viewer, host } = await mount();
    const before = rowHeight(host, 1);
    viewer.select({ row: 1, col: 0 });
    expect(viewer.beginEdit(LONG).ok).toBe(true);
    expect(viewer.commitEdit().ok).toBe(true);
    flushEditor(viewer);
    expect(rowHeight(host, 1)).toBe(before);
  });
});
