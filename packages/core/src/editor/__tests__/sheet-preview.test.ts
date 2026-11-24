// `renderPreview`: one hidden canvas of the window at scroll 0, sized by the element or the
// options, reused on the next paint, and false with no change when nothing can be painted.

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

let editors: XlsxEditorInstance[] = [];
afterEach(() => {
  for (const e of editors) e.destroy();
  editors = [];
});

function setup() {
  const doc = new Window().document as unknown as Document;
  const host = doc.createElement('div');
  const card = doc.createElement('div');
  doc.body.append(host, card);
  const editor = createXlsxEditor({ fonts: FONTS, translate: (k) => k });
  editors.push(editor);
  editor.attach(host);
  return { editor, card };
}

const workbook = () =>
  buildXlsx({
    sheets: [
      { name: 'S', rows: row(1, [c('A1', 1), c('B1', 2)]) },
      { name: 'T', rows: row(1, [c('A1', 9)]) },
    ],
  });

test('returns false and changes nothing before a workbook is ready', () => {
  const { editor, card } = setup();
  const keep = card.ownerDocument.createElement('span');
  card.append(keep);
  expect(editor.renderPreview(card, { width: 200, height: 100 })).toBe(false);
  expect(card.firstElementChild).toBe(keep);
});

test('paints one hidden canvas at the given size, and reuses it', async () => {
  const { editor, card } = setup();
  await editor.load(workbook());
  flushEditor(editor);
  expect(editor.renderPreview(card, { width: 200, height: 100 })).toBe(true);
  const canvas = card.querySelector('canvas')!;
  expect(card.childElementCount).toBe(1);
  expect(canvas.getAttribute('aria-hidden')).toBe('true');
  expect(canvas.style.width).toBe('200px');
  expect(canvas.style.height).toBe('100px');
  expect(editor.renderPreview(card, { sheet: 1, width: 120, height: 80, zoom: 1 })).toBe(true);
  expect(card.querySelector('canvas')).toBe(canvas);
  expect(canvas.style.width).toBe('120px');
});

test('refuses a sheet that does not exist, and an element with no size', async () => {
  const { editor, card } = setup();
  await editor.load(workbook());
  flushEditor(editor);
  for (const sheet of [2, -1, 0.5, Number.NaN])
    expect(editor.renderPreview(card, { sheet, width: 200, height: 100 })).toBe(false);
  // happy-dom lays nothing out, so the element itself is 0 by 0.
  expect(editor.renderPreview(card)).toBe(false);
  expect(card.childElementCount).toBe(0);
});

test('clamps a host size and zoom to the limits', async () => {
  const { editor, card } = setup();
  await editor.load(workbook());
  flushEditor(editor);
  expect(editor.renderPreview(card, { width: 1e9, height: 50, zoom: 99 })).toBe(true);
  expect(card.querySelector('canvas')!.style.width).toBe('2048px');
});

test('paints again when the number locale changes, and drops a preview that left the page', async () => {
  const { editor, card } = setup();
  await editor.load(workbook());
  flushEditor(editor);
  expect(editor.renderPreview(card, { width: 200, height: 100 })).toBe(true);
  const marker = card.ownerDocument.createElement('span');
  card.replaceChildren(marker);
  editor.setNumberLocale({ ...editor.snapshot().numberLocale, decimal: ',', group: '.' });
  expect(card.querySelector('canvas')).not.toBeNull();
  card.remove();
  card.replaceChildren(marker);
  editor.setNumberLocale(undefined);
  expect(card.firstElementChild).toBe(marker);
});
