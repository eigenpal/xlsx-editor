// The break opportunities of very long texts stay in the shared layout cache only while their
// workbook is open: a close, a destroy, and the load of another workbook forget them.

import { afterEach, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { Window } from 'happy-dom';
import { buildXlsx } from '../../../../../scripts/fixtures/builder.ts';
import type { FontSource } from '../../layout/index.ts';
import { breaksOf, keptLongTexts } from '../../layout/cell-text.ts';
import { createXlsxEditor, type XlsxEditorInstance } from '../xlsx-editor.ts';

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

const file = () =>
  buildXlsx({ sheets: [{ name: 'Data', rows: '<row r="1"><c r="A1"><v>1</v></c></row>' }] });

/** An open editor, with one long text in the cache. */
async function withLongText(): Promise<XlsxEditorInstance> {
  const doc = new Window().document as unknown as Document;
  const host = doc.createElement('div');
  doc.body.appendChild(host);
  const viewer = createXlsxEditor({ fonts: FONTS });
  viewers.push(viewer);
  viewer.attach(host);
  await viewer.load(file());
  breaksOf('word '.repeat(16 * 1024));
  expect(keptLongTexts()).toBe(1);
  return viewer;
}

describe('the cache of long texts', () => {
  test('a close forgets the long texts', async () => {
    const viewer = await withLongText();
    viewer.close();
    expect(keptLongTexts()).toBe(0);
  });

  test('a destroy forgets the long texts', async () => {
    const viewer = await withLongText();
    viewer.destroy();
    expect(keptLongTexts()).toBe(0);
  });

  test('the load of another workbook forgets the long texts', async () => {
    const viewer = await withLongText();
    await viewer.load(file());
    expect(keptLongTexts()).toBe(0);
  });
});
