/** @jsxImportSource react */
// `XlsxEditor.SheetPreview`: paints a canvas once the workbook is ready, paints the edit after an
// edit, names itself with `label`, and is hidden from assistive technology without one.

import { afterAll, afterEach, beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { act, cleanup, render, waitFor } from '@testing-library/react';
import { buildXlsx, c, row } from '../../../../scripts/fixtures/builder.ts';
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

// happy-dom lays nothing out: give the preview a size of 200 by 100.
const proto = HTMLElement.prototype;
const sizes = ['clientWidth', 'clientHeight'] as const;
const saved = sizes.map((k) => Object.getOwnPropertyDescriptor(Element.prototype, k));
beforeAll(() => {
  sizes.forEach((k, i) =>
    Object.defineProperty(proto, k, {
      configurable: true,
      get(this: HTMLElement) {
        if (this.classList.contains('xlsx-sheet-preview')) return i === 0 ? 200 : 100;
        return saved[i]?.get?.call(this) ?? 0;
      },
    })
  );
});
afterAll(() => sizes.forEach((k) => delete (proto as unknown as Record<string, unknown>)[k]));
afterEach(cleanup);

const workbook = buildXlsx({ sheets: [{ name: 'S', rows: row(1, [c('A1', 1)]) }] });

test('paints after load and after an edit, with the label as its name', async () => {
  let editor: XlsxEditorInstance | null = null;
  const view = render(
    <XlsxEditor.Root document={workbook} fonts={FONTS} onReady={(e) => (editor = e)}>
      <XlsxEditor.SheetPreview label="Sheet S" className="card" />
      <XlsxEditor.Viewport>
        <XlsxEditor.SheetTabs />
      </XlsxEditor.Viewport>
    </XlsxEditor.Root>
  );
  const preview = view.getByRole('img', { name: 'Sheet S' });
  expect(preview.className).toBe('xlsx-sheet-preview card');
  await waitFor(() => expect(preview.querySelector('canvas')).not.toBeNull(), { timeout: 5000 });
  const calls: unknown[] = [];
  const paint = editor!.renderPreview.bind(editor!);
  editor!.renderPreview = (el, options) => (calls.push(options), paint(el, options));
  await act(async () => {
    editor!.selectReference('B2');
    editor!.enterValue('7');
  });
  await waitFor(() => expect(calls.length).toBeGreaterThan(0));
});

test('is hidden without a label, and empty before a workbook', () => {
  const view = render(
    <XlsxEditor.Root fonts={FONTS}>
      <XlsxEditor.SheetPreview sheet={0} zoom={1} />
    </XlsxEditor.Root>
  );
  const preview = view.container.querySelector('.xlsx-sheet-preview')!;
  expect(preview.getAttribute('aria-hidden')).toBe('true');
  expect(preview.childElementCount).toBe(0);
});
