/** @jsxImportSource react */
// The clipboard events of the mounted editor in happy-dom (`bun run test:dom` preloads it):
// copy writes plain text and the private token to the clipboard data, paste of that data
// pastes the copy with its cell format, and a paste of HTML only is refused.

import { afterEach, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { act, cleanup, fireEvent, render, waitFor } from '@testing-library/react';
import { buildXlsx, c, row } from '../../../../scripts/fixtures/builder.ts';
import {
  CLIPBOARD_TYPE,
  XlsxEditor,
  type CopyResult,
  type FontSource,
  type PasteData,
  type XlsxEditorInstance,
} from '../index.ts';

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
  sheets: [{ name: 'S', rows: row(1, [c('A1', 'one'), c('B1', 2, { s: 1 })]) }],
  styles: {
    cellXfs: [
      '<xf numFmtId="10" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>',
    ],
  },
});

/** Clipboard data that keeps what a copy sets. */
function fakeClipboard(initial: Record<string, string> = {}) {
  const data = new Map(Object.entries(initial));
  return {
    data,
    clipboardData: {
      setData: (type: string, value: string) => void data.set(type, value),
      getData: (type: string) => data.get(type) ?? '',
      get types() {
        return [...data.keys()];
      },
    },
  };
}

async function open() {
  let viewer: XlsxEditorInstance | null = null;
  const view = render(
    <div style={{ height: 400 }}>
      <XlsxEditor document={workbook} fonts={FONTS} onReady={(v) => (viewer = v)} />
    </div>
  );
  await waitFor(() => expect(viewer?.snapshot().status).toBe('ready'), { timeout: 5000 });
  const grid = view.container.querySelector('.xlsx-scroller') as HTMLElement;
  // The grid acts on clipboard events only while it has the focus.
  grid.focus();
  return { viewer: viewer!, grid };
}

describe('clipboard events', () => {
  test('the adapter exports the clipboard type and the copy and paste shapes', () => {
    expect(CLIPBOARD_TYPE).toBe('application/x-xlsx-editor');
    const data: PasteData = { text: 'x' };
    const copy: CopyResult = { ok: true, content: { text: 'x', token: 't' } };
    expect(copy.ok && copy.content.text).toBe(data.text!);
  });

  test('copy then paste on the grid pastes the copy with its cell format', async () => {
    const { viewer, grid } = await open();
    act(() => viewer.select({ row: 0, col: 0 }, { row: 0, col: 1 }));
    const clip = fakeClipboard();
    expect(fireEvent.copy(grid, { clipboardData: clip.clipboardData })).toBe(false);
    expect(clip.data.get('text/plain')).toBe('one\t200.00%');
    expect(clip.data.get('application/x-xlsx-editor')).toBeTruthy();
    act(() => viewer.select({ row: 2, col: 0 }));
    expect(fireEvent.paste(grid, { clipboardData: clip.clipboardData })).toBe(false);
    act(() => viewer.select({ row: 2, col: 1 }));
    expect(viewer.snapshot().activeCell?.content).toBe('200%');
    expect(viewer.snapshot().activeFormat?.numberFormat.id).toBe(10);
    act(() => void expect(viewer.undo().ok).toBe(true));
    expect(viewer.snapshot().activeCell?.content).toBe('');
  });

  test('a paste of HTML only is refused with paste-html', async () => {
    const { viewer, grid } = await open();
    const clip = fakeClipboard({ 'text/html': '<b>x</b>' });
    expect(fireEvent.paste(grid, { clipboardData: clip.clipboardData })).toBe(false);
    expect(viewer.snapshot().refusal?.code).toBe('paste-html');
  });
});
