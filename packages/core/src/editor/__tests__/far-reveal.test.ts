// A jump to a far cell scrolls all the way (`extentReaching` in core's `scroll-extent.ts`). The
// scroll container cuts a scroll offset to its extent, as a browser does, and the scroll event
// that follows reaches the editor before its next frame. Without the fix, the reveal set an
// offset past the old extent, the container cut it, and the view took the cut offset back, so
// the jump stopped short of the cell. On a sheet past `MAX_EXTENT_PX`, the container scrolls in
// steps, and its last offset must reach the last row (`surfaceExtent`). The engine is driven
// directly, in happy-dom.

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

const WIDTH = 800;

test('a jump to a far cell grows the extent before it scrolls, so the scroll is not cut short', async () => {
  const win = new Window();
  const doc = win.document as unknown as Document;
  const host = doc.createElement('div');
  doc.body.append(host);
  const editor = createXlsxEditor({ fonts: FONTS });
  editors.push(editor);
  editor.attach(host);
  await editor.load(buildXlsx({ sheets: [{ name: 'S', rows: row(1, [c('A1', 1)]) }] }));
  const scroller = host.querySelector<HTMLElement>('.xlsx-scroller')!;
  const extent = host.querySelector<HTMLElement>('.xlsx-extent')!;
  Object.defineProperty(scroller, 'clientWidth', { get: () => WIDTH });
  Object.defineProperty(scroller, 'clientHeight', { get: () => 400 });
  // A browser cuts the scroll offset to the extent that the container holds now.
  let left = 0;
  const room = () => Math.max(0, parseFloat(extent.style.width || '0') - WIDTH);
  Object.defineProperty(scroller, 'scrollLeft', {
    get: () => left,
    set: (value: number) => {
      left = Math.max(0, Math.min(value, room()));
    },
  });
  flushEditor(editor);
  const before = room();
  // AZ is past the extent before the jump.
  expect(before).toBeLessThan(2000);

  editor.select({ row: 0, col: 51 });
  // The scroll event of the container arrives before the next frame paints.
  scroller.dispatchEvent(new win.Event('scroll') as unknown as Event);
  flushEditor(editor);

  expect(editor.snapshot().selectionRef).toBe('AZ1');
  // The offset went past the old extent, to the scroll that the reveal asked for: the view kept
  // it, and the extent grew to hold it.
  expect(left).toBeGreaterThan(before);
  expect(parseFloat(extent.style.width)).toBeGreaterThanOrEqual(left + WIDTH);
});

test('a jump to the last row of the sheet scrolls to it, also where the container scrolls in steps', async () => {
  const win = new Window();
  const doc = win.document as unknown as Document;
  const host = doc.createElement('div');
  doc.body.append(host);
  const editor = createXlsxEditor({ fonts: FONTS });
  editors.push(editor);
  editor.attach(host);
  await editor.load(buildXlsx({ sheets: [{ name: 'S', rows: row(1, [c('A1', 1)]) }] }));
  const scroller = host.querySelector<HTMLElement>('.xlsx-scroller')!;
  const extent = host.querySelector<HTMLElement>('.xlsx-extent')!;
  const HEIGHT = 400;
  Object.defineProperty(scroller, 'clientWidth', { get: () => WIDTH });
  Object.defineProperty(scroller, 'clientHeight', { get: () => HEIGHT });
  // A browser cuts the scroll offset to the extent that the container holds now.
  let top = 0;
  const room = () => Math.max(0, parseFloat(extent.style.height || '0') - HEIGHT);
  Object.defineProperty(scroller, 'scrollTop', {
    get: () => top,
    set: (value: number) => {
      top = Math.max(0, Math.min(value, room()));
    },
  });
  flushEditor(editor);

  editor.select({ row: 1048575, col: 3 });
  // The extent of a sheet of 1048576 rows passes `MAX_EXTENT_PX`, so it scrolls in steps.
  expect(parseFloat(extent.style.height)).toBeLessThanOrEqual(15_000_000 + HEIGHT);
  // The scroll event of the container arrives before the next frame paints.
  scroller.dispatchEvent(new win.Event('scroll') as unknown as Event);
  flushEditor(editor);

  expect(editor.snapshot().selectionRef).toBe('D1048576');
  // The container is at its end, and the window that it paints holds the last row: the active
  // cell box lies inside the grid area.
  expect(top).toBe(room());
  const active = host.querySelector<HTMLElement>('.xlsx-active-cell')!;
  const overlay = host.querySelector<HTMLElement>('.xlsx-active-cell')!.parentElement!;
  const y = parseFloat(active.style.top);
  expect(y).toBeGreaterThan(0);
  expect(y + parseFloat(active.style.height)).toBeLessThanOrEqual(
    parseFloat(overlay.style.height) + 0.5
  );
});
