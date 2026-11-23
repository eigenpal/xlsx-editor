// A jump to a far cell scrolls all the way (`extentReaching` in core's `scroll-extent.ts`). The
// scroll container cuts a scroll offset to its extent, as a browser does, and the scroll event
// that follows reaches the editor before its next frame. Without the fix, the reveal set an
// offset past the old extent, the container cut it, and the view took the cut offset back, so
// the jump stopped short of the cell. The engine is driven directly, in happy-dom.

import { afterEach, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { buildXlsx, c, row } from '../../../../scripts/fixtures/builder.ts';
import {
  createXlsxEditor,
  flushEditor,
  type XlsxEditorInstance,
} from '../../../core/src/editor/xlsx-editor.ts';
import type { FontSource } from '../index.ts';

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

const editors: XlsxEditorInstance[] = [];
afterEach(() => {
  for (const editor of editors.splice(0)) editor.destroy();
  document.body.replaceChildren();
});

const WIDTH = 800;

test('a jump to a far cell grows the extent before it scrolls, so the scroll is not cut short', async () => {
  const host = document.createElement('div');
  document.body.append(host);
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
  scroller.dispatchEvent(new Event('scroll'));
  flushEditor(editor);

  expect(editor.snapshot().selectionRef).toBe('AZ1');
  // The offset went past the old extent, to the scroll that the reveal asked for: the view kept
  // it, and the extent grew to hold it.
  expect(left).toBeGreaterThan(before);
  expect(parseFloat(extent.style.width)).toBeGreaterThanOrEqual(left + WIDTH);
});
