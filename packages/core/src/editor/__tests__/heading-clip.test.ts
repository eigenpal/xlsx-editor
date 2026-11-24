// The row and column headings of a sheet with frozen panes (`x:pane` with `@state="frozen"`):
// a heading of the scrolled pane that the pane shows in part is clipped to the band of its pane
// (`grid-surface.ts`), so it never paints over a heading of the frozen pane. The engine is
// driven directly, in happy-dom.

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

/** An editor on a sheet with row 1 and column A frozen, in a container of 800 by 397 pixels. */
async function frozenEditor() {
  const win = new Window();
  const doc = win.document as unknown as Document;
  const host = doc.createElement('div');
  doc.body.append(host);
  const editor = createXlsxEditor({ fonts: FONTS });
  editors.push(editor);
  editor.attach(host);
  await editor.load(
    buildXlsx({
      sheets: [
        {
          name: 'S',
          rows: row(1, [c('A1', 1)]),
          headings: true,
          viewChildren:
            '<pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="frozen"/>',
        },
      ],
    })
  );
  const scroller = host.querySelector<HTMLElement>('.xlsx-scroller')!;
  Object.defineProperty(scroller, 'clientWidth', { get: () => 800 });
  // Not a whole number of rows, so the scrolled pane shows its first row in part at the end.
  Object.defineProperty(scroller, 'clientHeight', { get: () => 397 });
  flushEditor(editor);
  return { editor, host };
}

const px = (v: string) => parseFloat(v);

/** The top inset of a `clip-path: inset(...)`, and its left inset. */
function insets(el: HTMLElement): { top: number; left: number } {
  const m = /inset\(([^)]*)\)/.exec(el.style.clipPath);
  if (!m) return { top: 0, left: 0 };
  const v = m[1]!.trim().split(/\s+/).map(px);
  const [top = 0, right = top, , left = right] = v;
  return { top, left };
}

test('a row heading of the scrolled pane does not show over the frozen row heading', async () => {
  const { editor, host } = await frozenEditor();
  editor.select({ row: 1048575, col: 1 });
  flushEditor(editor);
  const heads = [...host.querySelectorAll<HTMLElement>('.xlsx-row-heads .xlsx-head')];
  const frozen = heads.find((h) => h.dataset.row === '0')!;
  const frozenEnd = px(frozen.style.top) + px(frozen.style.height);
  for (const h of heads) {
    if (h === frozen) continue;
    // The part of each scrolled heading that shows starts at or below the frozen heading.
    expect(px(h.style.top) + insets(h).top).toBeGreaterThanOrEqual(frozenEnd - 0.01);
  }
});

test('a column heading of the scrolled pane does not show over the frozen column heading', async () => {
  const { editor, host } = await frozenEditor();
  editor.select({ row: 1, col: 16383 });
  flushEditor(editor);
  const heads = [...host.querySelectorAll<HTMLElement>('.xlsx-column-heads .xlsx-head')];
  const frozen = heads.find((h) => h.dataset.col === '0')!;
  const frozenEnd = px(frozen.style.left) + px(frozen.style.width);
  for (const h of heads) {
    if (h === frozen) continue;
    expect(px(h.style.left) + insets(h).left).toBeGreaterThanOrEqual(frozenEnd - 0.01);
  }
});

test('a heading that shows whole has no clip', async () => {
  const { host } = await frozenEditor();
  // At the opening scroll, the first rows and columns show whole.
  for (const n of ['0', '1', '2', '3']) {
    const r = host.querySelector<HTMLElement>(`.xlsx-row-heads [data-row="${n}"]`)!;
    const col = host.querySelector<HTMLElement>(`.xlsx-column-heads [data-col="${n}"]`)!;
    expect([r.style.clipPath, col.style.clipPath]).toEqual(['', '']);
  }
});
