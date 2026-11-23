// The expiry rule of an internal paste (design D2 of the change `workbook-structure-editing`):
// the snapshot is used only while the sheets that a copied formula names and the copied cell
// formats are unchanged; an expired token stays refused. Also: a copied sheet range that reads
// two ways, the target range of `grid-full`, the bounded scan of pasted text, and the signed
// number rule of the text export, copied line breaks, a pending cut after an undo, Escape that
// forgets a pending cut, the announcements of a copy and a cut, and the announcement of a
// `locked` refusal by another rule.

import { afterEach, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { Window } from 'happy-dom';
import {
  buildXlsx,
  c,
  row,
  type FixtureWorkbook,
} from '../../../../../scripts/fixtures/builder.ts';
import { createXlsxEditor, type XlsxEditorInstance } from '../index.ts';
import { flushEditor } from './flush-editor.ts';
import { exportCellText, signedNumberText, splitPlainText } from '../clipboard.ts';
import { MAX_COLS } from '../../store/index.ts';
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

async function mount(wb: FixtureWorkbook) {
  const win = new Window();
  const doc = win.document as unknown as Document;
  const host = doc.createElement('div');
  doc.body.appendChild(host);
  const translate = (key: string, params?: Readonly<Record<string, string | number>>) =>
    params?.ref !== undefined ? `${key} [${params.ref}]` : key;
  const viewer = createXlsxEditor({ fonts: FONTS, translate });
  viewers.push(viewer);
  viewer.attach(host);
  const scroller = host.querySelector('.xlsx-scroller') as HTMLElement;
  Object.defineProperty(scroller, 'clientWidth', { get: () => 800 });
  Object.defineProperty(scroller, 'clientHeight', { get: () => 400 });
  await viewer.load(buildXlsx(wb));
  flushEditor(viewer);
  hosts.set(viewer, host);
  return viewer;
}

const hosts = new WeakMap<XlsxEditorInstance, HTMLElement>();
const liveText = (v: XlsxEditorInstance) => hosts.get(v)?.querySelector('.xlsx-live')?.textContent;

function content(v: XlsxEditorInstance, r: number, col: number) {
  v.select({ row: r, col });
  return v.snapshot().activeCell?.content;
}

const code = (r: { ok: boolean; refusal?: { code: string } }) => (r.ok ? null : r.refusal?.code);

function copyAt(v: XlsxEditorInstance, r: number, col: number) {
  v.select({ row: r, col });
  const out = v.copy();
  if (!out.ok) throw new Error(out.refusal.code);
  return out.content;
}

describe('the sheets that a copy depends on', () => {
  const book = () =>
    mount({
      sheets: [
        {
          name: 'S',
          rows: row(1, [c('A1', 1), c('B1', 2, { f: 'A1+1' }), c('C1', 3, { f: 'O!A1' })]),
        },
        { name: 'O', rows: row(1, [c('A1', 1)]) },
        { name: 'U', rows: '' },
      ],
    });

  test('a rename of a sheet that no copied formula names keeps the copy', async () => {
    const v = await book();
    const data = copyAt(v, 0, 1);
    expect(v.renameSheet(2, 'Renamed').ok).toBe(true);
    expect(v.renameSheet(1, 'O2').ok).toBe(true);
    v.select({ row: 1, col: 1 });
    expect(v.paste({ token: data.token }).ok).toBe(true);
    expect(content(v, 1, 1)).toBe('=A2+1');
  });

  test('a rename of the sheet that a copied formula names refuses the paste', async () => {
    const v = await book();
    const data = copyAt(v, 0, 2);
    expect(v.renameSheet(2, 'Renamed').ok).toBe(true);
    v.select({ row: 1, col: 2 });
    expect(v.paste({ token: data.token }).ok).toBe(true);
    expect(v.renameSheet(1, 'O2').ok).toBe(true);
    v.select({ row: 2, col: 2 });
    const r = v.paste({ token: data.token });
    expect(code(r)).toBe('copy-expired');
    expect(r.ok ? null : r.refusal.subject).toBe('S!C1');
    expect(content(v, 2, 2)).toBe('');
  });

  test('a new sheet with a name that a copied formula names refuses the paste', async () => {
    const v = await mount({
      sheets: [{ name: 'S', rows: row(1, [c('A1', 0, { f: 'IFERROR(Later!A1,0)' })]) }],
    });
    const data = copyAt(v, 0, 0);
    expect(v.addSheet('Later').ok).toBe(true);
    expect(v.setActiveSheet(0)).toBe(true);
    v.select({ row: 1, col: 0 });
    expect(code(v.paste({ token: data.token }))).toBe('copy-expired');
  });

  test('a [0] prefix names a sheet of this workbook, so a new sheet with its name expires it', async () => {
    const v = await mount({
      sheets: [{ name: 'S', rows: row(1, [c('A1', 0, { f: 'IFERROR([0]Later!A1,0)' })]) }],
    });
    const data = copyAt(v, 0, 0);
    expect(v.addSheet('Later').ok).toBe(true);
    expect(v.setActiveSheet(0)).toBe(true);
    v.select({ row: 1, col: 0 });
    expect(code(v.paste({ token: data.token }))).toBe('copy-expired');
  });
});

describe('the cell formats that a copy depends on', () => {
  const plain = () => mount({ sheets: [{ name: 'S', rows: row(1, [c('A1', 1), c('B1', 2)]) }] });

  test('an undo and another format edit refuse the paste of the copied format', async () => {
    const v = await plain();
    v.select({ row: 0, col: 0 });
    expect(v.toggleFormat('bold').ok).toBe(true);
    const data = copyAt(v, 0, 0);
    expect(v.undo().ok).toBe(true);
    v.select({ row: 0, col: 1 });
    expect(v.toggleFormat('italic').ok).toBe(true);
    v.select({ row: 0, col: 2 });
    const r = v.paste({ token: data.token, text: data.text });
    expect(code(r)).toBe('copy-expired');
    expect(content(v, 0, 2)).toBe('');
    // The same token stays refused.
    expect(code(v.paste({ token: data.token, text: data.text }))).toBe('copy-expired');
  });

  test('a format edit of other cells keeps the copy and its format', async () => {
    const v = await plain();
    v.select({ row: 0, col: 0 });
    expect(v.toggleFormat('bold').ok).toBe(true);
    const data = copyAt(v, 0, 0);
    v.select({ row: 0, col: 1 });
    expect(v.toggleFormat('italic').ok).toBe(true);
    v.select({ row: 2, col: 0 });
    expect(v.paste({ token: data.token }).ok).toBe(true);
    v.select({ row: 2, col: 0 });
    expect(v.snapshot().activeFormat).toMatchObject({ bold: true, italic: false });
  });

  test('a later format edit of the source keeps the copy, and the paste has the copied format', async () => {
    // The edit appends a new cell format and leaves the copied index as it was.
    const v = await plain();
    v.select({ row: 0, col: 0 });
    expect(v.toggleFormat('bold').ok).toBe(true);
    const data = copyAt(v, 0, 0);
    v.select({ row: 0, col: 0 });
    expect(v.toggleFormat('italic').ok).toBe(true);
    v.select({ row: 2, col: 0 });
    expect(v.paste({ token: data.token }).ok).toBe(true);
    v.select({ row: 2, col: 0 });
    expect(v.snapshot().activeFormat).toMatchObject({ bold: true, italic: false });
  });
});

describe('a copied sheet range that reads two ways', () => {
  test('a copy of A1:S!B2 refuses the paste with paste-unreadable-formula', async () => {
    const v = await mount({
      sheets: [{ name: 'S', rows: row(1, [c('A1', 1), c('C1', 2, { f: 'SUM(A1:S!B2)' })]) }],
    });
    const data = copyAt(v, 0, 2);
    v.select({ row: 3, col: 3 });
    const r = v.paste({ token: data.token });
    expect(code(r)).toBe('paste-unreadable-formula');
    expect(r.ok ? null : r.refusal.subject).toBe('S!C1');
    expect(content(v, 3, 3)).toBe('');
  });
});

describe('a paste past the last column', () => {
  test('grid-full names the target range, not the source', async () => {
    const v = await mount({
      sheets: [
        { name: 'S', rows: [row(1, [c('A1', 1), c('B1', 2)]), row(2, [c('A2', 3)])].join('') },
      ],
    });
    v.select({ row: 0, col: 0 }, { row: 1, col: 1 });
    const out = v.copy();
    if (!out.ok) throw new Error(out.refusal.code);
    v.select({ row: 0, col: MAX_COLS - 1 });
    const r = v.paste({ token: out.content.token });
    expect(code(r)).toBe('grid-full');
    expect(r.ok ? null : r.refusal.subject).toBe('S!XFD1:XFD2');
    const t = v.paste({ text: 'a\tb' });
    expect(t.ok ? null : t.refusal.subject).toBe('S!XFD1');
  });
});

describe('the bounded scan of pasted text', () => {
  test('a text with more breaks and tabs than the limit is refused before any split', async () => {
    const big = 'a\t'.repeat(2_000_000);
    expect(splitPlainText(big)).toBeNull();
    expect(splitPlainText('a\tb\nc', 3)).toEqual([['a', 'b'], ['c']]);
    expect(splitPlainText('a\tb\nc\td', 2)).toBeNull();
    const v = await mount({ sheets: [{ name: 'S', rows: '' }] });
    expect(code(v.paste({ text: big }))).toBe('range-too-large');
  });
});

describe('the signed number rule of the text export', () => {
  test('only one signed number with a digit after its sign keeps its sign', () => {
    for (const text of [
      '-1,234.50',
      '-1.2E+03',
      '-.5',
      '+1',
      '-1:30',
      '-1 1/2',
      '-5%',
      '-$1,234.50',
      '-5,50 €',
      '-¥5',
    ])
      expect(signedNumberText(text)).toBe(true);
    for (const text of [
      '-1+2',
      '-E5',
      '-',
      '-1-2',
      '-1E+3+4',
      '-(1)',
      '-$A$1',
      '-$1:1',
      '-$',
      '-5%€',
    ])
      expect(signedNumberText(text)).toBe(false);
    expect(exportCellText('-1+2', true)).toBe("'-1+2");
    expect(exportCellText('-1.2E+03', true)).toBe('-1.2E+03');
    expect(exportCellText('-$1,234.50', true)).toBe('-$1,234.50');
    expect(exportCellText('-$1,234.50', false)).toBe("'-$1,234.50");
  });
});

describe('copied line breaks (approximation copied-line-breaks)', () => {
  test('a tab or a line break inside a cell becomes a space in the copied text', async () => {
    expect(exportCellText('a\tb\r\nc\nd', false)).toBe('a b  c d');
    const v = await mount({
      sheets: [{ name: 'S', rows: row(1, [c('A1', 'a\nb'), c('B1', 'c')]) }],
    });
    v.select({ row: 0, col: 0 }, { row: 0, col: 1 });
    const r = v.copy();
    if (!r.ok) throw new Error(r.refusal.code);
    expect(r.content.text).toBe('a b\tc');
    // A paste of the copy in the editor keeps the break.
    v.select({ row: 2, col: 0 });
    expect(v.paste({ token: r.content.token }).ok).toBe(true);
    expect(content(v, 2, 0)).toBe('a\nb');
  });
});

describe('a pending cut', () => {
  test('an undo expires it', async () => {
    const book = { sheets: [{ name: 'S', rows: row(1, [c('A1', 1), c('B1', 2)]) }] };
    const v = await mount(book);
    expect(v.enterValue('5', { row: 4, col: 4 }).ok).toBe(true);
    v.select({ row: 0, col: 0 });
    const cut = v.cut();
    if (!cut.ok) throw new Error(cut.refusal.code);
    expect(v.undo().ok).toBe(true);
    v.select({ row: 2, col: 0 });
    expect(code(v.paste({ token: cut.content.token }))).toBe('cut-expired');
  });
});

describe('the cut announcement', () => {
  test('a cut says in the live region that its cells wait to move', async () => {
    const v = await mount({ sheets: [{ name: 'S', rows: row(1, [c('A1', 1), c('B1', 2)]) }] });
    v.select({ row: 0, col: 0 }, { row: 0, col: 1 });
    const r = v.cut();
    expect(r.ok).toBe(true);
    expect(liveText(v)).toBe('grid.cut [A1:B1]');
    // A copy says nothing of a move.
    v.select({ row: 3, col: 3 });
    expect(v.copy().ok).toBe(true);
    expect(liveText(v)).not.toContain('grid.cut');
    // A copy says what it copied, in a short text.
    expect(liveText(v)).toBe('grid.copied [D4]');
  });

  test('a plain paste ends a pending cut and says that the cells stay', async () => {
    const v = await mount({ sheets: [{ name: 'S', rows: row(1, [c('A1', 1), c('B1', 2)]) }] });
    v.select({ row: 0, col: 0 });
    const r = v.cut();
    if (!r.ok) throw new Error(r.refusal.code);
    // The text of the cut with no private type, as a clipboard that drops the type gives it.
    v.select({ row: 2, col: 0 });
    expect(v.paste({ text: r.content.text }).ok).toBe(true);
    expect(liveText(v)).toBe('grid.cutEnded');
    expect(content(v, 0, 0)).toBe('1');
    expect(content(v, 2, 0)).toBe('1');
    // The cut no longer waits: a later paste of its token takes the plain lane too.
    v.select({ row: 3, col: 0 });
    expect(v.paste({ token: r.content.token, text: 'x' }).ok).toBe(true);
    expect(content(v, 0, 0)).toBe('1');
    expect(content(v, 3, 0)).toBe('x');
  });

  test('a refused plain paste keeps the pending cut', async () => {
    const v = await mount({ sheets: [{ name: 'S', rows: row(1, [c('A1', 1)]) }] });
    v.select({ row: 0, col: 0 });
    const r = v.cut();
    if (!r.ok) throw new Error(r.refusal.code);
    // Two rows from the last row of the grid leave the grid.
    v.select({ row: 1048575, col: 0 });
    expect(code(v.paste({ text: 'x\ny' }))).toBe('grid-full');
    expect(liveText(v)).not.toBe('grid.cutEnded');
    // The cut still waits: its token moves the cell.
    v.select({ row: 2, col: 0 });
    expect(v.paste({ token: r.content.token }).ok).toBe(true);
    expect(content(v, 0, 0)).toBe('');
    expect(content(v, 2, 0)).toBe('1');
  });

  test('a plain paste after an expired cut says nothing of the cut', async () => {
    const v = await mount({ sheets: [{ name: 'S', rows: row(1, [c('A1', 1)]) }] });
    v.select({ row: 0, col: 0 });
    const r = v.cut();
    if (!r.ok) throw new Error(r.refusal.code);
    expect(v.enterValue('5', { row: 4, col: 4 }).ok).toBe(true);
    v.select({ row: 2, col: 0 });
    expect(v.paste({ text: 'x' }).ok).toBe(true);
    expect(liveText(v)).not.toBe('grid.cutEnded');
    // The cut stays expired: its token is still refused.
    expect(code(v.paste({ token: r.content.token }))).toBe('cut-expired');
    expect(content(v, 0, 0)).toBe('1');
  });

  test('Escape on the grid forgets a pending cut: the paste takes the plain lane', async () => {
    const v = await mount({ sheets: [{ name: 'S', rows: row(1, [c('A1', 1), c('B1', 2)]) }] });
    const scroller = hosts.get(v)!.querySelector('.xlsx-scroller') as HTMLElement;
    const W = scroller.ownerDocument.defaultView as unknown as typeof globalThis;
    const escape = () => {
      const e = new W.KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true });
      scroller.dispatchEvent(e);
      return e.defaultPrevented;
    };
    v.select({ row: 0, col: 0 });
    const r = v.cut();
    if (!r.ok) throw new Error(r.refusal.code);
    expect(escape()).toBe(true);
    expect(liveText(v)).toBe('grid.cutEnded');
    // With no pending cut, Escape goes on to the page.
    expect(escape()).toBe(false);
    v.select({ row: 2, col: 0 });
    expect(v.paste({ token: r.content.token, text: 'x' }).ok).toBe(true);
    // The source stays, and the paste wrote the plain text.
    expect(content(v, 0, 0)).toBe('1');
    expect(content(v, 2, 0)).toBe('x');
  });

  test('Escape keeps an expired cut: every later paste is refused with cut-expired', async () => {
    const v = await mount({ sheets: [{ name: 'S', rows: row(1, [c('A1', 1), c('B1', 2)]) }] });
    const scroller = hosts.get(v)!.querySelector('.xlsx-scroller') as HTMLElement;
    const W = scroller.ownerDocument.defaultView as unknown as typeof globalThis;
    const escape = () => {
      const e = new W.KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true });
      scroller.dispatchEvent(e);
      return e.defaultPrevented;
    };
    v.select({ row: 0, col: 0 });
    const r = v.cut();
    if (!r.ok) throw new Error(r.refusal.code);
    // An edit after the cut expires it, before and after its first refused paste.
    v.select({ row: 4, col: 4 });
    expect(v.enterValue('5').ok).toBe(true);
    expect(escape()).toBe(false);
    v.select({ row: 2, col: 0 });
    const first = v.paste({ token: r.content.token, text: 'x' });
    expect(first.ok ? null : first.refusal.code).toBe('cut-expired');
    expect(escape()).toBe(false);
    const again = v.paste({ token: r.content.token, text: 'x' });
    expect(again.ok ? null : again.refusal.code).toBe('cut-expired');
    expect(content(v, 0, 0)).toBe('1');
    expect(content(v, 2, 0)).toBe('');
  });

  test('a locked refusal by another rule is announced again', async () => {
    const unlocked =
      '<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" ' +
      'applyProtection="1"><protection locked="0"/></xf>';
    const v = await mount({
      sheets: [
        {
          name: 'S',
          rows: row(1, [c('A1', 45000, { s: 1 }), c('B1', 'x', { s: 2 })]),
          beforeMerges: '<sheetProtection sheet="1"/>',
        },
      ],
      styles: { cellXfs: [unlocked, unlocked.replace('numFmtId="14"', 'numFmtId="0"')] },
    });
    v.select({ row: 0, col: 0 });
    const r = v.copy();
    if (!r.ok) throw new Error(r.refusal.code);
    v.select({ row: 0, col: 1 });
    const pasted = v.paste({ token: r.content.token, text: r.content.text });
    expect(pasted.ok ? null : pasted.refusal.lock).toBe('format');
    expect(liveText(v)).toBe('editing.lockedRefused.pasteFormats');
    const sized = v.setColumnWidth(1, 1, 20);
    expect(sized.ok ? null : sized.refusal.lock).toBe('size');
    expect(liveText(v)).toBe('editing.lockedRefused.size');
  });
});
