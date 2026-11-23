// The editor while the styles part stays as bytes past a read limit: paste, cut and paste, Clear
// formats, and Clear all give the default format 0. Every other format command is refused, and
// the toolbar disables its controls with the reason `stylesUnread`.

import { afterEach, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { Window } from 'happy-dom';
import { buildXlsx, c, row } from '../../../../../scripts/fixtures/builder.ts';
import {
  createXlsxEditor,
  mergeCommandState,
  toolbarCommandState,
  type XlsxEditorInstance,
} from '../index.ts';
import { flushEditor } from './flush-editor.ts';
import { attr, cellKey, readXlsx, sheetIndex, xmlPartOf } from '../../store/index.ts';
import type { FontSource } from '../../layout/index.ts';

const NS_X = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
/** Elements nested past the 256 levels of the XML part limit. */
const DEEP = `${'<x:a>'.repeat(300)}${'</x:a>'.repeat(300)}`;

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

/**
 * A workbook whose styles part passes the depth limit, with A1 and B1 in the cell format 1, and
 * C1 with no value in the cell format 1. `protection` is an `x:sheetProtection` element.
 */
function workbook(protection = ''): Uint8Array {
  const parts = unzipSync(
    buildXlsx({
      sheets: [{ name: 'S', rows: row(1, [c('A1', 7), c('B1', 'x')]), beforeMerges: protection }],
    })
  );
  const sheet = strFromU8(parts['xl/worksheets/sheet1.xml']!);
  parts['xl/worksheets/sheet1.xml'] = strToU8(
    sheet
      .replace('<c r="A1"', '<c r="A1" s="1"')
      .replace('<c r="B1"', '<c r="B1" s="1"')
      .replace('</row>', '<c r="C1" s="1"/></row>')
  );
  parts['xl/styles.xml'] = strToU8(`<x:styleSheet xmlns:x="${NS_X}">${DEEP}</x:styleSheet>`);
  return zipSync(parts);
}

async function mount(protection = ''): Promise<XlsxEditorInstance> {
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
  await viewer.load(workbook(protection));
  flushEditor(viewer);
  expect(viewer.snapshot().partsPastLimit.map((p) => p.kind)).toEqual(['styles']);
  return viewer;
}

/**
 * The shown content and the saved `@s` of each cell of row 1, or `-` for no saved cell. The save
 * keeps the styles part as it was read.
 */
function saved(v: XlsxEditorInstance): Record<string, [string, string]> {
  const bytes = v.save()!;
  expect(strFromU8(unzipSync(bytes)['xl/styles.xml']!)).toContain(DEEP);
  const read = readXlsx(bytes);
  if (!read.ok) throw new Error(read.error.reason);
  const index = sheetIndex(xmlPartOf(read.pkg, '/xl/worksheets/sheet1.xml')!);
  const out: Record<string, [string, string]> = {};
  for (const [col, name] of ['A1', 'B1', 'C1', 'D1'].entries()) {
    v.select({ row: 0, col });
    const el = index.cells.get(cellKey(0, col))?.el;
    out[name] = [v.snapshot().activeCell?.content ?? '', el ? (attr(el, 's') ?? '0') : '-'];
  }
  return out;
}

const code = (r: { ok: boolean; refusal?: { code: string } }) => (r.ok ? null : r.refusal?.code);

describe('editing while the styles part stays as bytes', () => {
  test('a paste and a cut and paste give the default format 0', async () => {
    const v = await mount();
    v.select({ row: 0, col: 0 });
    const copy = v.copy();
    if (!copy.ok) throw new Error(copy.refusal.code);
    v.select({ row: 0, col: 2 });
    expect(code(v.paste({ token: copy.content.token }))).toBeNull();
    v.select({ row: 0, col: 0 });
    const cut = v.cut();
    if (!cut.ok) throw new Error(cut.refusal.code);
    v.select({ row: 0, col: 3 });
    expect(code(v.paste({ token: cut.content.token }))).toBeNull();
    const cells = saved(v);
    expect(cells.C1).toEqual(['7', '0']);
    expect(cells.D1).toEqual(['7', '0']);
    // The cut source keeps no value and no format other than 0.
    expect(cells.A1?.[0]).toBe('');
    expect(['0', '-']).toContain(cells.A1![1]);
  });

  test('Clear all and Clear formats give the default format 0', async () => {
    const v = await mount();
    v.select({ row: 0, col: 1 });
    expect(code(v.clearAll())).toBeNull();
    v.select({ row: 0, col: 0 });
    expect(code(v.clearFormats())).toBeNull();
    const cells = saved(v);
    expect(cells.A1).toEqual(['7', '0']);
    expect(cells.B1?.[0]).toBe('');
    expect(['0', '-']).toContain(cells.B1![1]);
  });

  test('every other format command is refused', async () => {
    const v = await mount();
    v.select({ row: 0, col: 0 });
    expect(code(v.toggleFormat('bold'))).toBe('unsupported-package');
    expect(code(v.applyFormat({ wrapText: true }))).toBe('unsupported-package');
  });

  test('the toolbar disables the format controls and keeps the clears enabled', async () => {
    const v = await mount();
    v.select({ row: 0, col: 0 }, { row: 0, col: 1 });
    const s = v.snapshot();
    for (const id of [
      'font.bold',
      'font.family',
      'fill.color',
      'borders.all',
      'alignment.wrap',
      'number.format',
      'number.increaseDecimals',
    ] as const) {
      const state = toolbarCommandState(s, id);
      expect([id, state.enabled, state.disabledReason]).toEqual([
        id,
        false,
        { key: 'disabledReason.stylesUnread' },
      ]);
    }
    for (const id of ['clear.formats', 'clear.all', 'clear.contents'] as const)
      expect([id, toolbarCommandState(s, id).enabled]).toEqual([id, true]);
    expect(mergeCommandState(s, 'mergeCenter').disabledReason).toEqual({
      key: 'disabledReason.stylesUnread',
    });
    expect(mergeCommandState(s, 'merge').enabled).toBe(true);
  });

  test('on a protected sheet, the toolbar disables Clear formats and Clear all, as the store refuses them', async () => {
    // `formatCells="0"` allows formatting, so only the unread protection of each cell locks it.
    const v = await mount('<sheetProtection sheet="1" formatCells="0"/>');
    // While the part is unread, each cell reads as locked, so one cell gives the reason `locked`
    // first. The range E1:F1 holds no cell with content, so no locked cell holds content.
    v.select({ row: 0, col: 4 }, { row: 0, col: 5 });
    const s = v.snapshot();
    expect(s.protection?.allowFormatCells).toBe(true);
    for (const id of ['clear.formats', 'clear.all'] as const) {
      const state = toolbarCommandState(s, id);
      expect([id, state.enabled, state.disabledReason]).toEqual([
        id,
        false,
        { key: 'disabledReason.stylesUnreadProtected' },
      ]);
    }
    expect(toolbarCommandState(s, 'font.bold').disabledReason).toEqual({
      key: 'disabledReason.stylesUnread',
    });
    // The engine refuses the same commands: the protection of each cell is in the unread part.
    expect(code(v.clearFormats())).toBe('unsupported-package');
    expect(code(v.clearAll())).toBe('unsupported-package');
    // A cell of the selection with content is locked, so Clear all keeps the reason `locked`.
    v.select({ row: 0, col: 0 });
    expect(toolbarCommandState(v.snapshot(), 'clear.all').disabledReason).toEqual({
      key: 'disabledReason.locked',
    });
    expect(code(v.clearAll())).toBe('locked');
  });
});
