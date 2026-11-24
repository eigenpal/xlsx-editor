// View > Headings: the `view.headings` slot shows and sets `x:sheetView/@showRowColHeaders` of the
// shown sheet (`CT_SheetView`, ECMA-376 Part 1, 18.3.1.87) through the store, so the grid follows
// it with no new open, undo and redo step it, and the save keeps it.

import { afterEach, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { Window } from 'happy-dom';
import { buildXlsx, c, row } from '../../../../../scripts/fixtures/builder.ts';
import { CHROME_MENUS, menuRowRole } from '../chrome-menus.ts';
import { chromeControlForSlot } from '../chrome-controls.ts';
import { runToolbarCommand, toolbarCommandState } from '../toolbar-commands.ts';
import { createXlsxEditor, type XlsxEditorInstance } from '../index.ts';
import { readXlsx, xmlPartOf } from '../../store/index.ts';
import { firstSheetView, viewShowsHeadings } from '../../store/store/sheet-view.ts';
import type { FontSource } from '../../layout/index.ts';
import { flushEditor } from './flush-editor.ts';

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

/** Open a workbook of two sheets: `A` hides its headings, and `B` shows them. */
async function open() {
  const win = new Window();
  const doc = win.document as unknown as Document;
  const host = doc.createElement('div');
  doc.body.appendChild(host);
  const viewer = createXlsxEditor({ fonts: FONTS });
  viewers.push(viewer);
  viewer.attach(host);
  await viewer.load(
    buildXlsx({
      sheets: [
        { name: 'A', rows: row(1, [c('A1', 1)]) },
        { name: 'B', rows: row(1, [c('A1', 2)]), headings: true },
      ],
    })
  );
  flushEditor(viewer);
  const heads = () => !(host.querySelector('.xlsx-column-heads') as HTMLElement).hidden;
  return { viewer, heads };
}

/** Whether the saved bytes of the editor show the headings of a sheet. */
function savedHeadings(viewer: XlsxEditorInstance, sheet: number): boolean {
  const read = readXlsx(viewer.save()!);
  if (!read.ok) throw new Error(JSON.stringify(read.error));
  const parts = [...read.pkg.parts.keys()].filter((p) => p.startsWith('/xl/worksheets/')).sort();
  return viewShowsHeadings(firstSheetView(xmlPartOf(read.pkg, parts[sheet])?.root));
}

describe('View > Headings', () => {
  test('the View menu places the row, a checked row of its own', () => {
    const view = CHROME_MENUS.find((m) => m.id === 'view')!;
    expect(view.entries.at(-1)).toEqual({ kind: 'item', slot: 'view.headings' });
    expect(view.entries.at(-2)).toEqual({ kind: 'separator' });
    const control = chromeControlForSlot('view.headings')!;
    expect(control.toggle).toBe(true);
    expect(menuRowRole('view.headings', true)).toBe('menuitemcheckbox');
  });

  test('the row is checked while the shown sheet shows its headings', async () => {
    const { viewer } = await open();
    expect(viewer.snapshot().headings).toBe(false);
    const hidden = toolbarCommandState(viewer.snapshot(), 'view.headings');
    expect({ enabled: hidden.enabled, active: hidden.active }).toEqual({
      enabled: true,
      active: false,
    });
    expect(viewer.setActiveSheet(1)).toBe(true);
    expect(viewer.snapshot().headings).toBe(true);
    expect(toolbarCommandState(viewer.snapshot(), 'view.headings').active).toBe(true);
  });

  test('a press shows the headings with no new open, and undo and redo step it', async () => {
    const { viewer, heads } = await open();
    expect(heads()).toBe(false);
    const before = viewer.snapshot().revision;
    expect(runToolbarCommand(viewer, 'view.headings').ok).toBe(true);
    flushEditor(viewer);
    expect(viewer.snapshot().headings).toBe(true);
    expect(heads()).toBe(true);
    expect(viewer.snapshot().revision).toBe(before + 1);
    expect(viewer.snapshot().canUndo).toBe(true);
    expect(savedHeadings(viewer, 0)).toBe(true);
    expect(viewer.undo().ok).toBe(true);
    flushEditor(viewer);
    expect(viewer.snapshot().headings).toBe(false);
    expect(heads()).toBe(false);
    expect(savedHeadings(viewer, 0)).toBe(false);
    expect(viewer.redo().ok).toBe(true);
    flushEditor(viewer);
    expect(viewer.snapshot().headings).toBe(true);
    expect(heads()).toBe(true);
  });

  test('a press changes only the shown sheet, and an undo shown on another sheet follows', async () => {
    const { viewer } = await open();
    expect(viewer.setActiveSheet(1)).toBe(true);
    expect(viewer.setShowHeadings(false).ok).toBe(true);
    expect(viewer.snapshot().headings).toBe(false);
    expect(savedHeadings(viewer, 0)).toBe(false);
    expect(savedHeadings(viewer, 1)).toBe(false);
    expect(viewer.setActiveSheet(0)).toBe(true);
    expect(viewer.snapshot().headings).toBe(false);
    expect(viewer.undo().ok).toBe(true);
    // The undo shows the sheet of its step, with its headings back.
    expect(viewer.snapshot().activeSheet).toBe(1);
    expect(viewer.snapshot().headings).toBe(true);
  });

  test('the asked state makes no undo step', async () => {
    const { viewer } = await open();
    expect(viewer.setShowHeadings(false).ok).toBe(true);
    expect(viewer.snapshot().canUndo).toBe(false);
  });

  test('view mode disables the row with its reason, and keeps its check state', async () => {
    const { viewer } = await open();
    viewer.setActiveSheet(1);
    viewer.setMode('view');
    const state = toolbarCommandState(viewer.snapshot(), 'view.headings');
    expect(state.enabled).toBe(false);
    expect(state.disabledReason?.key).toBe('disabledReason.viewMode');
    expect(state.active).toBe(true);
    expect(runToolbarCommand(viewer, 'view.headings').ok).toBe(false);
    const refused = viewer.setShowHeadings(false);
    expect(refused.ok ? null : refused.refusal.code).toBe('view-mode');
    expect(viewer.snapshot().headings).toBe(true);
  });

  test('a protected sheet and a protected workbook leave the row enabled', async () => {
    const win = new Window();
    const host = (win.document as unknown as Document).createElement('div');
    const viewer = createXlsxEditor({ fonts: FONTS });
    viewers.push(viewer);
    viewer.attach(host);
    await viewer.load(
      buildXlsx({
        sheets: [
          { name: 'S', rows: row(1, [c('A1', 1)]), beforeMerges: '<sheetProtection sheet="1"/>' },
        ],
        workbookProtection: 'lockStructure="1" lockWindows="1"',
      })
    );
    expect(toolbarCommandState(viewer.snapshot(), 'view.headings').enabled).toBe(true);
    expect(runToolbarCommand(viewer, 'view.headings').ok).toBe(true);
    expect(viewer.snapshot().headings).toBe(true);
  });

  test('an open cell edit commits first', async () => {
    const { viewer } = await open();
    expect(viewer.beginEdit('typed').ok).toBe(true);
    expect(viewer.snapshot().editing).toBe(true);
    expect(viewer.setShowHeadings(true).ok).toBe(true);
    expect(viewer.snapshot().editing).toBe(false);
    expect(viewer.snapshot().activeCell?.content).toBe('typed');
  });
});
