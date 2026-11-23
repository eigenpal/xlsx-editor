// The clipboard of the editor (design D2 of the change `workbook-structure-editing`) in a
// happy-dom window of its own: the text export with its escapes, the internal paste with moved
// formulas and cell formats, cut, the plain-text paste, the routes and the refusals, and one
// undo step for a paste.

import { afterEach, describe, expect, spyOn, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { Window } from 'happy-dom';
import { buildXlsx, c, row, type FixtureSheet } from '../../../../../scripts/fixtures/builder.ts';
import {
  createXlsxEditor,
  disabledReasonText,
  refusalText,
  toolbarCommandState,
  type XlsxEditorInstance,
} from '../index.ts';
import { flushEditor } from './flush-editor.ts';
import { exportCellText, splitPlainText } from '../clipboard.ts';
import {
  attr,
  cellKey,
  readXlsx,
  sheetIndex,
  WorkbookDocumentStore,
  xmlPartOf,
} from '../../store/index.ts';
import type { FontSource } from '../../layout/index.ts';
import en from '../../../../i18n/en.json';

/** The English catalog, with parameters filled in. */
function english(key: string, params?: Readonly<Record<string, string | number>>): string {
  let node: unknown = en;
  for (const part of key.split('.')) node = (node as Record<string, unknown>)[part];
  return String(node).replace(/\{(\w+)\}/g, (all, k: string) => String(params?.[k] ?? all));
}

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

const STYLES = {
  cellXfs: [
    '<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>',
  ],
};

async function mount(sheet: Partial<FixtureSheet> & { rows: string }, styles = STYLES) {
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
  await viewer.load(buildXlsx({ sheets: [{ name: 'S', ...sheet }], styles }));
  flushEditor(viewer);
  return viewer;
}

/** The content and the number format id of a cell, as the snapshot shows them. */
function at(v: XlsxEditorInstance, r: number, col: number) {
  v.select({ row: r, col });
  const s = v.snapshot();
  return { content: s.activeCell?.content, format: s.activeFormat?.numberFormat.id };
}

/** The cell type attribute of a saved cell. */
function savedType(v: XlsxEditorInstance, r: number, col: number) {
  const read = readXlsx(v.save()!);
  if (!read.ok) throw new Error('save');
  const part = xmlPartOf(read.pkg, '/xl/worksheets/sheet1.xml')!;
  return attr(sheetIndex(part).cells.get(cellKey(r, col))?.el, 't');
}

const code = (r: { ok: boolean; refusal?: { code: string } }) => (r.ok ? null : r.refusal?.code);

function copied(v: XlsxEditorInstance, cut = false) {
  const r = cut ? v.cut() : v.copy();
  if (!r.ok) throw new Error(r.refusal.code);
  return r.content;
}

describe('clipboard text export', () => {
  test('a copy writes tab-separated shown text with CR LF rows and formula-injection escapes', async () => {
    const v = await mount({
      rows: [
        row(1, [c('A1', '=1+1'), c('B1', -5), c('C1', '@cmd')]),
        row(2, [c('A2', 45000, { s: 1 }), c('C2', 'a\tb')]),
      ].join(''),
    });
    v.select({ row: 0, col: 0 }, { row: 1, col: 2 });
    expect(copied(v).text).toBe("'=1+1\t-5\t'@cmd\r\n03-15-23\t\ta b");
  });

  test('the escape keeps a signed number and marks text that only looks like one', () => {
    expect(exportCellText('-1,234.50', true)).toBe('-1,234.50');
    expect(exportCellText('-2+3+cmd|x', true)).toBe("'-2+3+cmd|x");
    expect(exportCellText('-5', false)).toBe("'-5");
    expect(exportCellText('+1', false)).toBe("'+1");
    expect(exportCellText('\tx', false)).toBe("' x");
    expect(exportCellText('\rx', false)).toBe("' x");
    expect(exportCellText('a\r\nb', false)).toBe('a  b');
  });

  test('plain text splits at CR LF, LF, CR, and tab, and a final line break adds no row', () => {
    expect(splitPlainText('1\ta\r\n2\tb\n3\rc\n')).toEqual([['1', 'a'], ['2', 'b'], ['3'], ['c']]);
    expect(splitPlainText('x')).toEqual([['x']]);
    expect(splitPlainText('a\nb\nc', 2)).toBeNull();
  });
});

describe('internal paste', () => {
  test('a relative formula moves by the offset, and the cell format comes with it', async () => {
    const v = await mount({ rows: row(1, [c('A1', 1), c('B1', 1, { f: 'A1', s: 1 })]) });
    v.select({ row: 0, col: 1 });
    const content = copied(v);
    v.select({ row: 2, col: 1 });
    expect(v.paste({ text: content.text, token: content.token }).ok).toBe(true);
    expect(at(v, 2, 1)).toEqual({ content: '=A3', format: 14 });
    // The source does not change on a copy.
    expect(at(v, 0, 1)).toEqual({ content: '=A1', format: 14 });
  });

  test('a cut clears the source with its format in the paste transaction', async () => {
    const v = await mount({ rows: row(1, [c('A1', 7, { s: 1 }), c('B1', 1, { f: 'A1' })]) });
    v.select({ row: 0, col: 0 });
    const content = copied(v, true);
    // A cut changes nothing until the paste.
    expect(at(v, 0, 0)).toEqual({ content: '1900-01-07', format: 14 });
    v.select({ row: 0, col: 2 });
    expect(v.paste({ text: content.text, token: content.token }).ok).toBe(true);
    expect(at(v, 0, 2)).toEqual({ content: '1900-01-07', format: 14 });
    expect(at(v, 0, 0)).toEqual({ content: '', format: 0 });
    // References to the cut cell do not move (design D2).
    expect(at(v, 0, 1).content).toBe('=A1');
    // One undo puts back both the source and the target.
    expect(v.undo().ok).toBe(true);
    expect(at(v, 0, 0)).toEqual({ content: '1900-01-07', format: 14 });
    expect(at(v, 0, 2)).toEqual({ content: '', format: 0 });
  });

  test('a cut whose source clear the store refuses is refused at the cut', async () => {
    const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
    // The prefix that `@Requires` names is declared, as the reader requires.
    const X14 = 'urn:example:x14';
    // A shared master whose child has its formula inside mc:AlternateContent: the clear of the
    // master alone is refused, so the cut is refused now, not its paste later.
    const rows =
      '<row r="2"><c r="A2"><f t="shared" ref="A2:A3" si="0">B2+1</f><v>2</v></c></row>' +
      `<row r="3"><c r="A3"><mc:AlternateContent xmlns:mc="${MC}" xmlns:x14="${X14}"><mc:Choice Requires="x14">` +
      '<f t="shared" si="0"/><v>3</v></mc:Choice></mc:AlternateContent></c></row>';
    const v = await mount({ rows });
    // The formula bar of the wrapped child shows the master's formula moved to it, as the grid
    // shows that child's cached value.
    expect(at(v, 2, 0).content).toBe('=B3+1');
    v.select({ row: 1, col: 0 });
    // Clear contents and Clear all are disabled with the reason that the clear gives.
    expect(v.snapshot().selectionFormulaRangeBlocked).toBe('S!A2:A3');
    // The reason names the wrapped child, whose formula the editor cannot rewrite.
    expect(v.snapshot().selectionFormulaRangeCauseCell).toBe('S!A3');
    for (const id of ['clear.contents', 'clear.all'] as const) {
      const state = toolbarCommandState(v.snapshot(), id);
      expect(state.enabled).toBe(false);
      expect(state.disabledReason ? disabledReasonText(state.disabledReason, (k) => k) : '').toBe(
        id === 'clear.all' ? 'disabledReason.formulaCellClearAll' : 'disabledReason.formulaCell'
      );
      expect(state.disabledReason?.params).toEqual({ causeCell: 'S!A3', masterCell: 'S!A2' });
    }
    expect(toolbarCommandState(v.snapshot(), 'clear.formats').enabled).toBe(true);
    // No commit could change the master, so typing on it opens no cell editor.
    const open = v.beginEdit('5');
    expect(open.ok ? null : [open.refusal.code, open.refusal.causeCell]).toEqual([
      'shared-formula',
      'S!A3',
    ]);
    expect(v.snapshot().editing).toBe(false);
    const cut = v.cut();
    expect(
      cut.ok
        ? null
        : [cut.refusal.code, cut.refusal.causeCell, cut.refusal.masterCell, cut.refusal.command]
    ).toEqual(['shared-formula', 'S!A3', 'S!A2', 'cut']);
    if (!cut.ok)
      expect(refusalText(cut.refusal, (key) => key)).toBe('editing.formulaCellRefused.cut');
    // A copy of the same cell is not a clear, so it is done.
    const copy = v.copy();
    expect(copy.ok).toBe(true);
    // A clear and a paste over the master name their own command.
    const clear = v.clearValues();
    expect(clear.ok ? null : [clear.refusal.code, clear.refusal.command]).toEqual([
      'shared-formula',
      'clear',
    ]);
    // The message names the master to leave out, and the wrapped child that blocks it.
    expect(clear.ok ? '' : refusalText(clear.refusal, english)).toBe(
      'The cells were not cleared: S!A2 is a shared formula master, and the editor cannot rewrite the formula of S!A3. Clear a range without that master.'
    );
    // Clear all is refused the same way, and Clear formats, which its message names, is done.
    const all = v.clearAll();
    expect(all.ok ? null : [all.refusal.code, all.refusal.command]).toEqual([
      'shared-formula',
      'clearAll',
    ]);
    expect(v.clearFormats().ok).toBe(true);
    if (!copy.ok) return;
    const paste = v.paste({ token: copy.content.token });
    expect(paste.ok ? null : [paste.refusal.code, paste.refusal.command]).toEqual([
      'shared-formula',
      'paste',
    ]);
    if (!paste.ok)
      expect(refusalText(paste.refusal, (key) => key)).toBe('editing.formulaCellRefused.paste');
  });

  test('a shared master whose formula cannot be read names it as unreadable', async () => {
    // The master's text is not a formula, so no clear or edit can unshare its range.
    const rows =
      '<row r="2"><c r="A2"><f t="shared" ref="A2:A3" si="0">SUM((</f><v>2</v></c></row>' +
      '<row r="3"><c r="A3"><f t="shared" si="0"/><v>3</v></c></row>';
    const v = await mount({ rows });
    v.select({ row: 1, col: 0 });
    expect(v.snapshot().selectionFormulaRangeCauseCell).toBe('S!A2');
    expect(v.snapshot().selectionFormulaRangeCauseUnreadable).toBe(true);
    const state = toolbarCommandState(v.snapshot(), 'clear.contents');
    expect(state.disabledReason ? disabledReasonText(state.disabledReason, english) : '').toBe(
      'The editor cannot clear the shared formula master S!A2: its formula cannot be read. Select a range without that master.'
    );
    const open = v.beginEdit('5');
    expect(open.ok ? '' : refusalText(open.refusal, english)).toBe(
      'The edit was not applied: S!A2 is a shared formula master whose formula cannot be read. Edit another cell.'
    );
  });

  test('a cut formula keeps its references, and a second paste is a copy', async () => {
    const v = await mount({ rows: row(1, [c('A1', 1), c('B1', 1, { f: 'A1' })]) });
    v.select({ row: 0, col: 1 });
    const content = copied(v, true);
    v.select({ row: 4, col: 1 });
    expect(v.paste({ token: content.token }).ok).toBe(true);
    expect(at(v, 4, 1).content).toBe('=A1');
    expect(at(v, 0, 1).content).toBe('');
    v.select({ row: 5, col: 1 });
    expect(v.paste({ token: content.token }).ok).toBe(true);
    expect(at(v, 5, 1).content).toBe('=A6');
  });

  test('a cut pasted on another sheet keeps the cells its formulas named; a copy stays relative', async () => {
    const win = new Window();
    const host = (win.document as unknown as Document).createElement('div');
    (win.document as unknown as Document).body.appendChild(host);
    const v = createXlsxEditor({ fonts: FONTS });
    viewers.push(v);
    v.attach(host);
    await v.load(
      buildXlsx({
        sheets: [
          {
            name: 'S 1',
            rows: row(1, [c('A1', 1), c('B1', 1, { f: 'A1+S2!A1' }), c('C1', 1, { f: 'A1' })]),
          },
          { name: 'S2', rows: row(1, [c('A1', 99)]) },
        ],
      })
    );
    flushEditor(v);
    v.select({ row: 0, col: 1 });
    const cut = copied(v, true);
    expect(v.setActiveSheet(1)).toBe(true);
    v.select({ row: 4, col: 4 });
    expect(v.paste({ token: cut.token }).ok).toBe(true);
    expect(at(v, 4, 4).content).toBe("='S 1'!A1+'S2'!A1");
    expect(v.setActiveSheet(0)).toBe(true);
    v.select({ row: 0, col: 2 });
    const copy = copied(v);
    expect(v.setActiveSheet(1)).toBe(true);
    v.select({ row: 5, col: 4 });
    expect(v.paste({ token: copy.token }).ok).toBe(true);
    expect(at(v, 5, 4).content).toBe('=C6');
  });

  test('an edit between the cut and the paste refuses the paste with cut-expired', async () => {
    const v = await mount({ rows: row(1, [c('A1', 7)]) });
    v.select({ row: 0, col: 0 });
    const content = copied(v, true);
    expect(v.enterValue('8', { row: 3, col: 3 }).ok).toBe(true);
    v.select({ row: 0, col: 2 });
    const r = v.paste({ token: content.token });
    expect(code(r)).toBe('cut-expired');
    expect(r.ok ? null : r.refusal.subject).toBe('S!A1');
    expect(v.snapshot().refusal?.code).toBe('cut-expired');
    expect(at(v, 0, 2).content).toBe('');
    expect(at(v, 0, 0).content).toBe('7');
    // The same token stays refused with the same code, and never takes the plain lane.
    v.select({ row: 0, col: 2 });
    expect(code(v.paste({ token: content.token, text: content.text }))).toBe('cut-expired');
    expect(at(v, 0, 2).content).toBe('');
    // Text with no token takes the plain lane, and a new cut pastes.
    v.select({ row: 0, col: 2 });
    expect(v.paste({ text: content.text }).ok).toBe(true);
    expect(at(v, 0, 2).content).toBe('7');
    v.select({ row: 3, col: 3 });
    const again = copied(v, true);
    v.select({ row: 5, col: 5 });
    expect(v.paste({ token: again.token }).ok).toBe(true);
    expect(at(v, 5, 5).content).toBe('8');
    expect(at(v, 3, 3).content).toBe('');
  });

  test('one undo restores every target cell, its value and its format', async () => {
    const v = await mount({
      rows: [
        row(1, [c('A1', 1, { s: 1 }), c('B1', 'x')]),
        row(2, [c('A2', true)]),
        row(4, [c('C4', 'old'), c('D4', 9, { s: 1 })]),
        row(5, [c('C5', 3)]),
      ].join(''),
    });
    v.select({ row: 0, col: 0 }, { row: 1, col: 1 });
    const content = copied(v);
    v.select({ row: 3, col: 2 });
    expect(v.paste({ token: content.token }).ok).toBe(true);
    expect(at(v, 3, 2)).toEqual({ content: '1900-01-01', format: 14 });
    expect(at(v, 3, 3)).toEqual({ content: 'x', format: 0 });
    expect(at(v, 4, 2)).toEqual({ content: 'TRUE', format: 0 });
    // An empty source cell clears its target.
    expect(at(v, 4, 3)).toEqual({ content: '', format: 0 });
    // The paste selects the pasted range.
    v.select({ row: 3, col: 2 });
    expect(v.undo().ok).toBe(true);
    expect(at(v, 3, 2)).toEqual({ content: 'old', format: 0 });
    expect(at(v, 3, 3)).toEqual({ content: '1900-01-09', format: 14 });
    expect(at(v, 4, 2)).toEqual({ content: '3', format: 0 });
    expect(at(v, 4, 3)).toEqual({ content: '', format: 0 });
    expect(v.undo().ok).toBe(false);
  });

  test('rich text and an error value paste as plain text (approximation pasted-values)', async () => {
    const win = new Window();
    const host = (win.document as unknown as Document).createElement('div');
    (win.document as unknown as Document).body.appendChild(host);
    const v = createXlsxEditor({ fonts: FONTS });
    viewers.push(v);
    v.attach(host);
    await v.load(
      buildXlsx({
        rawStrings: ['<si><r><rPr><b/></rPr><t>Bold</t></r><r><t> plain</t></r></si>'],
        sheets: [
          { name: 'S', rows: row(1, [c('A1', 0, { t: 's' }), c('B1', '#N/A', { t: 'e' })]) },
        ],
      })
    );
    flushEditor(v);
    v.select({ row: 0, col: 0 }, { row: 0, col: 1 });
    const content = copied(v);
    v.select({ row: 2, col: 0 });
    expect(v.paste({ token: content.token }).ok).toBe(true);
    expect(at(v, 2, 0).content).toBe('Bold plain');
    expect(at(v, 2, 1).content).toBe('#N/A');
    expect(savedType(v, 2, 0)).toBe('s');
    expect(savedType(v, 2, 1)).toBe('s');
  });

  test('the paste selects the pasted range', async () => {
    const v = await mount({ rows: row(1, [c('A1', 1), c('B1', 2)]) });
    v.select({ row: 0, col: 0 }, { row: 0, col: 1 });
    const content = copied(v);
    v.select({ row: 2, col: 0 });
    expect(v.paste({ token: content.token }).ok).toBe(true);
    expect(v.snapshot().selectionRef).toBe('A3:B3');
  });

  test('a redo of a paste selects the pasted range again, in both lanes', async () => {
    const v = await mount({ rows: row(1, [c('A1', 1), c('B1', 2)]) });
    v.select({ row: 0, col: 0 }, { row: 0, col: 1 });
    const content = copied(v);
    for (const data of [{ token: content.token }, { text: '1\t2\r\n3\t4' }]) {
      v.select({ row: 2, col: 0 });
      expect(v.paste(data).ok).toBe(true);
      expect(v.undo().ok).toBe(true);
      expect(v.snapshot().selectionRef).toBe('A3');
      expect(v.redo().ok).toBe(true);
      const want = 'token' in data ? 'A3:B3' : 'A3:B4';
      expect([v.snapshot().selectionRef, v.snapshot().activeCell?.ref]).toEqual([want, 'A3']);
      expect(v.undo().ok).toBe(true);
    }
  });

  test('a cut never reaches a workbook that opens later: the paste takes the plain lane', async () => {
    const v = await mount({ rows: row(1, [c('A1', 7)]) });
    v.select({ row: 0, col: 0 });
    const content = copied(v, true);
    // The new store starts at revision 0, as the store of the cut did.
    await v.load(buildXlsx({ sheets: [{ name: 'Other', rows: row(1, [c('A1', 'keep')]) }] }));
    flushEditor(v);
    v.select({ row: 0, col: 3 });
    expect(v.paste({ token: content.token, text: content.text }).ok).toBe(true);
    expect(at(v, 0, 0).content).toBe('keep');
    expect(at(v, 0, 3).content).toBe('7');
    // Without text, the token alone changes nothing.
    v.select({ row: 2, col: 3 });
    expect(v.paste({ token: content.token }).ok).toBe(true);
    expect(at(v, 2, 3).content).toBe('');
  });

  test('a token that is not the last copy takes the plain lane', async () => {
    const v = await mount({ rows: row(1, [c('A1', 5, { s: 1 })]) });
    v.select({ row: 0, col: 0 });
    const first = copied(v);
    const second = copied(v);
    expect(second.token).not.toBe(first.token);
    v.select({ row: 0, col: 2 });
    expect(v.paste({ text: '5', token: first.token }).ok).toBe(true);
    // The plain lane writes the value and keeps the target format.
    expect(at(v, 0, 2)).toEqual({ content: '5', format: 0 });
  });
});

describe('plain-text paste', () => {
  test('two rows read as typed input, and the cell formats stay', async () => {
    const v = await mount({ rows: row(1, [c('B1', 'q', { s: 1 })]) });
    v.select({ row: 0, col: 0 });
    expect(v.paste({ text: '1\ta\r\n2\tb' }).ok).toBe(true);
    expect(at(v, 0, 0).content).toBe('1');
    expect(at(v, 0, 1)).toEqual({ content: 'a', format: 14 });
    expect(at(v, 1, 0).content).toBe('2');
    expect(at(v, 1, 1).content).toBe('b');
    expect(savedType(v, 0, 0)).toBeUndefined();
    expect(savedType(v, 0, 1)).toBe('s');
    expect(v.undo().ok).toBe(true);
    expect(at(v, 0, 0).content).toBe('');
    expect(at(v, 0, 1).content).toBe('q');
  });

  test('an escaped formula pastes as text, and = starts a formula', async () => {
    const v = await mount({ rows: '' });
    v.select({ row: 0, col: 0 });
    expect(v.paste({ text: "'=1+1\t=1+1\tTRUE" }).ok).toBe(true);
    // A text that reads as a formula shows with its leading apostrophe, so an edit keeps it text.
    expect(at(v, 0, 0).content).toBe("'=1+1");
    expect(savedType(v, 0, 0)).toBe('s');
    expect(at(v, 0, 1).content).toBe('=1+1');
    expect(savedType(v, 0, 1)).toBeUndefined();
    expect(savedType(v, 0, 2)).toBe('b');
  });
});

describe('clipboard refusals', () => {
  test('HTML only is refused with paste-html; an empty clipboard changes nothing', async () => {
    const v = await mount({ rows: row(1, [c('A1', 1)]) });
    v.select({ row: 0, col: 0 });
    expect(code(v.paste({ html: true }))).toBe('paste-html');
    expect(code(v.paste({ text: '', html: true }))).toBe('paste-html');
    expect(v.snapshot().refusal?.code).toBe('paste-html');
    expect(v.paste({}).ok).toBe(true);
    expect(v.snapshot().canUndo).toBe(false);
  });

  test('a copy or a paste that meets a merged range is refused with merged-range', async () => {
    const v = await mount({ rows: row(1, [c('A1', 1), c('C1', 2)]), merges: ['C3:D4'] });
    v.select({ row: 0, col: 0 }, { row: 0, col: 1 });
    const content = copied(v);
    v.select({ row: 3, col: 1 });
    expect(code(v.paste({ token: content.token }))).toBe('merged-range');
    expect(code(v.paste({ text: '1\t2' }))).toBe('merged-range');
    v.select({ row: 2, col: 2 });
    expect(code(v.copy())).toBe('merged-range');
    expect(code(v.cut())).toBe('merged-range');
    expect(v.snapshot().canUndo).toBe(false);
  });

  test('a cut refused before its copy checks has the cut command and the cut message', async () => {
    const v = await mount({ rows: row(1, [c('A1', 1)]), merges: ['C3:D4'] });
    const cutOf = (): [string, string | undefined, string] | null => {
      const r = v.cut();
      return r.ok ? null : [r.refusal.code, r.refusal.command, refusalText(r.refusal, (k) => k)];
    };
    v.select({ row: 2, col: 2 });
    expect(cutOf()).toEqual(['merged-range', 'cut', 'editing.cutRefused.merged-range']);
    v.select({ row: 0, col: 0 }, { row: 1, col: 16383 });
    expect(cutOf()).toEqual(['copy-too-large', 'cut', 'editing.cutRefused.copy-too-large']);
    v.select({ row: 0, col: 0 });
    v.setMode('view');
    expect(cutOf()).toEqual(['view-mode', 'cut', 'editing.cutRefused.view-mode']);
    v.setMode('edit');
    expect(v.snapshot().canUndo).toBe(false);
  });

  test('a paste that leaves the grid is refused with grid-full, a moved reference with its own code', async () => {
    const v = await mount({ rows: row(2, [c('B2', 1, { f: 'A1' })]) });
    v.select({ row: 0, col: 16383 });
    expect(code(v.paste({ text: '1\t2' }))).toBe('grid-full');
    v.select({ row: 1, col: 1 });
    const content = copied(v);
    // =A1 one row up, or one column left, would name a cell outside the sheet.
    for (const target of [
      { row: 0, col: 1 },
      { row: 1, col: 0 },
      { row: 0, col: 0 },
    ]) {
      v.select(target);
      const r = v.paste({ token: content.token });
      expect(code(r)).toBe('paste-reference-outside');
      expect(r.ok ? null : r.refusal.subject).toBe('S!B2');
    }
    expect(v.snapshot().canUndo).toBe(false);
  });

  test('a copy or a paste of more than one whole row of cells is refused', async () => {
    const v = await mount({ rows: '' });
    v.select({ row: 0, col: 0 }, { row: 1, col: 16383 });
    expect(code(v.copy())).toBe('copy-too-large');
    v.select({ row: 0, col: 0 });
    expect(code(v.paste({ text: Array(16_385).fill('1').join('\n') }))).toBe('range-too-large');
    expect(code(v.paste({ text: `${'\t'.repeat(9000)}\n${'\t'.repeat(9000)}` }))).toBe(
      'range-too-large'
    );
  });

  test('an array formula is refused on copy with its own code', async () => {
    const v = await mount({
      rows: row(1, ['<c r="A1"><f t="array" ref="A1:A2">1+1</f><v>2</v></c>', c('B1', 2)]),
    });
    // A clear that meets the array range is refused, so the toolbar disables it.
    v.select({ row: 1, col: 0 });
    // No commit could change a member, so typing on it opens no cell editor.
    expect(code(v.beginEdit('5'))).toBe('shared-formula');
    expect(v.snapshot().selectionFormulaRangeBlocked).toBe('S!A1:A2');
    expect(toolbarCommandState(v.snapshot(), 'clear.contents').enabled).toBe(false);
    expect(v.clearValues().ok).toBe(false);
    // A selection outside the range clears.
    v.select({ row: 0, col: 1 });
    expect(v.snapshot().selectionFormulaRangeBlocked).toBeNull();
    expect(toolbarCommandState(v.snapshot(), 'clear.contents').enabled).toBe(true);
    v.select({ row: 0, col: 0 }, { row: 0, col: 1 });
    const r = v.copy();
    expect(code(r)).toBe('copy-array-formula');
    // A copy names the array range as an edit and a clear of its member name it.
    expect(r.ok ? null : r.refusal.subject).toBe('S!A1:A2');
    v.select({ row: 1, col: 0 });
    const copied = v.copy();
    const cleared = v.clearValues();
    expect(copied.ok ? null : copied.refusal.subject).toBe('S!A1:A2');
    expect(cleared.ok ? null : cleared.refusal.subject).toBe('S!A1:A2');
  });

  test('a member of an array formula inside mc:AlternateContent is refused on copy', async () => {
    const mc = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
    const X14 = 'urn:example:x14';
    const master =
      `<c r="A1"><mc:AlternateContent xmlns:mc="${mc}" xmlns:x14="${X14}"><mc:Choice Requires="x14">` +
      '<f t="array" ref="A1:A2">{1;2}</f><v>1</v></mc:Choice></mc:AlternateContent></c>';
    const v = await mount({ rows: row(1, [master]) + row(2, [c('A2', 2)]) });
    v.select({ row: 1, col: 0 });
    expect(code(v.copy())).toBe('copy-array-formula');
  });

  test('a copy or a cut of a formula inside mc:AlternateContent is refused, not pasted as a value', async () => {
    const mc = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
    const X14 = 'urn:example:x14';
    const wrapped =
      `<c r="A2"><mc:AlternateContent xmlns:mc="${mc}" xmlns:x14="${X14}"><mc:Choice Requires="x14">` +
      '<f>B1*2</f><v>4</v></mc:Choice></mc:AlternateContent></c>';
    const v = await mount({ rows: row(1, [c('A1', 1), c('B1', 2)]) + row(2, [wrapped]) });
    v.select({ row: 0, col: 0 }, { row: 1, col: 0 });
    // A cut gives the reason of its copy, before the clear of its source.
    for (const [r, code, command, key] of [
      [v.copy(), 'copy-wrapped-formula', undefined, 'editing.refused.copy-wrapped-formula'],
      [v.cut(), 'copy-wrapped-formula', 'cut', 'editing.cutRefused.copy-wrapped-formula'],
    ] as const) {
      expect(r.ok ? null : [r.refusal.code, r.refusal.subject, r.refusal.command]).toEqual([
        code,
        'S!A2',
        command,
      ]);
      // A refused cut names the cut, not a copy.
      if (!r.ok) expect(refusalText(r.refusal, (k) => k)).toBe(key);
    }
    // The cut wrote nothing: the source keeps its formula.
    expect(at(v, 1, 0).content).toBe('=B1*2');
    // A cut runs the checks of the copy in the same order: an empty `x:f` before the wrapped cell
    // gives the same reason to both.
    const first = await mount({
      rows: row(1, ['<c r="A1"><f/></c>', wrapped.replace('A2', 'B1')]),
    });
    first.select({ row: 0, col: 0 }, { row: 0, col: 1 });
    for (const r of [first.copy(), first.cut()])
      expect(r.ok ? null : [r.refusal.code, r.refusal.subject]).toEqual([
        'copy-unreadable-formula',
        'S!A1',
      ]);
    // A range without that cell copies.
    v.select({ row: 0, col: 0 });
    expect(v.copy().ok).toBe(true);
  });
});

test('a cut on a sheet whose rows the store cannot edit names the sheet, not other cells', async () => {
  const mc = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
  const X14 = 'urn:example:x14';
  const wrapped =
    `<mc:AlternateContent xmlns:mc="${mc}" xmlns:x14="${X14}"><mc:Choice Requires="x14">` +
    `${row(2, [c('A2', 2)])}</mc:Choice></mc:AlternateContent>`;
  const v = await mount({ rows: row(1, [c('A1', 1)]) + wrapped });
  v.select({ row: 0, col: 0 });
  const r = v.cut();
  expect(r.ok ? null : [r.refusal.code, r.refusal.command]).toEqual(['invalid-sheet', 'cut']);
  expect(r.ok ? '' : refusalText(r.refusal, (k) => k)).toBe('editing.cutRefused.invalid-sheet');
});

describe('a paste on a protected sheet', () => {
  // xf 1: a date format, unlocked. xf 2: General, unlocked. Formatting cells is not allowed.
  const UNLOCKED = {
    cellXfs: [
      '<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyProtection="1"><protection locked="0"/></xf>',
      '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyProtection="1"><protection locked="0"/></xf>',
    ],
  };
  const sheet = {
    rows: row(1, [c('A1', 45000, { s: 1 }), c('B1', 'x', { s: 2 }), c('C1', 7, { s: 2 })]),
    beforeMerges: '<sheetProtection sheet="1"/>',
  };

  test('a cut of a locked wrapped formula gives the reason of its copy', async () => {
    const mc = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
    const X14 = 'urn:example:x14';
    const wrapped =
      `<c r="A2"><mc:AlternateContent xmlns:mc="${mc}" xmlns:x14="${X14}"><mc:Choice Requires="x14">` +
      '<f>1+1</f><v>2</v></mc:Choice></mc:AlternateContent></c>';
    const v = await mount({ ...sheet, rows: sheet.rows + row(2, [wrapped]) }, UNLOCKED);
    v.select({ row: 1, col: 0 });
    // A2 is locked, so the clear of the cut is refused too; the copy reason comes first.
    for (const r of [v.copy(), v.cut()])
      expect(r.ok ? null : [r.refusal.code, r.refusal.subject]).toEqual([
        'copy-wrapped-formula',
        'S!A2',
      ]);
  });

  test('a paste that would change the format of an unlocked cell is refused with the format lock', async () => {
    const v = await mount(sheet, UNLOCKED);
    v.select({ row: 0, col: 0 });
    const content = copied(v);
    v.select({ row: 0, col: 1 });
    const r = v.paste({ token: content.token, text: content.text });
    expect(r.ok ? null : [r.refusal.code, r.refusal.lock]).toEqual(['locked', 'format']);
    // The message names the paste, not a format command.
    expect(r.ok ? null : r.refusal.command).toBe('paste');
    expect(r.ok ? null : refusalText(r.refusal, (key) => key)).toBe(
      'editing.lockedRefused.pasteFormats'
    );
    expect(at(v, 0, 1)).toEqual({ content: 'x', format: 0 });
    expect(v.snapshot().canUndo).toBe(false);
  });

  test('a paste in the format that the unlocked cell already shows is allowed', async () => {
    const v = await mount(sheet, UNLOCKED);
    v.select({ row: 0, col: 2 });
    const content = copied(v);
    v.select({ row: 0, col: 1 });
    expect(v.paste({ token: content.token, text: content.text }).ok).toBe(true);
    expect(at(v, 0, 1).content).toBe('7');
  });

  test('a cut that its paste could not clear is refused at the cut, with the rule', async () => {
    const v = await mount({ ...sheet, rows: sheet.rows + row(2, [c('A2', 1)]) }, UNLOCKED);
    const cut = (r: number, col: number) => {
      v.select({ row: r, col });
      const out = v.cut();
      return out.ok ? null : out.refusal;
    };
    // A2 is locked: a paste of the cut would clear it.
    const locked = cut(1, 0);
    expect([locked?.code, locked?.lock, locked?.command]).toEqual(['locked', 'cell', 'cut']);
    expect(refusalText(locked!, (key) => key)).toBe('editing.lockedRefused.cut');
    // B1 is unlocked, but the sheet does not allow its format to be cleared.
    const formats = cut(0, 1);
    expect([formats?.lock, formats?.command]).toEqual(['format', 'cut']);
    expect(refusalText(formats!, (key) => key)).toBe('editing.lockedRefused.cutFormats');
    // A copy after the refused cut, with no other selection, is the next command: the refusal
    // no longer shows.
    expect(v.snapshot().refusal?.command).toBe('cut');
    expect(v.copy().ok).toBe(true);
    expect(v.snapshot().refusal).toBeNull();
    // A copy of a locked cell stays allowed.
    v.select({ row: 1, col: 0 });
    expect(v.copy().ok).toBe(true);
    // With formatting allowed, the cut of an unlocked cell applies, and its paste clears it.
    const free = await mount(
      { ...sheet, beforeMerges: '<sheetProtection sheet="1" formatCells="0"/>' },
      UNLOCKED
    );
    free.select({ row: 0, col: 2 });
    const content = copied(free, true);
    free.select({ row: 0, col: 1 });
    expect(free.paste({ token: content.token, text: content.text }).ok).toBe(true);
    expect([at(free, 0, 1).content, at(free, 0, 2).content]).toEqual(['7', '']);
  });

  test('a paste at a locked cell names the cell rule, also when the format would change', async () => {
    const rows = sheet.rows + row(2, [c('A2', 1)]);
    // A copy of an unlocked cell in another format, on a sheet that locks formatting.
    const v = await mount({ ...sheet, rows }, UNLOCKED);
    v.select({ row: 0, col: 0 });
    const content = copied(v);
    v.select({ row: 1, col: 0 });
    const r = v.paste({ token: content.token, text: content.text });
    expect(r.ok ? null : [r.refusal.code, r.refusal.lock]).toEqual(['locked', 'cell']);
    expect(r.ok ? null : refusalText(r.refusal, (key) => key)).toBe('editing.lockedRefused.paste');
    // A cut of an unlocked cell, on a sheet that allows formatting.
    const free = await mount(
      { rows, beforeMerges: '<sheetProtection sheet="1" formatCells="0"/>' },
      UNLOCKED
    );
    free.select({ row: 0, col: 2 });
    const cut = copied(free, true);
    free.select({ row: 1, col: 0 });
    const p = free.paste({ token: cut.token, text: cut.text });
    expect(p.ok ? null : [p.refusal.code, p.refusal.lock]).toEqual(['locked', 'cell']);
    expect(at(free, 0, 2).content).toBe('7');
  });

  test('a cut moved by one column over its own cells applies, with formatting allowed', async () => {
    const rows = row(1, [c('B1', 1, { s: 2 }), c('C1', 2, { s: 2 }), c('D1', 3, { s: 2 })]);
    const v = await mount(
      { rows, beforeMerges: '<sheetProtection sheet="1" formatCells="0"/>' },
      UNLOCKED
    );
    v.select({ row: 0, col: 1 }, { row: 0, col: 2 });
    const cut = copied(v, true);
    v.select({ row: 0, col: 2 });
    expect(v.paste({ token: cut.token, text: cut.text }).ok).toBe(true);
    expect([1, 2, 3].map((col) => at(v, 0, col).content)).toEqual(['', '1', '2']);
    // Each cell keeps its protection: B1, which the cut cleared, stays unlocked.
    const locked = [1, 2, 3].map(
      (col) => (v.select({ row: 0, col }), v.snapshot().activeCellLocked)
    );
    expect(locked).toEqual([false, false, false]);
  });

  test('a cut over its own cells clears only the cells that its paste does not write', async () => {
    const v = await mount({ rows: row(1, [c('B1', 1), c('C1', 2)]) + row(2, [c('B2', 3)]) });
    v.select({ row: 0, col: 1 }, { row: 1, col: 2 });
    const cut = copied(v, true);
    v.select({ row: 1, col: 2 });
    const transact = spyOn(WorkbookDocumentStore.prototype, 'transact');
    try {
      expect(v.paste({ token: cut.token, text: cut.text }).ok).toBe(true);
      const ops = transact.mock.calls[0]![0];
      const clears = ops.filter((op) => op.op === 'clearValues' || op.op === 'clearFormats');
      // The source B1:C2 less the target C2:D3: B1:C1 above it, and B2 at its left.
      expect(clears.map((op) => ('ref' in op ? op.ref : ''))).toEqual([
        'B1:C1',
        'B1:C1',
        'B2',
        'B2',
      ]);
    } finally {
      transact.mockRestore();
    }
    expect([at(v, 0, 1).content, at(v, 1, 2).content, at(v, 2, 3).content]).toEqual(['', '1', '']);
  });

  test('the clear of a cut keeps the protection of each source cell', async () => {
    // E1 and F1, the target, are unlocked cells that hold nothing.
    const target = ['<c r="E1" s="2"/>', '<c r="F1" s="2"/>'];
    const rows = row(1, [c('B1', 1, { s: 1 }), c('C1', 2, { s: 2 }), ...target]);
    const v = await mount(
      { rows, beforeMerges: '<sheetProtection sheet="1" formatCells="0"/>' },
      UNLOCKED
    );
    v.select({ row: 0, col: 1 }, { row: 0, col: 2 });
    const cut = copied(v, true);
    v.select({ row: 0, col: 4 });
    expect(v.paste({ token: cut.token, text: cut.text }).ok).toBe(true);
    // B1 and C1 hold nothing and show format 0, and stay unlocked, so a value can be typed.
    for (const col of [1, 2]) {
      expect(at(v, 0, col)).toEqual({ content: '', format: 0 });
      v.select({ row: 0, col });
      expect(v.snapshot().activeCellLocked).toBe(false);
    }
    expect(at(v, 0, 4)).toEqual({ content: '1900-01-01', format: 14 });
    v.select({ row: 0, col: 1 });
    expect(v.enterValue('5').ok).toBe(true);
    // Undo puts back the cells and the styles part.
    v.undo();
    v.undo();
    expect([at(v, 0, 1).content, at(v, 0, 4).content]).toEqual(['1900-01-01', '']);
  });
});

describe('clipboard events', () => {
  test('copy and paste events on the grid use the clipboard data; the cell editor keeps its own', async () => {
    const v = await mount({ rows: row(1, [c('A1', 3, { s: 1 })]) });
    const grid = (viewers[0] as unknown as { els: { scroller: HTMLElement } }).els.scroller;
    const store = new Map<string, string>();
    const clipboardData = {
      setData: (type: string, value: string) => void store.set(type, value),
      getData: (type: string) => {
        if (type === 'text/html') throw new Error('the HTML was read');
        return store.get(type) ?? '';
      },
      get types() {
        return [...store.keys()];
      },
    };
    const fire = (type: string, target: EventTarget = grid) => {
      const e = new (grid.ownerDocument.defaultView as unknown as typeof globalThis).Event(type, {
        bubbles: true,
        cancelable: true,
      });
      Object.defineProperty(e, 'clipboardData', { value: clipboardData });
      target.dispatchEvent(e);
      return e.defaultPrevented;
    };
    // The grid acts on clipboard events only while it has the focus.
    grid.focus();
    v.select({ row: 0, col: 0 });
    expect(fire('copy')).toBe(true);
    // The shown `01-03-00` reads back as 2000, so the copy gives the ISO 8601 date.
    expect(store.get('text/plain')).toBe('1900-01-03');
    expect(store.get('application/x-xlsx-editor')).toBeTruthy();
    v.select({ row: 1, col: 0 });
    expect(fire('paste')).toBe(true);
    expect(at(v, 1, 0)).toEqual({ content: '1900-01-03', format: 14 });
    // An event from the open cell editor passes through.
    v.beginEdit('x');
    const input = grid.querySelector('textarea')!;
    expect(fire('paste', input)).toBe(false);
    v.cancelEdit();
    grid.focus();
    // HTML only: the paste reads the list of types, never the HTML, and is refused.
    store.clear();
    store.set('text/html', '<b>x</b>');
    expect(fire('paste')).toBe(true);
    expect(v.snapshot().refusal?.code).toBe('paste-html');
  });

  test('an event at the stage acts only while the grid has the focus', async () => {
    const v = await mount({ rows: row(1, [c('A1', 'Copied')]) });
    const { scroller: grid, stage } = (
      viewers[0] as unknown as { els: { scroller: HTMLElement; stage: HTMLElement } }
    ).els;
    const doc = grid.ownerDocument;
    const store = new Map<string, string>();
    const clipboardData = {
      setData: (type: string, value: string) => void store.set(type, value),
      getData: (type: string) => store.get(type) ?? '',
      get types() {
        return [...store.keys()];
      },
    };
    const fire = (type: string, target: EventTarget) => {
      const e = new (doc.defaultView as unknown as typeof globalThis).Event(type, {
        bubbles: true,
        cancelable: true,
      });
      Object.defineProperty(e, 'clipboardData', { value: clipboardData });
      target.dispatchEvent(e);
      return e.defaultPrevented;
    };
    // A field outside the editor has the focus: the stage's events pass through.
    const field = doc.createElement('input');
    doc.body.append(field);
    try {
      field.focus();
      v.select({ row: 0, col: 0 });
      expect(fire('copy', stage)).toBe(false);
      expect(fire('copy', field)).toBe(false);
      // The page selection can stay in the grid after the focus leaves it: an event at the
      // grid itself does not act either.
      expect(fire('copy', grid)).toBe(false);
      expect(fire('paste', grid)).toBe(false);
      expect(store.size).toBe(0);
      // The grid has the focus, and the browser fires at the stage: the grid copies and pastes.
      grid.focus();
      expect(doc.activeElement).toBe(grid);
      expect(fire('copy', stage)).toBe(true);
      expect(store.get('text/plain')).toBe('Copied');
      v.select({ row: 1, col: 0 });
      expect(fire('paste', stage)).toBe(true);
      expect(at(v, 1, 0).content).toBe('Copied');
      // An event on the rest of the page, or at a field, is never the grid's.
      expect(fire('paste', doc.body)).toBe(false);
      expect(fire('paste', field)).toBe(false);
    } finally {
      field.remove();
    }
  });

  test('the page selection moves into the grid when the grid takes the focus', async () => {
    await mount({ rows: row(1, [c('A1', 'Copied')]) });
    const { scroller: grid } = (viewers[0] as unknown as { els: { scroller: HTMLElement } }).els;
    const doc = grid.ownerDocument;
    // A selection on host text: a press on the grid cannot move it, because the grid cancels
    // the press.
    const text = doc.createElement('p');
    text.textContent = 'Host text';
    doc.body.append(text);
    try {
      doc.getSelection()!.selectAllChildren(text);
      grid.focus();
      const selection = doc.getSelection()!;
      expect(selection.anchorNode).toBe(grid);
      expect(selection.isCollapsed).toBe(true);
      // A selection that is already in the grid stays where it is.
      const inside = doc.createElement('span');
      inside.textContent = 'Inside';
      grid.append(inside);
      selection.selectAllChildren(inside);
      grid.blur();
      grid.focus();
      expect(selection.anchorNode).toBe(inside);
      inside.remove();
    } finally {
      text.remove();
    }
  });
});
