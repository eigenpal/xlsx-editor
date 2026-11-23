// A cell whose Markup Compatibility content (ECMA-376 Part 3) is not its value: a clear and a cut of it
// are refused with `cell-alternate-content`, whose message names the cell. A cut of a cell whose
// only value is in `mc:AlternateContent`, which the copy cannot carry, is refused with
// `cut-wrapped-value`.

import { afterEach, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { Window } from 'happy-dom';
import { buildXlsx, row, type FixtureWorkbook } from '../../../../../scripts/fixtures/builder.ts';
import {
  createXlsxEditor,
  disabledReasonText,
  refusalText,
  toolbarCommandState,
  type XlsxEditorInstance,
} from '../index.ts';
import en from '../../../../i18n/en.json';
import { flushEditor } from './flush-editor.ts';
import { readXlsx, xmlPartOf } from '../../store/package/package.ts';
import { serializeXmlPart } from '../../store/package-write/serialize.ts';
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
let viewerHosts: HTMLElement[] = [];
afterEach(() => {
  for (const v of viewers) v.destroy();
  viewers = [];
  viewerHosts = [];
});

async function mount(wb: FixtureWorkbook) {
  const win = new Window();
  const doc = win.document as unknown as Document;
  const host = doc.createElement('div');
  doc.body.appendChild(host);
  const viewer = createXlsxEditor({ fonts: FONTS });
  viewers.push(viewer);
  viewerHosts.push(host as unknown as HTMLElement);
  viewer.attach(host);
  await viewer.load(buildXlsx(wb));
  flushEditor(viewer);
  return viewer;
}

/** The English catalog, with parameters filled in. */
function english(key: string, params?: Readonly<Record<string, string | number>>): string {
  let node: unknown = en;
  for (const part of key.split('.')) node = (node as Record<string, unknown>)[part];
  return String(node).replace(/\{(\w+)\}/g, (all, k: string) => String(params?.[k] ?? all));
}

const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
// `mc:Choice/@Requires` names a prefix, which must be declared (ECMA-376 Part 3, 7.6).
const X14 = 'urn:example:x14';
const OTHER =
  `<mc:AlternateContent xmlns:mc="${MC}" xmlns:x14="${X14}"><mc:Choice Requires="x14">` +
  '<e:x xmlns:e="urn:example:extension">keep</e:x></mc:Choice></mc:AlternateContent>';
const WB: FixtureWorkbook = {
  sheets: [
    {
      name: 'S',
      rows: row(1, ['<c r="A1"><v>1</v></c>']) + row(2, [`<c r="A2"><v>2</v>${OTHER}</c>`]),
    },
  ],
};

describe('a cell with markup compatibility content that is not its value', () => {
  test('refuses a clear with a message that names the cell, and changes nothing', async () => {
    const v = await mount(WB);
    v.select({ row: 0, col: 0 }, { row: 1, col: 0 });
    const r = v.clearValues();
    expect(r.ok ? null : { code: r.refusal.code, subject: r.refusal.subject }).toEqual({
      code: 'cell-alternate-content',
      subject: 'S!A2',
    });
    expect(r.ok ? '' : refusalText(r.refusal, english)).toBe(
      'The cells were not cleared: cell S!A2 holds data in a form that the editor cannot update. The editor cannot change the value of that cell. Clear a range without that cell.'
    );
    v.select({ row: 0, col: 0 });
    expect(v.snapshot().activeCell?.content).toBe('1');
  });

  test('opens no cell editor on that cell, because no commit of it could apply', async () => {
    const v = await mount(WB);
    v.select({ row: 1, col: 0 });
    const r = v.beginEdit('7');
    expect(r.ok ? null : { code: r.refusal.code, subject: r.refusal.subject }).toEqual({
      code: 'cell-alternate-content',
      subject: 'S!A2',
    });
    expect(v.snapshot().editing).toBeFalsy();
    // A cell next to it opens as before.
    v.select({ row: 0, col: 0 });
    expect(v.beginEdit('7').ok).toBe(true);
  });

  test('announces a formula range refusal once while the user keeps typing on its cell', async () => {
    const v = await mount({
      sheets: [
        { name: 'S', rows: row(1, ['<c r="A1"><f t="array" ref="A1:A2">1+1</f><v>2</v></c>']) },
      ],
    });
    const host = viewerHosts.at(-1)!;
    const live = host.querySelector('.xlsx-live') as HTMLElement;
    const scroller = host.querySelector('.xlsx-scroller') as HTMLElement;
    const key = (k: string) =>
      scroller.dispatchEvent(
        new (host.ownerDocument.defaultView as unknown as typeof globalThis).KeyboardEvent(
          'keydown',
          { key: k, bubbles: true, cancelable: true }
        )
      );
    v.select({ row: 1, col: 0 });
    key('x');
    expect(v.snapshot().refusal?.code).toBe('shared-formula');
    const first = live.textContent;
    expect(first).toBe(
      'The edit was not applied: it would change the formula range at S!A1:A2, which the editor cannot rewrite or remove. Edit another cell.'
    );
    key('y');
    expect(live.textContent).toBe(first);
    expect(v.snapshot().editing).toBeFalsy();
  });

  test('announces the refusal once while the user keeps typing on that cell', async () => {
    const v = await mount(WB);
    const host = viewerHosts.at(-1)!;
    const live = host.querySelector('.xlsx-live') as HTMLElement;
    const scroller = host.querySelector('.xlsx-scroller') as HTMLElement;
    const key = (k: string) =>
      scroller.dispatchEvent(
        new (host.ownerDocument.defaultView as unknown as typeof globalThis).KeyboardEvent(
          'keydown',
          { key: k, bubbles: true, cancelable: true }
        )
      );
    v.select({ row: 1, col: 0 });
    key('x');
    expect(v.snapshot().refusal?.code).toBe('cell-alternate-content');
    const first = live.textContent;
    expect(first).toBe(
      'The edit was not applied: cell S!A2 holds data in a form that the editor cannot update. The editor cannot change the value of that cell. Edit another cell.'
    );
    // Each key leaves the live region as it was, so a screen reader does not read it again.
    key('y');
    expect(live.textContent).toBe(first);
    key('z');
    expect(live.textContent).toBe(first);
    expect(v.snapshot().editing).toBeFalsy();
    // A command after the refused keys is announced, and its message names the command.
    expect(v.clearValues().ok).toBe(false);
    const cleared = live.textContent;
    expect(cleared).toBe(
      'The cells were not cleared: cell S!A2 holds data in a form that the editor cannot update. The editor cannot change the value of that cell. Clear a range without that cell.'
    );
    // The same command again is a repeat; another command is announced.
    v.clearValues();
    expect(live.textContent).toBe(cleared);
    expect(v.clearAll().ok).toBe(false);
    expect(live.textContent).toBe(
      'The cells were not cleared: cell S!A2 holds data in a form that the editor cannot update. The editor cannot change the value of that cell. To remove only the formats, use Clear formats.'
    );
  });

  test('names Clear all and the cut in their refusals, over a range and over one cell', async () => {
    const v = await mount(WB);
    v.select({ row: 0, col: 0 }, { row: 1, col: 0 });
    const r = v.clearAll();
    expect(r.ok ? null : { code: r.refusal.code, command: r.refusal.command }).toEqual({
      code: 'cell-alternate-content',
      command: 'clearAll',
    });
    expect(r.ok ? '' : refusalText(r.refusal, english)).toBe(
      'The cells were not cleared: cell S!A2 holds data in a form that the editor cannot update. The editor cannot change the value of that cell. To remove only the formats, use Clear formats.'
    );
    // A cut of that cell alone also names the cut.
    v.select({ row: 1, col: 0 });
    const cut = v.cut();
    expect(cut.ok ? 'ok' : cut.refusal.command).toBe('cut');
    expect(cut.ok ? '' : refusalText(cut.refusal, english)).toBe(
      'The cells were not cut: cell S!A2 holds data in a form that the editor cannot update. The editor cannot change the value of that cell, so a cut cannot clear it. Cut a range without that cell.'
    );
  });

  test('disables Clear contents and Clear all while the selection holds that cell', async () => {
    const v = await mount(WB);
    v.select({ row: 0, col: 0 }, { row: 1, col: 0 });
    expect(v.snapshot().selectionContentBlocked).toBe('S!A2');
    const reason = (id: 'clear.contents' | 'clear.all') => {
      const state = toolbarCommandState(v.snapshot(), id);
      expect(state.enabled).toBe(false);
      return state.disabledReason ? disabledReasonText(state.disabledReason, english) : '';
    };
    expect(reason('clear.contents')).toBe(
      'The editor cannot clear cell S!A2. That cell holds data in a form that the editor cannot update. Select a range without that cell.'
    );
    // While Clear formats works, Clear all points to it, as its refusal message does.
    expect(reason('clear.all')).toBe(
      'The editor cannot clear cell S!A2. That cell holds data in a form that the editor cannot update. To remove only the formats, use Clear formats.'
    );
    // While Clear formats is refused too, Clear all does not point to it.
    const refused = { ...v.snapshot(), clearRefusal: { code: 'clear-too-large' as const } };
    expect(toolbarCommandState(refused, 'clear.all').disabledReason?.key).toBe(
      'disabledReason.cellAlternateContent'
    );
    // Clear formats changes no value, and a selection without that cell clears as before.
    expect(toolbarCommandState(v.snapshot(), 'clear.formats').enabled).toBe(true);
    v.select({ row: 0, col: 0 });
    expect(v.snapshot().selectionContentBlocked).toBeNull();
    expect(toolbarCommandState(v.snapshot(), 'clear.contents').enabled).toBe(true);
  });

  test('refuses a cut of that cell at once, so no paste of it is refused later', async () => {
    const v = await mount(WB);
    v.select({ row: 0, col: 0 }, { row: 1, col: 0 });
    const r = v.cut();
    expect(r.ok ? null : { code: r.refusal.code, subject: r.refusal.subject }).toEqual({
      code: 'cell-alternate-content',
      subject: 'S!A2',
    });
    expect(r.ok ? '' : refusalText(r.refusal, english)).toBe(
      'The cells were not cut: cell S!A2 holds data in a form that the editor cannot update. The editor cannot change the value of that cell, so a cut cannot clear it. Cut a range without that cell.'
    );
    // A copy of the same cells is not refused: it leaves the source as it is.
    const copied = v.copy();
    expect(copied.ok).toBe(true);
    // A paste of the range over that cell is refused, and its message names the paste.
    v.select({ row: 1, col: 0 });
    const token = copied.ok ? copied.content.token : '';
    const pasted = v.paste({ token });
    const text =
      'The cells were not pasted: cell S!A2 holds data in a form that the editor cannot update. The editor cannot change the value of that cell, so a paste cannot replace it. Paste to a range without that cell.';
    expect(pasted.ok ? '' : refusalText(pasted.refusal, english)).toBe(text);
    // So is a paste of one cell onto it.
    v.select({ row: 0, col: 0 });
    const one = v.copy();
    v.select({ row: 1, col: 0 });
    const single = v.paste({ token: one.ok ? one.content.token : '' });
    expect(single.ok ? 'ok' : single.refusal.command).toBe('paste');
    expect(single.ok ? '' : refusalText(single.refusal, english)).toBe(text);
  });

  test('cuts a cell with that content and no value, and leaves it as it is', async () => {
    const v = await mount({
      sheets: [
        {
          name: 'S',
          rows: row(1, ['<c r="A1"><v>1</v></c>']) + row(2, [`<c r="A2">${OTHER}</c>`]),
        },
      ],
    });
    v.select({ row: 0, col: 0 }, { row: 1, col: 0 });
    const cut = v.cut();
    expect(cut.ok).toBe(true);
    v.select({ row: 0, col: 2 });
    expect(v.paste({ token: cut.ok ? cut.content.token : '' }).ok).toBe(true);
    const read = readXlsx(v.save()!);
    if (!read.ok) throw new Error('reopen');
    const xml = serializeXmlPart(xmlPartOf(read.pkg, '/xl/worksheets/sheet1.xml')!);
    expect(xml).toContain(`<c r="A2">${OTHER}</c>`);
    expect(xml).toContain('<c r="C1"><v>1</v></c>');
  });

  test('refuses a cut of a cell whose only value is wrapped, and keeps the value', async () => {
    const wrapped =
      `<mc:AlternateContent xmlns:mc="${MC}" xmlns:x14="${X14}"><mc:Choice Requires="x14"><v>5</v></mc:Choice>` +
      '<mc:Fallback><v>5</v></mc:Fallback></mc:AlternateContent>';
    const v = await mount({
      sheets: [{ name: 'S', rows: row(1, [`<c r="A1">${wrapped}</c>`]) }],
    });
    v.select({ row: 0, col: 0 });
    const r = v.cut();
    expect(r.ok ? null : { code: r.refusal.code, subject: r.refusal.subject }).toEqual({
      code: 'cut-wrapped-value',
      subject: 'S!A1',
    });
    // Every refusal of a cut is marked as a cut (`CUT_REFUSAL_CODES`).
    expect(r.ok ? 'ok' : r.refusal.command).toBe('cut');
    expect(r.ok ? '' : refusalText(r.refusal, english)).toBe(
      'The cells were not cut: cell S!A1 holds a value that the editor cannot show, so a paste cannot carry it. Clear that cell, or cut other cells.'
    );
    // The save keeps the wrapped value in its cell.
    const read = readXlsx(v.save()!);
    if (!read.ok) throw new Error('reopen');
    expect(serializeXmlPart(xmlPartOf(read.pkg, '/xl/worksheets/sheet1.xml')!)).toContain(
      '<v>5</v></mc:Choice>'
    );
  });

  test('refuses a cut of a wrapped value after ten other children', async () => {
    const wrapped =
      `<mc:AlternateContent xmlns:mc="${MC}" xmlns:x14="${X14}"><mc:Choice Requires="x14"><v>5</v></mc:Choice>` +
      '</mc:AlternateContent>';
    const ten = `<e:x xmlns:e="urn:example:extension">k</e:x>`.repeat(10);
    const v = await mount({
      sheets: [{ name: 'S', rows: row(1, [`<c r="A1">${ten}${wrapped}</c>`]) }],
    });
    v.select({ row: 0, col: 0 });
    const r = v.cut();
    expect(r.ok ? null : r.refusal.code).toBe('cut-wrapped-value');
  });

  test('a refused merge over an array formula names the merge and its next step', async () => {
    const v = await mount({
      sheets: [
        {
          name: 'S',
          rows:
            row(1, ['<c r="D1"><v>1</v></c>']) +
            row(2, ['<c r="D2"><f t="array" ref="D2:D3">1+1</f><v>2</v></c>']) +
            row(3, ['<c r="D3"><v>2</v></c>']),
        },
      ],
    });
    v.select({ row: 0, col: 3 }, { row: 1, col: 3 });
    const r = v.mergeCells({ dropValues: true });
    expect(r.ok ? null : { code: r.refusal.code, command: r.refusal.command }).toEqual({
      code: 'shared-formula',
      command: 'merge',
    });
    expect(r.ok ? '' : refusalText(r.refusal, english)).toBe(
      'The cells were not merged: the merge would change the formula range at S!D2:D3, which the editor cannot rewrite or remove. Merge a range without cells of that range.'
    );
  });

  test('a refused merge over a wrapped cell names the merge and its next step', async () => {
    const wrapped =
      `<mc:AlternateContent xmlns:mc="${MC}" xmlns:x14="${X14}"><mc:Choice Requires="x14">` +
      '<f>1+1</f><v>2</v></mc:Choice></mc:AlternateContent>';
    // A value in an element in another namespace is content that the merge would hide too.
    const foreign = '<e:x xmlns:e="urn:example:extension"><v>2</v></e:x>';
    for (const d2 of [wrapped, foreign]) {
      const v = await mount({
        sheets: [
          {
            name: 'S',
            rows: row(1, ['<c r="D1"><v>1</v></c>']) + row(2, [`<c r="D2">${d2}</c>`]),
          },
        ],
      });
      v.select({ row: 0, col: 3 }, { row: 1, col: 3 });
      // The merge asks first, and names the cell.
      const ask = v.mergeCells();
      expect(ask.ok ? null : { code: ask.refusal.code, subject: ask.refusal.subject }).toEqual({
        code: 'merge-drops-values',
        subject: 'D2',
      });
      const r = v.mergeCells({ dropValues: true });
      expect(r.ok ? null : { code: r.refusal.code, command: r.refusal.command }).toEqual({
        code: 'cell-alternate-content',
        command: 'merge',
      });
      expect(r.ok ? '' : refusalText(r.refusal, english)).toBe(
        'The cells were not merged: cell S!D2 holds data in a form that the editor cannot update. The editor cannot change the value of that cell. Merge a range without that cell.'
      );
    }
  });

  test('a refused commit keeps the cell editor open and says to cancel the cell edit', async () => {
    // The child formula of A2 would move out of the grid, so an edit of its master is refused at
    // the commit, after the cell editor opened.
    const v = await mount({
      sheets: [
        {
          name: 'S',
          rows:
            row(1, ['<c r="A1"><f t="shared" ref="A1:A2" si="0">B1048576</f><v>1</v></c>']) +
            row(2, ['<c r="A2"><f t="shared" si="0"/><v>2</v></c>']),
        },
      ],
    });
    v.select({ row: 0, col: 0 });
    expect(v.beginEdit('5').ok).toBe(true);
    const r = v.commitEdit();
    expect(r.ok ? null : { code: r.refusal.code, command: r.refusal.command }).toEqual({
      code: 'shared-formula',
      command: 'commit',
    });
    const text = r.ok ? '' : refusalText(r.refusal, english);
    expect(text).toEndWith('Cancel the cell edit.');
    expect(text).not.toContain('nother cell');
    expect(v.snapshot().editing).toBe(true);

    // A save that commits first names the save, and the cell editor still stays open.
    expect(v.save()).toBeNull();
    expect(v.snapshot().refusal?.command).toBe('save');
    expect(v.snapshot().editing).toBe(true);

    // The cancel ends the edit, and the cell keeps its content.
    v.cancelEdit();
    expect(v.snapshot().editing).toBe(false);
    expect(v.save()).not.toBeNull();
  });
});
