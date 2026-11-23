// The custom number format entry: the check of a typed code by the store's own
// rules, the preview of the active cell's value through the display path, the workbook's
// custom codes, and the apply through `number.custom`: one `x:numFmt` with an id of 164 or
// more, an equal code reused, a built-in code written as its id, one undo step, and the code
// kept through a save and a reopen (ECMA-376 Part 1, 18.8.30 and 18.8.31).

import { afterEach, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { Window } from 'happy-dom';
import { unzipSync, strFromU8 } from 'fflate';
import { buildXlsx, c, row } from '../../../../../scripts/fixtures/builder.ts';
import en from '../../../../i18n/en.json';
import {
  createXlsxEditor,
  MAX_LISTED_NUMBER_FORMATS,
  NUMBER_FORMAT_CODE_PROBLEM_KEYS,
  NUMBER_FORMAT_SAMPLE,
  numberFormatDialogApplyStep,
  numberFormatCodeProblem,
  numberFormatDialogCurrent,
  numberFormatDialogStart,
  previewNumberFormat,
  runToolbarCommand,
  toolbarCommandState,
  workbookNumberFormatCodes,
  type XlsxEditorInstance,
} from '../index.ts';
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

let viewers: XlsxEditorInstance[] = [];
afterEach(() => {
  for (const v of viewers) v.destroy();
  viewers = [];
});

async function open(bytes: Uint8Array): Promise<XlsxEditorInstance> {
  const win = new Window();
  const doc = win.document as unknown as Document;
  const host = doc.createElement('div');
  doc.body.appendChild(host);
  const viewer = createXlsxEditor({ fonts: FONTS });
  viewers.push(viewer);
  viewer.attach(host);
  await viewer.load(bytes);
  flushEditor(viewer);
  return viewer;
}

const NUM_FMTS = [
  '<numFmt numFmtId="170" formatCode="0.000"/>',
  '<numFmt numFmtId="165" formatCode="&quot;kg&quot; #,##0"/>',
  // The same code under another id is listed once.
  '<numFmt numFmtId="171" formatCode="0.000"/>',
  // A code that the store would refuse is not listed.
  '<numFmt numFmtId="172" formatCode="&quot;open"/>',
];

const workbook = (numFmts: readonly string[] = NUM_FMTS) =>
  buildXlsx({
    sheets: [
      {
        name: 'S',
        rows: row(1, [c('A1', 1234.5), c('B1', 'text'), c('C1', true)]) + row(2, [c('A2', -2)]),
      },
    ],
    styles: { numFmts },
  });

/** The styles part of saved bytes. */
const stylesOf = (bytes: Uint8Array): string => strFromU8(unzipSync(bytes)['xl/styles.xml']!);

/** A catalog value by its dotted key. */
function catalog(key: string): unknown {
  let node: unknown = en;
  for (const part of key.split('.')) node = (node as Record<string, unknown> | undefined)?.[part];
  return node;
}

describe('numberFormatCodeProblem', () => {
  test('names each code that the store refuses, and passes the others', () => {
    expect(numberFormatCodeProblem('')?.key).toBe('numberFormat.problem.empty');
    expect(numberFormatCodeProblem('0'.repeat(256))).toEqual({
      key: 'numberFormat.problem.tooLong',
      params: { limit: '255' },
    });
    expect(numberFormatCodeProblem('0\u0001')?.key).toBe('numberFormat.problem.invalidCharacter');
    expect(numberFormatCodeProblem('"open')?.key).toBe('numberFormat.problem.unreadable');
    expect(numberFormatCodeProblem('[Red')?.key).toBe('numberFormat.problem.unreadable');
    expect(numberFormatCodeProblem('0;0;0;@;0')?.key).toBe('numberFormat.problem.unreadable');
    for (const ok of ['0.00', '#,##0.00 "EUR"', '[Red]0;[Blue]-0', 'yyyy-mm-dd', 'General', '@'])
      expect({ ok, problem: numberFormatCodeProblem(ok) }).toEqual({ ok, problem: null });
  });

  test('every problem key is in the catalog', () => {
    for (const key of NUMBER_FORMAT_CODE_PROBLEM_KEYS)
      expect({ key, found: typeof catalog(key) }).toEqual({ key, found: 'string' });
  });
});

describe('previewNumberFormat', () => {
  test('formats the active cell through the code, in the display path', async () => {
    const viewer = await open(workbook());
    expect(previewNumberFormat(viewer, '#,##0.00')).toEqual({
      ok: true,
      text: '1,234.50',
      sample: false,
    });
    expect(previewNumberFormat(viewer, '[Red]0.0')).toEqual({
      ok: true,
      text: '1234.5',
      color: 'FF0000',
      sample: false,
    });
    viewer.setNumberLocale({ decimal: ',', group: '.' });
    expect(previewNumberFormat(viewer, '#,##0.00')).toMatchObject({ text: '1.234,50' });
  });

  test('shows text and booleans as the grid does, and a sample for an empty cell', async () => {
    const viewer = await open(workbook());
    viewer.select({ row: 0, col: 1 });
    expect(previewNumberFormat(viewer, '0.00')).toMatchObject({ text: 'text', sample: false });
    expect(previewNumberFormat(viewer, '"<"@">"')).toMatchObject({ text: '<text>' });
    viewer.select({ row: 0, col: 2 });
    expect(previewNumberFormat(viewer, '0.00')).toMatchObject({ text: 'TRUE' });
    viewer.select({ row: 4, col: 4 });
    expect(previewNumberFormat(viewer, '0.0')).toEqual({
      ok: true,
      text: NUMBER_FORMAT_SAMPLE.toFixed(1),
      sample: true,
    });
  });

  test('a built-in code previews as the apply writes it: its id, in the locale', async () => {
    const viewer = await open(workbook());
    viewer.setNumberLocale({ decimal: ',', group: '.', builtins: { 14: 'dd.mm.yyyy' } });
    // `mm-dd-yy` is the ECMA-376 code of id 14, which this locale shows as `dd.mm.yyyy`.
    expect(previewNumberFormat(viewer, 'mm-dd-yy')).toMatchObject({ text: '18.05.1903' });
    expect(runToolbarCommand(viewer, 'number.custom', 'mm-dd-yy').ok).toBe(true);
    expect(viewer.snapshot().activeFormat?.numberFormat).toMatchObject({
      id: 14,
      code: 'dd.mm.yyyy',
    });
  });

  test('a built-in code that the workbook redefines stays a custom code', async () => {
    const viewer = await open(workbook(['<numFmt numFmtId="14" formatCode="yyyy-mm-dd"/>']));
    expect(previewNumberFormat(viewer, 'mm-dd-yy')).toMatchObject({ text: '05-18-03' });
    expect(runToolbarCommand(viewer, 'number.custom', 'mm-dd-yy').ok).toBe(true);
    expect(viewer.snapshot().activeFormat?.numberFormat).toMatchObject({
      id: 164,
      code: 'mm-dd-yy',
    });
    // An equal redefinition keeps the id.
    const same = await open(workbook(['<numFmt numFmtId="14" formatCode="mm-dd-yy"/>']));
    expect(runToolbarCommand(same, 'number.custom', 'mm-dd-yy').ok).toBe(true);
    expect(same.snapshot().activeFormat?.numberFormat.id).toBe(14);
  });

  test('a formula result after an edit is marked stale, as the grid marks it', async () => {
    const viewer = await open(
      buildXlsx({ sheets: [{ name: 'S', rows: row(1, [c('A1', 4, { f: 'B1*2' }), c('B1', 2)]) }] })
    );
    expect(previewNumberFormat(viewer, '0.0')).toEqual({ ok: true, text: '4.0', sample: false });
    viewer.select({ row: 0, col: 1 });
    viewer.beginEdit('5');
    expect(viewer.commitEdit().ok).toBe(true);
    viewer.select({ row: 0, col: 0 });
    expect(previewNumberFormat(viewer, '0.0')).toEqual({
      ok: true,
      text: '4.0',
      sample: false,
      stale: true,
    });
  });

  test('rich text shows as stored, as the grid paints it', async () => {
    const viewer = await open(
      buildXlsx({
        sheets: [{ name: 'S', rows: row(1, [c('A1', 0, { t: 's' })]) }],
        rawStrings: ['<si><r><rPr><b/></rPr><t>Bold</t></r><r><t> plain</t></r></si>'],
      })
    );
    expect(previewNumberFormat(viewer, '"x"@')).toEqual({
      ok: true,
      text: 'Bold plain',
      sample: false,
    });
  });

  test('a refused code has its problem and no text', async () => {
    const viewer = await open(workbook());
    expect(previewNumberFormat(viewer, '"open')).toEqual({
      ok: false,
      problem: { key: 'numberFormat.problem.unreadable' },
    });
  });

  test('formats the sample with no workbook open', () => {
    const viewer = createXlsxEditor({ fonts: FONTS });
    viewers.push(viewer);
    expect(previewNumberFormat(viewer, '0.0')).toMatchObject({ ok: true, sample: true });
  });
});

/** The start of a dialog that opens now. */
const startOf = (viewer: XlsxEditorInstance) => numberFormatDialogStart(viewer)!;

describe('numberFormatDialogApplyStep', () => {
  test('an unchanged code on one cell, or on a merged cell, writes nothing', async () => {
    const viewer = await open(
      buildXlsx({ sheets: [{ name: 'S', rows: row(1, [c('A1', 5)]), merges: ['B2:C3'] }] })
    );
    expect(numberFormatDialogApplyStep(viewer, startOf(viewer), 'General')).toEqual({
      kind: 'unchanged',
    });
    viewer.select({ row: 1, col: 1 });
    expect(viewer.snapshot().selectionRef).toBe('B2:C3');
    expect(numberFormatDialogApplyStep(viewer, startOf(viewer), 'General')).toEqual({
      kind: 'unchanged',
    });
    expect(numberFormatDialogApplyStep(viewer, startOf(viewer), '0.0')).toEqual({
      kind: 'apply',
    });
  });

  test('an unchanged code on a range applies, also General over an id with no code', async () => {
    const viewer = await open(
      buildXlsx({
        sheets: [{ name: 'S', rows: row(1, [c('A1', 5, { s: 1 })]) + row(2, [c('A2', 6)]) }],
        styles: {
          cellXfs: ['<xf numFmtId="44" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/>'],
        },
      })
    );
    viewer.select({ row: 1, col: 0 }, { row: 0, col: 0 });
    const plain = startOf(viewer);
    expect(numberFormatDialogApplyStep(viewer, plain, plain.initial)).toEqual({ kind: 'apply' });
    viewer.select({ row: 0, col: 0 }, { row: 1, col: 0 });
    expect(viewer.snapshot().activeFormat?.numberFormat.id).toBe(44);
    expect(numberFormatDialogApplyStep(viewer, startOf(viewer), 'General')).toEqual({
      kind: 'apply',
    });
  });

  test('an edit after the start makes an unchanged code apply', async () => {
    const viewer = await open(workbook());
    const start = numberFormatDialogStart(viewer)!;
    runToolbarCommand(viewer, 'number.custom', '0%');
    expect(numberFormatDialogApplyStep(viewer, start, start.initial)).toEqual({ kind: 'apply' });
  });

  test('a selection that moved after the start closes the Apply', async () => {
    const viewer = await open(workbook());
    const start = numberFormatDialogStart(viewer)!;
    expect(numberFormatDialogCurrent(viewer, start)).toBe(true);
    viewer.select({ row: 1, col: 0 });
    expect(numberFormatDialogCurrent(viewer, start)).toBe(false);
    expect(numberFormatDialogApplyStep(viewer, start, start.initial)).toEqual({ kind: 'closed' });
  });

  test('a new number locale after the start closes the Apply', async () => {
    const viewer = await open(workbook());
    const start = numberFormatDialogStart(viewer)!;
    viewer.setNumberLocale({ decimal: '.', group: ',' });
    expect(numberFormatDialogCurrent(viewer, start)).toBe(true);
    viewer.setNumberLocale({ decimal: ',', group: '.' });
    expect(numberFormatDialogCurrent(viewer, start)).toBe(false);
  });

  test('after another workbook loads, the Apply closes and writes nothing', async () => {
    const viewer = await open(workbook());
    const start = startOf(viewer);
    await viewer.load(workbook());
    flushEditor(viewer);
    expect(numberFormatDialogApplyStep(viewer, start, '0.0')).toEqual({ kind: 'closed' });
  });

  test('a custom code General on a range applies; the start commits an open edit', async () => {
    const viewer = await open(
      buildXlsx({
        sheets: [
          { name: 'S', rows: row(1, [c('A1', 5, { s: 1 })]) + row(2, [c('A2', 6, { s: 2 })]) },
        ],
        styles: {
          numFmts: ['<numFmt numFmtId="164" formatCode="General"/>'],
          cellXfs: [
            '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/>',
            '<xf numFmtId="2" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/>',
          ],
        },
      })
    );
    viewer.select({ row: 0, col: 0 }, { row: 1, col: 0 });
    const start = numberFormatDialogStart(viewer)!;
    expect(start.initial).toBe('General');
    expect(numberFormatDialogApplyStep(viewer, start, 'General')).toEqual({ kind: 'apply' });
    viewer.select({ row: 2, col: 0 });
    viewer.beginEdit('0.25');
    expect(numberFormatDialogStart(viewer)?.initial).toBe('General');
    expect(viewer.snapshot().editing).toBe(false);
    expect(previewNumberFormat(viewer, '0%')).toMatchObject({ ok: true, text: '25%' });
  });

  test('a disabled slot gives no start, and an Apply closes', async () => {
    const viewer = await open(workbook());
    const start = numberFormatDialogStart(viewer)!;
    viewer.setMode('view');
    expect(numberFormatDialogStart(viewer)).toBeNull();
    expect(numberFormatDialogApplyStep(viewer, start, '0.0')).toEqual({ kind: 'closed' });
  });

  test('a code that the store refuses is its problem', async () => {
    const viewer = await open(workbook());
    expect(numberFormatDialogApplyStep(viewer, startOf(viewer), '"open')).toEqual({
      kind: 'problem',
      problem: numberFormatCodeProblem('"open')!,
    });
    expect(numberFormatDialogApplyStep(viewer, startOf(viewer), '')).toEqual({
      kind: 'problem',
      problem: numberFormatCodeProblem('')!,
    });
  });
});

describe('workbookNumberFormatCodes', () => {
  test('lists each applicable custom code once, by id', async () => {
    const viewer = await open(workbook());
    expect(workbookNumberFormatCodes(viewer)).toEqual(['"kg" #,##0', '0.000']);
  });

  test('is empty with no workbook, and capped', async () => {
    const viewer = createXlsxEditor({ fonts: FONTS });
    viewers.push(viewer);
    expect(workbookNumberFormatCodes(viewer)).toEqual([]);
    const many = Array.from(
      { length: MAX_LISTED_NUMBER_FORMATS + 5 },
      (_, i) =>
        `<numFmt numFmtId="${164 + i}" formatCode="0.${'0'.repeat((i % 20) + 1)}&quot;${i}&quot;"/>`
    );
    const big = await open(workbook(many));
    expect(workbookNumberFormatCodes(big).length).toBe(MAX_LISTED_NUMBER_FORMATS);
  });
});

describe('number.custom', () => {
  test('its value is the active cell code, and its state follows the format commands', async () => {
    const viewer = await open(workbook());
    expect(toolbarCommandState(viewer.snapshot(), 'number.custom')).toMatchObject({
      enabled: true,
      value: 'General',
    });
    viewer.setMode('view');
    expect(toolbarCommandState(viewer.snapshot(), 'number.custom').disabledReason?.key).toBe(
      'disabledReason.viewMode'
    );
  });

  test('writes one x:numFmt from id 164, reuses an equal code, and undoes in one step', async () => {
    const viewer = await open(workbook([]));
    expect(runToolbarCommand(viewer, 'number.custom', '0.0 "u"').ok).toBe(true);
    expect(viewer.snapshot().activeFormat?.numberFormat).toMatchObject({
      id: 164,
      code: '0.0 "u"',
    });
    viewer.select({ row: 1, col: 0 });
    expect(runToolbarCommand(viewer, 'number.custom', '0.0 "u"').ok).toBe(true);
    expect(viewer.snapshot().activeFormat?.numberFormat.id).toBe(164);
    const saved = viewer.save()!;
    expect(stylesOf(saved).match(/<numFmt /g)?.length).toBe(1);
    expect(stylesOf(saved)).toContain('<numFmt numFmtId="164" formatCode="0.0 &quot;u&quot;"/>');
    expect(viewer.undo().ok).toBe(true);
    expect(viewer.snapshot().activeFormat?.numberFormat.id).toBe(0);
    expect(viewer.undo().ok).toBe(true);
    viewer.select({ row: 0, col: 0 });
    expect(viewer.snapshot().activeFormat?.numberFormat.id).toBe(0);
  });

  test('a code equal to a built-in code is written as its id', async () => {
    const viewer = await open(workbook([]));
    expect(runToolbarCommand(viewer, 'number.custom', '0.00%').ok).toBe(true);
    expect(viewer.snapshot().activeFormat?.numberFormat).toMatchObject({ id: 10, code: '0.00%' });
    expect(stylesOf(viewer.save()!)).not.toContain('<numFmt ');
  });

  test('the locale code of the active built-in id writes that id back', async () => {
    const viewer = await open(
      buildXlsx({
        sheets: [{ name: 'S', rows: row(1, [c('A1', 45000, { s: 1 })]) + row(2, [c('A2', 1)]) }],
        styles: {
          cellXfs: ['<xf numFmtId="14" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/>'],
        },
      })
    );
    viewer.setNumberLocale({ decimal: ',', group: '.', builtins: { 14: 'dd.mm.yyyy' } });
    viewer.select({ row: 0, col: 0 }, { row: 1, col: 0 });
    const value = toolbarCommandState(viewer.snapshot(), 'number.custom').value!;
    expect(value).toBe('dd.mm.yyyy');
    expect(runToolbarCommand(viewer, 'number.custom', value)).toEqual({ ok: true });
    viewer.select({ row: 1, col: 0 });
    expect(viewer.snapshot().activeFormat?.numberFormat.id).toBe(14);
  });

  test('the locale code of a built-in id that the workbook redefines is a new code', async () => {
    const viewer = await open(
      buildXlsx({
        sheets: [{ name: 'S', rows: row(1, [c('A1', 45000, { s: 1 })]) }],
        styles: {
          numFmts: ['<numFmt numFmtId="14" formatCode="yyyy-mm-dd"/>'],
          cellXfs: ['<xf numFmtId="14" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/>'],
        },
      })
    );
    expect(toolbarCommandState(viewer.snapshot(), 'number.custom').value).toBe('yyyy-mm-dd');
    expect(runToolbarCommand(viewer, 'number.custom', 'mm-dd-yy')).toEqual({ ok: true });
    expect(viewer.snapshot().activeFormat?.numberFormat).toMatchObject({
      id: 164,
      code: 'mm-dd-yy',
    });
    expect(viewer.snapshot().canUndo).toBe(true);
  });

  test('an existing custom code is reused by its id', async () => {
    const viewer = await open(workbook());
    expect(runToolbarCommand(viewer, 'number.custom', '"kg" #,##0').ok).toBe(true);
    expect(viewer.snapshot().activeFormat?.numberFormat.id).toBe(165);
  });

  test('a refused code writes nothing', async () => {
    const viewer = await open(workbook());
    for (const code of ['', '"open', '0\u0001'])
      expect(runToolbarCommand(viewer, 'number.custom', code)).toMatchObject({
        ok: false,
        refusal: { code: 'bad-format' },
      });
    expect(viewer.snapshot().canUndo).toBe(false);
  });

  test('the code survives a save and a reopen, and escapes on save', async () => {
    const viewer = await open(workbook([]));
    const code = '[Blue]#,##0.00 "<&>"';
    expect(runToolbarCommand(viewer, 'number.custom', code).ok).toBe(true);
    const saved = viewer.save()!;
    expect(stylesOf(saved)).toContain('formatCode="[Blue]#,##0.00 &quot;&lt;&amp;&gt;&quot;"');
    const reopened = await open(saved);
    expect(reopened.snapshot().activeFormat?.numberFormat).toMatchObject({ id: 164, code });
    expect(previewNumberFormat(reopened, code)).toMatchObject({ text: '1,234.50 <&>' });
  });
});

describe('the dialog stylesheet', () => {
  test('each forced-colors border of the dialog comes after the rule that it overrides', () => {
    const css = readFileSync(join(import.meta.dir, '../../styles/editor.css'), 'utf8');
    const at = (text: string) => css.indexOf(text);
    const forced = (selector: string) =>
      css.lastIndexOf(
        `@media (forced-colors: active)`,
        at(`  ${selector} {\n    border-width: 2px;`)
      );
    expect(at('.xlsx-editor .xlsx-dialog__button {')).toBeLessThan(
      forced('.xlsx-editor .xlsx-dialog__button--primary')
    );
    expect(at('.xlsx-editor .xlsx-dialog__input {')).toBeLessThan(
      forced(".xlsx-editor .xlsx-dialog__input:where([aria-invalid='true'])")
    );
  });
});
