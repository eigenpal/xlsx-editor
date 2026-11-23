// The number format commands of the toolbar: the entries of the number format list
// write their codes, the currency symbol comes from the number locale, and the decimal step
// changes one place in each number section, or is refused with a reason.

import { afterEach, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { Window } from 'happy-dom';
import { buildXlsx, c, row } from '../../../../../scripts/fixtures/builder.ts';
import {
  CUSTOM_NUMBER_FORMAT_PREFIX,
  createXlsxEditor,
  disabledReasonText,
  numberFormatChoiceCode,
  numberFormatChoices,
  numberFormatValue,
  runToolbarCommand,
  sameNumberLocale,
  toolbarCommandState,
  type NumberLocale,
  type XlsxEditorInstance,
} from '../index.ts';
import {
  NUMBER_FORMAT_CHOICES,
  MAX_STEP_DECIMALS,
  decimalStepOf,
  isNumberFormatChoice,
  numberFormatFamilyOf,
} from '../number-format-commands.ts';
import { decimalStepPreview } from './decimal-step-preview.ts';
import { DECIMAL_SELECTION_PROBLEM_KEYS, DECIMAL_STEP_PROBLEM_KEYS } from '../decimal-reasons.ts';
import { englishText } from '../english.ts';
import { flushEditor } from './flush-editor.ts';
import { builtinOrCode, generalDecimals, stepDecimals } from '../number-format-commands.ts';
import { decimalStepOps, decimalStepsDetail, selectionDecimalSteps } from '../decimal-selection.ts';
import type { FontSource } from '../../layout/index.ts';
import { readXlsx, xmlPartOf } from '../../store/package/package.ts';
import { attr, child, children } from '../../store/package/query.ts';
import { workbookIndex } from '../../store/workbook/workbook-index.ts';

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

let editors: XlsxEditorInstance[] = [];
afterEach(() => {
  for (const e of editors) e.destroy();
  editors = [];
});

async function open(
  numFmts?: readonly string[],
  numberLocale?: NumberLocale,
  cellXfs?: readonly string[]
) {
  const win = new Window();
  const doc = win.document as unknown as Document;
  const host = doc.createElement('div');
  doc.body.appendChild(host);
  const editor = createXlsxEditor({ fonts: FONTS, ...(numberLocale ? { numberLocale } : {}) });
  editors.push(editor);
  editor.attach(host);
  await editor.load(
    buildXlsx({
      sheets: [
        {
          name: 'S',
          rows: row(1, [
            c('A1', 1234.5),
            c('B1', 2),
            ...(cellXfs ? [c('D1', 1.25, { s: 1 })] : []),
          ]),
        },
      ],
      styles: { ...(numFmts ? { numFmts } : {}), ...(cellXfs ? { cellXfs } : {}) },
    })
  );
  flushEditor(editor);
  return editor;
}

const code = (e: XlsxEditorInstance) => {
  const { id, code } = e.snapshot().activeFormat!.numberFormat;
  return { id, code };
};
const EN: NumberLocale = { decimal: '.', group: ',' };

/** The `x:numFmt` records of a saved copy, as `id=code`. */
function savedNumFmts(bytes: Uint8Array): string[] {
  const read = readXlsx(bytes);
  if (!read.ok) throw new Error('the saved copy does not open');
  const root = xmlPartOf(read.pkg, workbookIndex(read.pkg)!.stylesPart)!.root;
  return children(child(root, 'x:numFmts'), 'x:numFmt').map(
    (el) => `${attr(el, 'numFmtId')}=${attr(el, 'formatCode')}`
  );
}

describe('stepDecimals', () => {
  test.each([
    ['0', '0.0', null],
    ['0.00', '0.000', '0.0'],
    ['0.0', '0.00', '0'],
    ['#,##0', '#,##0.0', null],
    ['0%', '0.0%', null],
    ['0.0%', '0.00%', '0%'],
    ['0,', '0.0,', null],
    ['#,##0.00,,"M"', '#,##0.000,,"M"', '#,##0.0,,"M"'],
    ['0.##', '0.###', '0.#'],
    ['0.00E+00', '0.000E+00', '0.0E+00'],
    ['##0.0E+0', '##0.00E+0', '##0E+0'],
    ['"$"#,##0.00', '"$"#,##0.000', '"$"#,##0.0'],
    ['0.00" a.0"', '0.000" a.0"', '0.0" a.0"'],
    ['[$€-407] #,##0.00', '[$€-407] #,##0.000', '[$€-407] #,##0.0'],
    ['[$-409]0.00', '[$-409]0.000', '[$-409]0.0'],
    ['0.0;;', '0.00;;', '0;;'],
    ['0.0e+0', '0.00e+0', '0e+0'],
    ['0.0 "m"', '0.00 "m"', '0 "m"'],
    [
      '_-* #,##0.00_-;-* #,##0.00_-;_-* "-"??_-;_-@_-',
      '_-* #,##0.000_-;-* #,##0.000_-;_-* "-"???_-;_-@_-',
      '_-* #,##0.0_-;-* #,##0.0_-;_-* "-"?_-;_-@_-',
    ],
    ['#,##0.00_);[Red](#,##0.00)', '#,##0.000_);[Red](#,##0.000)', '#,##0.0_);[Red](#,##0.0)'],
    [
      '[Blue][>=100]0.0;[Red]"x"0.00;0;@',
      '[Blue][>=100]0.00;[Red]"x"0.000;0.0;@',
      '[Blue][>=100]0;[Red]"x"0.0;0;@',
    ],
    // A zero section with a fill and another literal than "-" is no room for places.
    ['0.00;-0.00;"zero"* ', '0.000;-0.000;"zero"* ', '0.0;-0.0;"zero"* '],
    // A run of `?` beside a literal in the zero section pads it: it gets or loses one `?`.
    ['[<0]"neg";[>0]0.00;"x"??', '[<0]"neg";[>0]0.000;"x"???', '[<0]"neg";[>0]0.0;"x"?'],
    ['0.00;-0.00;"-"??', '0.000;-0.000;"-"???', '0.0;-0.0;"-"?'],
  ])('%s steps to %s and %s, and keeps the rest of the code', (from, up, down) => {
    expect(stepDecimals(from, 1)).toEqual({ ok: true, code: up });
    const r = stepDecimals(from, -1);
    if (down === null) expect(r.ok).toBe(false);
    else expect(r).toEqual({ ok: true, code: down });
  });

  test('General takes the places that it shows for the number of the cell', () => {
    expect(generalDecimals(1234.5678)).toBe(4);
    expect(generalDecimals(2)).toBe(0);
    expect(stepDecimals('General', 1, 1234.5678)).toEqual({ ok: true, code: '0.00000' });
    expect(stepDecimals('General', -1, 1234.5678)).toEqual({ ok: true, code: '0.000' });
    expect(stepDecimals('General', -1, 1.5)).toEqual({ ok: true, code: '0' });
    expect(stepDecimals('General', -1, 2)).toEqual({ ok: false, problem: 'noDecimals' });
    // A number that General shows in exponent form keeps that form, with the places of the
    // mantissa: 1e20 shows 1E+20, 1.5e-10 shows 1.5E-10, and 123456789012 shows 1.23457E+11.
    expect(stepDecimals('General', 1, 1e20)).toEqual({ ok: true, code: '0.0E+00' });
    expect(stepDecimals('General', -1, 1e20)).toEqual({ ok: false, problem: 'noDecimals' });
    expect(stepDecimals('General', 1, 1.5e-10)).toEqual({ ok: true, code: '0.00E+00' });
    expect(stepDecimals('General', -1, 1.5e-10)).toEqual({ ok: true, code: '0E+00' });
    expect(stepDecimals('General', 1, 123456789012)).toEqual({ ok: true, code: '0.000000E+00' });
    expect(stepDecimals('General', -1, 123456789012)).toEqual({ ok: true, code: '0.0000E+00' });
    // General shows -0.000012345 in fixed form, so the step stays fixed.
    expect(stepDecimals('General', 1, -0.000012345)).toEqual({ ok: true, code: '0.0000000000' });
    expect(stepDecimals('General', -1, -0.000012345)).toEqual({ ok: true, code: '0.00000000' });
    // An empty, text, or other cell that holds no number.
    expect(stepDecimals('General', 1)).toEqual({ ok: true, code: '0.0' });
    expect(stepDecimals('General', -1)).toEqual({ ok: false, problem: 'noNumber' });
  });

  test.each([
    ['General;0', 1, 'generalSection'],
    ['General;-0.00', -1, 'generalSection'],
    ['mm-dd-yy', 1, 'date'],
    ['h:mm:ss', -1, 'date'],
    ['[h]:mm:ss', 1, 'date'],
    ['mm:ss.0', 1, 'date'],
    ['@', 1, 'text'],
    ['"abc"', -1, 'text'],
    ['# ?/?', 1, 'fraction'],
    ['0.00 ?/?', 1, 'fraction'],
    ['0', -1, 'noDecimals'],
    ['"unclosed', 1, 'unreadable'],
    // A control character that a saved code cannot hold, as a tab from `&#9;` in the file.
    ['0.00\t', 1, 'unreadable'],
    ['0.00\u007f', -1, 'unreadable'],
  ] as const)('%s with step %d is refused as %s', (from, step, problem) => {
    expect(stepDecimals(from, step)).toEqual({ ok: false, problem });
  });

  test('a decrease of a section with no integer digit keeps a digit', () => {
    expect(stepDecimals('.0', -1)).toEqual({ ok: true, code: '0' });
    expect(stepDecimals('.0%', -1)).toEqual({ ok: true, code: '0%' });
    expect(stepDecimals('.00', -1)).toEqual({ ok: true, code: '.0' });
  });

  test('a long code is refused, and a step never writes more than 255 characters', () => {
    expect(stepDecimals('0'.repeat(10_000), 1)).toEqual({ ok: false, problem: 'unreadable' });
    const full = `0.0${'"x"'.repeat(84)}`;
    expect(full.length).toBe(255);
    expect(stepDecimals(full, 1)).toEqual({ ok: false, problem: 'tooLong' });
    const most = `0.${'0'.repeat(MAX_STEP_DECIMALS)}`;
    expect(stepDecimals(most, 1)).toEqual({ ok: false, problem: 'maxDecimals' });
  });

  test('a format with no known code is refused with its own reason', () => {
    expect(decimalStepOf({ id: 7, code: 'General', hasCode: false }, 1)).toEqual({
      ok: false,
      problem: 'noCode',
    });
    expect(decimalStepOf({ id: 200, code: 'General', hasCode: false }, 1)).toEqual({
      ok: false,
      problem: 'missingCode',
    });
    // An id that the workbook gives a code steps that code.
    expect(decimalStepOf({ id: 7, code: '"$"#,##0.00_);("$"#,##0.00)', hasCode: true }, 1)).toEqual(
      {
        ok: true,
        code: '"$"#,##0.000_);("$"#,##0.000)',
      }
    );
    expect(decimalStepOf({ id: 0, code: 'General', hasCode: true }, 1)).toEqual({
      ok: true,
      code: '0.0',
    });
  });
});

describe('the number format list', () => {
  test('a dash currency symbol is not taken for the zero literal of accounting', () => {
    for (const currencySymbol of ['-', '"-"']) {
      const code = numberFormatChoiceCode('accounting', {
        decimal: '.',
        group: ',',
        currencySymbol,
      });
      expect(code).toBeDefined();
      const once = stepDecimals(code!, -1);
      const twice = once.ok ? stepDecimals(once.code, -1) : once;
      const back = twice.ok ? stepDecimals(twice.code, 1) : twice;
      expect(back).toEqual(once);
    }
  });

  test('each entry has a code that the parser reads, and shows as its entry', () => {
    for (const choice of NUMBER_FORMAT_CHOICES) {
      const c = numberFormatChoiceCode(choice.value, EN)!;
      expect(c).toBeDefined();
      const id = choice.id ?? 164;
      expect(numberFormatValue({ id, code: c }, EN)).toBe(choice.value);
    }
    expect(isNumberFormatChoice('nothing')).toBe(false);
    expect(isNumberFormatChoice('currency')).toBe(true);
  });

  test('currency and accounting take the currency symbol of the number locale', () => {
    const euro: NumberLocale = { decimal: ',', group: '.', currencySymbol: '€' };
    expect(numberFormatChoiceCode('currency', EN)).toBe('"$"#,##0.00');
    expect(numberFormatChoiceCode('currency', euro)).toBe('"€"#,##0.00');
    expect(numberFormatChoiceCode('accounting', euro)).toContain('_("€"* #,##0.00_)');
    // A quotation mark in the symbol is escaped character by character.
    expect(numberFormatChoiceCode('currency', { ...euro, currencySymbol: 'a"b' })).toBe(
      '\\a\\"\\b#,##0.00'
    );
    expect(numberFormatChoiceCode('currency', { ...euro, currencySymbol: 'CHF' })).toBe(
      '"CHF"#,##0.00'
    );
  });

  test.each(['', 'x'.repeat(9), '\n', '\u0000', '\u0001', '\uD800', '￾'])(
    'a currency symbol %j that is not valid leaves currency and accounting out',
    (symbol) => {
      const locale: NumberLocale = { ...EN, currencySymbol: symbol };
      expect(numberFormatChoiceCode('currency', locale)).toBeUndefined();
      expect(numberFormatChoiceCode('accounting', locale)).toBeUndefined();
      expect(numberFormatChoices(locale).map((c) => c.value)).not.toContain('currency');
      expect(numberFormatChoiceCode('number', locale)).toBe('0.00');
    }
  );

  test('two locales with another currency symbol differ', () => {
    expect(sameNumberLocale(EN, { ...EN, currencySymbol: '$' })).toBe(true);
    expect(sameNumberLocale(EN, { ...EN, currencySymbol: '£' })).toBe(false);
  });

  test('a code that differs from an entry only in its places belongs to that entry', () => {
    const accounting = numberFormatChoiceCode('accounting', EN)!;
    const stepped = stepDecimals(accounting, 1);
    expect(stepped.ok).toBe(true);
    for (const [code, family] of [
      ['0%', 'percent'],
      ['0.0%', 'percent'],
      ['0.000E+00', 'scientific'],
      ['"$"#,##0.000', 'currency'],
      [stepped.ok ? stepped.code : '', 'accounting'],
      ['0.00%', undefined],
      ['mm-dd-yy', undefined],
      ['General', undefined],
      ['#,##0.0 "kg"', undefined],
      ['0.#%', undefined],
      ['0.??', undefined],
    ] as const)
      expect({ code, family: numberFormatFamilyOf(code, EN)?.value }).toEqual({ code, family });
  });

  test('a custom id with no record matches no entry, also not General', () => {
    expect(numberFormatValue({ id: 200, code: 'General', hasCode: false }, EN)).toBe(
      `${CUSTOM_NUMBER_FORMAT_PREFIX}General`
    );
    expect(numberFormatValue({ id: 200, code: 'General', hasCode: true }, EN)).toBe('general');
  });

  test('an entry follows the built-in codes of the locale', () => {
    const locale: NumberLocale = { decimal: ',', group: '.', builtins: { 14: 'dd.mm.yyyy' } };
    expect(numberFormatValue({ id: 14, code: 'dd.mm.yyyy' }, locale)).toBe('shortDate');
    expect(numberFormatValue({ id: 14, code: 'mm-dd-yy' }, locale)).toBe(
      `${CUSTOM_NUMBER_FORMAT_PREFIX}mm-dd-yy`
    );
    expect(numberFormatChoiceCode('shortDate', locale)).toBe('dd.mm.yyyy');
    expect(builtinOrCode(14, 'dd.mm.yyyy', new Map())).toEqual({ id: 14 });
    expect(builtinOrCode(14, 'dd.mm.yyyy', new Map([[14, 'yyyy']]))).toEqual({
      code: 'dd.mm.yyyy',
    });
  });
});

describe('the number format commands on the editor', () => {
  test('each entry applies in one undo step and shows as the value of the list', async () => {
    const editor = await open();
    // The cell starts as General: begin with Number, so each entry changes the format.
    runToolbarCommand(editor, 'number.format', 'number');
    for (const choice of NUMBER_FORMAT_CHOICES.filter((c) => c.value !== 'number')) {
      const revision = editor.snapshot().revision;
      expect(runToolbarCommand(editor, 'number.format', choice.value).ok).toBe(true);
      expect(toolbarCommandState(editor.snapshot(), 'number.format').value).toBe(choice.value);
      if (choice.id !== null) expect(code(editor).id).toBe(choice.id);
      expect(editor.snapshot().revision).toBe(revision + 1);
    }
    expect(runToolbarCommand(editor, 'history.undo').ok).toBe(true);
    expect(toolbarCommandState(editor.snapshot(), 'number.format').value).toBe('time');
  });

  test('a pick of the entry that the cell shows writes nothing and adds no undo step', async () => {
    // D1 has `numFmtId="2"` with no `applyNumberFormat`: it shows Number.
    const xf = '<xf numFmtId="2" fontId="0" fillId="0" borderId="0" xfId="0"/>';
    const editor = await open(undefined, undefined, [xf]);
    editor.select({ row: 0, col: 3 });
    expect(toolbarCommandState(editor.snapshot(), 'number.format').value).toBe('number');
    const revision = editor.snapshot().revision;
    expect(runToolbarCommand(editor, 'number.format', 'number')).toEqual({ ok: true });
    expect(editor.snapshot().revision).toBe(revision);
    expect(editor.snapshot().canUndo).toBe(false);
    // An open edit with no change closes, as on a pick of another entry.
    expect(editor.beginEdit().ok).toBe(true);
    expect(runToolbarCommand(editor, 'number.format', 'number')).toEqual({ ok: true });
    expect(editor.snapshot().editing).toBe(false);
    expect(editor.snapshot().canUndo).toBe(false);
    // Another entry still writes.
    expect(runToolbarCommand(editor, 'number.format', 'percent').ok).toBe(true);
    expect(editor.snapshot().canUndo).toBe(true);
  });

  test('a pick of the shown entry on a disabled list gives the refusal, not ok', async () => {
    const xf = '<xf numFmtId="2" fontId="0" fillId="0" borderId="0" xfId="0"/>';
    const editor = await open(undefined, undefined, [xf]);
    editor.select({ row: 0, col: 3 });
    editor.setMode('view');
    expect(toolbarCommandState(editor.snapshot(), 'number.format').disabledReason?.key).toBe(
      'disabledReason.viewMode'
    );
    const r = runToolbarCommand(editor, 'number.format', 'number');
    expect(r.ok ? null : r.refusal.code).toBe('view-mode');
  });

  test('currency writes the locale symbol, and the saved copy keeps the code', async () => {
    const editor = await open(undefined, { decimal: '.', group: ',', currencySymbol: '€' });
    expect(runToolbarCommand(editor, 'number.format', 'currency').ok).toBe(true);
    expect(code(editor)).toEqual({ id: 164, code: '"€"#,##0.00' });
    expect(savedNumFmts(editor.save()!)).toEqual(['164="€"#,##0.00']);
  });

  test('a currency symbol that is not valid refuses currency and writes nothing', async () => {
    const editor = await open(undefined, { ...EN, currencySymbol: '\u0001' });
    const r = runToolbarCommand(editor, 'number.format', 'currency');
    expect(r.ok ? null : r.refusal.code).toBe('bad-format');
    expect(editor.snapshot().canUndo).toBe(false);
  });

  test('the date entry writes the locale code when the workbook gives id 14 another code', async () => {
    const locale: NumberLocale = { ...EN, builtins: { 14: 'dd.mm.yyyy' } };
    const editor = await open(['<numFmt numFmtId="14" formatCode="yyyy"/>'], locale);
    expect(runToolbarCommand(editor, 'number.format', 'shortDate').ok).toBe(true);
    expect(code(editor)).toEqual({ id: 164, code: 'dd.mm.yyyy' });
    expect(toolbarCommandState(editor.snapshot(), 'number.format').value).toBe('shortDate');
    // With no other code for id 14, the entry writes the id.
    const plain = await open(undefined, locale);
    runToolbarCommand(plain, 'number.format', 'shortDate');
    expect(code(plain)).toEqual({ id: 14, code: 'dd.mm.yyyy' });
  });

  test('increase and decrease step the code of the active cell, one undo step each', async () => {
    const editor = await open();
    const state = (id: 'number.increaseDecimals' | 'number.decreaseDecimals') =>
      toolbarCommandState(editor.snapshot(), id);
    // General shows 1234.5: an increase shows two places, which is built-in id 2.
    expect(runToolbarCommand(editor, 'number.increaseDecimals').ok).toBe(true);
    expect(code(editor)).toEqual({ id: 2, code: '0.00' });
    expect(runToolbarCommand(editor, 'number.increaseDecimals').ok).toBe(true);
    expect(code(editor).code).toBe('0.000');
    for (let i = 0; i < 3; i++) runToolbarCommand(editor, 'number.decreaseDecimals');
    expect(code(editor)).toEqual({ id: 1, code: '0' });
    expect(state('number.decreaseDecimals').disabledReason?.key).toBe(
      'disabledReason.decimalsNoDecimals'
    );
    const refused = runToolbarCommand(editor, 'number.decreaseDecimals');
    expect(refused.ok ? null : refused.refusal.code).toBe('decimals-unchanged');
    expect(runToolbarCommand(editor, 'history.undo').ok).toBe(true);
    expect(code(editor).code).toBe('0.0');
    // An empty cell in General: decrease has no places to remove, increase gives `0.0`.
    editor.select({ row: 0, col: 2 });
    expect(state('number.decreaseDecimals').disabledReason?.key).toBe(
      'disabledReason.decimalsNoShownPlaces'
    );
    expect(runToolbarCommand(editor, 'number.increaseDecimals').ok).toBe(true);
    expect(code(editor).code).toBe('0.0');
  });

  /** A row of cells with a date, a `[Red]` code, a percent, General, and text. */
  async function mixed(rows?: string, merges?: readonly string[]) {
    const win = new Window();
    const doc = win.document as unknown as Document;
    const host = doc.createElement('div');
    doc.body.appendChild(host);
    const editor = createXlsxEditor({ fonts: FONTS });
    editors.push(editor);
    editor.attach(host);
    const xf = (id: number) =>
      `<xf numFmtId="${id}" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>`;
    await editor.load(
      buildXlsx({
        sheets: [
          {
            name: 'S',
            rows:
              rows ??
              row(1, [
                c('A1', 45000, { s: 1 }),
                c('B1', -3, { s: 2 }),
                c('C1', 0.5, { s: 3 }),
                c('D1', 1.5),
                c('E1', 'text'),
                c('F1', 2),
              ]),
            ...(merges ? { merges } : {}),
          },
        ],
        styles: {
          numFmts: ['<numFmt numFmtId="164" formatCode="[Red]0.00"/>'],
          cellXfs: [xf(14), xf(164), xf(9), xf(12)],
        },
      })
    );
    flushEditor(editor);
    return editor;
  }

  /** One General number at A1 in a column of `width` characters, with the step of each way. */
  async function narrow(value: number, width: number) {
    const win = new Window();
    const doc = win.document as unknown as Document;
    const host = doc.createElement('div');
    doc.body.appendChild(host);
    const editor = createXlsxEditor({ fonts: FONTS });
    editors.push(editor);
    editor.attach(host);
    const cols = `<cols><col min="1" max="1" width="${width}" customWidth="1"/></cols>`;
    await editor.load(buildXlsx({ sheets: [{ name: 'S', rows: row(1, [c('A1', value)]), cols }] }));
    flushEditor(editor);
    editor.select({ row: 0, col: 0 });
    const previews = ([1, -1] as const).map((step) => {
      const r = decimalStepPreview(editor.snapshot(), step);
      return r?.ok ? r.code : 'refused';
    });
    const codes = ([1, -1] as const).map((step) => {
      const r = runToolbarCommand(
        editor,
        step > 0 ? 'number.increaseDecimals' : 'number.decreaseDecimals'
      );
      const code = r.ok ? editor.snapshot().activeFormat!.numberFormat.code : 'refused';
      editor.undo();
      return code;
    });
    return { codes, previews };
  }

  /** The number format code of each cell of row 1, from A to F. */
  const rowCodes = (editor: XlsxEditorInstance) =>
    [0, 1, 2, 3, 4, 5].map((col) => {
      editor.select({ row: 0, col });
      return editor.snapshot().activeFormat!.numberFormat.code;
    });

  test('a step on a selection steps each cell from its own format, in one undo step', async () => {
    const editor = await mixed();
    const before = rowCodes(editor);
    expect(before).toEqual(['mm-dd-yy', '[Red]0.00', '0%', 'General', 'General', 'General']);
    // The active cell holds a date, which cannot step: the other cells can.
    editor.select({ row: 0, col: 0 }, { row: 0, col: 5 });
    for (const id of ['number.increaseDecimals', 'number.decreaseDecimals'] as const)
      expect(toolbarCommandState(editor.snapshot(), id).enabled).toBe(true);
    expect(runToolbarCommand(editor, 'number.increaseDecimals').ok).toBe(true);
    // The date and the text keep their formats. D1 (1.5) and F1 (2) are in other columns, so
    // each General number steps from its own places.
    expect(rowCodes(editor)).toEqual(['mm-dd-yy', '[Red]0.000', '0.0%', '0.00', 'General', '0.0']);
    expect(runToolbarCommand(editor, 'history.undo').ok).toBe(true);
    expect(rowCodes(editor)).toEqual(before);
    // A decrease changes only the cells with places to remove: F1 (2) shows none.
    editor.select({ row: 0, col: 0 }, { row: 0, col: 5 });
    expect(runToolbarCommand(editor, 'number.decreaseDecimals').ok).toBe(true);
    expect(rowCodes(editor)).toEqual(['mm-dd-yy', '[Red]0.0', '0%', '0', 'General', 'General']);
  });

  test('a selection with no cell that can step gives the reason of the active cell', async () => {
    const editor = await mixed();
    // The date at A1 and the text at E1: neither steps.
    editor.select({ row: 0, col: 0 });
    const r = runToolbarCommand(editor, 'number.increaseDecimals');
    expect(r.ok ? null : r.refusal.code).toBe('decimals-unchanged');
    expect(
      toolbarCommandState(editor.snapshot(), 'number.increaseDecimals').disabledReason?.key
    ).toBe('disabledReason.decimalsDate');
    expect(editor.snapshot().canUndo).toBe(false);
  });

  test('300 alternating General numbers step to one code in one undo step', async () => {
    const rows = Array.from({ length: 300 }, (_, k) =>
      row(k + 1, [c(`A${k + 1}`, k % 2 ? 2.25 : 1.5)])
    );
    const editor = await mixed(rows.join(''));
    const codes = () =>
      [0, 1, 298, 299].map((r) => {
        editor.select({ row: r, col: 0 });
        return editor.snapshot().activeFormat!.numberFormat.code;
      });
    editor.select({ row: 0, col: 0 }, { row: 299, col: 0 });
    expect(runToolbarCommand(editor, 'number.increaseDecimals').ok).toBe(true);
    // The active A1 (1.5) gives the code of every General number.
    expect(codes()).toEqual(['0.00', '0.00', '0.00', '0.00']);
    expect(runToolbarCommand(editor, 'history.undo').ok).toBe(true);
    expect(codes()).toEqual(['General', 'General', 'General', 'General']);
    expect(editor.snapshot().canUndo).toBe(false);
  });

  test('General cells with no number in a range give no change and a disabled step', async () => {
    const rows = [2, 3, 4, 5].map((r) => row(r, [c(`A${r}`, 'text')])).join('');
    const editor = await mixed(rows);
    editor.select({ row: 1, col: 0 }, { row: 4, col: 0 });
    for (const id of ['number.increaseDecimals', 'number.decreaseDecimals'] as const)
      expect(toolbarCommandState(editor.snapshot(), id).disabledReason?.key).toBe(
        'disabledReason.decimalsNoNumberSelection'
      );
    const r = runToolbarCommand(editor, 'number.increaseDecimals');
    expect(r.ok ? null : r.refusal.code).toBe('decimals-unchanged');
    expect(editor.snapshot().canUndo).toBe(false);
    // An empty range creates no cell either.
    editor.select({ row: 6, col: 2 }, { row: 7, col: 3 });
    expect(toolbarCommandState(editor.snapshot(), 'number.increaseDecimals').enabled).toBe(false);
    // One empty cell still gets `0.0` on an increase.
    editor.select({ row: 6, col: 2 });
    expect(runToolbarCommand(editor, 'number.increaseDecimals').ok).toBe(true);
    expect(editor.snapshot().activeFormat!.numberFormat.code).toBe('0.0');
  });

  test('with no number at the active cell, the first cell with a number names the reason', async () => {
    const reasons = async (rows: string) => {
      const editor = await mixed(rows);
      editor.select({ row: 0, col: 0 }, { row: 0, col: 1 });
      return (['number.increaseDecimals', 'number.decreaseDecimals'] as const).map(
        (id) => toolbarCommandState(editor.snapshot(), id).disabledReason?.key
      );
    };
    // Text at A1 and a date at B1 (cell format 1 is id 14).
    expect(await reasons(row(1, [c('A1', 'text'), c('B1', 45000, { s: 1 })]))).toEqual([
      'disabledReason.decimalsDateSelection',
      'disabledReason.decimalsDateSelection',
    ]);
    // An empty A1 and a fraction at B1 (cell format 4 is id 12).
    const fraction = await (async () => {
      const editor = await mixed(row(1, [c('B1', 1.5, { s: 4 })]));
      editor.select({ row: 0, col: 0 }, { row: 0, col: 1 });
      return toolbarCommandState(editor.snapshot(), 'number.increaseDecimals').disabledReason?.key;
    })();
    expect(fraction).toBe('disabledReason.decimalsFractionSelection');
    // The text names a selected cell, as the reason can come from a cell that is not active.
    const date = await mixed(row(1, [c('A1', 'text'), c('B1', 45000, { s: 1 })]));
    date.select({ row: 0, col: 0 }, { row: 0, col: 1 });
    const reason = toolbarCommandState(date.snapshot(), 'number.increaseDecimals').disabledReason!;
    expect(disabledReasonText(reason, englishText)).toBe(
      'A selected cell shows a date or a time. The editor does not change its decimal places.'
    );
    // Text only: no selected cell holds a number.
    expect(await reasons(row(1, [c('A1', 'text'), c('B1', 'more')]))).toEqual([
      'disabledReason.decimalsNoNumberSelection',
      'disabledReason.decimalsNoNumberSelection',
    ]);
  });

  test('one cell speaks of this cell, and the same problem on more cells of a selected cell', async () => {
    const editor = await mixed(row(1, [c('A1', 45000, { s: 1 }), c('B1', 45001, { s: 1 })]));
    const text = () => {
      const reason = toolbarCommandState(
        editor.snapshot(),
        'number.increaseDecimals'
      ).disabledReason!;
      return {
        single: editor.snapshot().selectionSingle,
        key: reason.key,
        text: disabledReasonText(reason, englishText),
      };
    };
    editor.select({ row: 0, col: 0 });
    expect(text()).toEqual({
      single: true,
      key: DECIMAL_STEP_PROBLEM_KEYS.date,
      text: 'This number format shows a date or a time. The editor does not change its decimal places.',
    });
    editor.select({ row: 0, col: 0 }, { row: 0, col: 1 });
    expect(text()).toEqual({
      single: false,
      key: DECIMAL_SELECTION_PROBLEM_KEYS.date,
      text: 'A selected cell shows a date or a time. The editor does not change its decimal places.',
    });
  });

  test('an open edit with changed text disables both steps, and the step is refused', async () => {
    const editor = await mixed(row(1, [c('A1', 1.5)]));
    editor.select({ row: 0, col: 0 });
    expect(editor.beginEdit().ok).toBe(true);
    expect(editor.setEditText('50%').ok).toBe(true);
    expect(editor.snapshot().editChanged).toBe(true);
    const keys = () =>
      (['number.increaseDecimals', 'number.decreaseDecimals'] as const).map(
        (id) => toolbarCommandState(editor.snapshot(), id).disabledReason?.key
      );
    expect(keys()).toEqual(['disabledReason.editPending', 'disabledReason.editPending']);
    const r = runToolbarCommand(editor, 'number.increaseDecimals');
    expect(r.ok).toBe(false);
    if (!r.ok) expect(r.refusal.code).toBe('edit-pending');
  });

  test('on a protected sheet, the state and the command give the same reason first', async () => {
    const win = new Window();
    const doc = win.document as unknown as Document;
    const host = doc.createElement('div');
    doc.body.appendChild(host);
    const editor = createXlsxEditor({ fonts: FONTS });
    editors.push(editor);
    editor.attach(host);
    // A1 is not locked, so it takes typed text, but the protection does not allow formats.
    const unlocked =
      '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyProtection="1"><protection locked="0"/></xf>';
    await editor.load(
      buildXlsx({
        sheets: [
          {
            name: 'S',
            rows: row(1, [c('A1', 2, { s: 1 })]),
            beforeMerges: '<sheetProtection sheet="1"/>',
          },
        ],
        styles: { cellXfs: [unlocked] },
      })
    );
    flushEditor(editor);
    editor.select({ row: 0, col: 0 });
    const both = () => {
      const state = toolbarCommandState(editor.snapshot(), 'number.increaseDecimals');
      const r = runToolbarCommand(editor, 'number.increaseDecimals');
      return [state.disabledReason?.key, r.ok ? 'ok' : r.refusal.code];
    };
    expect(editor.beginEdit().ok).toBe(true);
    expect(editor.setEditText('50%').ok).toBe(true);
    expect(both()).toEqual(['disabledReason.editPending', 'edit-pending']);
    editor.cancelEdit();
    expect(both()).toEqual(['disabledReason.formatLocked', 'locked']);
    // A decrease on a cell with no places to remove gives the lock too, not `decimals-unchanged`.
    const decrease = () => {
      const state = toolbarCommandState(editor.snapshot(), 'number.decreaseDecimals');
      const r = runToolbarCommand(editor, 'number.decreaseDecimals');
      return [state.disabledReason?.key, r.ok ? 'ok' : r.refusal.code];
    };
    expect(decrease()).toEqual(['disabledReason.formatLocked', 'locked']);
  });

  test('General steps from the places that a narrow column shows', async () => {
    // A 9 character column shows 3.141593: six places, so the steps write 7 and 5 places.
    const pi = await narrow(3.14159265358, 9);
    expect(pi.codes).toEqual(['0.0000000', '0.00000']);
    // The preview of a host gives the code that the command writes.
    expect(pi.previews).toEqual(pi.codes);
    // An 8 character column shows 1.23E-05: the steps keep the exponent form.
    expect((await narrow(0.0000123456789, 8)).codes).toEqual(['0.000E+00', '0.0E+00']);
    // A 20 character column shows every place that General shows at full width.
    expect((await narrow(3.14159265358, 20)).codes).toEqual(['0.0000000000', '0.00000000']);
    // With no text measurer, as headless, the step reads the full text, 3.141592654, which
    // has 9 places: a decrease writes 8, and never more places than the text shows.
    const read = readXlsx(
      buildXlsx({ sheets: [{ name: 'S', rows: row(1, [c('A1', 3.14159265358)]) }] })
    );
    if (!read.ok) throw new Error('the workbook does not open');
    const index = workbookIndex(read.pkg)!;
    const one = {
      sheetIndex: 0,
      active: { row: 0, col: 0 },
      anchor: { row: 0, col: 0 },
      range: { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } },
    };
    const r = decimalStepOps(read.pkg, index, index.sheets[0]!, one, EN, -1);
    expect(r.ok && JSON.stringify(r.ops)).toContain('"code":"0.00000000"');
    // The preview of a snapshot with no shown text reads the full text too.
    const headless = selectionDecimalSteps(read.pkg, index, index.sheets[0]!.partName, one, EN);
    expect(decimalStepsDetail(headless).shownText).toBeNull();
    const wide = await narrow(3.14159265358, 20);
    expect(wide.previews).toEqual(wide.codes);
    expect(stepDecimals('General', -1, 3.14159265358)).toEqual({ ok: true, code: '0.00000000' });
  });

  test('a code with the most places names the limit in its reason', async () => {
    const most = `0.${'0'.repeat(MAX_STEP_DECIMALS)}`;
    const editor = await open([`<numFmt numFmtId="164" formatCode="${most}"/>`], undefined, [
      '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>',
    ]);
    editor.select({ row: 0, col: 3 });
    const reason = toolbarCommandState(
      editor.snapshot(),
      'number.increaseDecimals'
    ).disabledReason!;
    expect(reason).toEqual({ key: DECIMAL_STEP_PROBLEM_KEYS.maxDecimals, params: { limit: '30' } });
    expect(disabledReasonText(reason, englishText)).toBe(
      'This number format already has the most decimal places that the editor writes: 30.'
    );
  });

  test('the preview treats a General cell with no number as the command does', async () => {
    const preview = (editor: XlsxEditorInstance, step: 1 | -1) =>
      decimalStepPreview(editor.snapshot(), step);
    for (const a1 of [[], [c('A1', 'text')]]) {
      const editor = await mixed(row(1, [...a1, c('B1', 1.5)]));
      editor.select({ row: 0, col: 0 }, { row: 0, col: 1 });
      // More cells: the step skips A1, so the preview of A1 gives `noNumber`.
      expect(preview(editor, 1)).toEqual({ ok: false, problem: 'noNumber' });
      expect(runToolbarCommand(editor, 'number.increaseDecimals').ok).toBe(true);
      editor.select({ row: 0, col: 0 });
      expect(editor.snapshot().activeFormat!.numberFormat.code).toBe('General');
      editor.select({ row: 0, col: 1 });
      expect(editor.snapshot().activeFormat!.numberFormat.code).toBe('0.00');
    }
    // One cell: an increase gives `0.0`, and a decrease has no place to remove.
    const one = await mixed(row(1, [c('B1', 1.5)]));
    one.select({ row: 0, col: 0 });
    expect(preview(one, 1)).toEqual({ ok: true, code: '0.0' });
    expect(preview(one, -1)).toEqual({ ok: false, problem: 'noPlaces' });
    const reason = toolbarCommandState(one.snapshot(), 'number.decreaseDecimals').disabledReason;
    expect(reason?.key).toBe(DECIMAL_STEP_PROBLEM_KEYS.noPlaces);
  });

  test('a merged range steps as its anchor: the cells that it covers are not read', async () => {
    // A1:B1 is merged: A1 is 0.5 with `0%`, and the covered B1 is 1.25 with `[Red]0.00`.
    const editor = await mixed(row(1, [c('A1', 0.5, { s: 3 }), c('B1', 1.25, { s: 2 })]), [
      'A1:B1',
    ]);
    editor.select({ row: 0, col: 0 });
    expect(editor.snapshot().selection!.range.end).toEqual({ row: 0, col: 1 });
    const state = toolbarCommandState(editor.snapshot(), 'number.decreaseDecimals');
    expect(state.disabledReason?.key).toBe(DECIMAL_STEP_PROBLEM_KEYS.noDecimals);
    const r = runToolbarCommand(editor, 'number.decreaseDecimals');
    expect(r.ok ? 'ok' : r.refusal.code).toBe('decimals-unchanged');
    // An increase writes the anchor only: the covered B1 keeps its cell format (index 2).
    expect(runToolbarCommand(editor, 'number.increaseDecimals').ok).toBe(true);
    expect(editor.snapshot().activeFormat!.numberFormat.code).toBe('0.0%');
    const saved = readXlsx(editor.save()!);
    if (!saved.ok) throw new Error('the saved workbook does not open');
    const sheet = xmlPartOf(saved.pkg, workbookIndex(saved.pkg)!.sheets[0]!.partName)!.root;
    const b1 = children(child(sheet, 'x:sheetData')!, 'x:row')
      .flatMap((r) => children(r, 'x:c'))
      .find((cell) => attr(cell, 'r') === 'B1');
    expect(attr(b1, 's')).toBe('2');
  });

  test('one merged range counts as one cell: an increase seeds its empty anchor', async () => {
    const editor = await mixed(row(1, [c('D1', 'note')]), ['A2:B3']);
    editor.select({ row: 1, col: 0 }, { row: 2, col: 1 });
    const s = editor.snapshot().selection!;
    expect(s.range).toEqual({ start: { row: 1, col: 0 }, end: { row: 2, col: 1 } });
    expect(
      toolbarCommandState(editor.snapshot(), 'number.decreaseDecimals').disabledReason?.key
    ).toBe('disabledReason.decimalsNoShownPlaces');
    expect(runToolbarCommand(editor, 'number.increaseDecimals').ok).toBe(true);
    editor.select({ row: 1, col: 0 });
    expect(editor.snapshot().activeFormat!.numberFormat.code).toBe('0.0');
  });

  test('one merged range whose anchor holds text acts as one text cell', async () => {
    const editor = await mixed(row(2, [c('A2', 'text')]), ['A2:B3']);
    editor.select({ row: 1, col: 0 }, { row: 2, col: 1 });
    for (const id of ['number.increaseDecimals', 'number.decreaseDecimals'] as const)
      expect(toolbarCommandState(editor.snapshot(), id).disabledReason?.key).toBe(
        'disabledReason.decimalsNoNumber'
      );
    const r = runToolbarCommand(editor, 'number.increaseDecimals');
    expect(r.ok ? null : [r.refusal.code, r.refusal.detail]).toEqual([
      'decimals-unchanged',
      'noNumbers',
    ]);
  });

  test('a whole column steps the cells that exist and creates no other cell', async () => {
    const editor = await mixed();
    editor.select({ row: 0, col: 1 }, { row: 1_048_575, col: 1 });
    // The list is disabled, because a format would create too many cells. The step is enabled,
    // and its hint says why.
    const list = toolbarCommandState(editor.snapshot(), 'number.format');
    expect(list.disabledReason?.key).toBe('disabledReason.rangeTooLarge');
    const step = toolbarCommandState(editor.snapshot(), 'number.increaseDecimals');
    expect([step.enabled, step.hint]).toEqual([true, { key: 'hint.decimalsLargeSelection' }]);
    // One cell has no hint.
    editor.select({ row: 0, col: 1 });
    expect(toolbarCommandState(editor.snapshot(), 'number.increaseDecimals').hint).toBeUndefined();
    editor.select({ row: 0, col: 1 }, { row: 1_048_575, col: 1 });
    expect(runToolbarCommand(editor, 'number.increaseDecimals').ok).toBe(true);
    const read = readXlsx(editor.save()!);
    if (!read.ok) throw new Error('the saved copy does not open');
    const sheet = xmlPartOf(read.pkg, workbookIndex(read.pkg)!.sheets[0]!.partName!)!;
    const cells = children(child(sheet.root, 'x:sheetData'), 'x:row').flatMap((r) =>
      children(r, 'x:c')
    );
    expect(cells.map((el) => attr(el, 'r'))).toEqual(['A1', 'B1', 'C1', 'D1', 'E1', 'F1']);
    expect(rowCodes(editor)[1]).toBe('[Red]0.000');
  });

  test('date, fraction, and text codes disable both steps with their reasons', async () => {
    const editor = await open();
    for (const [value, key] of [
      ['shortDate', 'disabledReason.decimalsDate'],
      ['fraction', 'disabledReason.decimalsFraction'],
      ['text', 'disabledReason.decimalsText'],
    ] as const) {
      runToolbarCommand(editor, 'number.format', value);
      for (const id of ['number.increaseDecimals', 'number.decreaseDecimals'] as const) {
        const s = toolbarCommandState(editor.snapshot(), id);
        expect({ value, id, key: s.disabledReason?.key }).toEqual({ value, id, key });
      }
    }
  });

  test('a too long result names the limit in its reason', async () => {
    const editor = await open();
    expect(editor.applyFormat({ numberFormat: { code: `0.0${'"x"'.repeat(84)}` } }).ok).toBe(true);
    expect(
      toolbarCommandState(editor.snapshot(), 'number.increaseDecimals').disabledReason
    ).toEqual({ key: 'disabledReason.decimalsTooLong', params: { limit: '255' } });
  });

  test('a built-in id with a record, and ids with no code', async () => {
    const xf = (id: number) =>
      `<xf numFmtId="${id}" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>`;
    // Id 7 with a record steps its code.
    const recorded = await open(
      ['<numFmt numFmtId="7" formatCode="&quot;$&quot;#,##0.00"/>'],
      undefined,
      [xf(7)]
    );
    recorded.select({ row: 0, col: 3 });
    expect(runToolbarCommand(recorded, 'number.increaseDecimals').ok).toBe(true);
    expect(code(recorded).code).toBe('"$"#,##0.000');
    // Id 7 with no record has no code in the locale.
    const builtin = await open(undefined, undefined, [xf(7)]);
    builtin.select({ row: 0, col: 3 });
    expect(
      toolbarCommandState(builtin.snapshot(), 'number.increaseDecimals').disabledReason?.key
    ).toBe('disabledReason.decimalsNoCode');
    // A custom id with no record.
    const missing = await open(undefined, undefined, [xf(200)]);
    missing.select({ row: 0, col: 3 });
    expect(
      toolbarCommandState(missing.snapshot(), 'number.decreaseDecimals').disabledReason?.key
    ).toBe('disabledReason.decimalsMissingCode');
  });

  test('a typed currency amount and the Currency entry write one code, and step alike', async () => {
    const editor = await open();
    const at = (row: number) => {
      editor.select({ row, col: 0 });
      return code(editor);
    };
    editor.select({ row: 2, col: 0 });
    expect(editor.enterValue('$1,234.50').ok).toBe(true);
    editor.select({ row: 3, col: 0 });
    expect(editor.enterValue('1234.5').ok).toBe(true);
    expect(runToolbarCommand(editor, 'number.format', 'currency').ok).toBe(true);
    expect(at(2)).toEqual(at(3));
    expect(at(2).code).toBe('"$"#,##0.00');
    editor.select({ row: 2, col: 0 }, { row: 3, col: 0 });
    expect(runToolbarCommand(editor, 'number.increaseDecimals').ok).toBe(true);
    expect(at(2)).toEqual(at(3));
    expect(at(2).code).toBe('"$"#,##0.000');
  });

  test('a step never writes a built-in id that the workbook gives another code', async () => {
    const editor = await open(['<numFmt numFmtId="2" formatCode="0.000"/>']);
    runToolbarCommand(editor, 'number.increaseDecimals');
    expect(code(editor).code).toBe('0.00');
    expect(code(editor).id).not.toBe(2);
    // The list entry for numbers writes the code then, so it shows two places here.
    expect(runToolbarCommand(editor, 'number.format', 'number').ok).toBe(true);
    expect(code(editor).code).toBe('0.00');
  });

  test('a conditional formatting rule that shows another number format gives the hint', async () => {
    const win = new Window();
    const doc = win.document as unknown as Document;
    const host = doc.createElement('div');
    doc.body.appendChild(host);
    const editor = createXlsxEditor({ fonts: FONTS });
    editors.push(editor);
    editor.attach(host);
    await editor.load(
      buildXlsx({
        sheets: [
          {
            name: 'S',
            rows: row(1, [c('A1', 1.5)]),
            afterMerges:
              '<conditionalFormatting sqref="A1"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>0</formula></cfRule></conditionalFormatting>',
          },
        ],
        styles: { dxfs: ['<dxf><numFmt numFmtId="200" formatCode="0.000"/></dxf>'] },
      })
    );
    flushEditor(editor);
    const s = toolbarCommandState(editor.snapshot(), 'number.increaseDecimals');
    expect(s.enabled).toBe(true);
    expect(s.hint).toEqual({ key: 'hint.conditionalFormat' });
  });
});
