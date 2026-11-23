import { describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import {
  buildXlsx,
  c,
  row,
  type FixtureWorkbook,
} from '../../../../../scripts/fixtures/builder.ts';
import { readXlsx, xmlPartOf, type XlsxPackage } from '../../store/package/package.ts';
import { sheetByName, workbookIndex } from '../../store/workbook/workbook-index.ts';
import { cellKey, sheetIndex } from '../../store/workbook/sheet-index.ts';
import { parseCellRef } from '../../store/cell-ref.ts';
import { saveXlsx } from '../../store/package-write/save.ts';
import { WorkbookDocumentStore } from '../../store/store/document-store.ts';
import type { WorkbookOp } from '../../store/store/ops.ts';
import { cellValue, sheetStale, stringItem } from '../values.ts';
import { cellFormat, richTextRuns } from '../cell-format.ts';
import { sharedStringItems } from '../../store/workbook/workbook-index.ts';

function open(bytes: Uint8Array): XlsxPackage {
  const r = readXlsx(bytes);
  if (!r.ok) throw new Error(JSON.stringify(r.error));
  return r.pkg;
}

function valueAt(pkg: XlsxPackage, sheet: string, ref: string) {
  const index = workbookIndex(pkg)!;
  const part = xmlPartOf(pkg, sheetByName(index, sheet)!.partName)!;
  const a = parseCellRef(ref)!;
  const cell = sheetIndex(part).cells.get(cellKey(a.row, a.col))!.el;
  return cellValue(pkg, index, cell, sheetStale(index, part.root));
}

const one = (rows: string, extra: Partial<FixtureWorkbook> = {}): XlsxPackage =>
  open(buildXlsx({ sheets: [{ name: 'S', rows }], ...extra }));

describe('stored values for every cell type', () => {
  const pkg = open(
    new Uint8Array(readFileSync(`${import.meta.dir}/../../../../../fixtures/basic-values.xlsx`))
  );
  test.each([
    ['A1', { kind: 'string', text: 'Shared text' }],
    ['B1', { kind: 'string', text: 'Inline text' }],
    ['A2', { kind: 'number', value: 42 }],
    ['B2', { kind: 'number', value: -3.5 }],
    ['A3', { kind: 'boolean', value: true }],
    ['B3', { kind: 'boolean', value: false }],
    ['A4', { kind: 'error', code: '#DIV/0!' }],
    ['B4', { kind: 'error', code: '#N/A' }],
    ['A5', { kind: 'number', value: 84 }],
    ['B5', { kind: 'string', text: 'Shared text' }],
  ] as const)('%s', (ref, value) => {
    const r = valueAt(pkg, 'Values', ref);
    expect(r.value).toEqual(value);
    expect(r.stale).toBe(false);
  });

  test('a formula shows its cached value and is not evaluated', () => {
    const r = valueAt(pkg, 'Values', 'A5');
    expect(r).toEqual({ value: { kind: 'number', value: 84 }, formula: true, stale: false });
  });

  test('a date cell keeps its ISO text', () => {
    const r = valueAt(one(row(1, [c('A1', '2024-03-01', { t: 'd' })])), 'S', 'A1');
    expect(r.value).toEqual({ kind: 'date', iso: '2024-03-01' });
  });

  test('a shared string index out of range shows nothing and is reported', () => {
    const r = valueAt(one(row(1, [c('A1', 5, { t: 's' })]), { strings: ['x'] }), 'S', 'A1');
    expect(r.value).toEqual({ kind: 'empty' });
    expect(r.problem).toEqual({ code: 'shared-string-out-of-range' });
  });

  test('rich text runs join their text, and escapes decode', () => {
    const pkg2 = one(row(1, [c('A1', 0, { t: 's' })]), {
      rawStrings: ['<si><r><rPr><b/></rPr><t>a_x0009_</t></r><r><t>b</t></r></si>'],
    });
    const r = valueAt(pkg2, 'S', 'A1');
    expect(r.value.kind === 'string' ? r.value.text : '').toBe('a\tb');
    expect(r.value.kind === 'string' ? r.value.runs?.length : 0).toBe(2);
  });
});

describe('strict stored numbers and shared string indexes', () => {
  test.each([
    '0x10',
    '0b1',
    '0o7',
    ' 1',
    '1 ',
    '',
    'Infinity',
    'INF',
    'NaN',
    '1e',
    '.',
    '+',
    '1e999',
  ])('x:v %p is a bad number', (raw) => {
    const r = valueAt(one(`<row r="1"><c r="A1"><v>${raw}</v></c></row>`), 'S', 'A1');
    expect(r.value).toEqual({ kind: 'empty' });
    expect(r.problem).toEqual({ code: 'bad-number' });
  });

  test.each([
    ['-1.5E+3', -1500],
    ['.5', 0.5],
    ['7.', 7],
    ['+2e-2', 0.02],
  ] as const)('x:v %p is %p', (raw, value) => {
    const r = valueAt(one(`<row r="1"><c r="A1"><v>${raw}</v></c></row>`), 'S', 'A1');
    expect(r.value).toEqual({ kind: 'number', value });
  });

  test.each(['0x0', ' 0', '0 ', '', '+0', '0.0', '1e0'])(
    'shared string index %p is refused',
    (raw) => {
      const pkg = one(`<row r="1"><c r="A1" t="s"><v>${raw}</v></c></row>`, { strings: ['x'] });
      expect(valueAt(pkg, 'S', 'A1').problem).toEqual({ code: 'shared-string-out-of-range' });
    }
  );
});

describe('stale formula results (design D3)', () => {
  const formulas = row(1, [c('A1', 2, { f: '1+1' }), c('B1', 3)]);

  test('fullCalcOnLoad on x:calcPr makes every formula result stale, and only formulas', () => {
    const pkg = one(formulas, { calcPr: 'fullCalcOnLoad="1"' });
    expect(valueAt(pkg, 'S', 'A1').stale).toBe(true);
    expect(valueAt(pkg, 'S', 'B1').stale).toBe(false);
  });

  test('fullCalcOnLoad on x:sheetCalcPr makes only that sheet stale', () => {
    const pkg = open(
      buildXlsx({
        sheets: [
          { name: 'S', rows: formulas, calcPr: 'fullCalcOnLoad="1"' },
          { name: 'T', rows: formulas },
        ],
      })
    );
    expect(valueAt(pkg, 'S', 'A1').stale).toBe(true);
    expect(valueAt(pkg, 'T', 'A1').stale).toBe(false);
  });

  test('a formula with no cached value shows nothing and is stale', () => {
    const r = valueAt(one(row(1, [c('A1', undefined, { f: '1+1' })])), 'S', 'A1');
    expect(r).toEqual({ value: { kind: 'empty' }, formula: true, stale: true });
  });

  test('an x:f inside mc:AlternateContent in a cell makes it a stale formula cell', () => {
    // The stale flag counts this formula (`partHasFormula`), so the grid marks its cached value.
    const wrapped =
      '<row r="1"><c r="A1"><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14="urn:example:x14">' +
      '<mc:Choice Requires="x14"><f>1+1</f></mc:Choice></mc:AlternateContent><v>2</v></c></row>';
    const r = valueAt(one(wrapped, { calcPr: 'fullCalcOnLoad="1"' }), 'S', 'A1');
    expect(r).toEqual({ value: { kind: 'number', value: 2 }, formula: true, stale: true });
    expect(valueAt(one(wrapped), 'S', 'A1').stale).toBe(false);
  });

  test('the branch value of a wrapped formula survives an edit, save, and reopen', () => {
    // Each cell keeps its formula and cached value only inside `mc:AlternateContent`. A value
    // edit on another sheet marks the workbook stale; the saved and reopened copy reads the same
    // branch values, and each is now stale.
    const mc = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
    const wrap = (ref: string, choice: string, fallback: string) =>
      `<c r="${ref}"><mc:AlternateContent xmlns:mc="${mc}" xmlns:x14="urn:example:x14"><mc:Choice Requires="x14">${choice}` +
      `</mc:Choice><mc:Fallback>${fallback}</mc:Fallback></mc:AlternateContent></c>`;
    const rows =
      '<row r="1">' +
      wrap('A1', '<f>1+1</f><v>2</v>', '<f>1+1</f><v>2</v>') +
      wrap('B1', '<f>3+4</f>', '<v>7</v>') +
      wrap('C1', '<f>B1*2</f><v>5</v>', '<f>OLD()</f>') +
      '</row>';
    const store = new WorkbookDocumentStore(
      open(
        buildXlsx({
          sheets: [
            { name: 'S', rows },
            { name: 'T', rows: row(1, [c('A1', 1)]) },
          ],
        })
      )
    );
    const before = ['A1', 'B1', 'C1'].map((ref) => valueAt(store.package, 'S', ref));
    expect(before).toEqual([
      { value: { kind: 'number', value: 2 }, formula: true, stale: false },
      { value: { kind: 'number', value: 7 }, formula: true, stale: false },
      { value: { kind: 'empty' }, formula: true, stale: true },
    ]);
    const edit: WorkbookOp = {
      op: 'setCellValue',
      sheet: 'T',
      ref: 'A1',
      value: { type: 'number', value: 5 },
    };
    expect(store.transact([edit]).ok).toBe(true);
    const reopened = open(saveXlsx(store.package));
    expect(workbookIndex(reopened)!.fullCalcOnLoad).toBe(true);
    const after = ['A1', 'B1', 'C1'].map((ref) => valueAt(reopened, 'S', ref));
    expect(after).toEqual(before.map((r) => ({ ...r, stale: true })));
  });

  test('the stale state survives save and reopen', () => {
    const pkg = one(formulas, { calcPr: 'fullCalcOnLoad="1"' });
    expect(valueAt(open(saveXlsx(pkg)), 'S', 'A1').stale).toBe(true);
  });
});

describe('display text', () => {
  test('a date cell shows through its date format, and General numbers use the locale', async () => {
    const { cellFormat } = await import('../cell-format.ts');
    const { displayValue } = await import('../display.ts');
    const { EN_US } = await import('../number-format/index.ts');
    const pkg = one(row(1, [c('A1', '2024-03-01', { t: 'd', s: 1 }), c('B1', 1234.5)]), {
      styles: {
        numFmts: ['<numFmt numFmtId="164" formatCode="yyyy-mm-dd"/>'],
        cellXfs: ['<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0"/>'],
      },
    });
    const index = workbookIndex(pkg)!;
    const opts = { locale: EN_US, date1904: false };
    expect(
      displayValue(valueAt(pkg, 'S', 'A1').value, cellFormat(pkg, index, '1').format, opts).text
    ).toBe('2024-03-01');
    expect(
      displayValue(valueAt(pkg, 'S', 'B1').value, cellFormat(pkg, index, undefined).format, opts)
        .text
    ).toBe('1234.5');
    expect(
      displayValue({ kind: 'error', code: '#N/A' }, cellFormat(pkg, index, undefined).format, opts)
        .text
    ).toBe('#N/A');
  });
});

describe('a rich string item with many runs', () => {
  const pkg = open(
    buildXlsx({
      sheets: [{ name: 'S', rows: row(1, [c('A1', 0, { t: 's' })]) }],
      rawStrings: ['<si><r><rPr><b/></rPr><t>one</t></r><r><t>two</t></r></si>'],
    })
  );
  const index = workbookIndex(pkg)!;
  const item = sharedStringItems(pkg, index)[0];

  test('reads once', () => {
    expect(stringItem(item)).toBe(stringItem(item));
    expect(stringItem(item).text).toBe('onetwo');
  });

  test('resolves its fonts once for each cell font', () => {
    const runs = stringItem(item).runs!;
    const font = cellFormat(pkg, index, undefined).format.font;
    const resolved = richTextRuns(runs, font, pkg, index);
    expect(richTextRuns(runs, font, pkg, index)).toBe(resolved);
    expect(resolved.map((r) => r.font.bold)).toEqual([true, font.bold]);
    expect(richTextRuns(runs, { ...font, bold: true }, pkg, index)).not.toBe(resolved);
  });
});
