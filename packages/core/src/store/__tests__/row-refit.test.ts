// A value edit fits its row to its content again: a row that stores `x:row/@ht` without
// `@customHeight` has a fitted height, not a height set by hand (ECMA-376 Part 1, 18.3.1.73).
// One write (`applyOp`), a run of writes (`setCellValues`, as a paste writes them), and a clear
// (`clearRange`) drop the fitted height of each row that they change, whatever the row holds, as
// a format change does (`withoutFittedHeight`). Layout then fits the row to its content after
// the edit. A row with `@customHeight="1"` keeps its height, and undo puts the fitted height
// back.

import { describe, expect, test } from 'bun:test';
import { c, row, type FixtureWorkbook } from '../../../../../scripts/fixtures/builder.ts';
import { attr, child, children } from '../package/query.ts';
import { readXlsx, type XlsxPackage } from '../package/package.ts';
import { saveXlsx } from '../package-write/save.ts';
import type { WorkbookDocumentStore } from '../store/document-store.ts';
import type { WorkbookOp } from '../store/ops.ts';
import { sheetRoot, storeOf } from './structure-helpers.ts';

/** Each `x:row` as `r:ht:customHeight`, with `-` for an absent attribute. */
const rows = (pkg: XlsxPackage) =>
  children(child(sheetRoot(pkg, 'S'), 'x:sheetData'), 'x:row').map(
    (el) => `${attr(el, 'r')}:${attr(el, 'ht') ?? '-'}:${attr(el, 'customHeight') ?? '-'}`
  );

type Sheet = FixtureWorkbook['sheets'][number];

/** Row 1 with a fitted height, row 2 with a height set by hand. */
const sheet = (extra: Partial<Sheet> = {}, styles?: FixtureWorkbook['styles']) =>
  storeOf({
    ...(styles ? { styles } : {}),
    sheets: [
      {
        name: 'S',
        rows:
          row(1, [c('A1', 1), c('B1', 2)], 'ht="30"') +
          row(2, [c('A2', 'three'), c('B2', 'four')], 'ht="40" customHeight="1"'),
        ...extra,
      },
    ],
  });

const BEFORE = ['1:30:-', '2:40:1'];
const AFTER = ['1:-:-', '2:40:1'];
const write = (ref: string, value: string | number): WorkbookOp => ({
  op: 'setCellValue',
  sheet: 'S',
  ref,
  value: typeof value === 'number' ? { type: 'number', value } : { type: 'string', value },
});
const clear = (ref: string): WorkbookOp => ({ op: 'clearValues', sheet: 'S', ref });

function apply(s: WorkbookDocumentStore, ops: WorkbookOp[]) {
  const r = s.transact(ops);
  if (!r.ok) throw new Error(JSON.stringify(r.rejection));
}

const xfWithRotation = (rotation: number) =>
  `<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment textRotation="${rotation}"/></xf>`;

describe('a value edit drops a fitted height', () => {
  test('a single write drops a fitted height and keeps a height set by hand', () => {
    const s = sheet();
    apply(s, [write('A1', 'a longer text')]);
    apply(s, [write('A2', 'a longer text')]);
    expect(rows(s.package)).toEqual(AFTER);
    expect(s.undo().ok).toBe(true);
    expect(s.undo().ok).toBe(true);
    expect(rows(s.package)).toEqual(BEFORE);
  });

  test('a run of writes, as a paste writes them, drops it in one step', () => {
    const s = sheet();
    apply(s, [write('A1', 'x'), write('B1', 'y'), write('A2', 'x'), write('B2', 'y')]);
    expect(rows(s.package)).toEqual(AFTER);
    expect(s.undo().ok).toBe(true);
    expect(rows(s.package)).toEqual(BEFORE);
  });

  test('a clear drops it on the rows that it changes', () => {
    const s = sheet();
    apply(s, [clear('A1:A2')]);
    expect(rows(s.package)).toEqual(AFTER);
    expect(s.undo().ok).toBe(true);
    expect(rows(s.package)).toEqual(BEFORE);
  });

  test('a clear that empties a row with a fitted height drops the row', () => {
    const s = sheet();
    apply(s, [clear('A1:B2')]);
    expect(rows(s.package)).toEqual(['2:40:1']);
  });

  test('a write that changes nothing keeps it, and makes no undo step', () => {
    const s = sheet();
    apply(s, [write('A1', 1)]);
    apply(s, [write('A1', 1), write('B1', 2)]);
    expect(rows(s.package)).toEqual(BEFORE);
    expect(s.canUndo).toBe(false);
  });

  test('customHeight="0" counts as fitted, and @hidden stays', () => {
    const s = storeOf({
      sheets: [
        {
          name: 'S',
          rows:
            row(1, [c('A1', 1)], 'ht="30" customHeight="0"') +
            row(2, [c('A2', 2)], 'ht="30" hidden="1"'),
        },
      ],
    });
    apply(s, [write('A1', 5), write('A2', 6)]);
    const els = children(child(sheetRoot(s.package, 'S'), 'x:sheetData'), 'x:row');
    expect(els.map((el) => attr(el, 'ht') ?? '-')).toEqual(['-', '-']);
    expect(attr(els[1], 'hidden')).toBe('1');
  });

  test('a saved workbook reopens without the fitted height of the edited row', () => {
    const s = sheet();
    apply(s, [write('B1', 'a longer text')]);
    const read = readXlsx(saveXlsx(s.package));
    if (!read.ok) throw new Error('reopen');
    expect(rows(read.pkg)).toEqual(AFTER);
  });
});

describe('every changed row drops its fitted height, whatever it holds', () => {
  const RICH = '<c r="B1" t="inlineStr"><is><r><rPr><sz val="28"/></rPr><t>Big</t></r></is></c>';
  const cases: [string, () => WorkbookDocumentStore][] = [
    [
      'a sheet with a custom default height',
      () => sheet({ formatPr: 'defaultRowHeight="15" customHeight="1"' }),
    ],
    [
      'a row with a thick bottom border',
      () => sheet({ rows: row(1, [c('A1', 1), c('B1', 2)], 'ht="30" thickBot="1"') }),
    ],
    [
      'a row that a merged range crosses',
      () =>
        sheet({
          rows: row(1, [c('A1', 'merged'), c('B1', 2)], 'ht="30"') + row(2, [c('C2', 1)]),
          merges: ['A1:A2'],
        }),
    ],
    [
      'a row with an inline string of formatted runs',
      () => sheet({ rows: row(1, [c('A1', 1), RICH], 'ht="30"') }),
    ],
    [
      'a row with rotated text',
      () =>
        sheet(
          { rows: row(1, [c('A1', 'up', { s: 1 }), c('B1', 2)], 'ht="30"') },
          { cellXfs: [xfWithRotation(90)] }
        ),
    ],
  ];
  const edits: [string, WorkbookOp[]][] = [
    ['a write', [write('B1', 'x')]],
    ['a run of writes', [write('B1', 'x'), write('C1', 'y')]],
    ['a clear', [clear('B1')]],
  ];
  for (const [name, make] of cases)
    for (const [edit, ops] of edits)
      test(`${name}: ${edit}`, () => {
        const s = make();
        apply(s, ops);
        expect(rows(s.package)[0]).toBe('1:-:-');
      });
});
