// The store's guards for one cell write (tasks 7.2 and 7.3): text that XML 1.0 cannot hold
// never commits, a write that changes nothing commits nothing, `x:dimension` widens with a new
// cell, shared formula masters unshare in one pass, and the calculation side effects of
// designs D3 and D5 hold for a chain at another path and after undo and redo.

import { describe, expect, test } from 'bun:test';
import {
  buildXlsx,
  c,
  row,
  type FixtureWorkbook,
} from '../../../../../scripts/fixtures/builder.ts';
import {
  contentTypeOf,
  readXlsx,
  relationshipsOf,
  xmlPartOf,
  type XlsxPackage,
} from '../package/package.ts';
import { attr, child, textOf } from '../package/query.ts';
import { serializeXmlPart } from '../package-write/serialize.ts';
import { saveXlsx } from '../package-write/save.ts';
import { workbookIndex } from '../workbook/workbook-index.ts';
import { cellKey, sheetIndex } from '../workbook/sheet-index.ts';
import { parseCellRef } from '../cell-ref.ts';
import { WorkbookDocumentStore } from '../store/document-store.ts';
import type { CellInput, HistoryOp, WorkbookOp } from '../store/ops.ts';
import { applyHistory } from './history-helpers.ts';
import { smlElement, IdSource, plainAttr, textNode } from '../store/tree-edit.ts';

const SHEET = '/xl/worksheets/sheet1.xml';

function open(wb: FixtureWorkbook): XlsxPackage {
  const read = readXlsx(buildXlsx(wb));
  if (!read.ok) throw new Error(JSON.stringify(read.error));
  return read.pkg;
}

const cell = (pkg: XlsxPackage, ref: string) => {
  const a = parseCellRef(ref)!;
  return sheetIndex(xmlPartOf(pkg, SHEET)!).cells.get(cellKey(a.row, a.col))?.el;
};
const set = (ref: string, value: CellInput): WorkbookOp => ({
  op: 'setCellValue',
  sheet: 'S',
  ref,
  value,
});
const formula = (text: string): CellInput => ({ type: 'formula', text });
const sheetXml = (pkg: XlsxPackage) => serializeXmlPart(xmlPartOf(pkg, SHEET)!);
const codeOf = (r: ReturnType<WorkbookDocumentStore['transact']>) =>
  r.ok ? null : r.rejection.code;
const dimension = (pkg: XlsxPackage) =>
  attr(child(xmlPartOf(pkg, SHEET)!.root, 'x:dimension'), 'ref');

describe('text that XML 1.0 cannot hold', () => {
  test('a formula with such a character is refused, and nothing commits', () => {
    const store = new WorkbookDocumentStore(open({ sheets: [{ name: 'S', rows: '' }] }));
    for (const text of ['"a\u0001"', '"\ud800"', "'S\u0002'!A1", 'T[\u0003]', '1+\u0000'])
      expect(codeOf(store.transact([set('A1', formula(text))]))).toBe('invalid-character');
    expect([store.revision, store.canUndo]).toEqual([0, false]);
    expect(() => saveXlsx(store.package)).not.toThrow();
  });

  test('a string with such a character is written escaped, saves, and reads back', () => {
    const store = new WorkbookDocumentStore(open({ sheets: [{ name: 'S', rows: '' }] }));
    const text = 'a\u0001b\ud800c\u000bd';
    expect(store.transact([set('A1', { type: 'string', value: text })]).ok).toBe(true);
    const reopened = readXlsx(saveXlsx(store.package));
    expect(reopened.ok).toBe(true);
    if (!reopened.ok) return;
    const sst = xmlPartOf(reopened.pkg, workbookIndex(reopened.pkg)!.sharedStringsPart)!;
    expect(serializeXmlPart(sst)).toContain('a_x0001_b_xD800_c_x000B_d');
  });

  test('a row with such a character never commits, whatever op brings it', () => {
    const pkg = open({ sheets: [{ name: 'S', rows: row(1, [c('A1', 1)]) }] });
    const ids = new IdSource(xmlPartOf(pkg, SHEET)!);
    const v = smlElement(ids, 'v', 'cell', [], [textNode(ids, 'x\u0001')]);
    const bad = smlElement(ids, 'c', 'row', [plainAttr('r', 'A1'), plainAttr('t', 'str')], [v]);
    const node = smlElement(ids, 'row', 'sheetData', [plainAttr('r', '1')], [bad]);
    const r = applyHistory(pkg, [{ op: 'restoreRow', sheet: 'S', row: 1, node }]);
    expect(r.ok ? null : r.rejection.code).toBe('invalid-part');
  });
});

describe('formula text a user enters', () => {
  const store = () => new WorkbookDocumentStore(open({ sheets: [{ name: 'S', rows: '' }] }));

  test('R1C1 text is refused with its own reason', () => {
    const s = store();
    for (const text of ['R1C1', 'RC[1]', 'R[1]C[-1]', 'SUM(R1C1:R2C2)'])
      expect([text, codeOf(s.transact([set('A1', formula(text))]))]).toEqual([
        text,
        'r1c1-formula',
      ]);
    expect(s.transact([set('A1', formula('R1+C2'))]).ok).toBe(true);
  });

  test('text with an operator or an operand out of place is refused', () => {
    const s = store();
    for (const text of ['A1+', '+', '()', 'A1.5', '1+*2', '1 2', 'A1,B1'])
      expect([text, codeOf(s.transact([set('A1', formula(text))]))]).toEqual([
        text,
        text === 'A1.5' ? 'unreadable-formula' : 'formula-syntax',
      ]);
    expect([s.revision, s.canUndo]).toEqual([0, false]);
  });

  test('a formula in the file that breaks the grammar does not block an edit', () => {
    const s = new WorkbookDocumentStore(
      open({ sheets: [{ name: 'S', rows: row(1, [c('A1', 1, { f: 'B1+' })]) }] })
    );
    expect(s.transact([set('C1', { type: 'number', value: 2 })]).ok).toBe(true);
    expect(textOf(child(cell(s.package, 'A1')!, 'x:f'))).toBe('B1+');
  });
});

describe('a write that changes nothing', () => {
  test('an empty value on a missing cell adds no row, revision, or undo entry', () => {
    const store = new WorkbookDocumentStore(
      open({ sheets: [{ name: 'S', rows: row(1, [c('A1', 1)]) }] })
    );
    const before = store.package;
    expect(store.transact([set('C9', { type: 'empty' })]).ok).toBe(true);
    expect(store.package).toBe(before);
    expect([store.revision, store.canUndo]).toEqual([0, false]);
    expect(sheetXml(store.package)).not.toContain('r="9"');
  });

  test('the same value again adds no revision or undo entry', () => {
    const store = new WorkbookDocumentStore(
      open({
        sheets: [
          {
            name: 'S',
            rows: row(1, [
              c('A1', 1),
              '<c r="B1" t="s" s="0"><v>0</v></c>',
              c('C1', 1, { f: 'A1' }),
            ]),
          },
        ],
        strings: ['x'],
      })
    );
    for (const [ref, value] of [
      ['A1', { type: 'number', value: 1 }],
      ['B1', { type: 'string', value: 'x' }],
      ['D1', { type: 'boolean', value: true }],
      ['D1', { type: 'boolean', value: true }],
    ] as const)
      expect(store.transact([set(ref, value)]).ok).toBe(true);
    expect(store.revision).toBe(1);
    expect(store.undo().ok).toBe(true);
    expect(store.canUndo).toBe(false);
  });

  test('a date value keeps t="d" and the cell format, and a bad date is refused', () => {
    const store = new WorkbookDocumentStore(
      open({
        sheets: [{ name: 'S', rows: row(1, ['<c r="A1" s="0" t="d"><v>2024-01-02</v></c>']) }],
      })
    );
    expect(store.transact([set('A1', { type: 'date', iso: '2024-02-29T08:30:00' })]).ok).toBe(true);
    const a1 = cell(store.package, 'A1')!;
    expect([attr(a1, 't'), attr(a1, 's'), textOf(child(a1, 'x:v'))]).toEqual([
      'd',
      '0',
      '2024-02-29T08:30:00',
    ]);
    for (const iso of ['2023-02-29', '2024-13-01', '2024-1-1', '2024-01-01T25:00', 'x'])
      expect(codeOf(store.transact([set('A1', { type: 'date', iso })]))).toBe('bad-value');
  });
});

describe('x:dimension', () => {
  const wb = (dim: string): FixtureWorkbook => ({
    sheets: [{ name: 'S', dimension: dim, rows: row(2, [c('B2', 1), c('C2', 2)]) }],
  });

  test('a new cell outside the range widens it in the same undo step', () => {
    const store = new WorkbookDocumentStore(open(wb('B2:C2')));
    expect(store.transact([set('E7', { type: 'number', value: 3 })]).ok).toBe(true);
    expect(dimension(store.package)).toBe('B2:E7');
    expect(store.transact([set('A1', { type: 'string', value: 'a' })]).ok).toBe(true);
    expect(dimension(store.package)).toBe('A1:E7');
    expect(store.undo().ok).toBe(true);
    expect(dimension(store.package)).toBe('B2:E7');
    expect(store.undo().ok).toBe(true);
    expect(dimension(store.package)).toBe('B2:C2');
    expect(store.redo().ok).toBe(true);
    expect(dimension(store.package)).toBe('B2:E7');
    const reopened = readXlsx(saveXlsx(store.package));
    expect(reopened.ok && dimension(reopened.pkg)).toBe('B2:E7');
  });

  test('a cell inside the range, an edit of an existing cell, and no dimension change nothing', () => {
    const store = new WorkbookDocumentStore(open(wb('A1:C2')));
    store.transact([set('A2', { type: 'number', value: 1 })]);
    store.transact([set('C2', { type: 'number', value: 9 })]);
    expect(dimension(store.package)).toBe('A1:C2');
    const none = new WorkbookDocumentStore(
      open({ sheets: [{ name: 'S', rows: row(1, [c('A1', 1)]) }] })
    );
    none.transact([set('Z99', { type: 'number', value: 1 })]);
    expect(dimension(none.package)).toBeUndefined();
  });

  test('a single-cell range widens to a range', () => {
    const store = new WorkbookDocumentStore(open(wb('B2')));
    store.transact([set('B3', { type: 'number', value: 1 })]);
    expect(dimension(store.package)).toBe('B2:B3');
  });
});

describe('shared formulas', () => {
  test('clearing many masters unshares them in one pass, and undo and redo are one pass', () => {
    const rows: string[] = [];
    const n = 3000;
    for (let r = 1; r <= n; r += 1)
      rows.push(
        row(r, [
          c(`A${r}`, r),
          `<c r="B${r}"><f t="shared" ref="B${r}:C${r}" si="${r}">A${r}*2</f><v>1</v></c>`,
          `<c r="C${r}"><f t="shared" si="${r}"/><v>1</v></c>`,
        ])
      );
    const store = new WorkbookDocumentStore(open({ sheets: [{ name: 'S', rows: rows.join('') }] }));
    const before = sheetXml(store.package);
    let start = performance.now();
    expect(store.transact([{ op: 'clearValues', sheet: 'S', ref: `B1:B${n}` }]).ok).toBe(true);
    expect(performance.now() - start).toBeLessThan(1000);
    expect(textOf(child(cell(store.package, `C${n}`)!, 'x:f'))).toBe(`B${n}*2`);
    const cleared = sheetXml(store.package);
    expect(cleared).not.toContain('shared');
    start = performance.now();
    expect(store.undo().ok).toBe(true);
    expect(store.redo().ok).toBe(true);
    expect(store.undo().ok).toBe(true);
    expect(performance.now() - start).toBeLessThan(1000);
    expect(sheetXml(store.package)).toBe(before);
    expect(store.redo().ok).toBe(true);
    expect(sheetXml(store.package)).toBe(cleared);
  });

  test('a run of row restores in one pass equals the restores one at a time', () => {
    const pkg = open({
      sheets: [
        { name: 'S', rows: row(2, [c('A2', 1)]) + row(5, [c('A5', 2)]) + row(9, [c('A9', 3)]) },
      ],
    });
    const two = sheetIndex(xmlPartOf(pkg, SHEET)!).rows.get(1)!;
    const moved = {
      ...two,
      id: 'moved',
      attrs: two.attrs.map((a) => (a.local === 'r' ? { ...a, value: '7' } : a)),
    };
    const batch: HistoryOp[] = [
      { op: 'restoreRow', sheet: 'S', row: 5, node: null },
      { op: 'restoreRow', sheet: 'S', row: 7, node: { ...moved, children: [] } },
      { op: 'restoreRow', sheet: 'S', row: 7, node: { ...moved, children: [] } },
      { op: 'restoreRow', sheet: 'S', row: 1, node: null },
    ];
    const done = applyHistory(pkg, batch);
    if (!done.ok) throw new Error(done.rejection.detail);
    const rowsOf = (p: XlsxPackage) => [...sheetIndex(xmlPartOf(p, SHEET)!).rows.keys()];
    expect(rowsOf(done.pkg)).toEqual([1, 6, 8]);
    const undone = applyHistory(done.pkg, done.inverse);
    if (!undone.ok) throw new Error(undone.rejection.detail);
    expect(rowsOf(undone.pkg)).toEqual([1, 4, 8]);
    expect(sheetXml(undone.pkg)).toBe(sheetXml(pkg));
  });

  test('a child overwritten first, then the master edited: the other children unshare', () => {
    const rows = [1, 2, 3, 4]
      .map((r) =>
        row(r, [
          c(`A${r}`, r),
          r === 1
            ? '<c r="B1"><f t="shared" ref="B1:B4" si="0">A1*2</f><v>2</v></c>'
            : `<c r="B${r}"><f t="shared" si="0"/><v>${r * 2}</v></c>`,
        ])
      )
      .join('');
    const store = new WorkbookDocumentStore(open({ sheets: [{ name: 'S', rows }] }));
    expect(store.transact([set('B3', { type: 'number', value: 7 })]).ok).toBe(true);
    expect(store.transact([set('B1', formula('A1+1'))]).ok).toBe(true);
    expect(textOf(child(cell(store.package, 'B1')!, 'x:f'))).toBe('A1+1');
    expect(textOf(child(cell(store.package, 'B2')!, 'x:f'))).toBe('A2*2');
    expect(child(cell(store.package, 'B3')!, 'x:f')).toBeUndefined();
    expect(textOf(child(cell(store.package, 'B4')!, 'x:f'))).toBe('A4*2');
    expect(sheetXml(store.package)).not.toContain('shared');
    expect(readXlsx(saveXlsx(store.package)).ok).toBe(true);
  });
});

describe('the calculation side effects', () => {
  const CHAIN_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml';
  const CHAIN_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain';

  test('D5: a chain at another path, with a relative target, is dropped whole', () => {
    const pkg = open({
      sheets: [{ name: 'S', rows: row(1, [c('A1', 1), c('B1', 2, { f: 'A1*2' })]) }],
      extraParts: [
        [
          'xl/calc/chain.xml',
          CHAIN_TYPE,
          new TextEncoder().encode(
            '<?xml version="1.0" encoding="UTF-8"?><calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><c r="B1" i="1"/></calcChain>'
          ),
          CHAIN_REL,
        ],
      ],
    });
    expect(relationshipsOf(pkg, '/xl/workbook.xml').find((r) => r.type === CHAIN_REL)?.target).toBe(
      '/xl/calc/chain.xml'
    );
    const store = new WorkbookDocumentStore(pkg);
    expect(store.transact([set('B1', formula('A1*3'))]).ok).toBe(true);
    const saved = readXlsx(saveXlsx(store.package));
    expect(saved.ok).toBe(true);
    if (!saved.ok) return;
    expect(saved.pkg.parts.has('/xl/calc/chain.xml')).toBe(false);
    expect(relationshipsOf(saved.pkg, '/xl/workbook.xml').some((r) => r.type === CHAIN_REL)).toBe(
      false
    );
    expect(contentTypeOf(saved.pkg, '/xl/calc/chain.xml')).not.toBe(CHAIN_TYPE);
  });

  test('D3: after edit, undo, and redo, the reopened workbook is stale and holds the edit', () => {
    const store = new WorkbookDocumentStore(
      open({ sheets: [{ name: 'S', rows: row(1, [c('A1', 1), c('B1', 2, { f: 'A1*2' })]) }] })
    );
    expect(store.transact([set('A1', { type: 'number', value: 5 })]).ok).toBe(true);
    expect(store.undo().ok).toBe(true);
    expect(store.redo().ok).toBe(true);
    const saved = readXlsx(saveXlsx(store.package));
    expect(saved.ok).toBe(true);
    if (!saved.ok) return;
    expect(workbookIndex(saved.pkg)!.fullCalcOnLoad).toBe(true);
    const a1 = sheetIndex(xmlPartOf(saved.pkg, SHEET)!).cells.get(cellKey(0, 0))!.el;
    expect(textOf(child(a1, 'x:v'))).toBe('5');
  });
});
