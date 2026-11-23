// Sheet protection (ECMA-376 Part 1, 18.3.1.85) in the store: on a protected sheet, a value edit
// on a locked cell, a format edit (a cell format index too), a column width, a row height, and a
// structural edit are refused with `locked`, unless the cell is unlocked or the protection
// allows the action. The password hash is kept as it is.

import { describe, expect, test } from 'bun:test';
import { buildXlsx, c, row } from '../../../../../scripts/fixtures/builder.ts';
import { readXlsx, xmlPartOf } from '../package/package.ts';
import { attr, child } from '../package/query.ts';
import { saveXlsx } from '../package-write/save.ts';
import { WorkbookDocumentStore } from '../store/document-store.ts';
import type { WorkbookOp } from '../store/ops.ts';
import {
  lockReader,
  protectionWorkTestRecorder,
  readSheetProtection,
} from '../store/sheet-protection.ts';
import { rangeBounds } from '../store/clear-range.ts';
import { cellKey, sheetIndex } from '../workbook/sheet-index.ts';
import { formatIndex } from '../workbook/inherited-format.ts';
import { styleLists, workbookIndex } from '../workbook/workbook-index.ts';

const HASH = 'algorithmName="SHA-512" hashValue="aGFzaA==" saltValue="c2FsdA==" spinCount="100000"';

/**
 * Sheet `P` is protected; B1 has an unlocked format (xf 1), and column C is unlocked by its
 * column format, and A2 is locked and hidden (xf 2). xf 3 is locked. Sheet `Free` is not
 * protected, and its A1 has the unlocked format (xf 1).
 */
function storeOf(attrs = ''): WorkbookDocumentStore {
  const read = readXlsx(
    buildXlsx({
      sheets: [
        {
          name: 'P',
          cols: '<cols><col min="3" max="3" width="9" style="1" customWidth="1"/></cols>',
          rows: row(1, [c('A1', 1), c('B1', 2, { s: 1 })]) + row(2, [c('A2', 3, { s: 2 })]),
          beforeMerges: `<sheetProtection ${HASH} sheet="1" ${attrs}/>`,
        },
        { name: 'Free', rows: row(1, [c('A1', 1, { s: 1 })]) },
      ],
      styles: {
        cellXfs: [
          '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyProtection="1">' +
            '<protection locked="0"/></xf>',
          // xf 2 hides the formula of a locked cell; xf 3 keeps the default protection.
          '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyProtection="1">' +
            '<protection hidden="1"/></xf>',
          '<xf numFmtId="2" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>',
        ],
      },
    })
  );
  if (!read.ok) throw new Error(JSON.stringify(read.error));
  return new WorkbookDocumentStore(read.pkg);
}

const set = (ref: string, sheet = 'P'): WorkbookOp => ({
  op: 'setCellValue',
  sheet,
  ref,
  value: { type: 'number', value: 9 },
});
const code = (store: WorkbookDocumentStore, ops: WorkbookOp[]) => {
  const r = store.transact(ops);
  return r.ok ? 'ok' : r.rejection.code;
};
const structural = (op: 'insertRows' | 'deleteRows' | 'insertColumns' | 'deleteColumns') =>
  ({ op, sheet: 'P', start: 5, count: 1 }) as WorkbookOp;

describe('a protected sheet', () => {
  test('refuses a value edit on a locked cell, and allows an unlocked one', () => {
    const store = storeOf();
    expect(code(store, [set('A1')])).toBe('locked');
    expect(code(store, [set('D9')])).toBe('locked');
    expect(code(store, [set('B1')])).toBe('ok');
    // A new cell in column C takes the unlocked column format.
    expect(code(store, [set('C7')])).toBe('ok');
    expect(code(store, [{ op: 'clearValues', sheet: 'P', ref: 'A1:A2' }])).toBe('locked');
    expect(code(store, [{ op: 'clearValues', sheet: 'P', ref: 'B1:C9' }])).toBe('ok');
    // A batch with one locked cell is refused as a whole.
    const before = store.revision;
    expect(code(store, [set('B1'), set('A2')])).toBe('locked');
    expect(store.revision).toBe(before);
    expect(code(store, [set('A1', 'Free')])).toBe('ok');
  });

  test('a clear refuses a locked cell whose only content is a formula inside mc:AlternateContent', () => {
    const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
    const wrapped =
      `<c r="B1"><mc:AlternateContent xmlns:mc="${MC}" xmlns:x14="urn:example:x14"><mc:Choice Requires="x14">` +
      '<f>1+1</f></mc:Choice></mc:AlternateContent></c>';
    const read = readXlsx(
      buildXlsx({
        sheets: [
          { name: 'T', rows: row(1, [wrapped]), beforeMerges: '<sheetProtection sheet="1"/>' },
        ],
      })
    );
    if (!read.ok) throw new Error(JSON.stringify(read.error));
    const result = new WorkbookDocumentStore(read.pkg).transact([
      { op: 'clearValues', sheet: 'T', ref: 'A1:C1' },
    ]);
    expect(result.ok ? undefined : result.rejection.code).toBe('locked');
    expect(result.ok ? undefined : result.rejection.detail).toContain('B1');
  });

  test('refuses a format edit unless formatCells is false', () => {
    const fmt: WorkbookOp = { op: 'applyFormat', sheet: 'P', ref: 'B1', patch: { bold: true } };
    const clear: WorkbookOp = { op: 'clearFormats', sheet: 'P', ref: 'A1' };
    expect(code(storeOf(), [fmt])).toBe('locked');
    expect(code(storeOf(), [clear])).toBe('locked');
    expect(code(storeOf('formatCells="1"'), [fmt])).toBe('locked');
    expect(code(storeOf('formatCells="0"'), [fmt])).toBe('ok');
  });

  test('refuses a cell format index unless formatCells is false', () => {
    const unlock: WorkbookOp = { op: 'setCellFormatIndex', sheet: 'P', ref: 'A1', index: 1 };
    const r = storeOf().transact([unlock]);
    expect(r.ok ? null : [r.rejection.code, r.rejection.lock]).toEqual(['locked', 'format']);
    // A new format that unlocks the cell is refused, whatever formatCells states.
    const unlocked = storeOf('formatCells="0"').transact([unlock]);
    expect(unlocked.ok ? null : [unlocked.rejection.code, unlocked.rejection.lock]).toEqual([
      'locked',
      'cell',
    ]);
    // An unlock and a write of the locked cell in one transaction: the unlock is refused, so
    // the write never reaches the cell.
    const store = storeOf();
    const before = store.revision;
    expect(code(store, [unlock, set('A1')])).toBe('locked');
    expect(code(store, [{ ...unlock, ref: 'B1', index: 0 }, set('B1')])).toBe('locked');
    expect(store.revision).toBe(before);
    // The format that a cell already shows changes nothing, so it is not refused.
    expect(code(store, [{ ...unlock, ref: 'B1', index: 1 }, set('B1')])).toBe('ok');
    expect(code(store, [{ ...unlock, ref: 'C9', index: 1 }])).toBe('ok');
  });

  test('refuses a cell format index that changes the protection of the cell, whatever formatCells states', () => {
    const allowed = storeOf('formatCells="0"');
    const to = (ref: string, index: number): WorkbookOp => ({
      op: 'setCellFormatIndex',
      sheet: 'P',
      ref,
      index,
    });
    const lockOf = (ops: WorkbookOp[]) => {
      const r = allowed.transact(ops);
      return r.ok ? 'ok' : `${r.rejection.code}:${r.rejection.lock}`;
    };
    // An unlock and a write of the locked cell in one transaction: the write never applies.
    expect(lockOf([to('A1', 1), set('A1')])).toBe('locked:cell');
    // A new cell (D4) takes a locked format: an unlocked format is refused too.
    expect(lockOf([to('D4', 1)])).toBe('locked:cell');
    // A lock of an unlocked cell, and a hidden formula.
    expect(lockOf([to('B1', 0)])).toBe('locked:cell');
    expect(lockOf([to('A1', 2)])).toBe('locked:cell');
    // A new format with the same protection is allowed.
    expect(lockOf([to('A1', 3)])).toBe('ok');
    expect(lockOf([to('C9', 1)])).toBe('ok');
  });

  test('a clear of formats keeps the protection of each cell on a protected sheet', () => {
    const store = storeOf('formatCells="0"');
    const clear = (ref: string): WorkbookOp => ({ op: 'clearFormats', sheet: 'P', ref });
    // B1 is unlocked; C2 takes the unlocked format of column C; A2 hides its formula.
    expect(code(store, [clear('A1:C2')])).toBe('ok');
    const index = workbookIndex(store.package)!;
    const part = xmlPartOf(store.package, '/xl/worksheets/sheet1.xml')!;
    const lock = lockReader(store.package, index);
    const formatOf = (ref: string) => {
      const box = rangeBounds(ref)!;
      return formatIndex(attr(sheetIndex(part).cells.get(cellKey(box[0], box[1]))?.el, 's')) ?? 0;
    };
    const state = (ref: string) => [lock.locked(formatOf(ref)), lock.hidden(formatOf(ref))];
    expect([state('A1'), state('B1'), state('C2'), state('A2')]).toEqual([
      [true, false],
      [false, false],
      [false, false],
      [true, true],
    ]);
    // xf 1 and xf 2 are format 0 with a protection of their own, so the clear reuses them.
    expect([formatOf('B1'), formatOf('C2'), formatOf('A2')]).toEqual([1, 1, 2]);
    expect(code(store, [set('B1'), set('C2')])).toBe('ok');
    // A format 0 with the protection that no record holds is appended, and undo removes it.
    const read = readXlsx(
      buildXlsx({
        sheets: [
          {
            name: 'P',
            rows: row(1, [c('A1', 1, { s: 1 })]),
            beforeMerges: '<sheetProtection sheet="1" formatCells="0"/>',
          },
        ],
        styles: {
          cellXfs: [
            '<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"' +
              ' applyProtection="1"><protection locked="0"/></xf>',
          ],
        },
      })
    );
    if (!read.ok) throw new Error('read');
    const dated = new WorkbookDocumentStore(read.pkg);
    expect(code(dated, [clear('A1')])).toBe('ok');
    const xfs = styleLists(dated.package, workbookIndex(dated.package)!).cellXfs;
    expect(xfs).toHaveLength(3);
    expect(attr(xfs[2], 'numFmtId')).toBe('0');
    expect(attr(child(xfs[2], 'x:protection'), 'locked')).toBe('0');
    expect(code(dated, [set('A1')])).toBe('ok');
    dated.undo();
    dated.undo();
    expect(styleLists(dated.package, workbookIndex(dated.package)!).cellXfs).toHaveLength(2);
    // On a sheet that is not protected, the clear gives format 0: A1 loses its `@s`.
    const free = storeOf();
    expect(code(free, [{ op: 'clearFormats', sheet: 'Free', ref: 'A1' }])).toBe('ok');
    const freePart = xmlPartOf(free.package, '/xl/worksheets/sheet2.xml')!;
    const freeA1 = sheetIndex(freePart).cells.get(cellKey(0, 0))?.el;
    expect(freeA1).toBeDefined();
    expect(attr(freeA1, 's')).toBeUndefined();
  });

  test('refuses a column width unless formatColumns is false, and a row height unless formatRows is false', () => {
    const width: WorkbookOp = { op: 'setColumnWidth', sheet: 'P', first: 1, last: 2, width: 30 };
    const height: WorkbookOp = { op: 'setRowHeight', sheet: 'P', first: 1, last: 1, height: 40 };
    const fitted: WorkbookOp = {
      op: 'setFittedRowHeights',
      sheet: 'P',
      rows: [{ row: 1, height: 40 }],
    };
    const widths: WorkbookOp = {
      op: 'setColumnWidths',
      sheet: 'P',
      columns: [{ column: 1, width: 30 }],
    };
    for (const op of [width, height, fitted, widths]) {
      const r = storeOf().transact([op]);
      expect(r.ok ? null : [r.rejection.code, r.rejection.lock]).toEqual(['locked', 'size']);
      expect(code(storeOf('formatColumns="1" formatRows="1"'), [op])).toBe('locked');
    }
    expect(code(storeOf('formatColumns="0"'), [width])).toBe('ok');
    expect(code(storeOf('formatColumns="0"'), [height])).toBe('locked');
    expect(code(storeOf('formatRows="0"'), [height])).toBe('ok');
    expect(code(storeOf('formatRows="0"'), [fitted])).toBe('ok');
    expect(code(storeOf('formatColumns="0"'), [fitted])).toBe('locked');
    expect(code(storeOf('formatColumns="0"'), [widths])).toBe('ok');
    expect(code(storeOf('formatRows="0"'), [widths])).toBe('locked');
    expect(code(storeOf('formatRows="0"'), [width])).toBe('locked');
    // The sheet that is not protected allows both.
    expect(
      code(storeOf(), [
        { ...width, sheet: 'Free' },
        { ...height, sheet: 'Free' },
      ])
    ).toBe('ok');
  });

  test('reads what the protection allows, true by default meaning locked', () => {
    const read = (attrs: string) => {
      const store = storeOf(attrs);
      const root = xmlPartOf(store.package, '/xl/worksheets/sheet1.xml')!.root;
      return readSheetProtection(root);
    };
    expect(read('')).toMatchObject({ allowFormatColumns: false, allowFormatRows: false });
    expect(read('formatColumns="false" formatRows="0"')).toMatchObject({
      allowFormatColumns: true,
      allowFormatRows: true,
    });
  });

  test('refuses a structural edit unless its attribute is false', () => {
    for (const op of ['insertRows', 'deleteRows', 'insertColumns', 'deleteColumns'] as const) {
      expect(code(storeOf(), [structural(op)])).toBe('locked');
      expect(code(storeOf(`${op}="0"`), [structural(op)])).toBe('ok');
    }
    expect(code(storeOf('insertRows="0"'), [structural('deleteRows')])).toBe('locked');
  });

  test('a cell format without protection takes the lock of its cell style', () => {
    const read = readXlsx(
      buildXlsx({
        sheets: [
          {
            name: 'P',
            rows: row(1, [c('A1', 1, { s: 1 }), c('B1', 2, { s: 2 }), c('C1', 3, { s: 3 })]),
            beforeMerges: '<sheetProtection sheet="1"/>',
          },
        ],
        styles: {
          // Style 1 is unlocked.
          cellStyleXfs: [
            '<xf numFmtId="0" fontId="0" fillId="0" borderId="0"><protection locked="0"/></xf>',
          ],
          cellXfs: [
            // No protection: the cell style's, unlocked.
            '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="1"/>',
            // `applyProtection` with no child: its own default, locked.
            '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="1" applyProtection="1"/>',
            // No protection and the Normal style: locked.
            '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>',
          ],
        },
      })
    );
    if (!read.ok) throw new Error(JSON.stringify(read.error));
    const store = new WorkbookDocumentStore(read.pkg);
    expect(code(store, [set('A1')])).toBe('ok');
    expect(code(store, [set('B1')])).toBe('locked');
    expect(code(store, [set('C1')])).toBe('locked');
  });

  test('a delete that the protection allows deletes locked cells too', () => {
    const store = storeOf('deleteRows="0"');
    expect(code(store, [{ op: 'deleteRows', sheet: 'P', start: 1, count: 1 } as WorkbookOp])).toBe(
      'ok'
    );
  });

  test('a sheet with sheet="0" is not protected', () => {
    const read = readXlsx(
      buildXlsx({
        sheets: [
          {
            name: 'P',
            rows: row(1, [c('A1', 1)]),
            beforeMerges: '<sheetProtection sheet="0"/>',
          },
        ],
      })
    );
    if (!read.ok) throw new Error('read');
    expect(code(new WorkbookDocumentStore(read.pkg), [set('A1')])).toBe('ok');
  });

  test('keeps the password hash through an edit and a save', () => {
    const store = storeOf();
    expect(code(store, [set('B1')])).toBe('ok');
    const read = readXlsx(saveXlsx(store.package));
    if (!read.ok) throw new Error('reopen');
    const root = xmlPartOf(read.pkg, '/xl/worksheets/sheet1.xml')!.root;
    const p = child(root, 'x:sheetProtection');
    expect([attr(p, 'hashValue'), attr(p, 'saltValue'), attr(p, 'spinCount')]).toEqual([
      'aGFzaA==',
      'c2FsdA==',
      '100000',
    ]);
    // Undo of an allowed edit applies.
    expect(store.undo().ok).toBe(true);
  });
});

describe('a workbook protection', () => {
  const storeWith = (attrs: string | undefined) => {
    const read = readXlsx(
      buildXlsx({
        sheets: [
          { name: 'A', rows: row(1, [c('A1', 1)]) },
          { name: 'B', rows: '' },
        ],
        workbookProtection: attrs,
      })
    );
    if (!read.ok) throw new Error('read');
    return new WorkbookDocumentStore(read.pkg);
  };
  const sheetOps: WorkbookOp[] = [
    { op: 'addSheet', name: 'C', index: 2 },
    { op: 'renameSheet', sheet: 'A', name: 'Z' },
    { op: 'deleteSheet', sheet: 'B' },
    { op: 'moveSheet', sheet: 'A', index: 1 },
  ];

  test('refuses every sheet op when lockStructure is true', () => {
    const store = storeWith(
      'workbookAlgorithmName="SHA-512" workbookHashValue="aGFzaA==" workbookSaltValue="c2FsdA==" ' +
        'workbookSpinCount="100000" lockStructure="1"'
    );
    for (const op of sheetOps) {
      const r = store.transact([op]);
      expect(r.ok ? null : [r.rejection.code, r.rejection.lock]).toEqual(['locked', 'workbook']);
    }
    expect(store.revision).toBe(0);
    // A cell edit is not a sheet op: the workbook protection does not refuse it.
    expect(code(store, [set('A1', 'A')])).toBe('ok');
  });

  test('allows sheet ops when lockStructure is false or absent, or with no element', () => {
    for (const attrs of ['lockStructure="0"', 'lockWindows="1"', undefined])
      for (const op of sheetOps) expect(code(storeWith(attrs), [op])).toBe('ok');
  });
});

// The protection checks of a transaction run once, not for each op (`protectionChecks`). The
// counts are exact on every machine, so the test needs no wall-clock limit. Sheet `P` holds the
// rules of `storeOf`: column C is unlocked by its column format (xf 1), and B1 has xf 1.
describe('the protection checks of a transaction', () => {
  const N = 2000;
  const work = protectionWorkTestRecorder();
  const writes = (sheet: string, col = 'C'): WorkbookOp[] =>
    Array.from({ length: N }, (_, i) => set(`${col}${i + 1}`, sheet));

  test('N cell writes on a protected sheet resolve it, and read its protection, once', () => {
    const store = storeOf();
    work.reset();
    expect(code(store, writes('P'))).toBe('ok');
    expect(work.sheetResolves).toBe(1);
    expect(work.protectionReads).toBe(1);
    // Every new cell of column C takes xf 1: one read of its protection.
    expect(work.formatReads).toBe(1);
    expect(work.styleReads).toBe(0);
  });

  test('a cell format with no protection is read once, with its cell style once', () => {
    const store = storeOf();
    const formatted = (s: number): WorkbookOp[] =>
      Array.from({ length: N }, (_, i) => ({
        op: 'setCellFormatIndex',
        sheet: 'P',
        ref: `E${i + 1}`,
        index: s,
      }));
    // xf 3 has no `x:protection` and no `@applyProtection`: it takes the protection of cell
    // style 0, as format 0 of each new cell in column E does. The format changes nothing of
    // the protection, so `formatCells="0"` allows it.
    const allowed = storeOf('formatCells="0"');
    work.reset();
    expect(code(allowed, formatted(3))).toBe('ok');
    expect(work.sheetResolves).toBe(1);
    expect(work.protectionReads).toBe(1);
    expect(work.formatReads).toBe(2);
    expect(work.styleReads).toBe(1);
    // The same checks refuse the same run when the sheet does not allow format edits.
    expect(code(store, formatted(3))).toBe('locked');
  });

  test('N cell writes on a sheet with no protection resolve it, and read it, once', () => {
    const store = storeOf();
    work.reset();
    expect(code(store, writes('Free', 'A'))).toBe('ok');
    expect(work.sheetResolves).toBe(1);
    expect(work.protectionReads).toBe(1);
    expect(work.formatReads).toBe(0);
  });

  test('a locked cell after N allowed writes still refuses the whole transaction', () => {
    const store = storeOf();
    const before = store.revision;
    work.reset();
    expect(code(store, [...writes('P'), set('A1')])).toBe('locked');
    expect(store.revision).toBe(before);
    expect(work.sheetResolves).toBe(1);
    // A1 has format 0, locked by default: one more cell format read.
    expect(work.formatReads).toBe(2);
  });
});
