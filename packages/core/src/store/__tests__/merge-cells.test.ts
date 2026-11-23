// The merge ops: `mergeCells` writes one `x:mergeCell` in a schema-valid `x:mergeCells`, keeps
// the top-left value, and refuses a merge that would drop a value, meet a merged range, or name
// one cell. `unmergeCells` removes each merged range that its range meets. One undo gives back
// the worksheet part byte for byte, and a saved and read package keeps the result.

import { describe, expect, test } from 'bun:test';
import { c, row, type FixtureSheet } from '../../../../../scripts/fixtures/builder.ts';
import { attr, child, children } from '../package/query.ts';
import { readXlsx, xmlPartOf, type XlsxPackage } from '../package/package.ts';
import { serializeXmlPart } from '../package-write/serialize.ts';
import { saveXlsx } from '../package-write/save.ts';
import { sheetByName, workbookIndex } from '../workbook/workbook-index.ts';
import type { WorkbookOp } from '../store/ops.ts';
import { firstValueLost, mergedRefMeeting } from '../store/merge-cells.ts';
import { formatCellRef } from '../cell-ref.ts';
import { sheetRoot, storeOf } from './structure-helpers.ts';

const partText = (pkg: XlsxPackage) =>
  serializeXmlPart(xmlPartOf(pkg, sheetByName(workbookIndex(pkg)!, 'S')!.partName)!);

/** Each `x:mergeCell/@ref`, and the `@count` of `x:mergeCells`. */
function merges(pkg: XlsxPackage) {
  const list = child(sheetRoot(pkg, 'S'), 'x:mergeCells');
  const refs = children(list, 'x:mergeCell').map((m) => attr(m, 'ref'));
  return { count: list ? attr(list, 'count') : undefined, refs };
}

/** The local names of the worksheet children, in order. */
const order = (pkg: XlsxPackage) =>
  sheetRoot(pkg, 'S').children.flatMap((n) => (n.type === 'element' ? [n.local] : []));

const store = (sheet: Partial<FixtureSheet> = {}) =>
  storeOf({
    sheets: [
      {
        name: 'S',
        rows: row(1, [c('A1', 'top'), c('C1', 5)]) + row(2, [c('A2', undefined, { s: 0 })]),
        ...sheet,
      },
    ],
  });

const merge = (ref: string): WorkbookOp => ({ op: 'mergeCells', sheet: 'S', ref });
const unmerge = (ref: string): WorkbookOp => ({ op: 'unmergeCells', sheet: 'S', ref });

/** Commit, then undo gives back the part byte for byte, and redo the edit. */
function commit(s: ReturnType<typeof store>, ops: WorkbookOp[]): XlsxPackage {
  const before = partText(s.package);
  const r = s.transact(ops);
  if (!r.ok) throw new Error(JSON.stringify(r.rejection));
  const after = partText(s.package);
  expect(s.undo().ok).toBe(true);
  expect(partText(s.package)).toBe(before);
  expect(s.redo().ok).toBe(true);
  expect(partText(s.package)).toBe(after);
  return s.package;
}

const refusal = (s: ReturnType<typeof store>, op: WorkbookOp) => {
  const before = s.package;
  const r = s.transact([op]);
  expect(s.package).toBe(before);
  return r.ok ? null : r.rejection.code;
};

describe('mergeCells', () => {
  test('adds x:mergeCells in its schema place, before x:printOptions, with @count', () => {
    const pkg = commit(store(), [merge('A1:B2')]);
    expect(merges(pkg)).toEqual({ count: '1', refs: ['A1:B2'] });
    const names = order(pkg);
    expect(names.indexOf('mergeCells')).toBe(names.indexOf('sheetData') + 1);
    expect(names.indexOf('mergeCells')).toBeLessThan(names.indexOf('printOptions'));
  });

  test('goes after x:autoFilter and before x:conditionalFormatting', () => {
    const s = store({
      beforeMerges: '<autoFilter ref="A1:C1"/>',
      afterMerges:
        '<conditionalFormatting sqref="A1"><cfRule type="expression" priority="1"><formula>1</formula></cfRule></conditionalFormatting>',
    });
    const names = order(commit(s, [merge('A1:B2')]));
    expect(names.indexOf('mergeCells')).toBe(names.indexOf('autoFilter') + 1);
    expect(names.indexOf('mergeCells')).toBeLessThan(names.indexOf('conditionalFormatting'));
  });

  test('writes the range top-left first', () => {
    expect(merges(commit(store(), [merge('B2:A1')])).refs).toEqual(['A1:B2']);
  });

  test('adds to an existing list and updates @count', () => {
    const pkg = commit(store({ merges: ['D4:E5'] }), [merge('A1:B2')]);
    expect(merges(pkg)).toEqual({ count: '2', refs: ['D4:E5', 'A1:B2'] });
  });

  test('keeps the top-left value and the empty cells with a format', () => {
    const pkg = commit(store(), [merge('A1:B2')]);
    expect(partText(pkg)).toContain('top');
  });

  test('refuses a range where another cell holds a value, and names those cells', () => {
    expect(refusal(store(), merge('A1:C1'))).toBe('merge-drops-values');
    const s = store({ rows: row(1, [c('A1', 1), c('C1', 2)]) + row(3, [c('B3', 3)]) });
    const r = s.transact([merge('A1:D4')]);
    expect(r.ok ? null : r.rejection.subject).toBe('C1');
  });

  test('refuses a range where another cell holds a formula', () => {
    const s = store({ rows: row(1, [c('A1', 1), c('B1', undefined, { f: '1+1' })]) });
    expect(refusal(s, merge('A1:B1'))).toBe('merge-drops-values');
  });

  test('refuses a range where another cell holds only a formula inside mc:AlternateContent', () => {
    const mc = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
    const wrapped =
      `<c r="B1"><mc:AlternateContent xmlns:mc="${mc}" xmlns:x14="urn:example:x14"><mc:Choice Requires="x14">` +
      '<f>1+1</f></mc:Choice></mc:AlternateContent></c>';
    const s = store({ rows: row(1, [c('A1', 1), wrapped]) });
    expect(refusal(s, merge('A1:B1'))).toBe('merge-drops-values');
  });

  test('merges after a clear of the other cells in the same transaction, as one undo step', () => {
    const s = store();
    const pkg = commit(s, [{ op: 'clearValues', sheet: 'S', ref: 'B1:C1' }, merge('A1:C1')]);
    expect(merges(pkg).refs).toEqual(['A1:C1']);
    expect(partText(pkg)).not.toContain('<v>5</v>');
  });

  test('refuses one cell, an overlap, and a bad reference', () => {
    expect(refusal(store(), merge('B2'))).toBe('merge-one-cell');
    expect(refusal(store({ merges: ['B2:C3'] }), merge('A1:B2'))).toBe('merge-overlap');
    const o = store({ merges: ['B2:C3'] }).transact([merge('A1:B2')]);
    expect(o.ok ? null : o.rejection.subject).toBe('B2:C3');
    expect(refusal(store({ merges: ['A1:B2'] }), merge('A1:B2'))).toBe('merge-overlap');
    expect(refusal(store(), merge('A1:'))).toBe('bad-reference');
    expect(refusal(store(), { op: 'mergeCells', sheet: 'T', ref: 'A1:B2' })).toBe('unknown-sheet');
  });

  test('a protected sheet refuses a merge and an unmerge with the merge rule', () => {
    const s = store({
      beforeMerges: '<sheetProtection sheet="1" formatCells="0"/>',
      merges: ['D4:E5'],
    });
    for (const op of [merge('A1:B2'), unmerge('D4:E5')]) {
      const r = s.transact([op]);
      expect(r.ok ? null : [r.rejection.code, r.rejection.lock]).toEqual(['locked', 'merge']);
    }
  });

  test('a saved and read package keeps the merge', () => {
    const pkg = commit(store(), [merge('A1:B2')]);
    const read = readXlsx(saveXlsx(pkg));
    if (!read.ok) throw new Error('reopen');
    expect(merges(read.pkg).refs).toEqual(['A1:B2']);
  });
});

test('a ref longer than any range of the grid is no merged range, and is never split', () => {
  const s = store({ merges: ['A1:B2' + ':'.repeat(5000)] });
  expect(mergedRefMeeting(sheetRoot(s.package, 'S'), [0, 0, 9, 9])).toBeUndefined();
  expect(s.transact([merge('A1:B2')]).ok).toBe(true);
});

test('firstValueLost gives the first cell in reading order, with and without an open edit', () => {
  // Values in a sparse pattern: rows 1, 3, and 40, and a far cell.
  const s = store({
    rows:
      row(1, [c('A1', 1), c('C1', 2)]) +
      row(3, [c('B3', 3), c('D3', undefined, { s: 0 })]) +
      row(40, [c('A40', 4)]) +
      row(900, [c('Z900', 5)]),
  });
  const part = xmlPartOf(s.package, sheetByName(workbookIndex(s.package)!, 'S')!.partName)!;
  const has = new Set(['A1', 'C1', 'B3', 'A40', 'Z900']);
  const naive = (t: number, l: number, b: number, r: number, edit?: [number, number, boolean]) => {
    for (let row = t; row <= Math.min(b, 1000); row++)
      for (let col = l; col <= Math.min(r, 30); col++) {
        if (row === t && col === l) continue;
        const ref = formatCellRef({ row, col });
        const held = edit && edit[0] === row && edit[1] === col ? edit[2] : has.has(ref);
        if (held) return ref;
      }
    return undefined;
  };
  const boxes: [number, number, number, number][] = [
    [0, 0, 2, 3],
    [0, 1, 1_048_575, 1],
    [0, 0, 1_048_575, 16_383],
    [1, 0, 39, 2],
    [2, 1, 2, 3],
    [0, 2, 1_048_575, 30],
  ];
  for (const box of boxes)
    for (const edit of [undefined, [2, 3, true], [0, 2, false], [5, 1, true]] as const) {
      const override = edit && { row: edit[0], col: edit[1], holds: edit[2] };
      expect(firstValueLost(part, box, override)).toBe(naive(...box, edit as never));
    }
});

test('firstValueLost agrees with a plain scan on random boxes and edits', () => {
  let seed = 7;
  const rnd = (n: number) => (seed = (seed * 1103515245 + 12345) % 2 ** 31) % n;
  const values = new Map<number, Set<number>>();
  for (let i = 0; i < 80; i++) {
    const r = rnd(40);
    (values.get(r) ?? values.set(r, new Set()).get(r)!).add(rnd(12));
  }
  const rows = [...values.keys()]
    .sort((a, b) => a - b)
    .map((r) =>
      row(
        r + 1,
        [...values.get(r)!].sort((a, b) => a - b).map((col) => c(formatCellRef({ row: r, col }), 1))
      )
    )
    .join('');
  const s = store({ rows });
  const part = xmlPartOf(s.package, sheetByName(workbookIndex(s.package)!, 'S')!.partName)!;
  for (let i = 0; i < 400; i++) {
    const t = rnd(42);
    const l = rnd(13);
    const box = [t, l, t + rnd(42), l + rnd(14)] as const;
    const edit = rnd(3) === 0 ? { row: rnd(42), col: rnd(14), holds: rnd(2) === 0 } : undefined;
    let want: string | undefined;
    for (let r = box[0]; r <= box[2] && !want; r++)
      for (let col = box[1]; col <= box[3] && !want; col++) {
        if (r === box[0] && col === box[1]) continue;
        const held =
          edit && edit.row === r && edit.col === col ? edit.holds : !!values.get(r)?.has(col);
        if (held) want = formatCellRef({ row: r, col });
      }
    expect(firstValueLost(part, box, edit)).toBe(want);
  }
});

describe('unmergeCells', () => {
  test('removes each merged range that the range meets, and keeps the others', () => {
    const pkg = commit(store({ merges: ['A1:B2', 'D4:E5', 'G1:H1'] }), [unmerge('B2:D4')]);
    expect(merges(pkg)).toEqual({ count: '1', refs: ['G1:H1'] });
  });

  test('the last unmerge removes the list, also with other children, so the part stays valid', () => {
    const s = store({
      beforeMerges:
        '<mergeCells count="1"><mergeCell ref="A1:B2"/><e:keep xmlns:e="urn:probe"/></mergeCells>',
    });
    const pkg = commit(s, [unmerge('A1')]);
    expect(order(pkg)).not.toContain('mergeCells');
    expect(partText(pkg)).not.toContain('urn:probe');
  });

  test('one cell of a merged range unmerges the whole range, and the empty list goes', () => {
    const pkg = commit(store({ merges: ['A1:B2'] }), [unmerge('B2')]);
    expect(order(pkg)).not.toContain('mergeCells');
  });

  test('a range with no merged range changes nothing and makes no undo step', () => {
    const s = store({ merges: ['A1:B2'] });
    const before = s.package;
    expect(s.transact([unmerge('D4')]).ok).toBe(true);
    expect(s.package).toBe(before);
    expect(s.canUndo).toBe(false);
  });
});
