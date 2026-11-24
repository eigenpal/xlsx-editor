// The limits of a selection of several ranges count the union of the ranges: a cell that two
// ranges hold counts once, as one command writes it once. Seeded random cases compare
// `unionCover` with a count of each cell, and `clearFormatsCreatedIn` with the cells that the
// store adds when it clears the ranges one by one.

import { describe, expect, test } from 'bun:test';
import { buildXlsx, c } from '../../../../../scripts/fixtures/builder.ts';
import { unionCover } from '../selection-ranges.ts';
import { clearFormatsCreatedIn, clearFormatsRange } from '../../store/store/format-range.ts';
import {
  formatCellRef,
  readXlsx,
  sheetIndex,
  workbookIndex,
  xmlPartOf,
  type XlsxPackage,
} from '../../store/index.ts';
import type { GridRange } from '../../contracts/index.ts';

/** A seeded generator, so a failure repeats. */
function random(seed: number): (n: number) => number {
  let s = seed;
  return (n) => {
    s = (s * 1103515245 + 12345) & 0x7fffffff;
    return s % n;
  };
}

function ranges(rnd: (n: number) => number, size: number): GridRange[] {
  return Array.from({ length: 1 + rnd(6) }, () => {
    const row = rnd(size);
    const col = rnd(size);
    return { start: { row, col }, end: { row: row + rnd(8), col: col + rnd(8) } };
  });
}

const inside = (r: GridRange, row: number, col: number) =>
  row >= r.start.row && row <= r.end.row && col >= r.start.col && col <= r.end.col;

describe('unionCover', () => {
  test('gives the area, the cells, and the bands of the union', () => {
    const rnd = random(12345);
    for (let t = 0; t < 1000; t++) {
      const list = ranges(rnd, 20);
      const cover = unionCover(list);
      let area = 0;
      for (let row = 0; row < 32; row++)
        for (let col = 0; col < 32; col++) {
          const held = list.some((r) => inside(r, row, col));
          if (held) area += 1;
          expect(cover.holds(row, col)).toBe(held);
          const banded = cover.bands.some(
            (b) =>
              row >= b.top &&
              row <= b.bottom &&
              b.runs.some(([left, right]) => col >= left && col <= right)
          );
          expect(banded).toBe(held);
        }
      expect(cover.area).toBe(area);
    }
  });
});

describe('clearFormatsCreatedIn', () => {
  test('counts the cells that the store adds when it clears every range', () => {
    const rnd = random(777);
    let counted = 0;
    for (let t = 0; t < 150; t++) {
      const cols = Array.from({ length: 3 }, () => {
        const min = 1 + rnd(12);
        return `<col min="${min}" max="${min + rnd(4)}" width="9" style="${rnd(3)}"/>`;
      });
      let rows = '';
      for (let r = 1; r <= 25; r++) {
        if (rnd(3) === 0) continue;
        const own = rnd(2) ? ` s="${rnd(3)}" customFormat="1"` : '';
        let cells = '';
        for (let col = 0; col < 16; col++)
          if (rnd(4) === 0) cells += c(formatCellRef({ row: r - 1, col }), 1, { s: rnd(3) });
        rows += `<row r="${r}"${own}>${cells}</row>`;
      }
      const xf = '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>';
      const read = readXlsx(
        buildXlsx({
          styles: { cellXfs: [xf, xf, xf] },
          sheets: [{ name: 'S', rows, cols: `<cols>${cols.join('')}</cols>` }],
        })
      );
      if (!read.ok) throw new Error('the package did not open');
      const partOf = (pkg: XlsxPackage) => xmlPartOf(pkg, workbookIndex(pkg)!.sheets[0]!.partName)!;
      const part = partOf(read.pkg);
      const list = ranges(rnd, 28);
      const got = clearFormatsCreatedIn(part.root, sheetIndex(part), unionCover(list).bands);
      // The store clears the ranges one by one, as the command does in one transaction.
      let pkg = read.pkg;
      for (const r of list) {
        const done = clearFormatsRange(
          pkg,
          'S',
          `${formatCellRef(r.start)}:${formatCellRef(r.end)}`
        );
        if (!done.ok) throw new Error(JSON.stringify(done.rejection));
        pkg = done.pkg;
      }
      const added = sheetIndex(partOf(pkg)).cells.size - sheetIndex(part).cells.size;
      expect(got).toBe(added);
      if (got > 0) counted += 1;
    }
    // The cases reach the count, not only empty unions.
    expect(counted).toBeGreaterThan(30);
  });
});
