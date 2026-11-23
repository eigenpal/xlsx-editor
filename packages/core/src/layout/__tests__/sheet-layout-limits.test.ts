// Cache keys, indent in number fitting, merged ranges that overlap, the cost of a window with
// large merged ranges, and the work of a window over a small and a large sheet.

import { beforeAll, describe, expect, test } from 'bun:test';
import {
  buildXlsx,
  c,
  row,
  type FixtureWorkbook,
} from '../../../../../scripts/fixtures/builder.ts';
import { readXlsx, type XlsxPackage } from '../../store/package/package.ts';
import { layoutSheet, layoutWorkTestRecorder, type SheetLayout } from '../sheet-layout.ts';
import { PRINT_PROFILE } from '../metrics.ts';
import { mergeCoverage, mergedRanges } from '../panes.ts';
import type { TextMeasurer } from '../text/measurer.ts';
import { substituteMeasurer } from './test-fonts.ts';

let measurer: TextMeasurer;
beforeAll(async () => {
  measurer = await substituteMeasurer();
});

const VIEW = { width: 800, height: 600, scrollX: 0, scrollY: 0 };
const HUGE = { width: 1e6, height: 1e6, scrollX: 0, scrollY: 0 };

function open(wb: FixtureWorkbook): XlsxPackage {
  const r = readXlsx(buildXlsx(wb));
  if (!r.ok) throw new Error('read');
  return r.pkg;
}

const layout = (pkg: XlsxPackage, viewport = VIEW): SheetLayout =>
  layoutSheet(pkg, 'S', viewport, { measurer, profile: PRINT_PROFILE })!;

const firstLine = (l: SheetLayout, r: number, col: number) =>
  l.panes.flatMap((p) => p.texts).find((t) => t.row === r && t.col === col)!.text.lines[0]!;

describe('the geometry cache', () => {
  const sized = (sz: number) =>
    open({
      styles: {
        fonts: [`<font><sz val="${sz}"/><name val="Calibri"/></font>`],
        cellXfs: ['<xf numFmtId="0" fontId="1" fillId="0" borderId="0" applyFont="1"/>'],
      },
      sheets: [{ name: 'S', rows: row(1, [c('A1', 'x', { s: 1 })]) }],
    });

  test('a styles-only change with the same sheet part gives the new auto row height', () => {
    const small = sized(11);
    const large = sized(30);
    const smallHeight = layout(small).geometry.rows.sizeOf(0);
    const largeHeight = layout(large).geometry.rows.sizeOf(0);
    expect(largeHeight).toBeGreaterThan(smallHeight);
    // The large styles with the small package's worksheet part, which the cache has seen.
    const parts = new Map(large.parts);
    for (const [name, part] of small.parts) if (name.includes('worksheets')) parts.set(name, part);
    expect(layout({ parts }).geometry.rows.sizeOf(0)).toBe(largeHeight);
  });
});

describe('indent in number fitting', () => {
  const indented = (alignment: string) =>
    layout(
      open({
        styles: {
          cellXfs: [
            `<xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyAlignment="1"><alignment ${alignment}/></xf>`,
          ],
        },
        sheets: [{ name: 'S', rows: row(1, [c('A1', 123456.7, { s: 1 })]) }],
      })
    );
  const shown = (l: SheetLayout) =>
    firstLine(l, 0, 0)
      .spans.map((s) => s.text)
      .join('');

  test('a right-aligned number loses one digit width for each indent level', () => {
    // The text area of a default column is 48 points; one level takes 6 of them.
    const l = indented('horizontal="right" indent="1"');
    expect(shown(l)).toBe('123457');
    const line = firstLine(l, 0, 0);
    expect(line.x + line.width).toBe(53 - 2 - 6);
  });

  test('a centered number takes no indent', () => {
    expect(shown(indented('horizontal="center" indent="1"'))).toBe('123456.7');
    expect(shown(indented('horizontal="right"'))).toBe('123456.7');
  });
});

describe('overlapping merged ranges', () => {
  const merged = (merges: string[]) =>
    open({ sheets: [{ name: 'S', rows: row(1, [c('A1', 'x')]), merges }] });

  test('a merge whose top-left cell lies inside an earlier merge is skipped', () => {
    const l = layout(merged(['A1:C3', 'B2:D4', 'A4:B5', 'E1:F1']));
    expect(l.problems).toContainEqual({ code: 'merges-overlapping', detail: '1' });
    expect(l.panes[0]!.texts.map((t) => [t.row, t.col])).toEqual([[0, 0]]);
  });

  test('the merge index keeps the ranges that do not overlap', () => {
    const pkg = merged(['A1:C3', 'B2:D4', 'A4:B5', 'E1:F1']);
    const part = [...pkg.parts.values()].find((p) => p.partName.includes('worksheets'))!;
    if (part.kind !== 'xml') throw new Error('xml');
    const index = mergedRanges(part.xml.root);
    expect(index.overlapping).toBe(1);
    expect(index.ranges.map((m) => [m.firstRow, m.firstCol, m.lastRow, m.lastCol])).toEqual([
      [0, 0, 2, 2],
      [0, 4, 0, 5],
      [3, 0, 4, 1],
    ]);
  });
});

describe('merge coverage', () => {
  test('answers row by row, with ranges that nest or overlap', () => {
    const covered = mergeCoverage([
      { firstRow: 0, firstCol: 0, lastRow: 5, lastCol: 10 },
      { firstRow: 1, firstCol: 2, lastRow: 1, lastCol: 3 },
      { firstRow: 3, firstCol: 20, lastRow: 4, lastCol: 21 },
    ]);
    expect([covered(0, 10), covered(0, 11), covered(1, 3), covered(3, 21), covered(3, 22)]).toEqual(
      [true, false, true, true, false]
    );
    expect([covered(6, 0), covered(2, 5)]).toEqual([false, true]);
  });
});

describe('the cost of large merged ranges', () => {
  const columnName = (n: number) => {
    let s = '';
    for (let k = n; k >= 0; k = Math.floor(k / 26) - 1) s = String.fromCharCode(65 + (k % 26)) + s;
    return s;
  };

  test('a huge viewport over large merges stays fast', () => {
    // Each merge covers 128 columns of every row: a pane of 2000 by 2000 cells meets 17 of them.
    const merges = Array.from(
      { length: 128 },
      (_, i) => `${columnName(i * 128)}1:${columnName(i * 128 + 127)}1048576`
    );
    const pkg = open({ sheets: [{ name: 'S', rows: row(1, [c('A1', 'x')]), merges }] });
    const start = performance.now();
    const l = layout(pkg, HUGE);
    expect(performance.now() - start).toBeLessThan(1000);
    expect(l.panes[0]!.texts).toHaveLength(1);
  });

  test('thousands of overlapping whole-sheet merges stay fast', () => {
    const merges = Array.from({ length: 5000 }, (_, i) => `A${i + 1}:XFD${1048576 - i}`);
    const pkg = open({ sheets: [{ name: 'S', rows: row(1, [c('A1', 'x')]), merges }] });
    const start = performance.now();
    const l = layout(pkg, HUGE);
    expect(performance.now() - start).toBeLessThan(1000);
    expect(l.problems).toContainEqual({ code: 'merges-overlapping', detail: '4999' });
  });
});

describe('the work of a window', () => {
  const DRAWING =
    '<xdr:oneCellAnchor><xdr:from><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row>' +
    '<xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="952500" cy="952500"/><xdr:sp><xdr:nvSpPr>' +
    '<xdr:cNvPr id="2" name="S"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:prstGeom prst="rect">' +
    '<a:avLst/></a:prstGeom></xdr:spPr></xdr:sp><xdr:clientData/></xdr:oneCellAnchor>';
  /** `rows` rows of 20 numbers, a rule over them, one drawing, and with `merged` a merge per row. */
  const dense = (rows: number, merged: boolean) =>
    open({
      styles: { dxfs: ['<dxf><font><b/></font></dxf>'] },
      sheets: [
        {
          name: 'S',
          rows: Array.from({ length: rows }, (_, r) =>
            row(
              r + 1,
              Array.from({ length: 20 }, (_, k) => c(`${String.fromCharCode(65 + k)}${r + 1}`, r))
            )
          ).join(''),
          merges: merged ? Array.from({ length: rows }, (_, r) => `V${r + 1}:W${r + 1}`) : [],
          afterMerges:
            `<conditionalFormatting sqref="A1:T${rows}"><cfRule type="cellIs" dxfId="0" ` +
            'priority="1" operator="greaterThan"><formula>5</formula></cfRule></conditionalFormatting>',
          drawing: { anchors: DRAWING },
        },
      ],
    });
  const work = layoutWorkTestRecorder();
  /** The work of a second layout of `viewport`, whose geometry is cached. */
  const workOf = (pkg: XlsxPackage, viewport = VIEW) => {
    layout(pkg, viewport);
    work.reset();
    layout(pkg, viewport);
    const { searchSteps, reset: _reset, ...counts } = { ...work };
    // A binary search over n keys takes at most ceil(log2(n + 1)) steps.
    const keys = 20 * layout(pkg, viewport).geometry.rows.count;
    expect(searchSteps).toBeLessThanOrEqual(counts.searches * Math.ceil(Math.log2(keys + 1)));
    return counts;
  };

  test('is the same over a sheet of 100 rows and one of 20000 rows, for each counter', () => {
    const small = workOf(dense(100, true));
    for (const [name, count] of Object.entries(small))
      expect(`${name}: ${count > 0}`).toBe(`${name}: true`);
    expect(workOf(dense(20_000, true))).toEqual(small);
  });

  test('near the last row is the same as near the top, with no merged ranges', () => {
    const large = dense(20_000, false);
    // Rows 1000 and 19000, each with stored rows above and below the window. A merge per row is
    // left out: the merge search reads each range that starts above the window (a known cost).
    const rows = layout(large).geometry.rows;
    const near = workOf(large, { ...VIEW, scrollY: rows.offsetOf(1000) });
    expect(near.mergesRead).toBe(0);
    expect(workOf(large, { ...VIEW, scrollY: rows.offsetOf(19_000) })).toEqual(near);
  });
});
