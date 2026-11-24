import { beforeAll, describe, expect, spyOn, test } from 'bun:test';
import * as conditional from '../../cascade/conditional.ts';
import * as inherited from '../../store/workbook/inherited-format.ts';
import { readFileSync } from 'node:fs';
import {
  buildXlsx,
  c,
  row,
  type FixtureWorkbook,
} from '../../../../../scripts/fixtures/builder.ts';
import { readXlsx, type XlsxPackage } from '../../store/package/package.ts';
import { formatCellRef } from '../../store/cell-ref.ts';
import { saveXlsx, WorkbookDocumentStore } from '../../store/index.ts';
import { layoutSheet, type CellTextBox, type SheetLayout } from '../sheet-layout.ts';
import { PRINT_PROFILE, SCREEN_PROFILE } from '../metrics.ts';
import type { TextMeasurer } from '../text/measurer.ts';
import { substituteMeasurer } from './test-fonts.ts';
import { referenceLocale } from '../../../../../scripts/fidelity/references.ts';
import { fixtureFile } from '../../../../../scripts/fixtures/catalog.ts';

let measurer: TextMeasurer;
beforeAll(async () => {
  measurer = await substituteMeasurer();
});

const open = (bytes: Uint8Array): XlsxPackage => {
  const r = readXlsx(bytes);
  if (!r.ok) throw new Error('read');
  return r.pkg;
};
const fixture = (name: string) =>
  open(
    new Uint8Array(readFileSync(`${import.meta.dir}/../../../../../fixtures/${fixtureFile(name)}`))
  );

const VIEW = { width: 540, height: 720, scrollX: 0, scrollY: 0 };

function layout(pkg: XlsxPackage, sheet: string, viewport = VIEW, screen = false): SheetLayout {
  const out = layoutSheet(pkg, sheet, viewport, {
    measurer,
    profile: screen ? SCREEN_PROFILE : PRINT_PROFILE,
  });
  if (!out) throw new Error('no sheet');
  return out;
}

const texts = (l: SheetLayout): CellTextBox[] => l.panes.flatMap((p) => p.texts);
const textAt = (l: SheetLayout, r: number, col: number) =>
  texts(l).find((t) => t.row === r && t.col === col);
const shown = (t: CellTextBox | undefined) =>
  t?.text.lines.map((line) => line.spans.map((s) => s.text).join('')).join('\n');

function sheetOf(
  sheet: Omit<FixtureWorkbook['sheets'][number], 'name'>,
  wb: Partial<FixtureWorkbook> = {}
) {
  return open(buildXlsx({ ...wb, sheets: [{ name: 'S', ...sheet }] }));
}

describe('basic-values.pdf: the printed grid starts at the 36 point margin', () => {
  // Each expected value is a glyph origin of the reference PDF minus 36, the grid origin.
  test('rows are 14 points, and text sits on the measured baselines', () => {
    const l = layout(fixture('basic-values'), 'Values');
    expect(l.geometry.rows.sizeOf(0)).toBe(14);
    expect(l.geometry.columns.sizeOf(0)).toBe(96);
    expect(l.geometry.columns.sizeOf(1)).toBe(53);
    const a1 = textAt(l, 0, 0)!;
    expect([a1.text.lines[0]!.x, a1.text.lines[0]!.baseline]).toEqual([39 - 36, 47 - 36]);
  });

  test('general alignment: text left, numbers right, booleans and errors centered', () => {
    const l = layout(fixture('basic-values'), 'Values');
    const x = (r: number, col: number) => textAt(l, r, col)!.text.lines[0]!.x + 36;
    expect(x(0, 1)).toBe(135); // Inline text
    expect(x(1, 0)).toBe(118); // 42
    expect(x(2, 0)).toBe(73); // TRUE
    expect(x(2, 1)).toBe(146); // FALSE
    expect(x(3, 0)).toBe(67); // #DIV/0!
    expect(x(3, 1)).toBe(148); // #N/A
    expect(x(4, 0)).toBe(118); // 84
    expect(textAt(l, 1, 0)!.text.lines[0]!.baseline + 36).toBe(61);
  });

  test('a General number that does not fit shows the longest rounding that fits', () => {
    const l = layout(fixture('basic-values'), 'Values');
    const c2 = textAt(l, 1, 2)!;
    expect(shown(c2)).toBe('1234568');
    expect(c2.text.lines[0]!.x + 36).toBe(194);
  });

  test('number-formats.pdf: right-aligned numbers end 3 points inside the cell', () => {
    const l = layout(fixture('number-formats'), 'Formats');
    const b2 = textAt(l, 1, 1)!;
    expect(b2.text.lines[0]!.x + b2.text.lines[0]!.width + 36).toBe(286);
  });

  // The PDF x of the first glyph that is not a space, for each case cell (scripts/fixtures/cases.ts).
  const glyphX = (l: SheetLayout, ref: string): number => {
    const col = ref.charCodeAt(0) - 65;
    const span = textAt(l, Number(ref.slice(1)) - 1, col)!.text.lines[0]!.spans[0]!;
    const i = span.text.search(/\S/);
    return span.positions[i]! + 36;
  };
  const withLocale = (name: string) => {
    const locale = referenceLocale(name);
    if (!locale) throw new Error(`${name}: no reference record`);
    const out = layoutSheet(fixture(name), 'Formats', VIEW, {
      measurer,
      profile: PRINT_PROFILE,
      locale,
    });
    if (!out) throw new Error('no sheet');
    return out;
  };

  test('number-formats.pdf: `?` and `_x` take the width of a digit and of `x`', () => {
    const l = withLocale('number-formats');
    expect(glyphX(l, 'E2')).toBe(535); // 1,5 through ?.??
    expect(glyphX(l, 'E11')).toBe(517); // -2 2/5 through # ??/??
    expect(glyphX(l, 'B22')).toBe(264); // 1/3 through # ??/??
    // The accounting zero: the dash before the width of `)`.
    const dash = textAt(l, 20, 4)!.text.lines[0]!.spans[0]!;
    expect(dash.positions[dash.text.indexOf('-')]! + 36).toBe(550);
  });

  test('number-formats.pdf: a number through `@` aligns left', () => {
    const l = withLocale('number-formats');
    expect(glyphX(l, 'B40')).toBe(147);
    expect(glyphX(l, 'E20')).toBe(417);
  });

  test('number-formats-more.pdf: a cell full of `#` centers them', () => {
    const l = withLocale('number-formats-more');
    for (const ref of ['B8', 'B9', 'B10']) expect(glyphX(l, ref)).toBe(149);
  });
});

describe('the spec scenarios', () => {
  test('overflow stops at a value: A1 shows across B1 and is clipped at C1', () => {
    const pkg = sheetOf({
      rows: row(1, [c('A1', 'A long left-aligned text that overflows'), c('C1', 5)]),
    });
    const l = layout(pkg, 'S');
    const a1 = textAt(l, 0, 0)!;
    expect(a1.clip.x).toBe(0);
    expect(a1.clip.x + a1.clip.width).toBe(l.geometry.columns.offsetOf(2));
    expect(a1.text.right).toBeGreaterThan(l.geometry.columns.offsetOf(2));
  });

  test('overflow stops at a formula: a formula with no cached value stops it and keeps its stale mark', () => {
    const pkg = sheetOf({
      rows: row(1, [
        c('A1', 'A long left-aligned text that overflows'),
        c('B1', undefined, { f: '1+1' }),
      ]),
    });
    const l = layout(pkg, 'S');
    const a1 = textAt(l, 0, 0)!;
    expect(a1.clip.x + a1.clip.width).toBe(l.geometry.columns.offsetOf(1));
    const b1 = textAt(l, 0, 1)!;
    expect(b1.stale).toBe(true);
    expect(shown(b1)).toBe('');
  });

  test('formula-overflow: a cached empty string is text, so B3 has one empty line', () => {
    const l = layout(fixture('formula-overflow'), 'Overflow');
    const b3 = textAt(l, 2, 1)!;
    expect(b3.stale).toBe(false);
    expect(b3.text.lines).toHaveLength(1);
    expect(b3.text.lines[0]!.spans).toHaveLength(0);
    const a3 = textAt(l, 2, 0)!;
    expect(a3.clip.x + a3.clip.width).toBe(l.geometry.columns.offsetOf(1));
  });

  test('grid-structure: a formula entered in B6 stops the text of A6 after save and reopen', () => {
    const store = new WorkbookDocumentStore(fixture('grid-structure'));
    const r = store.transact([
      {
        op: 'setCellValue',
        sheet: 'Grid',
        ref: 'B6',
        value: { type: 'formula', text: 'SUM(B4,2)' },
      },
    ]);
    if (!r.ok) throw new Error(JSON.stringify(r.rejection));
    const l = layout(open(saveXlsx(store.package)), 'Grid');
    const a6 = textAt(l, 5, 0)!;
    expect(shown(a6)).toBe('after hidden');
    expect(a6.clip.x + a6.clip.width).toBe(l.geometry.columns.offsetOf(1));
    const b6 = textAt(l, 5, 1)!;
    expect(b6.stale).toBe(true);
    expect(shown(b6)).toBe('');
  });

  test('wrapped text stays in its cell', () => {
    const pkg = sheetOf(
      { rows: row(1, [c('A1', 'A long text that wraps inside its cell', { s: 1 })]) },
      {
        styles: {
          cellXfs: [
            '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment wrapText="1"/></xf>',
          ],
        },
      }
    );
    const l = layout(pkg, 'S');
    const a1 = textAt(l, 0, 0)!;
    expect(a1.text.lines.length).toBeGreaterThan(1);
    expect(a1.clip.width).toBe(53);
    for (const line of a1.text.lines) expect(line.width).toBeLessThanOrEqual(48);
  });

  test('a merged range is one rectangle with the text of its top-left cell', () => {
    const pkg = sheetOf(
      {
        rows: [row(2, [c('B2', 'Merged', { s: 1 }), c('C2', 'hidden')])].join(''),
        merges: ['B2:D4'],
      },
      {
        styles: {
          fills: [
            '<fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/></patternFill></fill>',
          ],
          cellXfs: ['<xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0" applyFill="1"/>'],
        },
      }
    );
    const l = layout(pkg, 'S');
    const boxes = l.panes.flatMap((p) => p.cells);
    expect(boxes).toHaveLength(1);
    expect(boxes[0]!.merge).toEqual({ firstRow: 1, firstCol: 1, lastRow: 3, lastCol: 3 });
    expect(boxes[0]!.rect).toEqual({ x: 53, y: 14, width: 159, height: 42 });
    expect(texts(l).map((t) => [t.row, t.col])).toEqual([[1, 1]]);
    expect(shown(texts(l)[0])).toBe('Merged');
  });

  test('a hidden column has zero width and lays out no cells', () => {
    const pkg = sheetOf({
      rows: row(1, [c('B1', 'b'), c('C1', 'c'), c('D1', 'd')]),
      cols: '<cols><col min="3" max="3" width="10" hidden="1"/></cols>',
    });
    const l = layout(pkg, 'S');
    expect(l.panes[0]!.columns.map((col) => col.index).slice(0, 4)).toEqual([0, 1, 3, 4]);
    expect(textAt(l, 0, 2)).toBeUndefined();
    expect(textAt(l, 0, 3)!.text.lines[0]!.x).toBe(53 * 2 + 3);
  });

  test('frozen rows keep their place while the body scrolls down 50 rows', () => {
    const rows = Array.from({ length: 60 }, (_, i) =>
      row(i + 1, [c(`A${i + 1}`, `r${i + 1}`)])
    ).join('');
    const pkg = sheetOf({
      rows,
      viewChildren: '<pane ySplit="2" topLeftCell="A3" activePane="bottomLeft" state="frozen"/>',
    });
    const l = layout(pkg, 'S', { width: 300, height: 140, scrollX: 0, scrollY: 50 * 14 });
    expect(l.panes).toHaveLength(2);
    expect(shown(textAt(l, 0, 0))).toBe('r1');
    expect(textAt(l, 0, 0)!.text.lines[0]!.baseline).toBe(11);
    expect(textAt(l, 52, 0)!.text.lines[0]!.baseline).toBe(28 + 11);
    expect(textAt(l, 2, 0)).toBeUndefined();
    expect(l.frozenEdges).toEqual({ y: 28 });
  });
});

describe('virtualized layout', () => {
  test('scrolling to the last row lays out only the visible rows', () => {
    const pkg = sheetOf({
      rows: [row(1, [c('A1', 1)]), row(1_048_576, [c('A1048576', 'last')])].join(''),
    });
    const total = layoutSheet(pkg, 'S', VIEW, { measurer, profile: PRINT_PROFILE })!.geometry.rows
      .total;
    const l = layout(pkg, 'S', { ...VIEW, height: 100, scrollY: total - 100 });
    const pane = l.panes[0]!;
    expect(pane.rows.at(-1)!.index).toBe(1_048_575);
    expect(pane.rows.length).toBeLessThanOrEqual(9);
    expect(shown(textAt(l, 1_048_575, 0))).toBe('last');
  });

  test('each pane gives the viewport point of the sheet origin, which moves with the scroll', () => {
    const pkg = sheetOf({
      rows: row(1, [c('A1', 1)]),
      viewChildren:
        '<pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="frozen"/>',
    });
    const at = (scroll: number) => layout(pkg, 'S', { ...VIEW, scrollX: scroll, scrollY: scroll });
    for (const l of [at(0), at(60)]) {
      expect(l.panes).toHaveLength(4);
      for (const p of l.panes) {
        const col = p.columns[0]!;
        const r = p.rows[0]!;
        expect(p.sheetOrigin.x).toBeCloseTo(col.start - l.geometry.columns.offsetOf(col.index));
        expect(p.sheetOrigin.y).toBeCloseTo(r.start - l.geometry.rows.offsetOf(r.index));
      }
    }
    // The frozen panes keep their origin; the body origin moves by the scroll amount.
    const [corner, , , body] = at(0).panes;
    const [corner2, , , body2] = at(60).panes;
    expect(corner2!.sheetOrigin).toEqual(corner!.sheetOrigin);
    expect(body2!.sheetOrigin.x - body!.sheetOrigin.x).toBeCloseTo(-60);
    expect(body2!.sheetOrigin.y - body!.sheetOrigin.y).toBeCloseTo(-60);
  });

  test('text of a cell left of the window shows across into it', () => {
    const pkg = sheetOf({ rows: row(1, [c('A1', 'Overflowing text that reaches into column C')]) });
    const l = layout(pkg, 'S', { ...VIEW, scrollX: 60 });
    const a1 = textAt(l, 0, 0)!;
    expect(a1.clip.x).toBe(0);
    expect(a1.text.lines[0]!.x).toBeLessThan(0);
  });

  test('a formula with a stale cached value is marked', () => {
    const pkg = sheetOf(
      { rows: row(1, [c('A1', 2, { f: '1+1' })]) },
      { calcPr: 'fullCalcOnLoad="1"' }
    );
    expect(textAt(layout(pkg, 'S'), 0, 0)!.stale).toBe(true);
  });

  test('a sheet name that does not exist gives no layout', () => {
    expect(
      layoutSheet(fixture('basic-values'), 'Missing', VIEW, { measurer, profile: PRINT_PROFILE })
    ).toBeUndefined();
  });

  test('the screen profile makes rows 15 points', () => {
    expect(layout(fixture('basic-values'), 'Values', VIEW, true).geometry.rows.sizeOf(0)).toBe(15);
  });
});

describe('grid-structure.pdf: rows, hidden cells, merged ranges, and drawings', () => {
  // Expected values are page coordinates of the reference PDF minus 36, the grid origin.
  const page = (v: number) => v - 36;

  test('rows with a height print scaled, and a hidden row and column take no space', () => {
    const l = layout(fixture('grid-structure'), 'Grid');
    expect([0, 1, 2, 3, 4, 5].map((r) => l.geometry.rows.sizeOf(r))).toEqual([
      14, 9, 18, 31, 0, 14,
    ]);
    expect([1, 2, 3, 5].map((r) => textAt(l, r, 0)!.text.lines[0]!.baseline)).toEqual(
      [56, 74, 105, 119].map(page)
    );
    expect(textAt(l, 5, 1)!.text.lines[0]!.x).toBe(page(92));
    expect(textAt(l, 5, 2)).toBeUndefined();
    expect(textAt(l, 5, 3)!.text.lines[0]!.x).toBe(page(152));
  });

  test('a merged range centers its text, and a wrapped one stacks lines 14 points apart', () => {
    const l = layout(fixture('grid-structure'), 'Grid');
    const merged = textAt(l, 7, 1)!;
    expect(merged.text.lines[0]!.x).toBe(page(111));
    // The PDF has the baseline at 160; layout has it 1 point lower (the approximation
    // "Measured text offsets"; the test pins the exact difference, so a fix shows).
    expect(merged.text.lines[0]!.baseline - page(160)).toBe(1);
    const box = l.panes[0]!.cells.find((b) => b.row === 7 && b.col === 1)!;
    expect(box.rect).toEqual({ x: page(89), y: page(136), width: 113, height: 42 });
    const wrapped = textAt(l, 11, 5)!;
    expect(wrapped.text.lines.map((line) => [line.x, line.baseline])).toEqual([
      [page(258), page(203)],
      [page(258), page(217)],
    ]);
  });

  test('drawings sit on the same grid origin as the cells', () => {
    const l = layout(fixture('grid-structure'), 'Grid');
    const rects = l.panes[0]!.drawings.map((d) => [
      d.kind,
      d.rect.x,
      d.rect.y,
      d.rect.width,
      d.rect.height,
    ]);
    expect(rects[0]).toEqual(['picture', page(89), page(248), 60, 70]);
    expect(rects[2]).toEqual(['shape', page(89), page(346), 100, 50]);
    expect(rects[3]).toEqual(['shape', page(336), page(436), 100, 100]);
    const ellipse = l.panes[0]!.drawings[1]!;
    expect(ellipse.rect.x).toBe(page(212));
  });
});

describe('text-layout.pdf: every cell text rule', () => {
  const page = (v: number) => v - 36;
  let l: SheetLayout;
  beforeAll(() => {
    l = layout(fixture('text-layout'), 'Text');
  });
  const first = (r: number, col: number) => textAt(l, r - 1, col)!.text.lines[0]!;
  const starts = (r: number, col: number) =>
    textAt(l, r - 1, col)!.text.lines.map((line) => [line.x, line.baseline]);

  test('rows: Carlito 20 and a 14 point rich run make their rows taller in print', () => {
    expect(Array.from({ length: 14 }, (_, r) => l.geometry.rows.sizeOf(r))).toEqual([
      14, 14, 14, 14, 56, 14, 67, 14, 14, 14, 14, 24, 14, 17,
    ]);
  });

  test('overflow: left text runs right, right text runs left, centered text both ways', () => {
    expect([first(1, 0).x, first(1, 0).baseline]).toEqual([page(39), page(47)]);
    expect(first(2, 3).x).toBe(page(190));
    expect(first(3, 2).x).toBe(page(151));
    expect(textAt(l, 0, 0)!.clip.x + textAt(l, 0, 0)!.clip.width).toBe(page(180));
  });

  test('alignment, fill, and indent', () => {
    expect([1, 2, 3, 4, 5].map((col) => first(4, col).x)).toEqual(
      [130, 228, 255, 339, 422].map(page)
    );
    expect(first(4, 0).x).toBe(page(39));
    expect(
      first(4, 3)
        .spans.map((s) => s.text)
        .join('')
    ).toBe('abababababab');
  });

  test('justify, distributed, wrapping, and vertical alignment', () => {
    expect(starts(5, 0).map(([, y]) => y)).toEqual([103, 117, 131, 145].map(page));
    expect(first(5, 0).x + first(5, 0).width).toBe(page(108 - 2));
    expect(starts(5, 1).map(([x]) => x)).toEqual([119, 111, 111, 134].map(page));
    expect(starts(5, 2)).toEqual([
      [page(183), page(117)],
      [page(183), page(131)],
      [page(183), page(145)],
    ]);
    expect(first(5, 3).baseline).toBe(page(103));
    expect([first(5, 4).x, first(5, 4).baseline]).toEqual([page(327), page(124)]);
    expect(starts(5, 5).map(([, y]) => y)).toEqual([103, 117, 131].map(page));
  });

  test('centered across empty cells, shrunk to whole points, and fonts set the margin', () => {
    expect(first(6, 0).x).toBe(page(109));
    const shrunk = first(7, 4);
    expect([shrunk.x, shrunk.spans[0]!.style.sizePt]).toEqual([page(326), 6]);
    expect([first(12, 0).x, first(12, 0).baseline]).toEqual([page(40), page(304)]);
    expect(first(13, 0).x).toBe(page(38));
    expect(first(14, 0).spans.map((s) => [s.x, s.style.sizePt])).toEqual([
      [page(39), 14],
      [page(109), 11],
    ]);
    expect(first(14, 0).baseline).toBe(page(336));
  });

  test('numbers that do not fit, booleans, and errors', () => {
    const cells: [number, number, string, number][] = [
      [8, 6, '#####', 471],
      [8, 7, '1234.57', 505],
      [9, 6, '0', 490],
      [9, 7, '1.2E+11', 507],
      [10, 6, '#####', 471],
      [10, 7, '98765.4', 505],
      [11, 6, 'TRUE', 472],
      [11, 7, '#N/A', 512],
    ];
    for (const [r, col, text, x] of cells) {
      expect(shown(textAt(l, r - 1, col))).toBe(text);
      expect(first(r, col).x).toBe(page(x));
    }
  });

  test('a 45 degree rotation starts its baseline where the PDF does', () => {
    // The PDF draws "Rotated 45" with the matrix 7.778 7.778 -7.778 7.778 46.071 565.929.
    const rot = textAt(l, 6, 0)!.text.rotation!;
    expect(rot.originX).toBeCloseTo(page(46.071), 2);
    expect(rot.originY).toBeCloseTo(page(792 - 565.929), 2);
  });

  test('90 degrees: general aligns right, and the baseline starts 2 points above the bottom', () => {
    // The PDF draws "Up 90" with the matrix 0 11 -11 0 174 565.
    const rot = textAt(l, 6, 1)!.text.rotation!;
    expect(rot.angle).toBe(90);
    expect([rot.originX, rot.originY]).toEqual([page(174), page(792 - 565)]);
  });

  test('clockwise 45 degrees: general aligns right, at the measured baseline start', () => {
    // The PDF draws "Down 45" with the matrix 7.778 -7.778 7.778 7.778 212.929 594.929.
    const rot = textAt(l, 6, 2)!.text.rotation!;
    expect(rot.angle).toBe(-45);
    expect(rot.originX).toBeCloseTo(page(212.929), 1);
    expect(rot.originY).toBeCloseTo(page(792 - 594.929), 2);
  });

  test('a font under 11 points keeps the descent of the default font at the bottom', () => {
    // The PDF has the Carlito 8 baseline of A13 at 320, and the 6 point shrunk text of E7 at 226.
    expect(first(13, 0).baseline).toBe(page(320));
    expect(first(7, 4).baseline).toBe(page(226));
  });

  test('stacked text ends 2 points above the bottom, on the measured baselines', () => {
    // The PDF draws A, B, and C on the baselines 196, 210, and 224.
    const lines = textAt(l, 6, 3)!.text.lines;
    expect(lines.map((line) => line.baseline)).toEqual([196, 210, 224].map(page));
  });
});

// The approximation "Measured text offsets": text-layout.pdf measures these, and layout differs
// from it by the amounts below. Each test pins the exact difference, so a fix shows as a failure
// that updates the test and removes the item from the approximation.
describe('text-layout.pdf: measured text offsets', () => {
  const page = (v: number) => v - 36;
  let l: SheetLayout;
  beforeAll(() => {
    l = layout(fixture('text-layout'), 'Text');
  });

  test('stacked text: each letter sits 1 point right', () => {
    // The PDF draws A, B, and C at x 285.
    const lines = textAt(l, 6, 3)!.text.lines;
    expect(lines.map((line) => line.x - page(285))).toEqual([1, 1, 1]);
  });
});

describe('conditional formatting in layout', () => {
  test('B1 of preserved-parts lays out bold, as its reference PDF draws it', () => {
    const l = layoutSheet(fixture('preserved-parts'), 'Kept', VIEW_CF, {
      measurer,
      profile: PRINT_PROFILE,
    })!;
    const spans = (col: number) =>
      l.panes
        .flatMap((p) => p.texts)
        .filter((t) => t.row === 0 && t.col === col)
        .flatMap((t) => t.text.lines.flatMap((line) => line.spans));
    expect(spans(1).map((s) => [s.text, s.style.bold])).toEqual([['5', true]]);
    expect(spans(0).every((s) => !s.style.bold)).toBe(true);
  });
});

describe('conditional formatting in layout, conditional-rules', () => {
  // pdffonts, pdftotext -bbox, and the fill operators of fixtures/conditional-rules.pdf: the face,
  // decoration, color, and fill of each cell of column B.
  const layout = () =>
    layoutSheet(fixture('conditional-rules'), 'Rules', VIEW_CF, {
      measurer,
      profile: PRINT_PROFILE,
    })!;
  const spansAt = (l: SheetLayout, r: number) =>
    l.panes
      .flatMap((p) => p.texts)
      .filter((t) => t.row === r && t.col === 1)
      .flatMap((t) => t.text.lines.flatMap((line) => line.spans));

  test('each value of column B has the face, decoration, and color that the PDF draws', () => {
    const l = layout();
    const face = (r: number) =>
      spansAt(l, r)
        .map((s) => {
          const st = s.style;
          const marks = `${st.bold ? 'B' : ''}${st.italic ? 'I' : ''}${st.strike ? 'S' : ''}${st.underline !== 'none' ? 'U' : ''}`;
          return `${s.text}:${marks}${st.color === '000000' ? '' : `:${st.color}`}`;
        })
        .join(' ');
    expect(Array.from({ length: 26 }, (_, r) => face(r))).toEqual([
      '5:',
      '2:I:FF0000',
      // The default locale; the reference PDF prints `7,00` in the locale of its machine.
      '7.00:',
      '4:',
      'abc:B',
      'ABC:S',
      'TRUE:B',
      '9:BI',
      '9:I',
      '6:',
      '#DIV/0!:',
      '8:B',
      '8:B',
      '8:',
      '5:B',
      '',
      '',
      '',
      '3:U',
      '3:',
      '3::ED7D31',
      '3:',
      '3:',
      '3:',
      '5:B',
      '5:B',
    ]);
  });

  test('each cell of column B has the fill that the PDF draws, also cells without a value', () => {
    const l = layout();
    const fill = (r: number) => {
      const f = l.panes
        .flatMap((p) => p.cells)
        .find((cell) => cell.row === r && cell.col === 1)?.fill;
      return !f || f.kind === 'none' ? '' : f.kind === 'solid' ? f.color : f.kind;
    };
    const row = (r: number) => fill(r);
    expect(Array.from({ length: 26 }, (_, r) => row(r))).toEqual([
      'FFFF00', '', '', '', '', '', '', '00B050', '', 'FFFFFF', '', '', '', '', '',
      'FFFF00', 'FFFF00', 'FFFF00', '', 'FFC000', 'B4C7E7', 'FFFF00', '9999FF', '9999FF', '', '',
    ]); // prettier-ignore
    // B21: the PDF fills B4C6E7; theme 4 with tint 0.6 differs by 1 in green ("Tint rounding").
  });

  test('right-aligned italic numbers start where regular ones do (the harness pins the 2-point gap)', () => {
    // The PDF draws B2, B8, and B9 at x 247 and B1 at x 249; layout keeps them at one x.
    const l = layout();
    const x = (r: number) => spansAt(l, r)[0]!.x;
    expect([x(1), x(7), x(8)]).toEqual([x(0), x(0), x(0)]);
  });
});

describe('conditional formatting over blank cells in layout', () => {
  test('rules over a huge window stop at the check budget of the layout and report it', () => {
    // 1.5 million tiny visible cells, and 1,024 ranges in columns A and XFD that meet each row
    // and span each visible cell but miss it: without the budget, one layout took about 1 second.
    const ranges = Array.from({ length: 256 }, () => 'A1:A1048576 XFD1:XFD1048576').join(' ');
    const rules = [1, 2]
      .map(
        (p) =>
          `<conditionalFormatting sqref="${ranges}"><cfRule type="cellIs" dxfId="0" priority="${p}" operator="equal"><formula>0</formula></cfRule></conditionalFormatting>`
      )
      .join('');
    const pkg = open(
      buildXlsx({
        sheets: [
          {
            name: 'S',
            rows: Array.from(
              { length: 2000 },
              (_, i) => `<row r="${i + 1}" ht="1" customHeight="1"/>`
            ).join(''),
            formatPr: 'defaultRowHeight="1" customHeight="1"',
            cols: '<cols><col min="1" max="16384" width="0.2" customWidth="1"/></cols>',
            afterMerges: rules,
          },
        ],
        styles: { dxfs: ['<dxf><font><b/></font></dxf>'] },
      })
    );
    const l = layoutSheet(
      pkg,
      'S',
      { width: 1500, height: 1000, scrollX: 0, scrollY: 0 },
      { measurer, profile: SCREEN_PROFILE }
    )!;
    expect(l.problems.filter((p) => p.code === 'conditional-budget').length).toBe(1);
    expect(l.conditionalCut).toBe(true);
  });

  test('a cut layout says so even when cell problems fill the problem list', () => {
    const ranges = Array.from({ length: 256 }, () => 'A1:A1048576 XFD1:XFD1048576').join(' ');
    const rules = `<conditionalFormatting sqref="${ranges}"><cfRule type="cellIs" dxfId="0" priority="1" operator="equal"><formula>0</formula></cfRule></conditionalFormatting>`;
    // 300 cells of row 1 with a cell format index that the workbook does not have: their
    // problems fill the list before the budget runs out.
    const bad = () =>
      `<row r="1" ht="1" customHeight="1">${Array.from({ length: 300 }, (_, i) => `<c r="${formatCellRef({ row: 0, col: i + 1 })}" s="999"><v>1</v></c>`).join('')}</row>`;
    const pkg = open(
      buildXlsx({
        sheets: [
          {
            name: 'S',
            rows: Array.from({ length: 2000 }, (_, i) =>
              i === 0 ? bad() : `<row r="${i + 1}" ht="1" customHeight="1"/>`
            ).join(''),
            formatPr: 'defaultRowHeight="1" customHeight="1"',
            cols: '<cols><col min="1" max="16384" width="0.2" customWidth="1"/></cols>',
            afterMerges: rules + rules.replace('priority="1"', 'priority="2"'),
          },
        ],
        styles: { dxfs: ['<dxf><font><b/></font></dxf>'] },
      })
    );
    const l = layoutSheet(
      pkg,
      'S',
      { width: 1500, height: 1000, scrollX: 0, scrollY: 0 },
      { measurer, profile: SCREEN_PROFILE }
    )!;
    expect(l.problems.length).toBe(200);
    expect(l.problems.some((p) => p.code === 'conditional-budget')).toBe(false);
    expect(l.conditionalCut).toBe(true);
  });

  test('ordinary rules over a wide window at the smallest zoom stay inside the budget', () => {
    // 30 rules over A1:Z100000 and a window of 14,400 by 8,100 points: blank cells outside
    // column Z check no range.
    const rules = Array.from(
      { length: 30 },
      (_, i) =>
        `<conditionalFormatting sqref="A1:Z100000"><cfRule type="cellIs" dxfId="0" priority="${i + 1}" operator="equal"><formula>${i + 1000}</formula></cfRule></conditionalFormatting>`
    ).join('');
    const pkg = open(
      buildXlsx({
        sheets: [{ name: 'S', rows: '<row r="1"/>', afterMerges: rules }],
        styles: { dxfs: ['<dxf><font><b/></font></dxf>'] },
      })
    );
    const l = layoutSheet(
      pkg,
      'S',
      { width: 14_400, height: 8_100, scrollX: 0, scrollY: 0 },
      { measurer, profile: SCREEN_PROFILE }
    )!;
    expect(l.conditionalCut).toBe(false);
  });

  test('blank cells read the column formats once for each pane, not once for each cell', () => {
    // 16,000 `x:col` elements and one false rule over the window: a scan for each blank cell took
    // more than 2 seconds; one scan for each pane takes a few milliseconds.
    const cols = Array.from(
      { length: 16_000 },
      (_, i) => `<col min="${i + 1}" max="${i + 1}" width="9"/>`
    ).join('');
    const pkg = open(
      buildXlsx({
        sheets: [
          {
            name: 'S',
            cols: `<cols>${cols}</cols>`,
            rows: row(1, [c('A1', 1)]),
            // The rule meets every row of the window and is false, so each blank cell checks it.
            afterMerges:
              '<conditionalFormatting sqref="A1:ZZ9999"><cfRule type="cellIs" dxfId="0" priority="1" operator="equal"><formula>12345</formula></cfRule></conditionalFormatting>',
          },
        ],
        styles: { dxfs: ['<dxf><font><b/></font></dxf>'] },
      })
    );
    const spy = spyOn(inherited, 'columnFormats');
    try {
      layoutSheet(
        pkg,
        'S',
        { width: 1500, height: 1000, scrollX: 0, scrollY: 0 },
        { measurer, profile: PRINT_PROFILE }
      );
      // One read for the pane, and none for each blank cell: the count does not depend on time.
      expect(spy.mock.calls.length).toBeLessThan(10);
    } finally {
      spy.mockRestore();
    }
  });

  test('a cell lowercases its text once, and long string operands are lowercased when read', () => {
    const long = 'x'.repeat(8_000);
    const rules = Array.from(
      { length: 64 },
      (_, i) =>
        `<conditionalFormatting sqref="A1:XFD1048576"><cfRule type="cellIs" dxfId="0" priority="${i + 1}" operator="between"><formula>"a${long}"</formula><formula>"b${long}"</formula></cfRule></conditionalFormatting>`
    ).join('');
    const cells = Array.from({ length: 10 }, (_, i) => c(`A${i + 1}`, `Text ${i}`));
    const lowercases = (afterMerges: string) => {
      const pkg = open(
        buildXlsx({
          sheets: [{ name: 'S', rows: cells.map((x, i) => row(i + 1, [x])).join(''), afterMerges }],
          styles: { dxfs: ['<dxf><font><b/></font></dxf>'] },
        })
      );
      const lay = () =>
        layoutSheet(
          pkg,
          'S',
          { width: 1500, height: 1000, scrollX: 0, scrollY: 0 },
          { measurer, profile: PRINT_PROFILE }
        );
      // The first layout reads the rules; the count covers the second.
      lay();
      const spy = spyOn(String.prototype, 'toLowerCase');
      try {
        lay();
        return spy.mock.calls.length;
      } finally {
        spy.mockRestore();
      }
    };
    // 64 between rules over each visible cell: once for each of the 10 text cells, not twice
    // for each rule and cell.
    expect(lowercases(rules) - lowercases('')).toBeLessThanOrEqual(10);
  });

  test('blank cells of a row that no shown rule meets check no rule', () => {
    const pkg = open(
      buildXlsx({
        sheets: [
          {
            name: 'S',
            rows: row(1, [c('A1', 1)]),
            afterMerges:
              '<conditionalFormatting sqref="ZZ9999"><cfRule type="cellIs" dxfId="0" priority="1" operator="equal"><formula>0</formula></cfRule></conditionalFormatting>',
          },
        ],
        styles: { dxfs: ['<dxf><font><b/></font></dxf>'] },
      })
    );
    const spy = spyOn(conditional, 'conditionalFormat');
    try {
      layoutSheet(
        pkg,
        'S',
        { width: 1500, height: 1000, scrollX: 0, scrollY: 0 },
        { measurer, profile: PRINT_PROFILE }
      );
      // A1 is the one cell with a value; no blank cell of the window calls the walk.
      expect(spy.mock.calls.length).toBe(1);
    } finally {
      spy.mockRestore();
    }
  });

  const rule =
    '<conditionalFormatting sqref="A1:C3"><cfRule type="cellIs" dxfId="0" priority="1" operator="equal"><formula>0</formula></cfRule></conditionalFormatting>';
  const dxfs = ['<dxf><fill><patternFill><bgColor rgb="FFFFFF00"/></patternFill></fill></dxf>'];

  test('an empty merged range under a true rule paints the fill of the rule', () => {
    const pkg = open(
      buildXlsx({
        sheets: [{ name: 'S', rows: row(3, [c('C3', 1)]), merges: ['A1:B2'], afterMerges: rule }],
        styles: { dxfs },
      })
    );
    const l = layoutSheet(pkg, 'S', VIEW_CF, { measurer, profile: PRINT_PROFILE })!;
    const merged = l.panes.flatMap((p) => p.cells).filter((cell) => cell.merge);
    expect(merged.map((cell) => cell.fill)).toEqual([{ kind: 'solid', color: 'FFFF00' }]);
  });

  test('a blank cell with a row format and a true rule gets one record with the rule fill', () => {
    const pkg = open(
      buildXlsx({
        sheets: [
          {
            name: 'S',
            rows: '<row r="1" s="1" customFormat="1"><c r="C1" s="1"><v>1</v></c></row>',
            afterMerges: rule,
          },
        ],
        styles: {
          fills: [
            '<fill><patternFill patternType="solid"><fgColor rgb="FF112233"/></patternFill></fill>',
          ],
          cellXfs: ['<xf numFmtId="0" fontId="0" fillId="2" borderId="0" applyFill="1"/>'],
          dxfs,
        },
      })
    );
    const l = layoutSheet(pkg, 'S', VIEW_CF, { measurer, profile: PRINT_PROFILE })!;
    const at = l.panes.flatMap((p) => p.cells).filter((cell) => cell.row === 0 && cell.col === 0);
    expect(at.map((cell) => cell.fill)).toEqual([{ kind: 'solid', color: 'FFFF00' }]);
    // A blank cell of the row that no rule covers keeps the row fill.
    const d = l.panes.flatMap((p) => p.cells).filter((cell) => cell.row === 0 && cell.col === 3);
    expect(d.map((cell) => cell.fill)).toEqual([{ kind: 'solid', color: '112233' }]);
  });
});

const VIEW_CF = { width: 540, height: 720, scrollX: 0, scrollY: 0 };
