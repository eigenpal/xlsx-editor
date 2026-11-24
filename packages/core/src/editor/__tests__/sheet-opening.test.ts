// The opening view of a sheet (`sheetOpening`): with a frozen `x:pane` (`CT_Pane`, ECMA-376 Part 1,
// 18.3.1.66), the scroll comes from `x:pane/@topLeftCell`, the top-left visible cell of the
// bottom-right pane, and the active cell from the `x:selection` (18.3.1.78) whose `@pane` is
// `x:pane/@activePane`. Without one, both come from `x:sheetView` (18.3.1.87): the selection of
// its top-left pane, else its first selection.

import { expect, test } from 'bun:test';
import { buildXlsx, c, row } from '../../../../../scripts/fixtures/builder.ts';
import { readXlsx } from '../../store/index.ts';
import { sheetOpening } from '../sheet-opening.ts';
import { scrollToTopLeft } from '../view-map.ts';
import type { SheetGeometry } from '../../layout/index.ts';

const SHEET = '/xl/worksheets/sheet1.xml';

function opening(viewChildren: string, viewAttrs?: string) {
  const bytes = buildXlsx({
    sheets: [{ name: 'S', rows: row(1, [c('A1', 1)]), viewChildren, viewAttrs }],
  });
  const read = readXlsx(bytes);
  if (!read.ok) throw new Error('the package did not open');
  return sheetOpening(read.pkg, SHEET);
}

test('a frozen sheet opens at the top-left cell of its scrolling pane', () => {
  const cols = opening('<pane xSplit="2" topLeftCell="E1" activePane="topRight" state="frozen"/>');
  expect(cols.topLeft).toEqual({ row: 0, col: 4 });
  const rows = opening(
    '<pane ySplit="2" topLeftCell="A6" activePane="bottomLeft" state="frozen"/>'
  );
  expect(rows.topLeft).toEqual({ row: 5, col: 0 });
  const split = opening(
    '<pane xSplit="1" ySplit="1" topLeftCell="D9" activePane="bottomRight" state="frozenSplit"/>'
  );
  expect(split.topLeft).toEqual({ row: 8, col: 3 });
});

test('a frozen pane with no top-left cell falls back to the sheet view', () => {
  const o = opening('<pane xSplit="2" activePane="topRight" state="frozen"/>', 'topLeftCell="E3"');
  expect(o.topLeft).toEqual({ row: 2, col: 4 });
});

test('a split pane that does not freeze keeps the view cell and the selection of the top-left pane', () => {
  // The editor does not show split panes, so the cell and selection of the bottom-right pane do
  // not apply: the active cell stays in the view that opens.
  const o = opening(
    '<pane xSplit="2000" ySplit="1000" topLeftCell="H20" activePane="bottomRight"/>' +
      '<selection activeCell="C4" sqref="C4"/>' +
      '<selection pane="bottomRight" activeCell="H25" sqref="H25"/>',
    'topLeftCell="B3"'
  );
  expect(o.topLeft).toEqual({ row: 2, col: 1 });
  expect(o.active).toEqual({ row: 3, col: 2 });
});

test('a frozen pane that freezes no row or column keeps the view cell and the selection of the top-left pane', () => {
  const o = opening(
    '<pane xSplit="0" ySplit="0" topLeftCell="H20" activePane="bottomRight" state="frozen"/>' +
      '<selection activeCell="C4" sqref="C4"/>' +
      '<selection pane="bottomRight" activeCell="H25" sqref="H25"/>',
    'topLeftCell="B3"'
  );
  expect(o.topLeft).toEqual({ row: 2, col: 1 });
  expect(o.active).toEqual({ row: 3, col: 2 });
});

test('a split pane takes the selection of the top-left pane, wherever it is listed', () => {
  const o = opening(
    '<pane xSplit="2000" activePane="bottomRight"/>' +
      '<selection pane="topRight" activeCell="Z1" sqref="Z1"/>' +
      '<selection activeCell="C4" sqref="C4"/>'
  );
  expect(o.active).toEqual({ row: 3, col: 2 });
});

test('a frozen pane with an invalid top-left cell falls back to the sheet view', () => {
  const o = opening('<pane xSplit="1" topLeftCell="ZZZZZZ1" state="frozen"/>', 'topLeftCell="C1"');
  expect(o.topLeft).toEqual({ row: 0, col: 2 });
});

test('the opening scroll puts the top-left cell next to the frozen band', () => {
  // Each column is 50 points wide and each row 20 points high.
  const axis = (size: number) => ({ offsetOf: (i: number) => i * size });
  const geometry = { columns: axis(50), rows: axis(20) } as unknown as SheetGeometry;
  const frozen = { rows: 2, cols: 2 };
  expect(scrollToTopLeft(geometry, frozen, { row: 0, col: 4 })).toEqual({ x: 100, y: 0 });
  expect(scrollToTopLeft(geometry, frozen, { row: 5, col: 0 })).toEqual({ x: 0, y: 60 });
  // A cell inside the frozen band scrolls nothing on that axis.
  expect(scrollToTopLeft(geometry, frozen, { row: 9, col: 1 })).toEqual({ x: 0, y: 140 });
});

test('the active cell comes from the selection of the active pane', () => {
  const o = opening(
    '<pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="frozen"/>' +
      '<selection pane="topRight" activeCell="C1" sqref="C1"/>' +
      '<selection pane="bottomLeft" activeCell="A4" sqref="A4"/>' +
      '<selection pane="bottomRight" activeCell="F7" sqref="F7"/>'
  );
  expect(o.active).toEqual({ row: 6, col: 5 });
});

test('the active pane and the pane of a selection default to the top-left pane', () => {
  const o = opening(
    '<pane xSplit="1" topLeftCell="B1" state="frozen"/>' +
      '<selection pane="topRight" activeCell="D2" sqref="D2"/>' +
      '<selection activeCell="A3" sqref="A3"/>'
  );
  expect(o.active).toEqual({ row: 2, col: 0 });
});

test('with no selection for the active pane, the first selection applies', () => {
  const o = opening(
    '<pane xSplit="1" topLeftCell="C1" activePane="topRight" state="frozen"/>' +
      '<selection activeCell="B5" sqref="B5"/>'
  );
  expect(o.active).toEqual({ row: 4, col: 1 });
});

test('a sheet with no pane keeps its view cell and the selection of the top-left pane', () => {
  const o = opening('<selection activeCell="C3" sqref="C3"/>', 'topLeftCell="B2"');
  expect(o.topLeft).toEqual({ row: 1, col: 1 });
  expect(o.active).toEqual({ row: 2, col: 2 });
});
