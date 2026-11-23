// The key table over a recording host: Shift with an arrow or Page key extends a band of whole
// rows or columns as a band, so the structural keys still act on the extended band.

import { describe, expect, test } from 'bun:test';
import type { GridPoint, GridSelection } from '../../contracts/index.ts';
import { handleGridKey, type GridKeyHost } from '../grid-keys.ts';
import { extendTo, selectBand, selectCell } from '../selection.ts';
import { selectionRef } from '../editor-model.ts';

function host(start: GridSelection) {
  const calls: string[] = [];
  const h = {
    selection: start,
    focusCell: start.active,
    merges: [],
    hidden: () => ({ row: () => false, col: () => false }),
    pageRows: 10,
    facts: () => ({ lastRow: 5, lastCol: 5 }),
    dataEdge: (from: GridPoint) => from,
    select(anchor: GridPoint, focus?: GridPoint) {
      const sel = selectCell(0, anchor, []);
      const end = focus ?? anchor;
      h.selection = {
        ...sel,
        range: {
          start: { row: Math.min(anchor.row, end.row), col: Math.min(anchor.col, end.col) },
          end: { row: Math.max(anchor.row, end.row), col: Math.max(anchor.col, end.col) },
        },
      };
      h.focusCell = end;
    },
    setSelection(selection: GridSelection, focus: GridPoint) {
      h.selection = selection;
      h.focusCell = focus;
    },
    extend(to: GridPoint) {
      h.setSelection(extendTo(h.selection, to, []), to);
    },
    selectAll: () => calls.push('selectAll'),
    zoomStep: (d: number) => calls.push(`zoom ${d}`),
    setZoom: () => calls.push('setZoom'),
    cycleSheet: () => calls.push('cycleSheet'),
    undo: () => calls.push('undo'),
    redo: () => calls.push('redo'),
    beginEdit: () => calls.push('beginEdit'),
    clearValues: () => calls.push('clearValues'),
    toggleFormat: () => calls.push('toggleFormat'),
    structure: (name: string) => calls.push(name),
  };
  return { h: h as GridKeyHost & typeof h, calls };
}

const grid = {} as EventTarget;
const press = (h: GridKeyHost, key: string, mods: Partial<KeyboardEvent> = {}) =>
  handleGridKey(
    h,
    {
      key,
      target: grid,
      shiftKey: false,
      ctrlKey: false,
      metaKey: false,
      altKey: false,
      isComposing: false,
      keyCode: 0,
      preventDefault: () => undefined,
      ...mods,
    } as unknown as KeyboardEvent,
    grid
  );

describe('Shift extends a band as a band', () => {
  test('rows: Shift+Down and Shift+Page Down keep whole rows, and Mod+- deletes them', () => {
    const { h, calls } = host(selectBand(0, 'rows', 1, 1));
    press(h, 'ArrowDown', { shiftKey: true });
    expect(selectionRef(h.selection)).toBe('A2:XFD3');
    press(h, 'PageDown', { shiftKey: true });
    expect(selectionRef(h.selection)).toBe('A2:XFD13');
    press(h, 'ArrowRight', { shiftKey: true });
    expect(selectionRef(h.selection)).toBe('A2:XFD13');
    press(h, '-', { ctrlKey: true });
    expect(calls).toEqual(['deleteRows']);
  });

  test('columns: Shift+Right keeps whole columns, and Mod+- deletes them', () => {
    const { h, calls } = host(selectBand(0, 'columns', 1, 1));
    press(h, 'ArrowRight', { shiftKey: true });
    expect(selectionRef(h.selection)).toBe('B1:C1048576');
    press(h, 'PageDown', { shiftKey: true });
    expect(selectionRef(h.selection)).toBe('B1:C1048576');
    press(h, 'ArrowLeft', { shiftKey: true });
    press(h, 'ArrowLeft', { shiftKey: true });
    expect(selectionRef(h.selection)).toBe('A1:B1048576');
    press(h, '-', { ctrlKey: true });
    expect(calls).toEqual(['deleteColumns']);
  });

  test('a cell range still extends as a range', () => {
    const { h } = host(selectCell(0, { row: 1, col: 1 }, []));
    press(h, 'ArrowDown', { shiftKey: true });
    expect(selectionRef(h.selection)).toBe('B2:B3');
  });
});
