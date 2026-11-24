// The state of the sheet commands (`sheetCommandState`): the one source of the enabled state
// and the reason of the add button and the sheet menu rows, as `toolbarCommandState` is for the
// toolbar.

import { describe, expect, test } from 'bun:test';
import { sheetCommandState, type SheetCommandInput } from '../sheet-command-state.ts';
import type { EditorSheet } from '../editor-model.ts';

const sheet = (index: number, visible = true): EditorSheet => ({
  index,
  name: `S${index}`,
  visible,
  worksheet: true,
});
const input = (over: Partial<SheetCommandInput> = {}): SheetCommandInput => ({
  mode: 'edit',
  // S1 is hidden: a move goes past it.
  sheets: [sheet(0), sheet(1, false), sheet(2)],
  workbookProtection: null,
  sheetEditRefusal: null,
  ...over,
});
const keyOf = (s: ReturnType<typeof sheetCommandState>) => s.disabledReason?.key ?? null;
const ALL = ['add', 'rename', 'delete', 'moveLeft', 'moveRight'] as const;

describe('sheetCommandState', () => {
  test('enables every command, and moves go past a hidden sheet', () => {
    expect(sheetCommandState(input(), 'add', -1)).toEqual({ enabled: true, disabledReason: null });
    expect(sheetCommandState(input(), 'rename', 0).enabled).toBe(true);
    expect(sheetCommandState(input(), 'delete', 2).enabled).toBe(true);
    expect(sheetCommandState(input(), 'moveRight', 0)).toMatchObject({ enabled: true, to: 2 });
    expect(sheetCommandState(input(), 'moveLeft', 2)).toMatchObject({ enabled: true, to: 0 });
  });

  test('a move at the end of the tabs, or a sheet with no tab, has no reason', () => {
    expect(sheetCommandState(input(), 'moveLeft', 0)).toEqual({
      enabled: false,
      disabledReason: null,
    });
    expect(sheetCommandState(input(), 'moveRight', 2).enabled).toBe(false);
    expect(sheetCommandState(input(), 'rename', 1).enabled).toBe(false);
  });

  test('view mode disables every command with viewMode', () => {
    for (const c of ALL) {
      const s = sheetCommandState(input({ mode: 'view' }), c, 0);
      expect([c, s.enabled, keyOf(s)]).toEqual([c, false, 'disabledReason.viewMode']);
    }
  });

  test('a workbook that locks its structure disables every command with workbookLocked', () => {
    const locked = input({ workbookProtection: { allowStructure: false } });
    for (const c of ALL) {
      const s = sheetCommandState(locked, c, 0);
      expect([c, s.enabled, keyOf(s)]).toEqual([c, false, 'disabledReason.workbookLocked']);
    }
    const open = input({ workbookProtection: { allowStructure: true } });
    expect(sheetCommandState(open, 'add', -1).enabled).toBe(true);
  });

  test('Delete of the last visible sheet is disabled with lastSheet', () => {
    const one = input({ sheets: [sheet(0), sheet(1, false)] });
    expect(keyOf(sheetCommandState(one, 'delete', 0))).toBe('disabledReason.lastSheet');
    expect(sheetCommandState(one, 'rename', 0).enabled).toBe(true);
  });

  test('a chart, pivot table, or chartsheet in the workbook disables Rename and Delete', () => {
    const about = { kind: 'chart' } as const;
    const blocked = input({
      sheetEditRefusal: { code: 'structure-blocked', subject: '/xl/charts/chart1.xml', about },
    });
    for (const c of ['rename', 'delete'] as const)
      expect([c, sheetCommandState(blocked, c, 0)]).toEqual([
        c,
        { enabled: false, disabledReason: { key: 'disabledReason.sheetEditBlocked', about } },
      ]);
    // Add and the moves stay enabled, and the last sheet keeps its own reason.
    expect(sheetCommandState(blocked, 'add', -1).enabled).toBe(true);
    expect(sheetCommandState(blocked, 'moveRight', 0).enabled).toBe(true);
    const one = input({ ...blocked, sheets: [sheet(0), sheet(1, false)] });
    expect(keyOf(sheetCommandState(one, 'delete', 0))).toBe('disabledReason.lastSheet');
  });

  test('macros also disable add and the moves, as they can read a sheet by its position', () => {
    const about = { kind: 'macros' } as const;
    const macros = input({ sheetEditRefusal: { code: 'structure-blocked', about } });
    for (const [c, at] of [
      ['add', -1],
      ['moveLeft', 2],
      ['moveRight', 0],
    ] as const)
      expect([c, sheetCommandState(macros, c, at)]).toEqual([
        c,
        { enabled: false, disabledReason: { key: 'disabledReason.sheetOrderBlocked', about } },
      ]);
    expect(keyOf(sheetCommandState(macros, 'rename', 0))).toBe('disabledReason.sheetEditBlocked');
    // View mode and a locked sheet list keep their own reasons.
    expect(keyOf(sheetCommandState({ ...macros, mode: 'view' }, 'add', -1))).toBe(
      'disabledReason.viewMode'
    );
    const locked = { ...macros, workbookProtection: { allowStructure: false } };
    expect(keyOf(sheetCommandState(locked, 'add', -1))).toBe('disabledReason.workbookLocked');
  });

  test('a chartsheet has every sheet command but add disabled, with chartsheet', () => {
    // S1 is a visible chartsheet between two worksheets.
    const chart = input({ sheets: [sheet(0), { ...sheet(1), worksheet: false }, sheet(2)] });
    for (const c of ['rename', 'delete', 'moveLeft', 'moveRight'] as const)
      expect([c, sheetCommandState(chart, c, 1)]).toEqual([
        c,
        { enabled: false, disabledReason: { key: 'disabledReason.notWorksheet' } },
      ]);
    expect(sheetCommandState(chart, 'add', -1).enabled).toBe(true);
    // A worksheet moves past the chartsheet tab.
    expect(sheetCommandState(chart, 'moveRight', 0)).toMatchObject({ enabled: true, to: 1 });
  });
});
