// The chrome registry and the toolbar's command table (task 7.6): every slot of
// `CHROME_GROUPS` runs a command, the state of a slot comes from the snapshot alone, a pressed
// state is what the press toggles, the structural commands carry the refusal of design D4
// before the press, and every label and disabled reason is a key of the shared catalog.

import { afterEach, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { Window } from 'happy-dom';
import {
  buildXlsx,
  c,
  row,
  type FixtureWorkbook,
} from '../../../../../scripts/fixtures/builder.ts';
import en from '../../../../i18n/en.json';
import { MENU_ONLY_SLOTS } from '../menu-commands.ts';
import {
  CHROME_GROUPS,
  CHROME_SLOT_IDS,
  COLOR_CHOICES,
  CUSTOM_NUMBER_FORMAT_PREFIX,
  DISABLED_REASON_KEYS,
  TOOLBAR_HINT_KEYS,
  MAX_COLS,
  MAX_ROWS,
  createXlsxEditor,
  chromeControlForSlot,
  disabledReasonText,
  refusalText,
  subjectText,
  numberFormatValue,
  runToolbarCommand,
  toolbarCommandState,
  toolbarCommandStates,
  type ChromeGroup,
  type ChromeSlotId,
  type RefusalSubject,
  type EditResult,
  type EditorSnapshot,
  type XlsxEditorInstance,
} from '../index.ts';
import { NUMBER_FORMAT_CHOICES } from '../number-format-commands.ts';
import { flushEditor } from './flush-editor.ts';
import type { FontSource } from '../../layout/index.ts';
import { STRUCTURE_SUBJECT_KINDS } from '../../store/index.ts';
import { NO_DECIMAL_STEPS } from '../decimal-selection.ts';

const fontsDir = join(
  dirname(Bun.resolveSync('@docx-editor.dev/fonts/package.json', import.meta.dir)),
  'assets'
);
const FONTS: FontSource[] = [
  {
    family: 'Carlito',
    bold: false,
    italic: false,
    bytes: new Uint8Array(readFileSync(join(fontsDir, 'Carlito-Regular.ttf'))),
    origin: 'substitute',
  },
];

/** The English catalog, with `{name}` parameters filled in. */
function english(key: string, params?: Readonly<Record<string, string | number>>): string {
  let node: unknown = en;
  for (const part of key.split('.')) node = (node as Record<string, unknown>)[part];
  return String(node).replace(/\{(\w+)\}/g, (_, k: string) => String(params?.[k] ?? ''));
}

let viewers: XlsxEditorInstance[] = [];
afterEach(() => {
  for (const v of viewers) v.destroy();
  viewers = [];
});

const CF =
  '<conditionalFormatting sqref="A1"><cfRule type="expression" dxfId="0" priority="1"><formula>1</formula></cfRule></conditionalFormatting>';

const workbook = (afterMerges = '') =>
  buildXlsx({
    sheets: [
      {
        name: 'S',
        rows: row(1, [c('A1', 1), c('B1', 2)]) + row(2, [c('A2', 3)]),
        afterMerges,
      },
    ],
    styles: { dxfs: ['<dxf/>'] },
  });

async function open(bytes: Uint8Array) {
  const win = new Window();
  const host = (win.document as unknown as Document).createElement('div');
  (win.document as unknown as Document).body.appendChild(host);
  const viewer = createXlsxEditor({ fonts: FONTS });
  viewers.push(viewer);
  viewer.attach(host);
  documents.set(viewer, win.document as unknown as Document);
  await viewer.load(bytes);
  flushEditor(viewer);
  return viewer;
}

const documents = new WeakMap<XlsxEditorInstance, Document>();
/** The document a test viewer is mounted in. */
const document_of = (viewer: XlsxEditorInstance): Document => documents.get(viewer)!;

const EN = { decimal: '.', group: ',' } as const;

/** A snapshot of a ready viewer with a one-cell selection and no refusal. */
const readySnapshot = (): EditorSnapshot => ({
  status: 'ready',
  failure: null,
  partsPastLimit: [],
  themeColors: [],
  sheets: [],
  activeSheet: 0,
  selection: {
    sheetIndex: 0,
    active: { row: 0, col: 0 },
    anchor: { row: 0, col: 0 },
    range: { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } },
  },
  selectionRef: 'A1',
  activeCell: null,
  zoom: 1,
  stale: false,
  editing: false,
  editText: null,
  editChanged: false,
  composing: false,
  canUndo: false,
  canRedo: false,
  refusal: null,
  editDropped: false,
  numberLocale: { decimal: '.', group: ',' },
  activeFormat: null,
  decimalSteps: NO_DECIMAL_STEPS,
  structureRefusal: null,
  formatRefusal: null,
  conditionalHidden: false,
  conditionalFormat: null,
  clearRefusal: null,
  protection: null,
  activeCellLocked: false,
  selectionContentLocked: false,
  selectionContentBlocked: null,
  selectionFormulaRangeBlocked: null,
  selectionFormulaRangeCauseCell: null,
  selectionFormulaRangeCauseUnreadable: null,
  selectionFormulaRangeMasterCell: null,
  selectionMerged: false,
  selectionSingle: true,
  selectionDropsValuesAt: null,
  workbookProtection: null,
  sheetEditRefusal: null,
  mode: 'edit',
  modeRequests: 0,
  revision: 0,
  workbookKind: 'workbook',
  macroEnabled: false,
});

/** A catalog value by its dotted key, or undefined. */
function catalog(key: string): unknown {
  let node: unknown = en;
  for (const part of key.split('.')) node = (node as Record<string, unknown> | undefined)?.[part];
  return node;
}

/** A viewer double that records the method each slot calls. */
function recorder(snapshot: EditorSnapshot) {
  const calls: string[] = [];
  const ok = (name: string) => (): EditResult => {
    calls.push(name);
    return { ok: true };
  };
  const viewer = {
    snapshot: () => snapshot,
    undo: ok('undo'),
    redo: ok('redo'),
    toggleFormat: (w: string) => ok(`toggle:${w}`)(),
    applyFormat: (p: object) => ok(`apply:${Object.keys(p).join(',')}`)(),
    insertRows: ok('insertRows'),
    deleteRows: ok('deleteRows'),
    insertColumns: ok('insertColumns'),
    deleteColumns: ok('deleteColumns'),
    clearValues: ok('clearValues'),
    clearFormats: ok('clearFormats'),
    clearAll: ok('clearAll'),
    saveInTasks: async () => {
      calls.push('save');
      return { bytes: new Uint8Array([1]), revision: 0, stillOpen: true };
    },
  } as unknown as XlsxEditorInstance;
  return { viewer, calls };
}

/** A picked value that each value slot takes. */
const VALUE: Partial<Record<ChromeSlotId, string>> = {
  'font.family': 'Arial',
  'font.size': '14',
  'font.color': 'FF0000',
  'fill.color': 'none',
  'number.format': 'percent',
  'number.custom': '#,##0.000',
};

describe('the chrome registry', () => {
  test('slot ids are unique and follow `${group}.${control}`', () => {
    expect(new Set(CHROME_SLOT_IDS).size).toBe(CHROME_SLOT_IDS.length);
    for (const id of CHROME_SLOT_IDS) expect(id).toMatch(/^[a-z]+\.[a-z][A-Za-z]*$/);
    expect(CHROME_SLOT_IDS).toContain('cells.deleteRows');
  });

  test('no slot is unwired: each one runs a viewer command', async () => {
    const viewer = await open(workbook());
    // A top-aligned cell, so the press on bottom alignment has something to change.
    const snap = viewer.snapshot();
    // Two decimal places, so both decimal steps have a place to change.
    const { viewer: fake, calls } = recorder({
      ...snap,
      activeFormat: {
        ...snap.activeFormat!,
        vertical: 'top',
        numberFormat: { id: 2, code: '0.00', hasCode: true },
      },
    });
    const saved: Uint8Array[] = [];
    // The decimal steps read the cells of the selection, which only an editor that the engine
    // made holds: they run on the viewer, after every other slot.
    const steps = new Set<ChromeSlotId>(['number.increaseDecimals', 'number.decreaseDecimals']);
    for (const slot of CHROME_SLOT_IDS) {
      // The menu bar's own slots have their tests in menu-commands.test.ts.
      if (steps.has(slot) || MENU_ONLY_SLOTS.has(slot)) continue;
      calls.length = 0;
      const r = runToolbarCommand(fake, slot, VALUE[slot], { save: (b) => saved.push(b) });
      expect({ slot, ok: r.ok }).toEqual({ slot, ok: true });
      expect({ slot, calls: calls.length }).toEqual({ slot, calls: 1 });
      const state = toolbarCommandState(viewer.snapshot(), slot, { save: () => undefined });
      expect(state.disabledReason?.key).not.toBe('disabledReason.notWired');
    }
    await Promise.resolve();
    expect(saved.length).toBe(1);
    for (const slot of steps) {
      const before = viewer.snapshot().activeFormat!.numberFormat.code;
      expect({ slot, ok: runToolbarCommand(viewer, slot).ok }).toEqual({ slot, ok: true });
      expect(viewer.snapshot().activeFormat!.numberFormat.code).not.toBe(before);
    }
  });

  test('every label, choice, and disabled reason is a key of the catalog', () => {
    const keys = [
      ...(CHROME_GROUPS as readonly ChromeGroup[]).flatMap((g) => [
        g.labelKey,
        ...g.controls.map((c) => c.labelKey),
      ]),
      ...NUMBER_FORMAT_CHOICES.map((c) => c.labelKey),
      ...COLOR_CHOICES.map((c) => c.labelKey),
      ...DISABLED_REASON_KEYS,
      ...TOOLBAR_HINT_KEYS,
    ];
    for (const key of keys)
      expect({ key, found: typeof catalog(key) }).toEqual({ key, found: 'string' });
  });

  test('an id that the registry does not hold is disabled as not wired', async () => {
    const viewer = await open(workbook());
    const state = toolbarCommandState(viewer.snapshot(), 'font.strike' as ChromeSlotId);
    expect(state.disabledReason?.key).toBe('disabledReason.notWired');
    expect(chromeControlForSlot('font.strike')).toBeNull();
  });
});

describe('toolbarCommandState', () => {
  test('every slot is disabled until a workbook is ready', () => {
    const viewer = createXlsxEditor({ fonts: FONTS });
    viewers.push(viewer);
    for (const slot of CHROME_SLOT_IDS) {
      const s = toolbarCommandState(viewer.snapshot(), slot, { save: () => undefined });
      expect(s.enabled).toBe(false);
      // Open needs no workbook, only the host's open.
      expect(s.disabledReason?.key).toBe(
        slot === 'file.open' ? 'disabledReason.noOpen' : 'disabledReason.editorNotReady'
      );
    }
  });

  test('a clear that would break a formula range disables Clear contents and Clear all', () => {
    const snap = {
      ...readySnapshot(),
      selectionFormulaRangeBlocked: 'S!A2:A3',
      clearRefusal: { code: 'clear-too-large' },
    } as EditorSnapshot;
    for (const id of ['clear.contents', 'clear.all'] as const) {
      const state = toolbarCommandState(snap, id);
      expect(state.enabled).toBe(false);
      // The value clear's reason comes first. Clear formats is refused too, so no pointer to it.
      expect(state.disabledReason).toEqual({
        key: 'disabledReason.formulaRange',
        params: { subject: 'S!A2:A3' },
      });
    }
    // Clear formats changes no value, so only its own refusal disables it.
    expect(toolbarCommandState(snap, 'clear.formats').disabledReason?.key).toBe(
      'disabledReason.clearTooLarge'
    );
    const free = { ...snap, clearRefusal: null } as EditorSnapshot;
    expect(toolbarCommandState(free, 'clear.formats').enabled).toBe(true);
    // While Clear formats works, Clear all points to it.
    expect(toolbarCommandState(free, 'clear.all').disabledReason?.key).toBe(
      'disabledReason.formulaRangeClearAll'
    );
    // The master to leave out is named beside the cell whose formula blocks it.
    const caused = { ...snap, selectionFormulaRangeCauseCell: 'S!A3' } as EditorSnapshot;
    Object.assign(caused, { selectionFormulaRangeMasterCell: 'S!A2' });
    expect(toolbarCommandState(caused, 'clear.all').disabledReason).toEqual({
      key: 'disabledReason.formulaCell',
      params: { causeCell: 'S!A3', masterCell: 'S!A2' },
    });
    expect(
      toolbarCommandState({ ...caused, clearRefusal: null } as EditorSnapshot, 'clear.all')
        .disabledReason?.key
    ).toBe('disabledReason.formulaCellClearAll');
  });

  test('the state list keeps the order of the ids and equals the single-slot states', () => {
    const viewer = createXlsxEditor({ fonts: FONTS });
    viewers.push(viewer);
    const ids = [...CHROME_SLOT_IDS].reverse();
    const host = { save: () => undefined };
    const states = toolbarCommandStates(viewer.snapshot(), ids, host);
    expect(states.map((s) => s.id)).toEqual(ids);
    expect(states).toEqual(ids.map((id) => toolbarCommandState(viewer.snapshot(), id, host)));
  });

  test('undo and redo follow the history, with their reasons', async () => {
    const viewer = await open(workbook());
    const state = (slot: ChromeSlotId) => toolbarCommandState(viewer.snapshot(), slot);
    expect(state('history.undo').disabledReason?.key).toBe('disabledReason.nothingToUndo');
    expect(state('history.redo').disabledReason?.key).toBe('disabledReason.nothingToRedo');
    expect(runToolbarCommand(viewer, 'font.bold').ok).toBe(true);
    expect(state('history.undo').enabled).toBe(true);
    expect(runToolbarCommand(viewer, 'history.undo').ok).toBe(true);
    expect(state('history.redo').enabled).toBe(true);
  });

  test('undo cancels a changed cell edit, closes an unchanged one and steps once', async () => {
    const viewer = await open(workbook());
    const state = (slot: ChromeSlotId) => toolbarCommandState(viewer.snapshot(), slot);
    const el = document_of(viewer).querySelector(
      '.xlsx-cell-editor'
    ) as unknown as HTMLTextAreaElement;
    const type = (text: string) => {
      el.value = text;
      el.dispatchEvent(
        new (el.ownerDocument!.defaultView as unknown as typeof globalThis).Event('input')
      );
    };
    // An unchanged edit with an empty history: undo is disabled, as its command would refuse.
    expect(viewer.beginEdit().ok).toBe(true);
    expect(state('history.undo').enabled).toBe(false);
    // Changed text: undo is enabled, and the press cancels the edit. The history stays.
    type('99');
    expect(state('history.undo').enabled).toBe(true);
    expect(runToolbarCommand(viewer, 'history.undo').ok).toBe(true);
    expect(viewer.snapshot().editing).toBe(false);
    expect(viewer.snapshot().activeCell?.content).toBe('1');
    expect(viewer.snapshot().revision).toBe(0);
    // A type-over: the same rule, and the cell keeps its content.
    viewer.select({ row: 1, col: 0 });
    expect(viewer.enterValue('5').ok).toBe(true);
    viewer.select({ row: 0, col: 0 });
    viewer.beginEdit('x');
    expect(runToolbarCommand(viewer, 'history.undo').ok).toBe(true);
    expect(viewer.snapshot().editing).toBe(false);
    expect(viewer.snapshot().activeCell?.content).toBe('1');
    expect(viewer.snapshot().canUndo).toBe(true);
    // Unchanged text: undo is enabled by the history, closes the edit, and steps it once.
    viewer.beginEdit();
    expect(viewer.snapshot().editChanged).toBe(false);
    expect(state('history.undo').enabled).toBe(true);
    expect(runToolbarCommand(viewer, 'history.undo').ok).toBe(true);
    expect(viewer.snapshot().editing).toBe(false);
    expect(viewer.snapshot().canUndo).toBe(false);
    // The undone edit is the redo step, and a changed edit disables redo: its commit clears it.
    expect(state('history.redo').enabled).toBe(true);
    viewer.beginEdit('7');
    expect(viewer.snapshot().editChanged).toBe(true);
    expect(state('history.redo').disabledReason?.key).toBe('disabledReason.editPending');
  });

  test('a structure refusal without a subject has its own reason', () => {
    const snapshot: EditorSnapshot = {
      ...readySnapshot(),
      structureRefusal: { code: 'structure-blocked' },
    };
    expect(toolbarCommandState(snapshot, 'cells.insertRows').disabledReason).toEqual({
      key: 'disabledReason.structureBlockedUnnamed',
    });
  });

  test('a custom number format code that reads like a choice is not that choice', () => {
    expect(numberFormatValue({ id: 10, code: '0.00%' }, EN)).toBe('percent');
    expect(numberFormatValue({ id: 164, code: 'percent' }, EN)).toBe(
      `${CUSTOM_NUMBER_FORMAT_PREFIX}percent`
    );
    // A workbook that gives id 10 another code shows that code: no choice.
    expect(numberFormatValue({ id: 10, code: '0%' }, EN)).toBe(`${CUSTOM_NUMBER_FORMAT_PREFIX}0%`);
  });

  test('a pressed state is what the press toggles', async () => {
    const viewer = await open(workbook());
    const toggles: ChromeSlotId[] = [
      'font.bold',
      'font.italic',
      'font.underline',
      'alignment.left',
      'alignment.center',
      'alignment.right',
      'alignment.top',
      'alignment.middle',
      'alignment.wrap',
    ];
    for (const slot of toggles) {
      const before = toolbarCommandState(viewer.snapshot(), slot).active;
      expect(runToolbarCommand(viewer, slot).ok).toBe(true);
      expect({ slot, active: toolbarCommandState(viewer.snapshot(), slot).active }).toEqual({
        slot,
        active: !before,
      });
      expect(runToolbarCommand(viewer, slot).ok).toBe(true);
      expect({ slot, active: toolbarCommandState(viewer.snapshot(), slot).active }).toEqual({
        slot,
        active: before,
      });
    }
  });

  test('one alignment of an axis is pressed at a time', async () => {
    const viewer = await open(workbook());
    runToolbarCommand(viewer, 'alignment.left');
    runToolbarCommand(viewer, 'alignment.right');
    const s = viewer.snapshot();
    expect(toolbarCommandState(s, 'alignment.left').active).toBe(false);
    expect(toolbarCommandState(s, 'alignment.right').active).toBe(true);
  });

  test('value slots show the active cell format and set it', async () => {
    const viewer = await open(workbook());
    const value = (slot: ChromeSlotId) => toolbarCommandState(viewer.snapshot(), slot).value;
    expect(value('number.format')).toBe('general');
    expect(value('fill.color')).toBe('none');
    for (const [slot, v] of Object.entries(VALUE) as [ChromeSlotId, string][]) {
      expect(runToolbarCommand(viewer, slot, v).ok).toBe(true);
      expect({ slot, value: value(slot) }).toEqual({ slot, value: v });
    }
    expect(runToolbarCommand(viewer, 'fill.color', '00B050').ok).toBe(true);
    expect(value('fill.color')).toBe('00B050');
  });

  test('while a workbook loads, every slot is refused with loading, also a toolbar refusal', async () => {
    const viewer = await open(workbook());
    const loading = viewer.load(workbook());
    const refusal = (r: EditResult) => (r.ok ? null : r.refusal);
    // A save with no host, a save with one, a value the slot does not take, and a command:
    // each result matches the refusal that the snapshot keeps.
    expect(refusal(runToolbarCommand(viewer, 'file.save'))).toEqual({ code: 'loading' });
    expect(viewer.snapshot().refusal).toEqual({ code: 'loading' });
    const saved: Uint8Array[] = [];
    const host = { save: (b: Uint8Array) => saved.push(b) };
    expect(refusal(runToolbarCommand(viewer, 'file.save', undefined, host))).toEqual({
      code: 'loading',
      command: 'save',
    });
    expect(saved.length).toBe(0);
    expect(viewer.snapshot().refusal).toEqual({ code: 'loading', command: 'save' });
    expect(refusal(runToolbarCommand(viewer, 'font.size', ''))).toEqual({ code: 'loading' });
    expect(viewer.snapshot().refusal).toEqual({ code: 'loading' });
    expect(refusal(runToolbarCommand(viewer, 'font.size', '14'))).toEqual({ code: 'loading' });
    expect(refusal(runToolbarCommand(viewer, 'history.undo'))).toEqual({ code: 'loading' });
    // A toggle reads no format of the old workbook: it is refused, and writes nothing.
    expect(viewer.snapshot().activeFormat).toBeNull();
    expect(refusal(runToolbarCommand(viewer, 'alignment.bottom'))).toEqual({ code: 'loading' });
    expect(refusal(runToolbarCommand(viewer, 'alignment.left'))).toEqual({ code: 'loading' });
    await loading;
    expect(refusal(runToolbarCommand(viewer, 'font.size', ''))?.code).toBe('bad-format');
  });

  test('a value that a slot does not take is refused and writes nothing', async () => {
    const viewer = await open(workbook());
    for (const [slot, v] of [
      ['font.color', 'red'],
      ['font.size', ''],
      ['number.format', '0.000'],
      ['number.custom', '"open'],
      ['number.custom', ''],
      ['font.family', '  '],
    ] as const) {
      const r = runToolbarCommand(viewer, slot, v);
      expect(r.ok).toBe(false);
    }
    expect(runToolbarCommand(viewer, 'font.color').ok).toBe(false);
    expect(viewer.snapshot().canUndo).toBe(false);
  });

  test('save is disabled without a host receiver, and hands the bytes to it', async () => {
    const viewer = await open(workbook());
    expect(toolbarCommandState(viewer.snapshot(), 'file.save').disabledReason?.key).toBe(
      'disabledReason.noSave'
    );
    let bytes: Uint8Array | null = null;
    const host = { save: (b: Uint8Array) => (bytes = b) };
    expect(toolbarCommandState(viewer.snapshot(), 'file.save', host).enabled).toBe(true);
    expect(runToolbarCommand(viewer, 'file.save', undefined, host).ok).toBe(true);
    while (!bytes) await new Promise((resolve) => setTimeout(resolve, 1));
    expect((bytes as Uint8Array).byteLength).toBeGreaterThan(0);
  });

  test('a save that an open edit blocks reports the refusal of the edit', async () => {
    const viewer = await open(workbook());
    expect(viewer.beginEdit('=1+', 'enter').ok).toBe(true);
    let bytes: Uint8Array | null = null;
    const host = { save: (b: Uint8Array) => (bytes = b) };
    const r = runToolbarCommand(viewer, 'file.save', undefined, host);
    expect(r.ok).toBe(false);
    expect(r.ok ? null : r.refusal.code).toBe(viewer.snapshot().refusal!.code);
    expect(r.ok ? null : r.refusal.code).not.toBe('not-editable');
    expect(bytes).toBeNull();
  });

  test('a preserved conditional formatting range disables the structural commands, by name', async () => {
    const viewer = await open(workbook(CF));
    viewer.select({ row: 0, col: 0 }, { row: 0, col: MAX_COLS - 1 });
    const subject = 'x:conditionalFormatting (/xl/worksheets/sheet1.xml)';
    for (const slot of [
      'cells.insertRows',
      'cells.deleteRows',
      'cells.insertColumns',
      'cells.deleteColumns',
    ] as const) {
      const s = toolbarCommandState(viewer.snapshot(), slot);
      expect(s.enabled).toBe(false);
      expect(s.disabledReason).toEqual({
        key: 'disabledReason.structureBlocked',
        about: { kind: 'conditional-formatting', name: 'S' },
      });
    }
    // The user text names the content in words; the part path stays in the refusal for logs.
    const reason = toolbarCommandState(viewer.snapshot(), 'cells.insertRows').disabledReason!;
    expect(disabledReasonText(reason, english)).toBe(
      'Rows and columns cannot be inserted or deleted: conditional formatting on sheet S can refer to cells that would move, and the editor cannot move such a reference.'
    );
    // The press gives the same answer as the state.
    const r = runToolbarCommand(viewer, 'cells.deleteRows');
    expect(r.ok ? null : r.refusal.subject).toBe(subject);
    expect(r.ok ? '' : r.refusal.detail).toContain('/xl/worksheets/sheet1.xml');
    expect(r.ok ? '' : refusalText(r.refusal, english)).toBe(
      'The edit was not applied: conditional formatting on sheet S can refer to content that the edit changes, and the editor cannot update that reference.'
    );
    // Values and formats stay available.
    expect(toolbarCommandState(viewer.snapshot(), 'font.bold').enabled).toBe(true);
    expect(toolbarCommandState(viewer.snapshot(), 'clear.contents').enabled).toBe(true);
  });

  test('an unsupported function call names the refused function', async () => {
    const viewer = await open(buildXlsx({ sheets: [{ name: 'S', rows: row(1, [c('A1', 1)]) }] }));
    viewer.select({ row: 1, col: 0 });
    const r = viewer.enterValue('=SUM(A1)+IFNA(A1,0)');
    expect(r.ok ? null : r.refusal.subject).toBe('IFNA');
    expect(r.ok ? '' : refusalText(r.refusal, english)).toBe(
      'The formula was not applied: the editor cannot write a call to IFNA. Use a predefined function.'
    );
  });

  test('an insert that pushes a value out of the grid names the content, and a formula as a formula', async () => {
    const viewer = await open(
      buildXlsx({
        sheets: [
          {
            name: 'S',
            rows: row(1, [c('A1', 1), c('XFD1', 2)]) + row(2, [c('XFD2', 0, { f: 'A1' })]),
          },
        ],
      })
    );
    const r = runToolbarCommand(viewer, 'cells.insertColumns');
    expect(r.ok ? null : r.refusal.code).toBe('grid-full');
    expect(r.ok ? '' : refusalText(r.refusal, english)).toContain('the content of S!XFD1');
    viewer.select({ row: 0, col: 0 });
    expect(viewer.clearValues().ok).toBe(true);
    viewer.select({ row: 0, col: 16383 });
    expect(viewer.clearValues().ok).toBe(true);
    viewer.select({ row: 0, col: 0 });
    const f = runToolbarCommand(viewer, 'cells.insertColumns');
    expect(f.ok ? '' : refusalText(f.refusal, english)).toContain('the formula in S!XFD2');
  });

  test('a structural command is enabled on a sheet with no blocker', async () => {
    const viewer = await open(workbook());
    expect(toolbarCommandState(viewer.snapshot(), 'cells.deleteRows').enabled).toBe(true);
    expect(runToolbarCommand(viewer, 'cells.insertRows').ok).toBe(true);
  });

  test('on a protected sheet, an unlocked cell clears its contents, and formatCells allows formats', async () => {
    const viewer = await open(
      buildXlsx({
        sheets: [
          {
            name: 'S',
            rows: row(1, [c('A1', 1, { s: 1 }), c('B1', 2)]),
            beforeMerges: '<sheetProtection sheet="1" formatCells="0"/>',
          },
        ],
        styles: {
          cellXfs: [
            '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyProtection="1">' +
              '<protection locked="0"/></xf>',
          ],
        },
      })
    );
    const s = viewer.snapshot();
    expect(s.activeCellLocked).toBe(false);
    for (const id of ['font.bold', 'clear.formats', 'clear.all', 'clear.contents'] as const)
      expect(toolbarCommandState(s, id).enabled).toBe(true);
    viewer.select({ row: 0, col: 1 });
    const locked = viewer.snapshot();
    expect(locked.activeCellLocked).toBe(true);
    expect(toolbarCommandState(locked, 'font.bold').enabled).toBe(true);
    expect(toolbarCommandState(locked, 'clear.contents').disabledReason?.key).toBe(
      'disabledReason.locked'
    );
    expect(toolbarCommandState(locked, 'clear.all').disabledReason?.key).toBe(
      'disabledReason.locked'
    );
    // A range with an unlocked active cell and a locked cell with content: the store refuses
    // the clear, and the state says so before the press.
    viewer.select({ row: 0, col: 0 }, { row: 0, col: 1 });
    const mixed = viewer.snapshot();
    expect([mixed.activeCellLocked, mixed.selectionContentLocked]).toEqual([false, true]);
    expect(toolbarCommandState(mixed, 'clear.contents').disabledReason?.key).toBe(
      'disabledReason.locked'
    );
    const refused = runToolbarCommand(viewer, 'clear.contents');
    expect(refused.ok ? null : refused.refusal.code).toBe('locked');
    // A range whose locked cells are empty clears.
    viewer.select({ row: 0, col: 0 }, { row: 5, col: 0 });
    expect(viewer.snapshot().selectionContentLocked).toBe(false);
    expect(runToolbarCommand(viewer, 'clear.contents').ok).toBe(true);
  });

  test('a whole column disables the format commands, and clearing stays available', async () => {
    const viewer = await open(workbook());
    viewer.select({ row: 0, col: 0 }, { row: MAX_ROWS - 1, col: 0 });
    const s = viewer.snapshot();
    const bold = toolbarCommandState(s, 'font.bold');
    expect(bold.disabledReason?.key).toBe('disabledReason.rangeTooLarge');
    expect(toolbarCommandState(s, 'number.format').enabled).toBe(false);
    expect(toolbarCommandState(s, 'clear.formats').enabled).toBe(true);
    // The command refuses the same range.
    const r = runToolbarCommand(viewer, 'font.bold');
    expect(r.ok ? null : r.refusal.code).toBe('range-too-large');
    // One whole row formats.
    viewer.select({ row: 0, col: 0 }, { row: 0, col: MAX_COLS - 1 });
    expect(toolbarCommandState(viewer.snapshot(), 'font.bold').enabled).toBe(true);
  });

  test('a protected sheet disables the locked commands with the locked reason', async () => {
    const viewer = await open(
      buildXlsx({
        sheets: [
          {
            name: 'S',
            rows: row(1, [c('A1', 1)]),
            beforeMerges: '<sheetProtection sheet="1" insertRows="0"/>',
          },
        ],
      })
    );
    const s = viewer.snapshot();
    expect(s.protection?.allowInsertRows).toBe(true);
    const reason = (id: ChromeSlotId) => toolbarCommandState(s, id).disabledReason?.key;
    // The tooltip names the protection rule that the command would be refused for.
    expect(reason('font.bold')).toBe('disabledReason.formatLocked');
    expect(reason('cells.deleteRows')).toBe('disabledReason.structureLocked');
    expect(reason('cells.insertColumns')).toBe('disabledReason.structureLocked');
    expect(toolbarCommandState(s, 'cells.insertRows').enabled).toBe(true);
    // The commands refuse alike, and a locked cell opens no editor.
    const r = runToolbarCommand(viewer, 'font.bold');
    expect(r.ok ? null : r.refusal.code).toBe('locked');
    const typed = viewer.beginEdit('7');
    expect(typed.ok ? null : typed.refusal.code).toBe('locked');
    expect(viewer.snapshot().editing).toBe(false);
    const entered = viewer.enterValue('7');
    expect(entered.ok ? null : entered.refusal.code).toBe('locked');
    expect(runToolbarCommand(viewer, 'cells.insertRows').ok).toBe(true);
    // Clearing formats needs formatCells; clearing contents needs an unlocked active cell.
    const now = viewer.snapshot();
    expect(now.activeCellLocked).toBe(true);
    for (const [id, key] of [
      ['clear.formats', 'disabledReason.formatLocked'],
      ['clear.all', 'disabledReason.locked'],
      ['clear.contents', 'disabledReason.locked'],
    ] as const) {
      expect(toolbarCommandState(now, id).disabledReason?.key).toBe(key);
      const refused = runToolbarCommand(viewer, id);
      expect(refused.ok ? null : refused.refusal.code).toBe('locked');
    }
  });

  test('on a dense sheet, the state counts only the cells inside the range, as the command does', async () => {
    // 20 000 cells in A1:J2000, so the sheet holds more cells than one command may create.
    const cols = 'ABCDEFGHIJ';
    let rows = '';
    for (let r = 1; r <= 2000; r += 1)
      rows += row(
        r,
        [...cols].map((col) => c(`${col}${r}`, r))
      );
    const viewer = await open(buildXlsx({ sheets: [{ name: 'S', rows }] }));
    // L1:M10000 holds no cell: a patch would create 20 000 cells.
    viewer.select({ row: 0, col: 11 }, { row: 9999, col: 12 });
    const state = toolbarCommandState(viewer.snapshot(), 'font.bold');
    expect(state.disabledReason?.key).toBe('disabledReason.rangeTooLarge');
    const r = runToolbarCommand(viewer, 'font.bold');
    expect(r.ok ? null : r.refusal.code).toBe('range-too-large');
    // A1:B10000 holds 4 000 cells: 16 000 new cells fit, and both agree.
    viewer.select({ row: 0, col: 0 }, { row: 9999, col: 1 });
    expect(toolbarCommandState(viewer.snapshot(), 'font.bold').enabled).toBe(true);
  });

  test('clearing a formatted column disables the clear commands, as the command refuses', async () => {
    const viewer = await open(
      buildXlsx({
        sheets: [
          {
            name: 'S',
            cols: '<cols><col min="2" max="2" width="9" style="1" customWidth="1"/></cols>',
            rows: row(1, [c('A1', 1)]),
          },
        ],
        styles: {
          cellXfs: ['<xf numFmtId="0" fontId="0" fillId="1" borderId="0" xfId="0" applyFill="1"/>'],
        },
      })
    );
    // Column B shows its format in each empty position: a clear writes s="0" in each one.
    viewer.select({ row: 0, col: 1 }, { row: MAX_ROWS - 1, col: 1 });
    for (const id of ['clear.formats', 'clear.all'] as const)
      expect(toolbarCommandState(viewer.snapshot(), id).disabledReason?.key).toBe(
        'disabledReason.clearTooLarge'
      );
    const r = viewer.clearFormats();
    expect(r.ok ? null : r.refusal.code).toBe('clear-too-large');
    // Column A has no format: a clear of it creates nothing and stays available.
    viewer.select({ row: 0, col: 0 }, { row: MAX_ROWS - 1, col: 0 });
    expect(toolbarCommandState(viewer.snapshot(), 'clear.formats').enabled).toBe(true);
    // A short range of column B clears, and undo takes the new cells away.
    viewer.select({ row: 0, col: 1 }, { row: 2, col: 1 });
    expect(viewer.clearFormats().ok).toBe(true);
    expect(viewer.snapshot().canUndo).toBe(true);
    expect(viewer.undo().ok).toBe(true);
    expect(viewer.snapshot().canUndo).toBe(false);
  });

  test('undo and redo select the range that they change', async () => {
    const viewer = await open(workbook());
    viewer.select({ row: 2, col: 2 });
    expect(viewer.enterValue('5').ok).toBe(true);
    viewer.select({ row: 0, col: 1 }, { row: 1, col: 3 });
    expect(runToolbarCommand(viewer, 'font.bold').ok).toBe(true);
    viewer.select({ row: 7, col: 0 });
    expect(viewer.undo().ok).toBe(true);
    expect(viewer.snapshot().selectionRef).toBe('B1:D2');
    expect(viewer.undo().ok).toBe(true);
    expect(viewer.snapshot().selectionRef).toBe('C3');
    expect(viewer.redo().ok).toBe(true);
    expect(viewer.snapshot().selectionRef).toBe('C3');
  });
});

const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
const enc = (text: string) => new TextEncoder().encode(text);
const cfOf = (formula: string) =>
  `<conditionalFormatting sqref="A1"><cfRule type="expression" dxfId="0" priority="1"><formula>${formula}</formula></cfRule></conditionalFormatting>`;
const data = { name: 'Data', rows: row(1, [c('A1', 1)]) + row(3, [c('A3', 3)]) };
const tableSheet = {
  name: 'Other',
  rows: row(1, [c('A1', 'A'), c('B1', 'B')]),
  afterPageSetup: '<tableParts count="1"><tablePart r:id="rIdT"/></tableParts>',
  rels: [['rIdT', `${R}table`, '../tables/table1.xml']] as const,
};
const tablePart = enc(
  '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="T" displayName="T" ref="A1:B3">' +
    '<tableColumns count="2"><tableColumn id="1" name="A"/>' +
    '<tableColumn id="2" name="B"><calculatedColumnFormula>Data!A1*2</calculatedColumnFormula></tableColumn>' +
    '</tableColumns></table>'
);

/** One workbook for each refusal kind of the structural commands, with `Data` shown. */
const REFUSAL_CASES: readonly (readonly [string, FixtureWorkbook, string, RefusalSubject?])[] = [
  [
    'conditional formatting on the edited sheet',
    { sheets: [{ ...data, afterMerges: cfOf('1') }], styles: { dxfs: ['<dxf/>'] } },
    'disabledReason.structureBlocked',
    { kind: 'conditional-formatting', name: 'Data' },
  ],
  [
    'data validation on the edited sheet',
    {
      sheets: [
        {
          ...data,
          afterMerges: '<dataValidations count="1"><dataValidation sqref="A1"/></dataValidations>',
        },
      ],
    },
    'disabledReason.structureBlocked',
    { kind: 'data-validation', name: 'Data' },
  ],
  [
    'a chart part',
    {
      sheets: [data],
      extraParts: [
        [
          'xl/charts/chart1.xml',
          'application/vnd.openxmlformats-officedocument.drawingml.chart+xml',
          enc('<c:chartSpace xmlns:c="urn:c"/>'),
        ],
      ],
    },
    'disabledReason.structureBlocked',
    { kind: 'chart' },
  ],
  [
    'a table on another sheet that names the edited sheet',
    {
      sheets: [data, tableSheet],
      extraParts: [
        [
          'xl/tables/table1.xml',
          'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml',
          tablePart,
        ],
      ],
    },
    'disabledReason.structureBlocked',
    { kind: 'table' },
  ],
  [
    'a query table that names the edited sheet',
    {
      sheets: [data],
      extraParts: [
        [
          'xl/queryTables/queryTable1.xml',
          'application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml',
          enc('<queryTable xmlns="urn:q" name="Q" connectionId="1" note="Data!A1"/>'),
        ],
      ],
    },
    'disabledReason.structureBlocked',
    { kind: 'workbook-content' },
  ],
  [
    'a workbook relationship outside the allowlist',
    {
      sheets: [data],
      extraParts: [
        ['xl/connections.xml', 'application/xml', enc('<x xmlns="urn:x"/>'), `${R}connections`],
      ],
    },
    'disabledReason.structureBlocked',
    { kind: 'workbook-content' },
  ],
  [
    'a sheet range that spans the edited sheet',
    {
      sheets: [
        { name: 'First', rows: '' },
        data,
        { name: 'Last', rows: '' },
        { name: 'Other', rows: '', afterMerges: cfOf('SUM(First:Last!A1)&gt;1') },
      ],
      styles: { dxfs: ['<dxf/>'] },
      activeTab: 1,
    },
    'disabledReason.structureBlocked',
    { kind: 'conditional-formatting', name: 'Other' },
  ],
];

describe('the structural state and the structural commands refuse alike', () => {
  const SLOTS = [
    'cells.insertRows',
    'cells.deleteRows',
    'cells.insertColumns',
    'cells.deleteColumns',
  ] as const;

  /** The state of each structural slot equals the refusal of its press, which changes nothing. */
  function expectSame(viewer: XlsxEditorInstance, key: string) {
    const refusal = viewer.snapshot().structureRefusal;
    expect(refusal).not.toBeNull();
    for (const slot of SLOTS) {
      const state = toolbarCommandState(viewer.snapshot(), slot);
      expect(state.enabled).toBe(false);
      expect(state.disabledReason?.key).toBe(key as never);
      if (refusal?.about !== undefined) expect(state.disabledReason?.about).toEqual(refusal.about);
      const before = viewer.save();
      const r = runToolbarCommand(viewer, slot);
      expect(r.ok ? null : r.refusal).toEqual(refusal);
      expect(viewer.save()).toEqual(before);
    }
  }

  for (const [what, wb, key, about] of REFUSAL_CASES)
    test(what, async () => {
      const viewer = await open(buildXlsx(wb));
      expect(viewer.snapshot().sheets[viewer.snapshot().activeSheet]?.name).toBe('Data');
      expect(viewer.snapshot().structureRefusal?.about).toEqual(about);
      expectSame(viewer, key);
      // The user text names the content in words, never a part path.
      const reason = toolbarCommandState(viewer.snapshot(), 'cells.insertRows').disabledReason!;
      const text = disabledReasonText(reason, english);
      expect(text).not.toContain('/xl/');
      expect(text).not.toContain('{');
    });

  test('a whole-sheet selection', async () => {
    const viewer = await open(workbook());
    viewer.select({ row: 0, col: 0 }, { row: MAX_ROWS - 1, col: MAX_COLS - 1 });
    expect(viewer.snapshot().structureRefusal).toEqual({ code: 'whole-sheet' });
    expectSame(viewer, 'disabledReason.wholeSheet');
    // A smaller selection enables the commands again.
    viewer.select({ row: 0, col: 0 }, { row: 0, col: 0 });
    expect(toolbarCommandState(viewer.snapshot(), 'cells.insertRows').enabled).toBe(true);
  });
});

describe('structure subject nouns', () => {
  test('every subject kind has a catalog noun, and every noun a kind', () => {
    const nouns = Object.keys(
      (en as { structureSubject: Record<string, string> }).structureSubject
    );
    expect([...STRUCTURE_SUBJECT_KINDS].sort() as string[]).toEqual(nouns.sort());
    for (const kind of STRUCTURE_SUBJECT_KINDS) {
      const text = subjectText({ kind, name: 'Data' }, english);
      expect(text.length).toBeGreaterThan(0);
      expect(text).not.toContain('{');
    }
  });

  test('a file name loses bidirectional controls and is cut to 64 characters', () => {
    const spoofed = subjectText({ kind: 'defined-name', name: 'Rates\u202Egnp.exe' }, english);
    expect(spoofed).toBe('the defined name Ratesgnp.exe');
    const long = subjectText({ kind: 'defined-name', name: 'N'.repeat(200) }, english);
    expect(long).toBe(`the defined name ${'N'.repeat(63)}…`);
  });

  test('a range-too-large refusal names the cell limit', () => {
    const text = refusalText({ code: 'range-too-large' }, english);
    expect(text).toContain('16,384 cells');
    expect(
      refusalText({ code: 'range-too-large' }, english, { decimal: ',', group: '.' })
    ).toContain('16.384 cells');
    expect(text).not.toContain('{');
    // A copy names the same limit, with the separator of the locale.
    expect(refusalText({ code: 'copy-too-large' }, english)).toContain('at most 16,384 cells');
    expect(
      refusalText({ code: 'copy-too-large' }, english, { decimal: ',', group: '.' })
    ).toContain('at most 16.384 cells');
  });

  test('a link of another sheet names the sheet in words', () => {
    const text = refusalText(
      {
        code: 'structure-blocked',
        subject: '/xl/ctrlProps/ctrlProp1.xml',
        about: { kind: 'control', name: 'Other' },
      },
      english
    );
    expect(text).toContain('a control on sheet Other that is linked to a cell');
  });

  test('a grid-full refusal of a drawing names it in words', () => {
    const text = refusalText(
      { code: 'grid-full', subject: '/xl/drawings/drawing1.xml', about: { kind: 'drawing' } },
      english
    );
    expect(text).toBe(
      'The edit was not applied: a drawing would move past the last row or column of the sheet.'
    );
  });
});
