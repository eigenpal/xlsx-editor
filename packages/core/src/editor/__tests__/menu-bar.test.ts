// The menu bar's core: the registry (`CHROME_MENUS`), the arrangement of host rows and menus
// (`arrangeMenuRows`, `arrangeMenuBar`), the keys of the bar and of a panel, the shortcut text,
// and the state and the run of the slots that only the menu bar shows (`menu-commands.ts`).

import { afterEach, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { Window } from 'happy-dom';
import en from '../../../../i18n/en.json';
import { buildXlsx, c, row } from '../../../../../scripts/fixtures/builder.ts';
import {
  CHROME_MENUS,
  chromeMenuShortcut,
  chromeMenuSlots,
  isApplePlatform,
  menuRowKeepsFocus,
  menuKeyWords,
  menuShortcutText,
  tidySeparators,
  type ChromeMenuEntry,
} from '../chrome-menus.ts';
import {
  CHROME_SLOT_IDS,
  chromeControlForSlot,
  chromeSlotId,
  defaultChromeGroups,
  type ChromeSlotId,
} from '../chrome-controls.ts';
import {
  MENU_NAMED_PARTS,
  arrangeMenuBar,
  arrangeMenuRows,
  menuBarTriggerKey,
  menuPanelKey,
  menuStep,
  menuTypeahead,
  menusThatFit,
} from '../menu-arrangement.ts';
import { MENU_ONLY_SLOTS, editorClipboardReader, menuSlotState } from '../menu-commands.ts';
import { CLIPBOARD_TYPE } from '../clipboard.ts';
import { runToolbarCommand, toolbarCommandState } from '../toolbar-commands.ts';
import { createXlsxEditor, type XlsxEditorInstance } from '../index.ts';
import type { FontSource } from '../../layout/index.ts';
import { flushEditor } from './flush-editor.ts';

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

let viewers: XlsxEditorInstance[] = [];
afterEach(() => {
  for (const v of viewers) v.destroy();
  viewers = [];
});

async function open(bytes?: Uint8Array) {
  const win = new Window();
  const doc = win.document as unknown as Document;
  const host = doc.createElement('div');
  doc.body.appendChild(host);
  const viewer = createXlsxEditor({ fonts: FONTS });
  viewers.push(viewer);
  viewer.attach(host);
  await viewer.load(
    bytes ?? buildXlsx({ sheets: [{ name: 'S', rows: row(1, [c('A1', 1), c('B1', 2)]) }] })
  );
  flushEditor(viewer);
  hosts.set(viewer, host);
  return viewer;
}

const hosts = new WeakMap<XlsxEditorInstance, HTMLElement>();

/** A system clipboard of text on the viewer's window: its writes, and a reader of the last one. */
function systemClipboard(viewer: XlsxEditorInstance) {
  const writes: string[] = [];
  const nav = hosts.get(viewer)!.ownerDocument.defaultView!.navigator;
  Object.defineProperty(nav, 'clipboard', {
    configurable: true,
    value: {
      writeText: async (text: string) => void writes.push(text),
      readText: async () => writes.at(-1) ?? '',
    },
  });
  return { writes, read: async () => ({ text: writes.at(-1) ?? '' }) };
}

/** A document copy command that fires the event at the focused element, as a browser does. */
function fireAtGrid(viewer: XlsxEditorInstance) {
  const doc = hosts.get(viewer)!.ownerDocument;
  const writes: string[] = [];
  (doc as unknown as { execCommand: (kind: string) => boolean }).execCommand = (kind) => {
    const event = new (doc.defaultView as unknown as typeof globalThis).Event(kind, {
      bubbles: true,
      cancelable: true,
    });
    Object.defineProperty(event, 'clipboardData', {
      value: { setData: (type: string) => void writes.push(type), getData: () => '', types: [] },
    });
    doc.activeElement!.dispatchEvent(event);
    return true;
  };
  return { writes };
}

/** Let the clipboard read and the paste after it settle. */
async function settle() {
  for (let i = 0; i < 5; i += 1) await Promise.resolve();
}

/** The input of the open cell editor of a viewer that `open` attached. */
function activeInput(viewer: XlsxEditorInstance): HTMLTextAreaElement {
  const input = hosts.get(viewer)?.querySelector<HTMLTextAreaElement>('textarea, input');
  if (!input) throw new Error('no cell editor input');
  return input;
}

/** The English text of a catalog key, or undefined. */
function catalog(key: string): unknown {
  let node: unknown = en;
  for (const part of key.split('.')) node = (node as Record<string, unknown> | undefined)?.[part];
  return node;
}

const slotsOf = (entries: readonly ChromeMenuEntry[]): string[] =>
  entries.flatMap((e) =>
    e.kind === 'item' ? [e.slot] : e.kind === 'submenu' ? slotsOf(e.items) : []
  );

describe('CHROME_MENUS', () => {
  test('the menus are File, Edit, View, Insert, and Format, each with a catalog label', () => {
    expect(CHROME_MENUS.map((m) => m.id)).toEqual(['file', 'edit', 'view', 'insert', 'format']);
    for (const menu of CHROME_MENUS) expect(typeof catalog(menu.labelKey)).toBe('string');
  });

  test('every row is a slot of the registry, and every submenu label is a catalog key', () => {
    const all = new Set<string>(CHROME_SLOT_IDS);
    for (const slot of chromeMenuSlots())
      expect({ slot, known: all.has(slot) }).toEqual({ slot, known: true });
    const walk = (entries: readonly ChromeMenuEntry[]): void => {
      for (const e of entries)
        if (e.kind === 'submenu') {
          expect(typeof catalog(e.labelKey)).toBe('string');
          walk(e.items);
        }
    };
    for (const menu of CHROME_MENUS) walk(menu.entries);
  });

  test('chromeMenuSlots lists the slots of every menu in order, submenus flattened', () => {
    expect<readonly string[]>(chromeMenuSlots()).toEqual(
      CHROME_MENUS.flatMap((m) => slotsOf(m.entries))
    );
    expect(chromeMenuSlots()).toContain('clear.contents');
  });

  test('the slots that only the menu bar shows are not in the preset toolbar', () => {
    const toolbar = new Set(
      defaultChromeGroups().flatMap((g) => g.controls.map((c) => `${g.id}.${c.id}`))
    );
    for (const slot of MENU_ONLY_SLOTS) expect(toolbar.has(slot)).toBe(false);
    expect(toolbar.has('file.save')).toBe(true);
    // Every slot that only the menu bar shows is in a menu.
    const placed = new Set(chromeMenuSlots());
    for (const slot of MENU_ONLY_SLOTS) expect(placed.has(slot)).toBe(true);
  });

  test('a separator never leads, trails, or repeats', () => {
    const check = (entries: readonly ChromeMenuEntry[]) =>
      expect(tidySeparators(entries)).toEqual([...entries]);
    for (const menu of CHROME_MENUS) check(menu.entries);
    expect(
      tidySeparators([
        { kind: 'separator' },
        { kind: 'item', slot: 'font.bold' },
        { kind: 'separator' },
        { kind: 'separator' },
        { kind: 'item', slot: 'font.italic' },
        { kind: 'separator' },
      ])
    ).toEqual([
      { kind: 'item', slot: 'font.bold' },
      { kind: 'separator' },
      { kind: 'item', slot: 'font.italic' },
    ]);
  });

  test('the slots outside the preset toolbar are the menu-only slots, each with a named row', () => {
    const preset = new Set(
      defaultChromeGroups().flatMap((g) => g.controls.map((c) => chromeSlotId(g, c)))
    );
    const outside = CHROME_SLOT_IDS.filter((slot) => !preset.has(slot));
    expect([...outside].sort()).toEqual([...MENU_ONLY_SLOTS].sort());
    const named = new Set<string>(Object.values(MENU_NAMED_PARTS).map((row) => row.slot));
    for (const slot of outside)
      expect({ slot, named: named.has(slot) }).toEqual({ slot, named: true });
  });

  test('the named rows stand for slots of the registry', () => {
    for (const [name, part] of Object.entries(MENU_NAMED_PARTS))
      expect({ name, control: chromeControlForSlot(part.slot) !== null }).toEqual({
        name,
        control: true,
      });
  });
});

describe('row rules in core', () => {
  test('a row finds its registry shortcut wherever it stands, and Zoom out names none', () => {
    expect(chromeMenuShortcut('history.undo')).toBe('Mod+Z');
    expect(chromeMenuShortcut('edit.selectAll')).toBe('Mod+A');
    expect(chromeMenuShortcut('alignment.left')).toBeUndefined();
    // Control or Command and - deletes whole rows or columns on such a selection.
    expect(chromeMenuShortcut('view.zoomOut')).toBeUndefined();
    expect(chromeMenuShortcut('view.zoomIn')).toBe('Mod+=');
  });

  test('only Open and the number format dialog keep the keyboard after a run', () => {
    expect(chromeMenuSlots().filter((slot) => menuRowKeepsFocus(slot))).toEqual([
      'file.open',
      'number.custom',
    ]);
  });
});

describe('menu styles', () => {
  const css = readFileSync(join(import.meta.dir, '../../styles/editor.css'), 'utf8');
  const flat = css.replace(/\s+/g, ' ');

  test('a checked row shows a check mark in its own column, not by color only', () => {
    expect(flat).toContain(".xlsx-menu-row::before { content: ''; flex: 0 0 14px; height: 0; }");
    expect(flat).toMatch(
      /\.xlsx-menu-row\[aria-checked='true'\]::before \{[^}]*border: solid currentColor;[^}]*border-width: 0 2px 2px 0;/
    );
    // The mark is in the text color, which forced colors keep: no forced-colors rule hides it.
    const forced = flat.slice(flat.lastIndexOf('@media (forced-colors: active)'));
    expect(forced).not.toMatch(/aria-checked='true'\]::before/);
  });
});

/** The modifier words of the English catalog. */
const WORDS = menuKeyWords((key) => String(catalog(key)));

describe('shortcut text', () => {
  test('the modifier words come from the catalog, so a locale names its own keys', () => {
    expect(WORDS).toEqual({ ctrl: 'Ctrl', shift: 'Shift', alt: 'Alt' });
    const other = { ctrl: 'Strg', shift: 'Umschalt', alt: 'Alt' };
    expect(menuShortcutText('Mod+Shift+Z', false, other)).toBe('Strg+Umschalt+Z');
    expect(menuShortcutText('Mod+Shift+Z', true, other)).toBe('⌘⇧Z');
  });

  test('Mod is Command on an Apple platform and Control elsewhere', () => {
    expect(menuShortcutText('Mod+Z', false, WORDS)).toBe('Ctrl+Z');
    expect(menuShortcutText('Mod+Z', true, WORDS)).toBe('⌘Z');
    expect(menuShortcutText('Mod+Shift+Z', true, WORDS)).toBe('⌘⇧Z');
    expect(menuShortcutText('Mod+Shift+Z', false, WORDS)).toBe('Ctrl+Shift+Z');
    expect(menuShortcutText('Mod+=', false, WORDS)).toBe('Ctrl+=');
    expect(menuShortcutText('Mod++', false, WORDS)).toBe('Ctrl++');
  });

  test('the platform reads the platform or the user agent', () => {
    expect(isApplePlatform({ platform: 'MacIntel' })).toBe(true);
    expect(isApplePlatform({ userAgent: 'Mozilla/5.0 (iPad; CPU OS 17_0)' })).toBe(true);
    expect(isApplePlatform({ platform: 'Win32' })).toBe(false);
    expect(isApplePlatform(undefined)).toBe(false);
  });
});

describe('arrangeMenuRows', () => {
  const edit = CHROME_MENUS.find((m) => m.id === 'edit')!.entries;
  type Child = { slotId?: string; hidden?: boolean; name?: string };
  const slotOf = (c: Child) => (c.slotId ? { slotId: c.slotId, hidden: c.hidden } : null);

  test('a host row for a registry slot replaces it in place; others follow', () => {
    const copy = { slotId: 'edit.copy', name: 'mine' };
    const extra = { name: 'extra' };
    const r = arrangeMenuRows<Child>(edit, [copy, extra], slotOf);
    expect(r.overrides.get('edit.copy')).toBe(copy);
    expect(r.extras).toEqual([extra]);
    expect(slotsOf(r.entries)).toEqual(slotsOf(edit));
  });

  test('a hidden row is left out with its lone separators, also in a submenu', () => {
    const hide = (slotId: string) => ({ slotId, hidden: true });
    const r = arrangeMenuRows<Child>(
      edit,
      [hide('edit.cut'), hide('edit.copy'), hide('edit.paste'), hide('clear.all')],
      slotOf
    );
    expect(slotsOf(r.entries)).not.toContain('edit.cut');
    expect(slotsOf(r.entries)).not.toContain('clear.all');
    expect(slotsOf(r.entries)).toContain('clear.contents');
    expect(tidySeparators(r.entries)).toEqual(r.entries);
  });

  test('a submenu whose rows are all hidden is left out', () => {
    const hide = (slotId: string) => ({ slotId, hidden: true });
    const r = arrangeMenuRows<Child>(
      edit,
      [hide('clear.contents'), hide('clear.formats'), hide('clear.all')],
      slotOf
    );
    expect(r.entries.some((e) => e.kind === 'submenu')).toBe(false);
  });

  test('a hidden row that names no registry row adds nothing and is reported', () => {
    const unmatched: string[] = [];
    const r = arrangeMenuRows<Child>(edit, [{ slotId: 'font.bold', hidden: true }], slotOf, (s) =>
      unmatched.push(s)
    );
    expect(r.extras).toEqual([]);
    expect(unmatched).toEqual(['font.bold']);
  });

  test('the last host row for a slot wins', () => {
    const a = { slotId: 'edit.copy', name: 'a' };
    const b = { slotId: 'edit.copy', name: 'b' };
    expect(arrangeMenuRows<Child>(edit, [a, b], slotOf).overrides.get('edit.copy')).toBe(b);
  });
});

describe('arrangeMenuBar', () => {
  const t = (key: string) => String(catalog(key) ?? key);
  const specs = (...list: { id: string; hidden?: boolean; after?: string; label?: string }[]) =>
    new Map(list.map((s) => [s.id, s]));

  test('the preset bar is the registry, changed in place', () => {
    const bar = arrangeMenuBar(specs({ id: 'edit', label: 'Change' }), true, t);
    expect(bar.map((m) => m.title)).toEqual(['File', 'Change', 'View', 'Insert', 'Format']);
  });

  test('a hidden menu goes, and a host menu follows its anchor or goes last', () => {
    const bar = arrangeMenuBar(
      specs(
        { id: 'view', hidden: true },
        { id: 'share', after: 'file', label: 'Share' },
        { id: 'help', label: 'Help' }
      ),
      true,
      t
    );
    expect(bar.map((m) => m.id)).toEqual(['file', 'share', 'edit', 'insert', 'format', 'help']);
  });

  test('an anchor that names no menu is a problem, and the menu goes last', () => {
    const problems: string[] = [];
    const bar = arrangeMenuBar(specs({ id: 'share', after: 'nope' }), true, t, (m) =>
      problems.push(m)
    );
    expect(bar.at(-1)?.id).toBe('share');
    expect(problems).toHaveLength(1);
    expect(problems[0]).toContain('after="nope"');
  });

  test('preset off shows only the host menus, and a built-in id keeps its title', () => {
    const bar = arrangeMenuBar(specs({ id: 'edit' }, { id: 'share', label: 'Share' }), false, t);
    expect(bar.map((m) => [m.id, m.title])).toEqual([
      ['edit', 'Edit'],
      ['share', 'Share'],
    ]);
    expect(bar[0]!.registry?.id).toBe('edit');
  });
});

describe('menu keys', () => {
  test('a trigger: the arrows step along the bar, reversed right to left', () => {
    expect(menuBarTriggerKey('ArrowRight', false)).toEqual({ kind: 'step', by: 1 });
    expect(menuBarTriggerKey('ArrowRight', true)).toEqual({ kind: 'step', by: -1 });
    expect(menuBarTriggerKey('ArrowDown', false)).toEqual({ kind: 'open', edge: 'first' });
    expect(menuBarTriggerKey('Enter', false)).toEqual({ kind: 'open', edge: 'first' });
    expect(menuBarTriggerKey('ArrowUp', false)).toEqual({ kind: 'open', edge: 'last' });
    expect(menuBarTriggerKey('Home', false)).toEqual({ kind: 'edge', edge: 'first' });
    expect(menuBarTriggerKey('x', false)).toBeNull();
  });

  test('a panel: the arrows, the ends, Escape, Tab, the bar, a submenu, and type-ahead', () => {
    const top = { rtl: false, submenu: false, onSubmenuRow: false };
    expect(menuPanelKey({ key: 'ArrowDown' }, top)).toEqual({ kind: 'step', by: 1 });
    expect(menuPanelKey({ key: 'End' }, top)).toEqual({ kind: 'edge', edge: 'last' });
    expect(menuPanelKey({ key: 'Escape' }, top)).toEqual({ kind: 'close' });
    expect(menuPanelKey({ key: 'Tab' }, top)).toEqual({ kind: 'leave' });
    expect(menuPanelKey({ key: 'ArrowRight' }, top)).toEqual({ kind: 'bar', by: 1 });
    expect(menuPanelKey({ key: 'ArrowLeft' }, top)).toEqual({ kind: 'bar', by: -1 });
    expect(menuPanelKey({ key: 'ArrowRight' }, { ...top, onSubmenuRow: true })).toEqual({
      kind: 'openSubmenu',
    });
    expect(menuPanelKey({ key: 'ArrowLeft' }, { ...top, submenu: true })).toEqual({
      kind: 'close',
    });
    expect(menuPanelKey({ key: 'ArrowLeft' }, { ...top, rtl: true })).toEqual({
      kind: 'bar',
      by: 1,
    });
    expect(menuPanelKey({ key: 'c' }, top)).toEqual({ kind: 'typeahead', char: 'c' });
    expect(menuPanelKey({ key: 'c', ctrlKey: true }, top)).toBeNull();
    expect(menuPanelKey({ key: ' ' }, top)).toBeNull();
  });

  test('type-ahead finds the next row that starts with the key, wrapping', () => {
    const labels = ['Undo', 'Redo', 'Cut', 'Copy', 'Paste'];
    expect(menuTypeahead(labels, -1, 'c')).toBe(2);
    expect(menuTypeahead(labels, 2, 'C')).toBe(3);
    expect(menuTypeahead(labels, 3, 'c')).toBe(2);
    expect(menuTypeahead(labels, 0, 'z')).toBe(-1);
  });

  test('a step wraps, and from no row goes to an end', () => {
    expect(menuStep(3, 2, 1)).toBe(0);
    expect(menuStep(3, 0, -1)).toBe(2);
    expect(menuStep(3, -1, 1)).toBe(0);
    expect(menuStep(3, -1, -1)).toBe(2);
    expect(menuStep(0, 0, 1)).toBe(-1);
  });

  test('the bar keeps at least one menu beside the More menu', () => {
    expect(menusThatFit([50, 50, 50], 0)).toBe(3);
    expect(menusThatFit([50, 50, 50], 200)).toBe(3);
    expect(menusThatFit([50, 50, 50], 120)).toBe(1);
    expect(menusThatFit([50, 50, 50], 10)).toBe(1);
    expect(menusThatFit([], 10)).toBe(0);
  });
});

describe('the slots of the menu bar', () => {
  test('Open needs the host open, also with no workbook, and waits for an open', async () => {
    const viewer = createXlsxEditor({ fonts: FONTS });
    viewers.push(viewer);
    const snap = viewer.snapshot();
    expect(menuSlotState(snap, 'file.open', {}).disabledReason?.key).toBe('disabledReason.noOpen');
    expect(menuSlotState(snap, 'file.open', { open: () => undefined }).enabled).toBe(true);
    expect(
      menuSlotState(snap, 'file.open', { open: () => undefined, opening: () => true })
        .disabledReason?.key
    ).toBe('disabledReason.loading');
  });

  test('Open, Paste, and the zoom run through the host and the editor', async () => {
    const viewer = await open();
    let opened = 0;
    expect(
      runToolbarCommand(viewer, 'file.open', undefined, { open: () => (opened += 1) }).ok
    ).toBe(true);
    expect(opened).toBe(1);
    expect(toolbarCommandState(viewer.snapshot(), 'edit.paste').disabledReason?.key).toBe(
      'disabledReason.noClipboard'
    );
    const pasted = runToolbarCommand(viewer, 'edit.paste', undefined, {
      readClipboard: async () => ({ text: 'pasted' }),
    });
    expect(pasted.ok).toBe(true);
    await Promise.resolve();
    await Promise.resolve();
    expect(viewer.snapshot().activeCell?.content).toBe('pasted');
    expect(toolbarCommandState(viewer.snapshot(), 'view.zoomReset').disabledReason?.key).toBe(
      'disabledReason.zoomActual'
    );
    runToolbarCommand(viewer, 'view.zoomIn');
    expect(viewer.snapshot().zoom).toBeGreaterThan(1);
    runToolbarCommand(viewer, 'view.zoomReset');
    expect(viewer.snapshot().zoom).toBe(1);
    runToolbarCommand(viewer, 'view.zoomOut');
    expect(viewer.snapshot().zoom).toBeLessThan(1);
  });

  test('Select all, Insert sheet, and the merge commands run on the editor', async () => {
    const viewer = await open();
    expect(toolbarCommandState(viewer.snapshot(), 'merge.merge').disabledReason?.key).toBe(
      'disabledReason.mergeOneCell'
    );
    expect(toolbarCommandState(viewer.snapshot(), 'merge.unmerge').disabledReason?.key).toBe(
      'disabledReason.noMergedRange'
    );
    runToolbarCommand(viewer, 'edit.selectAll');
    expect(viewer.snapshot().selectionRef).not.toBe('A1');
    viewer.select({ row: 2, col: 0 }, { row: 3, col: 1 });
    expect(runToolbarCommand(viewer, 'merge.mergeCenter').ok).toBe(true);
    expect(viewer.snapshot().selectionMerged).toBe(true);
    expect(runToolbarCommand(viewer, 'merge.unmerge').ok).toBe(true);
    expect(runToolbarCommand(viewer, 'sheet.add').ok).toBe(true);
    expect(viewer.snapshot().sheets.length).toBe(2);
  });

  test('a merge over values is disabled with the cell that holds one, never dropping values', async () => {
    const viewer = await open();
    viewer.select({ row: 0, col: 0 }, { row: 0, col: 1 });
    for (const slot of ['merge.merge', 'merge.mergeCenter'] as const) {
      const state = toolbarCommandState(viewer.snapshot(), slot);
      expect(state.enabled).toBe(false);
      expect(state.disabledReason).toEqual({
        key: 'disabledReason.mergeDropsValues',
        params: { cell: 'B1' },
      });
      expect(runToolbarCommand(viewer, slot).ok).toBe(false);
    }
    expect(viewer.snapshot().selectionMerged).toBe(false);
    viewer.select({ row: 0, col: 1 });
    expect(viewer.snapshot().activeCell?.content).toBe('2');
  });

  test('Copy copies through the editor where no clipboard event fires', async () => {
    const viewer = await open();
    const clip = systemClipboard(viewer);
    viewer.select({ row: 0, col: 0 }, { row: 0, col: 1 });
    expect(runToolbarCommand(viewer, 'edit.copy').ok).toBe(true);
    await Promise.resolve();
    expect(clip.writes).toEqual(['1\t2']);
    viewer.select({ row: 0, col: 0 });
    expect(runToolbarCommand(viewer, 'edit.cut').ok).toBe(true);
    await Promise.resolve();
    expect(clip.writes).toEqual(['1\t2', '1']);
  });

  test('Copy writes through the editor when the copy event fires but the grid does not write it', async () => {
    const viewer = await open();
    const clip = systemClipboard(viewer);
    const doc = hosts.get(viewer)!.ownerDocument;
    // The event goes to a node outside the grid, and nothing writes it.
    (doc as unknown as { execCommand: () => boolean }).execCommand = () => true;
    expect(runToolbarCommand(viewer, 'edit.copy').ok).toBe(true);
    await Promise.resolve();
    expect(clip.writes).toEqual(['1']);
  });

  test('Copy writes nothing more when the grid wrote the event, also when a host stops it', async () => {
    const viewer = await open();
    const clip = systemClipboard(viewer);
    const { writes: events } = fireAtGrid(viewer);
    const doc = hosts.get(viewer)!.ownerDocument;
    doc.addEventListener('copy', (e) => e.stopPropagation());
    expect(runToolbarCommand(viewer, 'edit.copy').ok).toBe(true);
    await Promise.resolve();
    expect(events).toEqual(['text/plain', CLIPBOARD_TYPE]);
    expect(clip.writes).toEqual([]);
  });

  test('menu Cut then menu Paste moves the cells, as the keys do', async () => {
    const viewer = await open();
    const clip = systemClipboard(viewer);
    viewer.select({ row: 0, col: 0 });
    expect(runToolbarCommand(viewer, 'edit.cut').ok).toBe(true);
    await Promise.resolve();
    viewer.select({ row: 3, col: 3 });
    runToolbarCommand(viewer, 'edit.paste', undefined, { readClipboard: clip.read });
    await settle();
    expect(viewer.snapshot().activeCell?.content).toBe('1');
    viewer.select({ row: 0, col: 0 });
    expect(viewer.snapshot().activeCell?.content ?? '').toBe('');
  });

  test('menu Paste after an expired cut pastes text equal to the cut as typed, with no refusal', async () => {
    const viewer = await open();
    const clip = systemClipboard(viewer);
    viewer.select({ row: 0, col: 0 });
    expect(runToolbarCommand(viewer, 'edit.cut').ok).toBe(true);
    await Promise.resolve();
    // An edit expires the cut. Then text equal to it comes from elsewhere.
    viewer.enterValue('9', { row: 4, col: 0 });
    clip.writes.push('1');
    viewer.select({ row: 3, col: 3 });
    runToolbarCommand(viewer, 'edit.paste', undefined, { readClipboard: clip.read });
    await settle();
    expect(viewer.snapshot().refusal).toBeNull();
    expect(viewer.snapshot().activeCell?.content).toBe('1');
    viewer.select({ row: 0, col: 0 });
    expect(viewer.snapshot().activeCell?.content).toBe('1');
  });

  describe('which clipboard text pastes the pending copy', () => {
    const formulaRows = row(1, [c('A1', 1)]) + row(2, [c('A2', 2, { f: 'A1+1' })]);
    const formulas = () =>
      open(
        buildXlsx({
          sheets: [
            { name: 'S', rows: formulaRows },
            { name: 'T', rows: row(1, [c('A1', 0)]) },
          ],
        })
      );
    /** Copy A1:A2 with the menu, then menu Paste `text` at C1, and give C2 as it ends. */
    const pasteAt = async (viewer: XlsxEditorInstance, text: string, between?: () => void) => {
      systemClipboard(viewer);
      viewer.select({ row: 0, col: 0 }, { row: 1, col: 0 });
      expect(runToolbarCommand(viewer, 'edit.copy').ok).toBe(true);
      await Promise.resolve();
      between?.();
      viewer.select({ row: 0, col: 2 });
      runToolbarCommand(viewer, 'edit.paste', undefined, { readClipboard: async () => ({ text }) });
      await settle();
      viewer.select({ row: 1, col: 2 });
      return viewer.snapshot().activeCell?.content;
    };

    test('the copy text with other line breaks and a trailing line break pastes the copy', async () => {
      expect(await pasteAt(await formulas(), '1\n2\n')).toBe('=C1+1');
      expect(await pasteAt(await formulas(), '1\r2')).toBe('=C1+1');
    });

    test('after an undo ends the outline, the copy text pastes as typed values', async () => {
      const viewer = await formulas();
      viewer.enterValue('7', { row: 6, col: 6 });
      expect(await pasteAt(viewer, '1\n2', () => void viewer.undo())).toBe('2');
    });

    test('on another sheet, the copy text pastes the copy while its outline shows', async () => {
      // The outline belongs to the store revision, not to the shown sheet: a sheet switch keeps
      // the copy pending, as it keeps it for the paste key.
      const viewer = await formulas();
      expect(await pasteAt(viewer, '1\n2', () => void viewer.setActiveSheet(1))).toBe('=C1+1');
      expect(viewer.snapshot().activeSheet).toBe(1);
    });

    test('a pending cut whose outline ended pastes its text as typed values, and moves nothing', async () => {
      const viewer = await formulas();
      const clip = systemClipboard(viewer);
      viewer.select({ row: 0, col: 0 });
      expect(runToolbarCommand(viewer, 'edit.cut').ok).toBe(true);
      await Promise.resolve();
      const grid = hosts.get(viewer)!.querySelector<HTMLElement>('.xlsx-scroller')!;
      grid.focus();
      const Keyboard = (grid.ownerDocument.defaultView as unknown as typeof globalThis)
        .KeyboardEvent;
      grid.dispatchEvent(new Keyboard('keydown', { key: 'Escape', bubbles: true }));
      viewer.select({ row: 3, col: 3 });
      runToolbarCommand(viewer, 'edit.paste', undefined, { readClipboard: clip.read });
      await settle();
      expect(viewer.snapshot().activeCell?.content).toBe('1');
      viewer.select({ row: 0, col: 0 });
      expect(viewer.snapshot().activeCell?.content).toBe('1');
    });

    test('other text pastes as typed values', async () => {
      expect(await pasteAt(await formulas(), '1\n2x')).toBe('2x');
    });

    test('after an edit ends the outline, the copy text pastes as typed values', async () => {
      const viewer = await formulas();
      expect(await pasteAt(viewer, '1\n2', () => viewer.enterValue('7', { row: 6, col: 6 }))).toBe(
        '2'
      );
    });

    test('after Escape ends the outline, the copy text pastes as typed values', async () => {
      const viewer = await formulas();
      const grid = hosts.get(viewer)!.querySelector<HTMLElement>('.xlsx-scroller')!;
      const escape = () => {
        grid.focus();
        const Keyboard = (grid.ownerDocument.defaultView as unknown as typeof globalThis)
          .KeyboardEvent;
        grid.dispatchEvent(new Keyboard('keydown', { key: 'Escape', bubbles: true }));
      };
      expect(await pasteAt(viewer, '1\n2', escape)).toBe('2');
    });
  });

  test('a copy that the grid refuses runs once, so its refusal shows once', async () => {
    const viewer = await open(
      buildXlsx({
        sheets: [
          {
            name: 'S',
            rows: row(1, [c('A1', 1), c('B1', 2)]),
            afterMerges: '<mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells>',
          },
        ],
      })
    );
    systemClipboard(viewer);
    fireAtGrid(viewer);
    viewer.select({ row: 0, col: 0 }, { row: 1, col: 0 });
    const copy = viewer.copy.bind(viewer);
    let runs = 0;
    viewer.copy = () => {
      runs += 1;
      return copy();
    };
    const rejections: string[] = [];
    viewer.on('rejection', (refusal) => void rejections.push(refusal.code));
    const result = runToolbarCommand(viewer, 'edit.copy');
    expect(runs).toBe(1);
    expect(rejections).toEqual(['merged-range']);
    // The command gives the refusal of the grid's copy.
    expect(result.ok).toBe(false);
    expect(!result.ok && result.refusal.code).toBe('merged-range');
    // Where the browser fires no event, the copy through the editor gives the same result.
    const doc = hosts.get(viewer)!.ownerDocument as unknown as { execCommand: () => boolean };
    doc.execCommand = () => false;
    const direct = runToolbarCommand(viewer, 'edit.copy');
    expect(!direct.ok && direct.refusal.code).toBe('merged-range');
  });

  test("menu Copy uses the editor's own grid, not a grid of an editor nested in its host", async () => {
    const win = new Window();
    const doc = win.document as unknown as Document;
    const host = doc.createElement('div');
    doc.body.appendChild(host);
    // The nested editor's host comes first in the outer host, so its grid comes first.
    const innerHost = doc.createElement('div');
    host.appendChild(innerHost);
    const workbook = () =>
      buildXlsx({ sheets: [{ name: 'S', rows: row(1, [c('A1', 1), c('B1', 2)]) }] });
    const inner = createXlsxEditor({ fonts: FONTS });
    const outer = createXlsxEditor({ fonts: FONTS });
    viewers.push(inner, outer);
    inner.attach(innerHost);
    outer.attach(host);
    await inner.load(workbook());
    await outer.load(workbook());
    flushEditor(inner);
    flushEditor(outer);
    hosts.set(outer, host);
    const clip = systemClipboard(outer);
    const { writes: events } = fireAtGrid(outer);
    expect(runToolbarCommand(outer, 'edit.copy').ok).toBe(true);
    await Promise.resolve();
    // The outer grid wrote the event, so no second copy wrote the system clipboard.
    expect(events).toEqual(['text/plain', CLIPBOARD_TYPE]);
    expect(clip.writes).toEqual([]);
    const outerGrid = host.querySelector(':scope > .xlsx-surface > .xlsx-scroller')!;
    expect(outerGrid.contains(doc.getSelection()!.anchorNode)).toBe(true);
  });

  test('menu Copy then menu Paste keeps the formula and the cell format', async () => {
    const xf =
      '<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>';
    const viewer = await open(
      buildXlsx({
        sheets: [
          { name: 'S', rows: row(1, [c('A1', 45000), c('B1', 45001, { f: 'A1+1', s: 1 })]) },
        ],
        styles: { cellXfs: [xf] },
      })
    );
    const clip = systemClipboard(viewer);
    viewer.select({ row: 0, col: 1 });
    const source = viewer.snapshot().activeFormat;
    expect(runToolbarCommand(viewer, 'edit.copy').ok).toBe(true);
    await Promise.resolve();
    viewer.select({ row: 1, col: 1 });
    runToolbarCommand(viewer, 'edit.paste', undefined, { readClipboard: clip.read });
    await settle();
    // The formula moves with the paste, as a paste of the key moves it: A1 becomes A2.
    expect(viewer.snapshot().activeCell?.content).toBe('=A2+1');
    expect(viewer.snapshot().activeFormat).toEqual(source);
  });

  test('in an open cell edit, Copy and Cut act on the selected text of the edit, as the keys do', async () => {
    const viewer = await open();
    const written: string[] = [];
    viewer.select({ row: 2, col: 0 });
    expect(viewer.beginEdit('hello').ok).toBe(true);
    const input = activeInput(viewer);
    Object.defineProperty(input.ownerDocument.defaultView!.navigator, 'clipboard', {
      configurable: true,
      value: { writeText: async (text: string) => void written.push(text) },
    });
    input.setSelectionRange(0, 2);
    expect(runToolbarCommand(viewer, 'edit.copy').ok).toBe(true);
    expect(viewer.snapshot().editing).toBe(true);
    input.setSelectionRange(0, 2);
    expect(runToolbarCommand(viewer, 'edit.cut').ok).toBe(true);
    expect(viewer.snapshot().editing).toBe(true);
    expect(input.value).toBe('llo');
    await Promise.resolve();
    expect(written).toEqual(['he', 'he']);
    expect(viewer.commitEdit().ok).toBe(true);
    viewer.select({ row: 2, col: 0 });
    expect(viewer.snapshot().activeCell?.content).toBe('llo');
  });

  test('in an open cell edit, Paste puts the text into the edit, as the key does', async () => {
    const viewer = await open();
    viewer.select({ row: 2, col: 0 });
    expect(viewer.beginEdit('hello').ok).toBe(true);
    const input = activeInput(viewer);
    input.setSelectionRange(1, 3);
    const r = runToolbarCommand(viewer, 'edit.paste', undefined, {
      readClipboard: async () => ({ text: 'XY' }),
    });
    expect(r.ok).toBe(true);
    await Promise.resolve();
    await Promise.resolve();
    expect(viewer.snapshot().editing).toBe(true);
    expect(input.value).toBe('hXYlo');
    expect(viewer.commitEdit().ok).toBe(true);
    viewer.select({ row: 2, col: 0 });
    expect(viewer.snapshot().activeCell?.content).toBe('hXYlo');
    // The cell below the edit is unchanged: the paste did not write cells.
    viewer.select({ row: 3, col: 0 });
    expect(viewer.snapshot().activeCell?.content ?? '').toBe('');
  });

  test('a clipboard with HTML and no text is refused with paste-html, as the paste key does', async () => {
    const viewer = await open();
    runToolbarCommand(viewer, 'edit.paste', undefined, {
      readClipboard: async () => ({ text: '', html: true }),
    });
    await Promise.resolve();
    await Promise.resolve();
    expect(viewer.snapshot().refusal?.code).toBe('paste-html');
  });

  test('an empty clipboard changes nothing, as the paste key does', async () => {
    const viewer = await open();
    runToolbarCommand(viewer, 'edit.paste', undefined, {
      readClipboard: async () => ({ text: '' }),
    });
    await Promise.resolve();
    await Promise.resolve();
    expect(viewer.snapshot().refusal).toBeNull();
    expect(viewer.snapshot().activeCell?.content).toBe('1');
  });

  test('the editor clipboard reader reads the window of the editor, with HTML only when no text', async () => {
    expect(editorClipboardReader(null)).toBeUndefined();
    const viewer = await open();
    const nav = hosts.get(viewer)!.ownerDocument.defaultView!.navigator;
    let text = 'cell';
    Object.defineProperty(nav, 'clipboard', {
      configurable: true,
      value: {
        readText: async () => text,
        read: async () => [{ types: ['text/html'] }],
      },
    });
    const read = editorClipboardReader(viewer)!;
    expect(await read()).toEqual({ text: 'cell' });
    text = '';
    expect(await read()).toEqual({ text: '', html: true });
  });

  test('a clipboard read that the browser refuses shows the refusal clipboard-read', async () => {
    const viewer = await open();
    const r = runToolbarCommand(viewer, 'edit.paste', undefined, {
      readClipboard: () => Promise.reject(new Error('refused')),
    });
    expect(r.ok).toBe(true);
    await Promise.resolve();
    await Promise.resolve();
    expect(viewer.snapshot().refusal?.code).toBe('clipboard-read');
    expect(typeof catalog('editing.refused.clipboard-read')).toBe('string');
  });

  test('Cut and Paste are disabled on a locked cell with content of a protected sheet', async () => {
    const viewer = await open();
    const host = { readClipboard: async () => ({ text: '' }) };
    const locked = {
      ...viewer.snapshot(),
      protection: viewer.snapshot().protection ?? ({} as never),
      selectionContentLocked: true,
    };
    for (const slot of ['edit.cut', 'edit.paste'] as const)
      expect(menuSlotState(locked, slot, host).disabledReason?.key).toBe('disabledReason.locked');
    expect(menuSlotState(locked, 'edit.copy', host).enabled).toBe(true);
  });

  test('view mode leaves Copy, Select all, and the zoom, and disables the rest', async () => {
    const viewer = await open();
    viewer.setMode('view');
    const snap = viewer.snapshot();
    const host = { open: () => undefined, readClipboard: async () => ({ text: '' }) };
    for (const slot of MENU_ONLY_SLOTS) {
      const state = menuSlotState(snap, slot as ChromeSlotId, host);
      const reads = ['file.open', 'edit.copy', 'edit.selectAll', 'view.zoomIn', 'view.zoomOut'];
      if (reads.includes(slot)) expect({ slot, on: state.enabled }).toEqual({ slot, on: true });
      else if (slot !== 'view.zoomReset')
        expect({ slot, key: state.disabledReason?.key }).toEqual({
          slot,
          key: 'disabledReason.viewMode',
        });
    }
  });
});
