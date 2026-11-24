// The menu bar as data, DOM-free: File, Edit, View, Insert, and Format. After `CHROME_MENUS` of
// docx-editor/packages/core/src/editor/chrome-controls.ts and the entry shapes of
// docx-editor/packages/core/src/editor/chrome-menu-entries.ts at ad686574 (EigenPal): the menu
// bar is an arrangement over the slots of `CHROME_GROUPS`, not a second registry, so a menu row
// and a toolbar control for the same command share one label, one icon, one command, and one
// enabled state (`toolbarCommandState`). The file name is PPTX Editor's
// (pptx-editor/packages/core/src/editor/chrome-menus.ts at 526d3b23), because this editor keeps
// its menu data in a file of its own, as PPTX Editor does.
//
// Differences from DOCX Editor:
// - The menus of a workbook editor. A row exists only for a command that the editor runs: there
//   is no print, page setup, find, or export row, and no View row for gridlines or frozen panes,
//   which the editor reads from the workbook but does not edit. View > Headings writes the
//   headings flag of the sheet view (`setShowHeadings`).
// - A row's shortcut is a key chord (`Mod+Z`), not a catalog key of fixed text, so the bar shows
//   it for the platform (`menuShortcutText`): Command where the Command key is the shortcut
//   modifier, Control elsewhere. A
//   row names only a chord that the editor binds.
// - A row has no `picker`: the editor has no size picker.

import type { ChromeSlotId } from './chrome-controls.ts';
import { ICON_PATHS } from './icon-paths.ts';

/**
 * A row that runs one chrome slot.
 *
 * @public
 */
export interface ChromeMenuItemEntry {
  readonly kind: 'item';
  readonly slot: ChromeSlotId;
  /** A label for this row in place of the slot's own label. */
  readonly labelKey?: string;
  /**
   * The key chord that the editor binds to this row's command, such as `Mod+Z`. `Mod` is
   * Command where the Command key is the shortcut modifier, and Control elsewhere. The bar shows
   * it at the end of the row
   * (`menuShortcutText`).
   */
  readonly shortcut?: string;
}

/**
 * A row that opens a nested panel of rows (Format > Alignment). It has its own label and no
 * slot, because opening a panel is no command.
 *
 * @public
 */
export interface ChromeMenuSubmenuEntry {
  readonly kind: 'submenu';
  readonly labelKey: string;
  readonly paths: readonly string[] | null;
  readonly items: readonly ChromeMenuEntry[];
}

/** A rule between groups of rows. @public */
export interface ChromeMenuSeparatorEntry {
  readonly kind: 'separator';
}

/** One row of a chrome menu. @public */
export type ChromeMenuEntry =
  ChromeMenuItemEntry | ChromeMenuSubmenuEntry | ChromeMenuSeparatorEntry;

/**
 * Every menu id, as a literal union. Public API: renaming one is a breaking change, as for a
 * group or a slot id.
 *
 * @public
 */
export type ChromeMenuId = 'file' | 'edit' | 'view' | 'insert' | 'format';

/** One menu of the menu bar. @public */
export interface ChromeMenu {
  readonly id: ChromeMenuId;
  readonly labelKey: string;
  readonly entries: readonly ChromeMenuEntry[];
}

const item = (slot: ChromeSlotId, shortcut?: string): ChromeMenuItemEntry =>
  shortcut === undefined ? { kind: 'item', slot } : { kind: 'item', slot, shortcut };
const SEPARATOR: ChromeMenuSeparatorEntry = { kind: 'separator' };

/**
 * The menu bar, in bar order: File, Edit, View, Insert, and Format.
 *
 * @public
 */
export const CHROME_MENUS: readonly ChromeMenu[] = [
  {
    id: 'file',
    labelKey: 'menu.file',
    // Open is the host's (`ToolbarHost.open`). Save hands the bytes to the host's `onSave`,
    // through the same save as the toolbar's control and Control or Command and S.
    entries: [item('file.open'), item('file.save', 'Mod+S')],
  },
  {
    id: 'edit',
    labelKey: 'menu.edit',
    entries: [
      item('history.undo', 'Mod+Z'),
      item('history.redo', 'Mod+Y'),
      SEPARATOR,
      item('edit.cut', 'Mod+X'),
      item('edit.copy', 'Mod+C'),
      item('edit.paste', 'Mod+V'),
      SEPARATOR,
      {
        kind: 'submenu',
        labelKey: 'formattingBar.groups.clear',
        paths: ICON_PATHS.delete_sweep,
        items: [item('clear.contents'), item('clear.formats'), item('clear.all')],
      },
      item('cells.deleteRows'),
      item('cells.deleteColumns'),
      SEPARATOR,
      item('edit.selectAll', 'Mod+A'),
    ],
  },
  {
    id: 'view',
    labelKey: 'menu.view',
    // Zoom out shows no chord: on a selection of whole rows or columns, Control or Command
    // and - deletes them (`grid-keys.ts`), so the row would name a key that can delete cells.
    entries: [
      item('view.zoomIn', 'Mod+='),
      item('view.zoomOut'),
      item('view.zoomReset', 'Mod+0'),
      SEPARATOR,
      item('view.headings'),
    ],
  },
  {
    id: 'insert',
    labelKey: 'menu.insert',
    entries: [item('cells.insertRows'), item('cells.insertColumns'), SEPARATOR, item('sheet.add')],
  },
  {
    id: 'format',
    labelKey: 'menu.format',
    entries: [
      item('font.bold', 'Mod+B'),
      item('font.italic', 'Mod+I'),
      item('font.underline', 'Mod+U'),
      SEPARATOR,
      {
        kind: 'submenu',
        labelKey: 'formattingBar.groups.alignment',
        paths: ICON_PATHS.format_align_left,
        items: [
          item('alignment.left'),
          item('alignment.center'),
          item('alignment.right'),
          SEPARATOR,
          item('alignment.top'),
          item('alignment.middle'),
          item('alignment.bottom'),
        ],
      },
      item('alignment.wrap'),
      SEPARATOR,
      item('merge.merge'),
      item('merge.mergeCenter'),
      item('merge.unmerge'),
      SEPARATOR,
      {
        kind: 'submenu',
        labelKey: 'formattingBar.groups.borders',
        paths: ICON_PATHS.border_all,
        items: [item('borders.outline'), item('borders.all'), item('borders.none')],
      },
      SEPARATOR,
      item('number.custom'),
      item('number.increaseDecimals'),
      item('number.decreaseDecimals'),
      SEPARATOR,
      item('clear.formats'),
    ],
  },
];

/** The slots of a list of entries, in menu order, submenus flattened. */
function flatten(entries: readonly ChromeMenuEntry[]): ChromeSlotId[] {
  return entries.flatMap((e) =>
    e.kind === 'item' ? [e.slot] : e.kind === 'submenu' ? flatten(e.items) : []
  );
}

/**
 * Every slot that a menu places, in menu order, submenus flattened, as DOCX Editor's function of
 * the same name.
 *
 * @public
 */
export function chromeMenuSlots(): readonly ChromeSlotId[] {
  return CHROME_MENUS.flatMap((menu) => flatten(menu.entries));
}

/**
 * The chord that `CHROME_MENUS` gives a slot's row, wherever the row stands, or undefined: the
 * shortcut of a host `Menu.Item` or a named row for that slot. @internal
 */
export function chromeMenuShortcut(slot: ChromeSlotId): string | undefined {
  const find = (entries: readonly ChromeMenuEntry[]): string | undefined => {
    for (const e of entries) {
      if (e.kind === 'item' && e.slot === slot && e.shortcut) return e.shortcut;
      if (e.kind === 'submenu') {
        const found = find(e.items);
        if (found) return found;
      }
    }
    return undefined;
  };
  for (const menu of CHROME_MENUS) {
    const found = find(menu.entries);
    if (found) return found;
  }
  return undefined;
}

/** The slots whose row moves the keyboard elsewhere: a file picker or a dialog. */
const KEEPS_FOCUS = new Set<ChromeSlotId>(['file.open', 'number.custom']);

/**
 * Whether the keyboard stays where a slot's row puts it after the row runs. Every other row gives
 * the keyboard back to the grid. @internal
 */
export function menuRowKeepsFocus(slot: ChromeSlotId): boolean {
  return KEEPS_FOCUS.has(slot);
}

// A cell format holds one horizontal and one vertical alignment: the `horizontal` and `vertical`
// attributes of `alignment` (CT_CellAlignment, ECMA-376 Part 1, 18.8.1) each take one value. So
// each set is one-of, not a set of independent switches. After the `selected` row prop of
// docx-editor/packages/react/src/editor/menu/parts.tsx at ad686574 (EigenPal).
const RADIO_SLOTS = new Set<ChromeSlotId>([
  'alignment.left',
  'alignment.center',
  'alignment.right',
  'alignment.top',
  'alignment.middle',
  'alignment.bottom',
]);

/**
 * The ARIA role of a slot's menu row: `menuitemradio` for a row of a one-of set (an alignment),
 * `menuitemcheckbox` for a row that toggles, and `menuitem` for a row that only acts. @internal
 */
export function menuRowRole(
  slot: ChromeSlotId,
  toggles: boolean
): 'menuitem' | 'menuitemcheckbox' | 'menuitemradio' {
  if (RADIO_SLOTS.has(slot)) return 'menuitemradio';
  return toggles ? 'menuitemcheckbox' : 'menuitem';
}

/** The key that a host row matches to replace a registry row: its slot. @internal */
export function chromeMenuRowKey(entry: Pick<ChromeMenuItemEntry, 'slot'>): string {
  return entry.slot;
}

// After pptx-editor/packages/core/src/editor/chrome-forms.ts (`tidySeparators`) at fef01c93,
// with no difference.
/** The entries without a leading, a trailing, or a repeated separator. @internal */
export function tidySeparators(entries: readonly ChromeMenuEntry[]): ChromeMenuEntry[] {
  const out: ChromeMenuEntry[] = [];
  for (const entry of entries) {
    if (
      entry.kind === 'separator' &&
      (out.length === 0 || out[out.length - 1]!.kind === 'separator')
    )
      continue;
    out.push(entry);
  }
  while (out.length && out[out.length - 1]!.kind === 'separator') out.pop();
  return out;
}

/** Whether the platform of a window uses the Command key as its shortcut modifier. @internal */
export function isApplePlatform(
  nav: { readonly platform?: string; readonly userAgent?: string } | undefined
): boolean {
  const text = `${nav?.platform ?? ''} ${nav?.userAgent ?? ''}`;
  return /Mac|iPhone|iPad|iPod/.test(text);
}

/**
 * The words of the modifier keys where the Command key is not the shortcut modifier, from the
 * catalog (`menu.keys.*`), so a locale names its own keys. @internal
 */
export interface MenuKeyWords {
  readonly ctrl: string;
  readonly shift: string;
  readonly alt: string;
}

/** The catalog keys of `MenuKeyWords`. @internal */
export const MENU_KEY_WORD_KEYS = {
  ctrl: 'menu.keys.ctrl',
  shift: 'menu.keys.shift',
  alt: 'menu.keys.alt',
} as const;

/** The words of the modifier keys through a translator. @internal */
export function menuKeyWords(t: (key: string) => string): MenuKeyWords {
  return {
    ctrl: t(MENU_KEY_WORD_KEYS.ctrl),
    shift: t(MENU_KEY_WORD_KEYS.shift),
    alt: t(MENU_KEY_WORD_KEYS.alt),
  };
}

/**
 * The text of a key chord for the platform: `Mod+Z` is `⌘Z` where the Command key is the
 * shortcut modifier, and `Ctrl+Z` elsewhere with the catalog's word for the key (`words`);
 * `Shift` is `⇧` and `Alt` is `⌥` where the Command key is the modifier. @internal
 */
export function menuShortcutText(chord: string, apple: boolean, words: MenuKeyWords): string {
  const parts = chord.split('+');
  // A chord that ends in `+` names the plus key: `Mod++`.
  if (chord.endsWith('++')) parts.splice(-2, 2, '+');
  const key = parts.pop() ?? '';
  const keyText = key.length === 1 ? key.toUpperCase() : key;
  const mods = parts.map((m) => {
    if (m === 'Mod') return apple ? '⌘' : words.ctrl;
    if (m === 'Shift') return apple ? '⇧' : words.shift;
    if (m === 'Alt') return apple ? '⌥' : words.alt;
    return m;
  });
  return apple ? `${mods.join('')}${keyText}` : [...mods, keyText].join('+');
}
