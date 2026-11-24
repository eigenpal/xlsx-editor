// The framework-free rules of the menu bar, shared by the React and Vue adapters (`menu/`), so
// the two cannot drift: which registry rows a host row replaces or hides, which menus the bar
// shows in which order, how many menus fit before the "More" menu, and what each key does on a
// trigger and in a panel. The grid context menu reuses the row rules (`arrangeMenuRows`).
// Adapter support, not a stable API, except `MENU_NAMED_PARTS`, whose names are the adapters'
// named row parts.
//
// After pptx-editor/packages/core/src/editor/menu-arrangement.ts at 526d3b23 (EigenPal):
// `MENU_NAMED_PARTS`, `arrangeMenuRows`, `arrangeMenuBar`, `menuBarArrow`, and
// `menuBarTriggerKey`. PPTX Editor keeps this arrangement in core, where DOCX Editor keeps it in
// each adapter (docx-editor/packages/react/src/editor/merge-arrangement.tsx at ad686574); this
// editor takes the core place, as it arranges its toolbar in core (`toolbar-arrange.ts`), so the
// React and Vue menu bars arrange the same way by construction. Differences from PPTX Editor:
// - A registry row has no value, so a named row is a slot only, and a row key is its slot.
// - `arrangeMenuRows` also applies the host rows inside a submenu, and leaves out a submenu
//   whose rows are all hidden.
// - `arrangeMenuBar` takes a translate function of this editor and reports each problem through
//   `onProblem`, which the adapter prints in development, as the toolbar does.
// - `menuPanelKey`, `menuTypeahead`, and `menusThatFit` are new here: the panel keys of the
//   WAI-ARIA menu pattern, its type-ahead, and the bar's overflow count, after DOCX Editor's menu
//   panel keys (docx-editor/packages/react/src/editor/menu/parts.tsx at ad686574) and this
//   editor's toolbar overflow (`groupsThatFit`).
// - The development problems of a composition that does nothing (`menuHiddenRowProblem`,
//   `menuRowInBarProblem`, and the reserved id of the "More" menu) are core sentences, after
//   docx-editor/packages/react/src/editor/menu/menu-warnings.ts at ad686574 (EigenPal), so the
//   two adapters print the same words with their own part names (`MenuPartNames`).

import type { ChromeSlotId } from './chrome-controls.ts';
import {
  CHROME_MENUS,
  chromeMenuRowKey,
  tidySeparators,
  type ChromeMenu,
  type ChromeMenuEntry,
} from './chrome-menus.ts';
import { groupsThatFit } from './toolbar-layout.ts';
import { windowScale } from './host-frame.ts';

// ---------------------------------------------------------------------------------------------
// Host rows over registry rows

/** The registry row that a named menu row part stands for. @public */
export interface MenuNamedRow {
  readonly slot: ChromeSlotId;
}

const row = (slot: ChromeSlotId): MenuNamedRow => ({ slot });

/**
 * The named rows of the menu bar (`XlsxEditor.Menu.Open`, `XlsxEditor.Menu.Cut`, ...) and the
 * row each one stands for, after DOCX Editor's and PPTX Editor's menu row parts. A named row is
 * shorthand for `XlsxEditor.Menu.Item` with its `slotId`, with the same props. Put one in the
 * menu that holds its row, such as `XlsxEditor.Menu.Edit`. The names follow PPTX Editor's where
 * the row is the same (`Undo`, `Cut`, `SelectAll`, `Bold`, `ClearFormatting`).
 *
 * @public
 */
export const MENU_NAMED_PARTS = {
  Open: row('file.open'),
  Save: row('file.save'),
  Undo: row('history.undo'),
  Redo: row('history.redo'),
  Cut: row('edit.cut'),
  Copy: row('edit.copy'),
  Paste: row('edit.paste'),
  SelectAll: row('edit.selectAll'),
  ClearContents: row('clear.contents'),
  ClearFormatting: row('clear.formats'),
  ClearAll: row('clear.all'),
  DeleteRows: row('cells.deleteRows'),
  DeleteColumns: row('cells.deleteColumns'),
  ZoomIn: row('view.zoomIn'),
  ZoomOut: row('view.zoomOut'),
  ZoomReset: row('view.zoomReset'),
  InsertRows: row('cells.insertRows'),
  InsertColumns: row('cells.insertColumns'),
  InsertSheet: row('sheet.add'),
  Bold: row('font.bold'),
  Italic: row('font.italic'),
  Underline: row('font.underline'),
  AlignLeft: row('alignment.left'),
  AlignCenter: row('alignment.center'),
  AlignRight: row('alignment.right'),
  AlignTop: row('alignment.top'),
  AlignMiddle: row('alignment.middle'),
  AlignBottom: row('alignment.bottom'),
  WrapText: row('alignment.wrap'),
  MergeCells: row('merge.merge'),
  MergeAndCenter: row('merge.mergeCenter'),
  UnmergeCells: row('merge.unmerge'),
  BordersOutline: row('borders.outline'),
  BordersAll: row('borders.all'),
  BordersNone: row('borders.none'),
  NumberFormat: row('number.custom'),
  IncreaseDecimals: row('number.increaseDecimals'),
  DecreaseDecimals: row('number.decreaseDecimals'),
} as const satisfies Record<string, MenuNamedRow>;

/** The name of a named menu row part. @public */
export type MenuRowPartName = keyof typeof MENU_NAMED_PARTS;

/** A host row that names a registry slot (`Menu.Item`, a named row). @internal */
export interface MenuSlotRowSpec {
  readonly slotId: string;
  readonly hidden?: boolean | undefined;
}

/** The rows of a menu after the host's children. @internal */
export interface ArrangedMenuRows<C> {
  /**
   * The registry rows still shown, with the separators that a hidden row left behind removed,
   * also inside a submenu. A submenu whose rows are all hidden is left out.
   */
  readonly entries: ChromeMenuEntry[];
  /** Host rows that replace registry rows, by row key; `null` removes the row. */
  readonly overrides: Map<string, C | null>;
  /** Host rows that replace no registry row: they follow the registry rows. */
  readonly extras: C[];
}

/**
 * Split a menu's host children against its registry rows: a slot row (`slotOf` answers for it)
 * whose slot matches a registry row replaces that row in place, or removes it when `hidden`;
 * other children are extra rows. The last host row for a slot wins. A hidden slot row that
 * matches nothing calls `onUnmatchedHidden` and adds no row. @internal
 */
export function arrangeMenuRows<C>(
  entries: readonly ChromeMenuEntry[],
  children: readonly C[],
  slotOf: (child: C) => MenuSlotRowSpec | null,
  onUnmatchedHidden?: (slotId: string) => void
): ArrangedMenuRows<C> {
  const keys = new Set<string>();
  const collect = (list: readonly ChromeMenuEntry[]): void => {
    for (const e of list)
      if (e.kind === 'item') keys.add(chromeMenuRowKey(e));
      else if (e.kind === 'submenu') collect(e.items);
  };
  collect(entries);
  const overrides = new Map<string, C | null>();
  const extras: C[] = [];
  for (const child of children) {
    const spec = slotOf(child);
    if (spec) {
      const key = chromeMenuRowKey({ slot: spec.slotId as ChromeSlotId });
      if (keys.has(key)) {
        overrides.set(key, spec.hidden ? null : child);
        continue;
      }
      if (spec.hidden) {
        onUnmatchedHidden?.(spec.slotId);
        continue;
      }
    }
    extras.push(child);
  }
  // A hidden row takes its separators with it when they would stand alone.
  const shown = (list: readonly ChromeMenuEntry[]): ChromeMenuEntry[] =>
    tidySeparators(
      list.flatMap((e): ChromeMenuEntry[] => {
        if (e.kind === 'item') return overrides.get(chromeMenuRowKey(e)) === null ? [] : [e];
        if (e.kind !== 'submenu') return [e];
        const items = shown(e.items);
        return items.length === 0 ? [] : [{ ...e, items }];
      })
    );
  return { entries: shown(entries), overrides, extras };
}

/**
 * How the problems name the parts of an adapter: `XlsxEditor.Menu` in React and `XlsxEditorMenu`
 * in Vue. @internal
 */
export interface MenuPartNames {
  readonly menu: string;
}

/** The part names of the React adapter. @internal */
export const REACT_MENU_PART_NAMES: MenuPartNames = { menu: 'XlsxEditor.Menu' };

/** The part names of the Vue adapter. @internal */
export const VUE_MENU_PART_NAMES: MenuPartNames = { menu: 'XlsxEditorMenu' };

/** The problem of a hidden slot row whose slot matches no row of its menu. @internal */
export function menuHiddenRowProblem(slotId: string, menu: string, names: MenuPartNames): string {
  return (
    `A hidden ${names.menu}.Item names "${slotId}", but the ${menu} menu has no row of that ` +
    `slot, so nothing is removed. Put the hidden row in the menu that shows the slot.`
  );
}

/** The problem of a slot row that is a direct child of the bar, not of a menu. @internal */
export function menuRowInBarProblem(slotId: string, names: MenuPartNames): string {
  return (
    `The ${names.menu}.Item "${slotId}" is a direct child of ${names.menu}, not of a menu, so ` +
    `it is left out. Put it in a menu part, such as ${names.menu}.Edit.`
  );
}

// ---------------------------------------------------------------------------------------------
// The menu bar

/** The id of the bar's "More" menu, which a host menu cannot take. @internal */
export const MENU_MORE_ID = 'xlsx-menu-more';

/** What the bar reads from a named menu or a `Menu.Menu` child. @internal */
export interface MenuBarSpec {
  readonly id: string;
  readonly label?: string | undefined;
  readonly labelKey?: string | undefined;
  readonly hidden?: boolean | undefined;
  readonly after?: string | undefined;
}

/** One menu of the bar, before the adapter renders its rows. @internal */
export interface ArrangedBarMenu<S extends MenuBarSpec> {
  readonly id: string;
  readonly title: string;
  /** The host's description of the menu, if any. */
  readonly spec: S | undefined;
  /** The registry menu, for a built-in menu: its rows come from `arrangeMenuRows`. */
  readonly registry: ChromeMenu | undefined;
}

/**
 * The menus of the bar, in order: with `preset`, the registry menus (minus hidden ones), changed
 * in place by the host's specs; then the host's other menus, each after its `after` menu or at
 * the end. With `preset` off, only the host's menus show, and a named built-in menu still brings
 * its title and rows. An `after` that names no shown menu is a problem, and the menu goes at the
 * end. @internal
 */
export function arrangeMenuBar<S extends MenuBarSpec>(
  specs: ReadonlyMap<string, S>,
  preset: boolean,
  t: (key: string) => string,
  onProblem?: (message: string) => void
): ArrangedBarMenu<S>[] {
  const titleOf = (spec: S | undefined, fallback: string): string =>
    spec?.label ?? (spec?.labelKey ? t(spec.labelKey) : fallback);
  const menus: ArrangedBarMenu<S>[] = [];
  if (preset)
    for (const menu of CHROME_MENUS) {
      const spec = specs.get(menu.id);
      if (spec?.hidden) continue;
      menus.push({ id: menu.id, title: titleOf(spec, t(menu.labelKey)), spec, registry: menu });
    }
  const builtIn = new Set<string>(CHROME_MENUS.map((m) => m.id));
  for (const spec of specs.values()) {
    if ((preset && builtIn.has(spec.id)) || spec.hidden) continue;
    if (spec.id === MENU_MORE_ID) {
      onProblem?.(
        `A host menu has the id "${MENU_MORE_ID}", which the bar keeps for its More menu, so ` +
          `the menu is left out. Give it another id.`
      );
      continue;
    }
    const registry = CHROME_MENUS.find((m) => m.id === spec.id);
    const entry: ArrangedBarMenu<S> = {
      id: spec.id,
      title: titleOf(spec, registry ? t(registry.labelKey) : spec.id),
      spec,
      registry,
    };
    const at = spec.after === undefined ? -1 : menus.findIndex((m) => m.id === spec.after);
    if (spec.after !== undefined && at === -1)
      onProblem?.(
        `The menu "${spec.id}" has after="${spec.after}", but no menu of the bar has that id. ` +
          `The menu is placed after every other menu.`
      );
    if (at === -1) menus.push(entry);
    else menus.splice(at + 1, 0, entry);
  }
  return menus;
}

/** The part name of a built-in menu, for problems (`Menu.File`). @internal */
export function menuPartName(id: string): string {
  return `Menu.${id[0]?.toUpperCase() ?? ''}${id.slice(1)}`;
}

/**
 * How many menus fit in a bar `width` pixels wide, from the start, with `gap` pixels between two
 * triggers: all of them, or as many as fit beside the "More" menu (`groupsThatFit`, the rule of
 * the toolbar). At least one menu stays in the bar. @internal
 */
export function menusThatFit(widths: readonly number[], width: number, gap = 0): number {
  if (widths.length === 0) return 0;
  return Math.max(1, groupsThatFit(widths, width, gap));
}

/**
 * The attribute of the block that holds the host's loose children of the bar. It never moves
 * into the "More" menu, so the bar charges its width as fixed. The block stands beside the
 * `menubar` element, not in it: its children are no rows of the menubar. @internal
 */
export const MENU_FIXED_ATTRIBUTE = 'data-menu-fixed';

/** The menus of the bar row `el`: the menus of its `menubar` element. */
const BAR_MENUS = ':scope > [role="menubar"] > [data-menu]';

/**
 * Measure the menus that the bar row `el` shows (the menus of its `menubar` element) into
 * `widths`, and return how many menus of `ids` fit beside the fixed block. While a menu has no
 * measured width, every menu shows, so the next render measures it. @internal
 */
export function measureMenuBar(
  el: HTMLElement,
  ids: readonly string[],
  widths: Map<string, number>
): number {
  const scale = windowScale(el).x;
  for (const m of el.querySelectorAll<HTMLElement>(BAR_MENUS))
    widths.set(m.dataset.menu ?? '', m.getBoundingClientRect().width / scale);
  if (ids.some((id) => !widths.has(id))) return ids.length;
  const style = el.ownerDocument.defaultView?.getComputedStyle(el);
  const px = (v: string | undefined) => {
    const n = parseFloat(v ?? '');
    return Number.isFinite(n) ? n : 0;
  };
  const gap = px(style?.columnGap);
  let width = el.clientWidth - px(style?.paddingLeft) - px(style?.paddingRight);
  const fixed = el.querySelector(`:scope > [${MENU_FIXED_ATTRIBUTE}]`);
  if (fixed && width > 0)
    width = Math.max(1, width - fixed.getBoundingClientRect().width / scale - gap);
  return menusThatFit(
    ids.map((id) => widths.get(id) ?? 0),
    width,
    gap
  );
}

/**
 * Watch the bar row `el`, each menu it shows, and its fixed block for a change of size. Returns
 * the stop. @internal
 */
export function observeMenuBar(el: HTMLElement, measure: () => void): () => void {
  const Observer = el.ownerDocument.defaultView?.ResizeObserver;
  const observer = Observer ? new Observer(measure) : null;
  observer?.observe(el);
  for (const m of el.querySelectorAll(`${BAR_MENUS}, :scope > [${MENU_FIXED_ATTRIBUTE}]`))
    observer?.observe(m);
  return () => observer?.disconnect();
}

// ---------------------------------------------------------------------------------------------
// Keys

/** The step ArrowLeft or ArrowRight takes along the bar (reversed right to left), or 0. @internal */
export function menuBarArrow(key: string, rtl: boolean): -1 | 0 | 1 {
  const by = key === 'ArrowRight' ? 1 : key === 'ArrowLeft' ? -1 : 0;
  return (rtl ? -by : by) as -1 | 0 | 1;
}

/** What a key on a bar trigger does. @internal */
export type MenuBarTriggerKey =
  | { readonly kind: 'step'; readonly by: -1 | 1 }
  | { readonly kind: 'open'; readonly edge: 'first' | 'last' }
  | { readonly kind: 'edge'; readonly edge: 'first' | 'last' }
  | null;

/**
 * A key on a menu bar trigger, as the WAI-ARIA menubar pattern: the arrows along the bar step,
 * ArrowDown, Enter, and Space open the menu on its first row, ArrowUp opens it on its last row,
 * and Home and End go to the ends of the bar. @internal
 */
export function menuBarTriggerKey(key: string, rtl: boolean): MenuBarTriggerKey {
  const by = menuBarArrow(key, rtl);
  if (by) return { kind: 'step', by };
  if (key === 'ArrowDown' || key === 'Enter' || key === ' ') return { kind: 'open', edge: 'first' };
  if (key === 'ArrowUp') return { kind: 'open', edge: 'last' };
  if (key === 'Home' || key === 'End')
    return { kind: 'edge', edge: key === 'Home' ? 'first' : 'last' };
  return null;
}

/** What a key in an open menu panel does. @internal */
export type MenuPanelKey =
  | { readonly kind: 'step'; readonly by: -1 | 1 }
  | { readonly kind: 'edge'; readonly edge: 'first' | 'last' }
  /** Open the submenu of the focused row. */
  | { readonly kind: 'openSubmenu' }
  /** Close this panel: a submenu goes back to its row, a bar menu to its trigger. */
  | { readonly kind: 'close' }
  /** Move to the next or the previous menu of the bar (from a top panel only). */
  | { readonly kind: 'bar'; readonly by: -1 | 1 }
  /** Leave the bar: Tab closes every panel. */
  | { readonly kind: 'leave' }
  | { readonly kind: 'typeahead'; readonly char: string }
  | null;

/**
 * A key in a menu panel, as the WAI-ARIA menu pattern: ArrowDown and ArrowUp step through the
 * rows and wrap, Home and End go to the ends, Escape closes the panel, Tab leaves the bar, a
 * printable character moves to the next row whose label starts with it. ArrowRight (ArrowLeft
 * right to left) opens the submenu of a submenu row, else moves to the next menu of the bar;
 * ArrowLeft closes a submenu, else moves to the previous menu of the bar. Enter and Space stay
 * with the row's own button. @internal
 */
export function menuPanelKey(
  e: {
    readonly key: string;
    readonly altKey?: boolean;
    readonly ctrlKey?: boolean;
    readonly metaKey?: boolean;
  },
  options: { readonly rtl: boolean; readonly submenu: boolean; readonly onSubmenuRow: boolean }
): MenuPanelKey {
  const { key } = e;
  if (key === 'ArrowDown') return { kind: 'step', by: 1 };
  if (key === 'ArrowUp') return { kind: 'step', by: -1 };
  if (key === 'Home' || key === 'End')
    return { kind: 'edge', edge: key === 'Home' ? 'first' : 'last' };
  if (key === 'Escape') return { kind: 'close' };
  if (key === 'Tab') return { kind: 'leave' };
  const by = menuBarArrow(key, options.rtl);
  if (by === 1) return options.onSubmenuRow ? { kind: 'openSubmenu' } : { kind: 'bar', by: 1 };
  if (by === -1) return options.submenu ? { kind: 'close' } : { kind: 'bar', by: -1 };
  if (e.altKey || e.ctrlKey || e.metaKey) return null;
  if ([...key].length === 1 && key !== ' ') return { kind: 'typeahead', char: key };
  return null;
}

/**
 * The index of the row that type-ahead moves to: the first row after `from` whose label starts
 * with `char` (any case), wrapping, or -1 when no row does. @internal
 */
export function menuTypeahead(labels: readonly string[], from: number, char: string): number {
  const c = char.toLocaleLowerCase();
  for (let i = 1; i <= labels.length; i += 1) {
    const at = (from + i + labels.length) % labels.length;
    if (labels[at]!.trim().toLocaleLowerCase().startsWith(c)) return at;
  }
  return -1;
}

/** The index `by` rows from `from` in a list of `count` rows, wrapping. @internal */
export function menuStep(count: number, from: number, by: -1 | 1): number {
  if (count === 0) return -1;
  if (from < 0) return by === 1 ? 0 : count - 1;
  return (from + by + count) % count;
}
