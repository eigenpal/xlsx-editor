// The rules of the sheet tabs that do not depend on a framework: the key of the sheet list and
// of each tab, the arrow keys of the tab menu, where the rename field lies over its tab, how the
// menu closes, and where focus goes when it closes. Moved from the React adapter's
// `SheetTabs.tsx`, which describes the behavior and its PPTX Editor sources, so the React and
// Vue sheet tabs act the same. Adapter support, not a stable API.

import type { EditorSheet } from './editor-model.ts';
import { windowScale } from './host-frame.ts';
import type { XlsxEditorInstance } from './xlsx-editor.ts';

/**
 * A key of the sheet list: it changes when a sheet is added, removed, moved, or renamed.
 * @internal
 */
export const sheetListKey = (
  sheets: readonly { readonly name: string; readonly visible: boolean }[]
): string => sheets.map((s) => `${s.visible ? 1 : 0}${s.name}`).join('\u0000');

/** The items of the tab menu. @internal */
export const SHEET_MENU_ITEM = '[role="menuitem"]';

/**
 * The key of a tab: its name, and how many tabs before it have the same name, so a tab keeps its
 * element when a sheet before it is added, deleted, or moved. A file can give two sheets the same
 * name: the count keeps their keys apart. @internal
 */
export function tabKey(tabs: readonly { readonly name: string }[], at: number): string {
  let same = 0;
  for (let i = 0; i < at; i += 1) if (tabs[i]!.name === tabs[at]!.name) same += 1;
  return `${same}:${tabs[at]!.name}`;
}

/** Move focus through the menu items for an arrow key, Home, or End. True when it moved. @internal */
export function moveInSheetMenu(menu: HTMLElement, key: string): boolean {
  const items = [...menu.querySelectorAll<HTMLElement>(SHEET_MENU_ITEM)];
  if (items.length === 0) return false;
  const at = items.indexOf(menu.ownerDocument.activeElement as HTMLElement);
  const next =
    key === 'ArrowDown'
      ? (at + 1) % items.length
      : key === 'ArrowUp'
        ? (at - 1 + items.length) % items.length
        : key === 'Home'
          ? 0
          : key === 'End'
            ? items.length - 1
            : -1;
  if (next < 0) return false;
  items[next]!.focus();
  return true;
}

/**
 * An open rename: the sheet name when it started, the box of its tab in the bar, and the room
 * before the add button. @internal
 */
export interface SheetRenameBox {
  readonly name: string;
  readonly left: number;
  readonly width: number;
  /** Undefined when the bar has no add button to measure. */
  readonly maxWidth: number | undefined;
}

/** The elements of the tab bar that place the rename field. @internal */
export interface SheetBarParts {
  readonly bar: HTMLElement | null;
  readonly list: HTMLElement | null;
  readonly add: HTMLElement | null;
}

/**
 * The rename box over the tab of the sheet named `name`, as the bar shows it now, or null when
 * the tab is gone. The field stays over the tab list, and ends before the add button, so the
 * button stays in reach. @internal
 */
export function sheetRenameBox(
  parts: SheetBarParts,
  sheets: readonly EditorSheet[],
  name: string
): SheetRenameBox | null {
  const sheet = sheets.find((s) => s.name === name);
  const tab =
    sheet && parts.list?.querySelector<HTMLElement>(`[data-sheet-index="${sheet.index}"]`);
  if (!tab) return null;
  const box = parts.bar?.getBoundingClientRect();
  // The boxes are in window pixels; the field's place is in the bar's CSS pixels, which differ
  // inside a host that scales the editor.
  const scale = parts.bar ? windowScale(parts.bar).x : 1;
  const tabBox = tab.getBoundingClientRect();
  const tabWidth = tabBox.width / scale;
  // The field stays over the tab list: a tab that scrolled out puts it at the list's edge.
  const measured = parts.list?.getBoundingClientRect();
  const listBox = measured && measured.width > 0 ? measured : undefined;
  const origin = box?.left ?? 0;
  const lo = listBox ? (listBox.left - origin) / scale : 0;
  const hi = listBox ? Math.max(lo, (listBox.right - origin) / scale - tabWidth) : Infinity;
  const left = Math.max(0, lo, Math.min((tabBox.left - origin) / scale, hi));
  // The field ends before the add button, so the button stays in reach.
  const add = parts.add?.getBoundingClientRect();
  const room = add ? Math.max(tabWidth, (add.left - origin) / scale - left - 2) : undefined;
  return { name, left, width: tabWidth, maxWidth: room };
}

/**
 * Where the tab menu opens in the bar `bar`, in its CSS pixels from the bar's left edge: at the
 * pointer (`clientX`), or at the tab's left edge when a key opens it. Inside a host that scales
 * the editor, window pixels are converted to the bar's pixels. @internal
 */
export function sheetMenuLeft(bar: HTMLElement | null, tab: HTMLElement, clientX?: number): number {
  const box = bar?.getBoundingClientRect();
  const x = clientX ?? tab.getBoundingClientRect().left;
  const scale = bar ? windowScale(bar).x : 1;
  return Math.max(0, (x - (box?.left ?? 0)) / scale);
}

/**
 * Close the tab menu on a press outside it, a window blur, a resize, a scroll of the page, and a
 * wheel outside it (PPTX Editor `listenContextMenuDismiss` and `listenAway`). A scroll that code
 * makes inside the page, such as the grid that restores its scroll position or the tab list
 * that shows the chosen tab, does not close it. The grid cancels `pointerdown`, so the press is
 * read in the capture phase, on the document that owns the bar. Returns the stop function.
 * @internal
 */
export function listenSheetMenuDismiss(
  bar: HTMLElement,
  inMenu: (node: Node) => boolean,
  close: () => void
): () => void {
  const doc = bar.ownerDocument;
  const view = doc.defaultView;
  const outside = (e: Event) => {
    if (!inMenu(e.target as Node)) close();
  };
  const onScroll = (e: Event) => {
    if (e.target === doc || e.target === doc.documentElement) close();
  };
  doc.addEventListener('pointerdown', outside, true);
  doc.addEventListener('wheel', outside, { capture: true, passive: true });
  doc.addEventListener('scroll', onScroll, true);
  view?.addEventListener('resize', close);
  view?.addEventListener('blur', close);
  return () => {
    doc.removeEventListener('pointerdown', outside, true);
    doc.removeEventListener('wheel', outside, { capture: true });
    doc.removeEventListener('scroll', onScroll, true);
    view?.removeEventListener('resize', close);
    view?.removeEventListener('blur', close);
  };
}

/** An open tab menu, as the focus rules read it. @internal */
export interface SheetMenuOpening {
  readonly sheet: number;
  /** The element that had focus when the menu opened: focus goes back to it on close. */
  readonly back: HTMLElement | null;
  /** Whether a key opened the menu. */
  readonly keyboard: boolean;
}

/**
 * Open the menu `el` inside the bar at `left`, and focus its first item. Returns the function
 * that the close runs: when focus is still in the menu, it goes back to the element that had it
 * at the open; when that element is gone, or was the page, a key open gives focus to the tab,
 * and a press to the grid. @internal
 */
export function openSheetMenu(
  el: HTMLElement,
  bar: HTMLElement | null,
  left: number,
  opening: SheetMenuOpening,
  editor: Pick<XlsxEditorInstance, 'focus'> | null
): () => void {
  const room = (bar?.clientWidth ?? 0) - el.offsetWidth;
  el.style.left = `${Math.max(0, Math.min(left, room))}px`;
  el.querySelector<HTMLElement>(SHEET_MENU_ITEM)?.focus();
  const back = opening.back;
  return () => {
    const current = el.ownerDocument.activeElement;
    const lost =
      !current || current === el.ownerDocument.body || !current.isConnected || el.contains(current);
    if (!lost) return;
    const page = back === el.ownerDocument.body;
    if (back?.isConnected && !page) back.focus({ preventScroll: true });
    // The element that had focus is gone, such as a rename field that the open ended, or it
    // was the page: a key open gives focus to the tab, and a press to the grid.
    else if (opening.keyboard)
      bar
        ?.querySelector<HTMLElement>(`[data-sheet-index="${opening.sheet}"]`)
        ?.focus({ preventScroll: true });
    else editor?.focus();
  };
}
