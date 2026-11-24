// The keys of a chrome popup that a pointer press opened. Chrome `mousedown` keeps the keyboard
// where it was (`guardToolbarMousedown`), so after a press on a menu trigger or a toolbar list,
// the grid, the cell editor, or another trigger still holds the focus while the popup is open.
// Without this rule, ArrowDown moves the grid selection behind the open menu, and a typed `14`
// overwrites the active cell under an open font size list.
//
// THE RULE, shared by both adapters: while the popup is open, listen on its ownerDocument in the
// CAPTURE phase. A navigation key that comes from this editor but from outside the popup belongs
// to the popup, as the WAI-ARIA menu and listbox patterns ask: the arrows, Home, End, Enter,
// Space, and a printable character. The first such key moves the focus into the popup, and the
// popup acts on it there (`PopupEntry`). The press itself moves no focus, so a pointer open still
// leaves an open cell edit alone until the user presses one of these keys. Escape stays with
// `listenForPopupEscape`, and Tab, F6, and keys with Control, Command, or Alt pass.
//
// A key from outside this editor (a host input, another editor) passes untouched, and so does a
// key in a modal dialog or during an input method composition.

import { isDomNode, isElementNode } from './host-realm.ts';
import { editorInstanceScope, hasOpenNestedPopup } from './popup-escape.ts';
import { menuBarArrow } from './menu-arrangement.ts';

/**
 * What a navigation key asks of a popup that does not hold the focus. `first` and `last` are
 * the ends of its items (ArrowDown, Enter, Space, and Home ask for the first; ArrowUp and End for
 * the last), `side` is ArrowLeft or ArrowRight as typed, and `char` is a printable character.
 *
 * @internal
 */
export type PopupEntry =
  | { readonly kind: 'first'; readonly key: string }
  | { readonly kind: 'last'; readonly key: string }
  | { readonly kind: 'side'; readonly key: 'ArrowLeft' | 'ArrowRight' }
  | { readonly kind: 'char'; readonly key: string };

/** The part of a key event that {@link popupEntryOf} reads. @internal */
export interface PopupEntryKey {
  readonly key: string;
  readonly altKey: boolean;
  readonly ctrlKey: boolean;
  readonly metaKey: boolean;
}

/** The entry that a key asks for, or null when the key is not one for the popup. @internal */
export function popupEntryOf(e: PopupEntryKey): PopupEntry | null {
  if (e.altKey || e.ctrlKey || e.metaKey) return null;
  const { key } = e;
  if (key === 'ArrowDown' || key === 'Enter' || key === ' ' || key === 'Home')
    return { kind: 'first', key };
  if (key === 'ArrowUp' || key === 'End') return { kind: 'last', key };
  if (key === 'ArrowLeft' || key === 'ArrowRight') return { kind: 'side', key };
  if ([...key].length === 1) return { kind: 'char', key };
  return null;
}

/**
 * What the popup did with an entry: `done` (the key is spent), `type` (a field of the popup took
 * the focus, and the character goes on into it), or null (the key is not the popup's).
 *
 * @internal
 */
export type PopupEntryResult = 'done' | 'type' | null;

/** How {@link listenForPointerPopupKeys} decides and acts. @internal */
export interface PointerPopupKeysOptions {
  /** The popup's root. Its `ownerDocument` receives the listener. */
  readonly popup: HTMLElement;
  /** Whether a node belongs to the popup: its keys go to its own handlers. */
  readonly contains: (node: Node) => boolean;
  /** The chrome root that holds the popup: a key from it counts as this editor's. */
  readonly chromeRoot?: () => Element | null | undefined;
  /** Move the focus into the popup and act on the key there. */
  readonly enter: (entry: PopupEntry, event: KeyboardEvent) => PopupEntryResult;
}

/** A modal dialog of the page or the editor owns its own keys. */
const MODAL = '[aria-modal="true"], [data-xlsx-modal], dialog[open]';

/**
 * Give the navigation keys from outside an open popup to the popup. Returns the disposer.
 *
 * Shared adapter glue, not for hosts.
 *
 * @internal
 */
export function listenForPointerPopupKeys(options: PointerPopupKeysOptions): () => void {
  const owner = options.popup.ownerDocument;
  const onKeyDown = (event: KeyboardEvent): void => {
    if (event.defaultPrevented || event.isComposing || event.keyCode === 229) return;
    const entry = popupEntryOf(event);
    if (!entry) return;
    const path = event.composedPath();
    const origin = path[0] ?? event.target;
    if (!isDomNode(origin) || options.contains(origin)) return;
    const element = isElementNode(origin) ? origin : origin.parentElement;
    if (element?.closest(MODAL)) return;
    const owners = [editorInstanceScope(options.popup), options.chromeRoot?.()];
    if (!owners.some((el) => el != null && path.includes(el))) return;
    const result = options.enter(entry, event);
    if (!result) return;
    event.stopPropagation();
    if (result === 'done') event.preventDefault();
  };
  owner.addEventListener('keydown', onKeyDown, true);
  return () => owner.removeEventListener('keydown', onKeyDown, true);
}

/**
 * The index of the item that a `char` entry moves to: the first item whose label starts with
 * `char` (any case), or -1. @internal
 */
export function itemStartingWith(labels: readonly string[], char: string): number {
  const c = char.toLocaleLowerCase();
  return labels.findIndex((label) => label.trim().toLocaleLowerCase().startsWith(c));
}

/**
 * Act on an entry in a toolbar popup that does not hold the focus: a list, a palette, or a menu
 * of rows. `items` are its focusable items in order, and `field` its text field, if any. A typed
 * character goes into the field, else moves to the first item that starts with it. Any other
 * entry focuses the item that `focusIn` picks (the selected item, else the first), or the last
 * item for `last`. @internal
 */
export function enterToolbarPopup(
  entry: PopupEntry,
  items: readonly HTMLElement[],
  field: HTMLElement | null,
  focusIn: () => void
): PopupEntryResult {
  if (entry.kind === 'char' && field) {
    field.focus();
    return 'type';
  }
  if (items.length === 0) return null;
  const at =
    entry.kind === 'char'
      ? itemStartingWith(
          items.map((el) => el.textContent ?? ''),
          entry.key
        )
      : -1;
  if (at >= 0) items[at]!.focus();
  else if (entry.kind === 'last') items.at(-1)!.focus();
  else focusIn();
  return 'done';
}

/** The rows of a menu panel: the menu items, checkbox items, and submenu triggers. @internal */
export const MENU_PANEL_ROWS =
  ':scope > [role^="menuitem"], :scope > [role="none"] > [role^="menuitem"], ' +
  ':scope > [role="group"] > [role^="menuitem"], ' +
  ':scope > [role="group"] > [role="none"] > [role^="menuitem"]';

/** The rows that a menu panel moves through, in order. @internal */
export function menuPanelRows(panel: HTMLElement): HTMLElement[] {
  return [...panel.querySelectorAll<HTMLElement>(MENU_PANEL_ROWS)];
}

/** The labels of menu rows, for type-ahead. @internal */
export function menuRowLabels(rows: readonly HTMLElement[]): string[] {
  return rows.map((r) => r.querySelector('.xlsx-menu-label')?.textContent ?? '');
}

/** The panel of a bar menu, beside its trigger, while it shows. @internal */
export function menuPanelOf(trigger: HTMLElement | null): HTMLElement | null {
  return trigger?.parentElement?.querySelector<HTMLElement>(':scope > [role="menu"]') ?? null;
}

/** Focus the first or the last row of a menu panel. @internal */
export function focusMenuPanelEdge(panel: HTMLElement | null, edge: 'first' | 'last'): void {
  if (!panel) return;
  const rows = menuPanelRows(panel);
  (edge === 'first' ? rows[0] : rows.at(-1))?.focus();
}

/**
 * Act on an entry in a menu panel of the bar that does not hold the focus. A typed character
 * moves to the first row that starts with it (else the first row), `side` moves to the previous
 * or the next menu of the bar (reversed right to left), and the ends focus the first or the last
 * row. @internal
 */
export function enterMenuPanel(
  entry: PopupEntry,
  panel: HTMLElement,
  options: { readonly rtl: boolean; readonly step: (by: -1 | 1) => void }
): PopupEntryResult {
  if (entry.kind === 'side') {
    const by = menuBarArrow(entry.key, options.rtl);
    if (by !== 0) options.step(by);
    return 'done';
  }
  const rows = menuPanelRows(panel);
  const at = entry.kind === 'char' ? itemStartingWith(menuRowLabels(rows), entry.key) : -1;
  const row = entry.kind === 'last' ? rows.at(-1) : (rows[at] ?? rows[0]);
  row?.focus();
  return 'done';
}

/**
 * A press on a menu bar trigger keeps the keyboard where it was. When that is another trigger of
 * the same bar, move it to `trigger`, so the next key acts on the menu that the press opened and
 * not on the menu of the trigger that still holds the keyboard. @internal
 */
export function followPressedTrigger(trigger: HTMLElement): void {
  const menubar = trigger.closest('[role="menubar"]');
  if (menubar?.contains(trigger.ownerDocument.activeElement)) trigger.focus();
}

/** The text field of a toolbar popup that a typed character goes into. */
const POPUP_FIELD = 'input[type="text"]';

/**
 * Give the navigation keys from outside an open toolbar popup to it (`listenForPointerPopupKeys`).
 * `root` holds the trigger and the popup (`[data-popup]`), and `selector` names its items. With
 * `nested`, a popup open inside this one takes the keys instead. Returns the disposer. @internal
 */
export function listenForToolbarPopupKeys(
  root: HTMLElement,
  selector: string,
  focusIn: () => void,
  nested = false
): () => void {
  return listenForPointerPopupKeys({
    popup: root,
    contains: (node) => root.contains(node),
    chromeRoot: () => root.closest('.xlsx-toolbar'),
    enter: (entry) => {
      const popup = root.querySelector<HTMLElement>('[data-popup]');
      if (!popup || (nested && hasOpenNestedPopup(popup))) return null;
      const items = [...popup.querySelectorAll<HTMLElement>(selector)];
      return enterToolbarPopup(
        entry,
        items,
        popup.querySelector<HTMLElement>(POPUP_FIELD),
        focusIn
      );
    },
  });
}
