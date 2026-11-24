// The state of a toolbar popup and the keys of its trigger: the Vue twin of the React adapter's
// `toolbar/popup.ts`, whose header explains the rules. Difference from React: `usePopup`
// returns refs.

import { shallowRef, useId, watch, type Ref } from 'vue';
import {
  focusPopupItem,
  listenForToolbarPopupKeys,
  placePopup,
  popupTriggerAction,
} from '@xlsx-editor.dev/core/editor';
import { useDropdownClose } from './useDropdownClose.ts';

/** How a popup behaves. */
export interface PopupOptions {
  /** The popup renders in flow, inside the overflow panel: no placement. */
  readonly inline?: boolean;
  /** Close the popup: its slot is disabled. */
  readonly closed?: () => boolean;
  /** A popup can open inside this one (the overflow panel). */
  readonly nested?: boolean;
}

/** The state of one popup, as `usePopup` returns it. */
export interface Popup {
  readonly open: Ref<boolean>;
  setOpen(open: boolean): void;
  readonly root: Ref<HTMLDivElement | null>;
  readonly id: string;
  /** Open from the keyboard: focus the selected item, or the first. */
  openWithKeyboard(): void;
  /**
   * Focus the selected item, or else the current item (`aria-current`: an entry that names the
   * value but is not picked), or the first.
   */
  focusIn(): void;
}

/**
 * The state of one popup: open, its root, and how the keyboard moves focus in. `selector`
 * names the items that take focus: the selected item first, else the current item
 * (`aria-current`), else the first item.
 */
export function usePopup(selector: string, options: PopupOptions = {}): Popup {
  const { inline = false, closed = () => false, nested = false } = options;
  const open = shallowRef(false);
  const root = shallowRef<HTMLDivElement | null>(null);
  let focusOnOpen = false;
  const id = useId();
  useDropdownClose(open, root, closed, nested);
  const focusIn = (): void => focusPopupItem(root.value, selector);
  // After the render that shows the popup. Placement: a popup that would run past the editor's
  // right edge opens from its trigger's end, and is then held inside the editor's left edge. A
  // popup that would run past the bottom of the window opens above its trigger when there is
  // more room there (`placePopup`). Then a keyboard open focuses an item.
  watch(
    open,
    (isOpen) => {
      if (!isOpen || !root.value) return;
      placePopup(root.value, inline, selector);
      if (!focusOnOpen) return;
      focusOnOpen = false;
      focusIn();
    },
    { flush: 'post' }
  );
  // A popup that a pointer press opened leaves the keyboard outside it: a navigation key from
  // the grid or the toolbar moves the keyboard into it (`listenForToolbarPopupKeys`). The watch
  // runs at once, so a close in a focus event stops the listener before the next key.
  watch(
    [open, root],
    ([isOpen, el], _old, onCleanup) => {
      if (isOpen && el && !closed())
        onCleanup(listenForToolbarPopupKeys(el, selector, focusIn, nested));
    },
    { flush: 'sync' }
  );
  return {
    open,
    setOpen: (next) => {
      open.value = next;
    },
    root,
    id,
    openWithKeyboard: () => {
      focusOnOpen = true;
      open.value = true;
    },
    focusIn,
  };
}

/**
 * The keys of a popup trigger. Enter, Space, and Down open the popup with the keyboard in it.
 * On an open popup, Down moves the keyboard into it, and Enter or Space close it. A disabled
 * trigger opens nothing.
 */
export function popupTriggerKeys(popup: Popup, enabled: boolean) {
  return {
    onKeydown: (e: KeyboardEvent) => {
      const action = popupTriggerAction(e.key, popup.open.value);
      if (!action) return;
      e.preventDefault();
      if (!enabled) return;
      if (action === 'open') popup.openWithKeyboard();
      else if (action === 'focusIn') popup.focusIn();
      else popup.setOpen(false);
    },
    // Space activates a button on key up in some browsers: the key down already acted.
    onKeyup: (e: KeyboardEvent) => {
      if (e.key === ' ') e.preventDefault();
    },
  };
}
