// The state of a toolbar popup and the keys of its trigger, for the toolbar parts and the
// overflow panel (`parts.tsx`, `Toolbar.tsx`). Each popup closes through the one close rule,
// `useDropdownClose`, and takes the navigation keys from outside it while a pointer press holds
// it open (core's `listenForToolbarPopupKeys`).

import { useEffect, useId, useLayoutEffect, useRef, useState, type KeyboardEvent } from 'react';
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
  readonly closed?: boolean;
  /** A popup can open inside this one (the overflow panel). */
  readonly nested?: boolean;
}

/**
 * The state of one popup: open, its root, and how the keyboard moves focus in. `selector`
 * names the items that take focus: the selected item first, else the current item
 * (`aria-current`), else the first item.
 */
export function usePopup(selector: string, options: PopupOptions = {}) {
  const { inline = false, closed = false, nested = false } = options;
  const [open, setOpen] = useState(false);
  const root = useRef<HTMLDivElement>(null);
  const focusOnOpen = useRef(false);
  const id = useId();
  useDropdownClose(open, setOpen, root, closed, nested);
  // Placement: a popup that would run past the editor's right edge opens from its trigger's
  // end, and is then held inside the editor's left edge. A popup that would run past the
  // bottom of the window opens above its trigger when there is more room there (`placePopup`).
  useLayoutEffect(() => {
    if (open && root.current) placePopup(root.current, inline, selector);
  }, [open, inline, selector]);
  /**
   * Focus the selected item, or else the current item (`aria-current`: an entry that names
   * the value but is not picked), or the first.
   */
  const focusIn = () => focusPopupItem(root.current, selector);
  useEffect(() => {
    if (!open || !focusOnOpen.current) return;
    focusOnOpen.current = false;
    focusIn();
  });
  // A popup that a pointer press opened leaves the keyboard outside it: a navigation key from
  // the grid or the toolbar moves the keyboard into it (`listenForToolbarPopupKeys`).
  useEffect(() => {
    const el = root.current;
    if (!open || closed || !el) return undefined;
    return listenForToolbarPopupKeys(el, selector, () => focusPopupItem(el, selector), nested);
  }, [open, closed, selector, nested]);
  /** Open from the keyboard: focus the selected item, or the first. */
  const openWithKeyboard = () => {
    focusOnOpen.current = true;
    setOpen(true);
  };
  return { open, setOpen, root, id, openWithKeyboard, focusIn };
}

/** The state of one popup, as `usePopup` returns it. */
export type Popup = ReturnType<typeof usePopup>;

/**
 * The keys of a popup trigger. Enter, Space, and Down open the popup with the keyboard in it.
 * On an open popup, Down moves the keyboard into it, and Enter or Space close it. A disabled
 * trigger opens nothing.
 */
export function popupTriggerKeys(popup: Popup, enabled: boolean) {
  return {
    onKeyDown: (e: KeyboardEvent<HTMLElement>) => {
      const action = popupTriggerAction(e.key, popup.open);
      if (!action) return;
      e.preventDefault();
      if (!enabled) return;
      if (action === 'open') popup.openWithKeyboard();
      else if (action === 'focusIn') popup.focusIn();
      else popup.setOpen(false);
    },
    // Space activates a button on key up in some browsers: the key down already acted.
    onKeyUp: (e: KeyboardEvent<HTMLElement>) => {
      if (e.key === ' ') e.preventDefault();
    },
  };
}
