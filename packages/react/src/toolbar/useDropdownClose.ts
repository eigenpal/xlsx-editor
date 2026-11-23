// Copied from docx-editor/packages/react/src/editor/toolbar/useDropdownClose.ts at a0825f1f
// (EigenPal), which equals the file as b0c962c1 last changed it. Differences: there is no
// navigation viewport, so `editorElements` is not passed and `useRef` is not imported (core's
// `editorInstanceScope` finds the grid in the `.xlsx-editor` container); the chrome root is
// the `.xlsx-toolbar` instead of the `.docx-editor`; a `nested` parameter passes `skip` with
// core's `hasOpenNestedPopup`, so the overflow panel leaves an Escape to a popup open inside
// it; and the doc comment is rewritten for the toolbar's popups, the grid, and `nested`.

import { useEffect, type RefObject } from 'react';
import {
  hasOpenNestedPopup,
  listenForOutsidePress,
  listenForPopupEscape,
  listenForPopupFocusLeave,
} from '@xlsx-editor.dev/core/editor';

/**
 * Close a toolbar popup on an outside press, on Escape, or when focus moves outside it. This
 * is the one close rule of every toolbar popup: the dropdowns, the color palettes, the menus,
 * and the overflow panel.
 *
 * Escape goes through core's `listenForPopupEscape`. An Escape from this editor (the popup,
 * the toolbar that holds it, or the instance container with the grid) closes the popup and
 * stops there, so the same key does not also cancel the cell editor. An Escape from host
 * content or from another editor closes the popup and keeps going. A close from inside the
 * popup returns focus to its trigger; a close from the grid leaves the keyboard in the grid,
 * with the selection where it was. Focus that moves outside the popup and outside the grid
 * closes it too. A hidden slot closes its popup. `nested` leaves an Escape to a popup open
 * inside this one (a dropdown in the overflow panel).
 */
export function useDropdownClose(
  open: boolean,
  setOpen: (open: boolean) => void,
  rootRef: RefObject<HTMLElement | null>,
  hidden = false,
  nested = false
): void {
  useEffect(() => {
    if (open && hidden) setOpen(false);
  }, [open, hidden, setOpen]);

  useEffect(() => {
    const root = rootRef.current;
    if (!open || hidden || !root) return undefined;
    const contains = (node: Node) => root.contains(node);
    const stopEscape = listenForPopupEscape({
      popup: root,
      contains,
      chromeRoot: () => root.closest('.xlsx-toolbar'),
      ...(nested ? { skip: () => hasOpenNestedPopup(root.querySelector('[data-popup]')) } : {}),
      close: (fromInside) => {
        setOpen(false);
        if (fromInside) root.querySelector<HTMLElement>('[aria-haspopup]')?.focus();
      },
    });
    const stopFocus = listenForPopupFocusLeave({
      popup: root,
      contains,
      close: () => setOpen(false),
    });
    const stopPress = listenForOutsidePress({ popup: root, contains, close: () => setOpen(false) });
    return () => {
      stopPress();
      stopEscape();
      stopFocus();
    };
  }, [open, hidden, setOpen, rootRef, nested]);
}
