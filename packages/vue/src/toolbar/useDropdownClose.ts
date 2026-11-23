// The Vue twin of the React adapter's `toolbar/useDropdownClose.ts`, which is copied from
// docx-editor/packages/react/src/editor/toolbar/useDropdownClose.ts at a0825f1f (EigenPal).
// The close rule is core's (`listenForPopupEscape`, `listenForPopupFocusLeave`, and
// `listenForOutsidePress`), so both adapters close a popup the same way.
//
// Differences from React: the open state, the hidden flag, and the popup element are refs, and
// a `watch` with an `onCleanup` replaces the effects.

import { watch, type Ref } from 'vue';
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
  open: Ref<boolean>,
  rootRef: Readonly<Ref<HTMLElement | null>>,
  hidden: () => boolean = () => false,
  nested = false
): void {
  watch(
    [open, hidden],
    ([isOpen, isHidden]) => {
      if (isOpen && isHidden) open.value = false;
    },
    { immediate: true }
  );

  watch(
    [open, hidden, rootRef],
    ([isOpen, isHidden, root], _old, onCleanup) => {
      if (!isOpen || isHidden || !root) return;
      const close = (): void => {
        open.value = false;
      };
      const contains = (node: Node) => root.contains(node);
      const stopEscape = listenForPopupEscape({
        popup: root,
        contains,
        chromeRoot: () => root.closest('.xlsx-toolbar'),
        ...(nested ? { skip: () => hasOpenNestedPopup(root.querySelector('[data-popup]')) } : {}),
        close: (fromInside) => {
          close();
          if (fromInside) root.querySelector<HTMLElement>('[aria-haspopup]')?.focus();
        },
      });
      const stopFocus = listenForPopupFocusLeave({ popup: root, contains, close });
      const stopPress = listenForOutsidePress({ popup: root, contains, close });
      onCleanup(() => {
        stopPress();
        stopEscape();
        stopFocus();
      });
    },
    { immediate: true, flush: 'post' }
  );
}
