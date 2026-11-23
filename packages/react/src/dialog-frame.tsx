/** @jsxImportSource react */
// The frame of an editor dialog. After docx-editor/packages/react/src/editor/dialog-parts.tsx
// at 03735f11 (EigenPal), `DialogFrame`: a native `dialog` opened with `showModal`, the first
// control focused, Tab kept inside (`trapTabWithin`), Escape closes (never mid-composition),
// a press outside the panel closes, and every key and press stops at the frame, so the grid
// and the toolbar never act on them. On close, focus goes back to the opener when it is still
// in the page. A frame outside `.xlsx-editor` scopes itself with that class (and `dark`), so
// the tokens apply.
//
// Differences:
// - The classes are `xlsx-dialog` and `.xlsx-editor`, this editor's names for `docx-dialog` and
//   `.docx-editor`.
// - The layout effect runs once, on mount (`[]`): this dialog has no session signal, and its
//   restore callback is read through a ref.
// - The kind attribute is `data-xlsx-dialog`, which `popup-escape.ts` reads as this editor's
//   modal.
// - The press outside must also start outside, as in
//   pptx-editor/packages/react/src/dialog-parts.tsx at fe9415e1, so a text selection dragged
//   out of the field does not close the dialog. The mousedown handler records where the press
//   started; the source only stops the event there.
// - `onRestoreFocus` lets the dialog choose where focus goes on close (the grid after an
//   apply), and replaces `restoreFocus`.
// - A key with `keyCode` 229 counts as composing, as `grid-keys.ts` reads it.
// - Focus on the page body at open is no opener, so a close does not restore focus to the body.
// - The cleanup blurs the focused control before the native close, so focus moves once, and
//   removes the `open` attribute when `close` is missing (DOM test environments).
// - The `role`, `style`, `className`, `describedBy`, `panelRef`, `dismissOutside`, session
//   signal, and busy props and the link focus fallback are left out, because this editor's one
//   dialog uses none of them.

import { useLayoutEffect, useRef, useState, type KeyboardEvent, type ReactNode } from 'react';
import { isDialogBackdropPoint, trapTabWithin } from '@xlsx-editor.dev/core/editor';
import { useChromeDark } from './color-mode.ts';

/** Props of {@link DialogFrame}. */
export interface DialogFrameProps {
  /** The `data-xlsx-dialog` kind: the styling hook of the dialog. */
  readonly kind: 'numberFormat';
  /** The accessible name. */
  readonly label: string;
  /** Close the dialog: Escape, or a press outside the panel. */
  onClose(): void;
  /**
   * Where focus goes when the dialog closes. Return true when it took focus; otherwise the
   * frame focuses the opener, when it is still in the page.
   */
  readonly onRestoreFocus?: () => boolean;
  /** The dialog's own keys, after Escape and Tab. */
  readonly onKeyDown?: (event: KeyboardEvent<HTMLDialogElement>) => void;
  /** The parts of the dialog: header, body, and footer. */
  readonly children: ReactNode;
}

/** An open modal dialog. Mount it to open it; unmount it to close it. */
export function DialogFrame({
  kind,
  label,
  onClose,
  onRestoreFocus,
  onKeyDown,
  children,
}: DialogFrameProps) {
  const ref = useRef<HTMLDialogElement | null>(null);
  const closeRef = useRef(onClose);
  closeRef.current = onClose;
  const restoreRef = useRef(onRestoreFocus);
  restoreRef.current = onRestoreFocus;
  const pressedOutside = useRef(false);
  // Through `className`, not `classList.add`: React rewrites `class` on each render.
  const [ownScope, setOwnScope] = useState(false);
  const dark = useChromeDark();
  useLayoutEffect(() => {
    const panel = ref.current;
    if (!panel) return undefined;
    setOwnScope(!panel.parentElement?.closest('.xlsx-editor'));
    // The page itself is no opener: restoring it would take focus from the editor.
    const active = panel.ownerDocument.activeElement as HTMLElement | null;
    const opener = active === panel.ownerDocument.body ? null : active;
    if (typeof panel.showModal === 'function') panel.showModal();
    else panel.setAttribute('open', ''); // DOM test environments; browsers use native modality.
    panel
      .querySelector<HTMLElement>(
        'input:not([disabled]),select:not([disabled]),button:not([disabled]),[tabindex="0"]'
      )
      ?.focus({ preventScroll: true });
    return () => {
      // Focus leaves the panel first, so the native close has nothing to restore, and focus
      // moves once: to where `onRestoreFocus` or the opener puts it.
      const active = panel.ownerDocument.activeElement as HTMLElement | null;
      if (active && panel.contains(active)) active.blur?.();
      if (typeof panel.close === 'function' && panel.open) panel.close();
      else panel.removeAttribute('open');
      if (restoreRef.current?.()) return;
      if (opener?.isConnected) opener.focus({ preventScroll: true });
    };
  }, []);
  return (
    <dialog
      ref={ref}
      aria-modal="true"
      aria-label={label}
      data-xlsx-dialog={kind}
      className={ownScope ? `xlsx-dialog xlsx-editor${dark ? ' dark' : ''}` : 'xlsx-dialog'}
      onCancel={(event) => {
        event.preventDefault();
        closeRef.current();
      }}
      onMouseDown={(event) => {
        pressedOutside.current = isDialogBackdropPoint(event.currentTarget, event);
        event.stopPropagation();
      }}
      onClick={(event) => {
        const close = pressedOutside.current && isDialogBackdropPoint(event.currentTarget, event);
        pressedOutside.current = false;
        if (close) closeRef.current();
        event.stopPropagation();
      }}
      onKeyDown={(event) => {
        // Some browsers send the key that ends a composition with `keyCode` 229 and
        // `isComposing` false, as the grid keys handle it.
        const composing = event.nativeEvent.isComposing || event.keyCode === 229;
        if (event.key === 'Escape' && !composing) {
          event.preventDefault();
          closeRef.current();
        } else if (trapTabWithin(event.currentTarget, event.nativeEvent)) event.preventDefault();
        else if (!composing && !event.defaultPrevented) onKeyDown?.(event);
        event.stopPropagation();
      }}
    >
      {children}
    </dialog>
  );
}
