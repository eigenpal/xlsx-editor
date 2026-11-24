/** @jsxImportSource vue */
// The frame of an editor dialog: the Vue twin of the React adapter's `dialog-frame.tsx`, which
// is after docx-editor/packages/react/src/editor/dialog-parts.tsx at 03735f11 (EigenPal),
// `DialogFrame`. A native `dialog` opened with `showModal`, the first control focused, Tab
// kept inside (`trapTabWithin`), Escape closes (never mid-composition), a press outside the
// panel closes, and every key and press stops at the frame, so the grid and the toolbar never
// act on them. On close, focus goes back to the opener when it is still in the page. A frame
// outside `.xlsx-editor` scopes itself with that class (and `dark`), so the tokens apply.
//
// The differences from DOCX Editor are the React frame's. Differences from React: `onMounted`
// and `onBeforeUnmount` replace the layout effect, and the parts are the default slot.

import { defineComponent, onBeforeUnmount, onMounted, shallowRef, type PropType } from 'vue';
import { isDialogBackdropPoint, trapTabWithin } from '@xlsx-editor.dev/core/editor';
import { useChromeDark } from './color-mode.ts';
import { useChromeDir } from './locale.ts';

/** An open modal dialog. Mount it to open it; unmount it to close it. */
export const DialogFrame = defineComponent({
  name: 'XlsxDialogFrame',
  props: {
    /** The `data-xlsx-dialog` kind: the styling hook of the dialog. */
    kind: { type: String as PropType<'numberFormat'>, required: true },
    /** The accessible name. */
    label: { type: String, required: true },
    /** Close the dialog: Escape, or a press outside the panel. */
    onClose: { type: Function as PropType<() => void>, required: true },
    /**
     * Where focus goes when the dialog closes. Return true when it took focus; otherwise the
     * frame focuses the opener, when it is still in the page.
     */
    onRestoreFocus: { type: Function as PropType<() => boolean>, default: undefined },
    /** The dialog's own keys, after Escape and Tab. */
    onKeyDown: { type: Function as PropType<(event: KeyboardEvent) => void>, default: undefined },
  },
  setup(props, { slots }) {
    const panelRef = shallowRef<HTMLDialogElement | null>(null);
    let pressedOutside = false;
    const ownScope = shallowRef(false);
    const dark = useChromeDark();
    const dir = useChromeDir();
    let opener: HTMLElement | null = null;
    onMounted(() => {
      const panel = panelRef.value;
      if (!panel) return;
      ownScope.value = !panel.parentElement?.closest('.xlsx-editor');
      // The page itself is no opener: restoring it would take focus from the editor.
      const active = panel.ownerDocument.activeElement as HTMLElement | null;
      opener = active === panel.ownerDocument.body ? null : active;
      if (typeof panel.showModal === 'function') panel.showModal();
      else panel.setAttribute('open', ''); // DOM test environments; browsers use native modality.
      panel
        .querySelector<HTMLElement>(
          'input:not([disabled]),select:not([disabled]),button:not([disabled]),[tabindex="0"]'
        )
        ?.focus({ preventScroll: true });
    });
    onBeforeUnmount(() => {
      const panel = panelRef.value;
      if (!panel) return;
      // Focus leaves the panel first, so the native close has nothing to restore, and focus
      // moves once: to where `onRestoreFocus` or the opener puts it.
      const active = panel.ownerDocument.activeElement as HTMLElement | null;
      if (active && panel.contains(active)) active.blur?.();
      if (typeof panel.close === 'function' && panel.open) panel.close();
      else panel.removeAttribute('open');
      if (props.onRestoreFocus?.()) return;
      if (opener?.isConnected) opener.focus({ preventScroll: true });
    });
    return () => (
      <dialog
        ref={panelRef}
        aria-modal="true"
        aria-label={props.label}
        data-xlsx-dialog={props.kind}
        dir={dir.value}
        class={
          ownScope.value ? `xlsx-dialog xlsx-editor${dark.value ? ' dark' : ''}` : 'xlsx-dialog'
        }
        onCancel={(event: Event) => {
          event.preventDefault();
          props.onClose();
        }}
        onMousedown={(event: MouseEvent) => {
          pressedOutside = isDialogBackdropPoint(event.currentTarget as HTMLDialogElement, event);
          event.stopPropagation();
        }}
        onClick={(event: MouseEvent) => {
          const close =
            pressedOutside &&
            isDialogBackdropPoint(event.currentTarget as HTMLDialogElement, event);
          pressedOutside = false;
          if (close) props.onClose();
          event.stopPropagation();
        }}
        onKeydown={(event: KeyboardEvent) => {
          // Some browsers send the key that ends a composition with `keyCode` 229 and
          // `isComposing` false, as the grid keys handle it.
          const composing = event.isComposing || event.keyCode === 229;
          if (event.key === 'Escape' && !composing) {
            event.preventDefault();
            props.onClose();
          } else if (trapTabWithin(event.currentTarget as HTMLElement, event))
            event.preventDefault();
          else if (!composing && !event.defaultPrevented) props.onKeyDown?.(event);
          event.stopPropagation();
        }}
      >
        {slots.default?.()}
      </dialog>
    );
  },
});
