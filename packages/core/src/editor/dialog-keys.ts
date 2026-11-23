// Copied from docx-editor/packages/core/src/editor/paragraph-dialog-fields.ts at 03735f11
// (EigenPal), function `trapTabWithin`, with its comment. Difference: it lives in its own
// module, because this editor has no paragraph dialog. Its comment names a paragraph; in this
// editor the surface behind the dialog is the cell grid.
//
// `isDialogBackdropPoint` is copied from pptx-editor/packages/core/src/editor/dialog-draft.ts
// at fe9415e1 (EigenPal), with its comment, so each adapter tests a backdrop press the same way.
// Note: the copied comment names two adapters; this editor has one until task 8.1.

/**
 * Keep Tab inside the dialog.
 *
 * `aria-modal` tells assistive tech the rest of the page is inert; it does not stop Tab,
 * so without this the third Tab lands on the document behind the dialog — which is the
 * editable surface, so the next keystroke types into the paragraph being formatted.
 *
 * Returns true when the event was handled, so a caller only has to call `preventDefault`.
 * @public
 */
export function trapTabWithin(panel: HTMLElement, event: KeyboardEvent): boolean {
  if (event.key !== 'Tab') return false;
  const focusable = [
    ...panel.querySelectorAll<HTMLElement>(
      'button, [href], input, select, textarea, [tabindex]:not([tabindex="-1"])'
    ),
  ].filter((node) => !node.hasAttribute('disabled') && node.tabIndex !== -1);
  if (focusable.length === 0) return false;
  const first = focusable[0]!;
  const last = focusable[focusable.length - 1]!;
  const active = panel.ownerDocument.activeElement;
  // Wrap at whichever end the user is walking off, and treat "focus is on the panel
  // itself" as being before the first control — that is where it sits when the dialog
  // has just opened.
  if (event.shiftKey && (active === first || active === panel)) {
    last.focus();
    return true;
  }
  if (!event.shiftKey && active === last) {
    first.focus();
    return true;
  }
  return false;
}

/**
 * True when a pointer event on a dialog element lies outside its box: on the backdrop, as
 * DOCX Editor's frame checks it. Both adapters use it. @internal
 */
export function isDialogBackdropPoint(
  dialog: Element,
  event: { readonly target: EventTarget | null; readonly clientX: number; readonly clientY: number }
): boolean {
  if (event.target !== dialog) return false;
  const r = dialog.getBoundingClientRect();
  return (
    event.clientX < r.left ||
    event.clientX > r.right ||
    event.clientY < r.top ||
    event.clientY > r.bottom
  );
}
