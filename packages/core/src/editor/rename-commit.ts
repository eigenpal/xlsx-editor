// The commit of an open sheet rename before a chrome command (design D4 of the change
// `workbook-structure-editing`). Chrome `mousedown` keeps the focus where it is, so a press on
// a toolbar control does not blur the rename field, and the command would read the old name.
// The press commits the field first, as a blur does, through an event on the field: the sheet
// tabs answer it with the outcome, so no module keeps state. A refused name keeps its reason in
// view, and the click of that press runs no command. So does the click of a press on a control
// that was disabled at the press, because the commit can enable it (Undo). A press on a sheet
// tab that writes the name leaves the focus to its click (`focusAfterRelease`).
//
// Moved from the React adapter, so the React and Vue toolbars and sheet tabs share it. Adapter
// support, not a stable API.

/** The event that asks the focused rename field to commit; its detail gets the outcome. @internal */
export const RENAME_COMMIT_EVENT = 'xlsx-sheet-rename-commit';

/** The detail of `RENAME_COMMIT_EVENT`: the field sets `written` to false on a refusal. @internal */
export interface RenameCommitDetail {
  written: boolean;
}

/** Whether a sheet rename field has focus in the document of `from`. @internal */
export function renameFocused(from: Node): boolean {
  const doc = from.ownerDocument ?? (from as Document);
  return doc.activeElement?.classList.contains('xlsx-sheet-rename') ?? false;
}

/**
 * Commit the sheet rename field that has focus in the document of `from`, as a blur does. True
 * when no rename field has focus, or when the name was written; false when it was refused.
 * @internal
 */
export function commitFocusedRename(from: Node): boolean {
  const doc = from.ownerDocument ?? (from as Document);
  const field = doc.activeElement;
  const view = doc.defaultView;
  if (!field || !view || !field.classList.contains('xlsx-sheet-rename')) return true;
  const detail: RenameCommitDetail = { written: true };
  field.dispatchEvent(new view.CustomEvent(RENAME_COMMIT_EVENT, { detail }));
  return detail.written;
}

/**
 * Run no command on the click of this press in the document of `from`: the press whose rename
 * was refused, or whose control was disabled at the press. One click only: the click removes
 * the listeners, and so do the next press, a key press, and the task after the release, so a
 * press that is released away from its control, with no click, blocks no later click, such as
 * a click from the keyboard. @internal
 */
export function blockNextClick(from: Node): void {
  const doc = from.ownerDocument ?? (from as Document);
  const view = doc.defaultView;
  const stop = (e: Event) => {
    e.preventDefault();
    e.stopPropagation();
    clear();
  };
  // The click of a release comes in the same task as the release, before this timer.
  const released = () => {
    if (view) view.setTimeout(clear, 0);
    else clear();
  };
  const clear = () => {
    doc.removeEventListener('click', stop, true);
    doc.removeEventListener('mousedown', clear, true);
    doc.removeEventListener('keydown', clear, true);
    doc.removeEventListener('mouseup', released, true);
    doc.removeEventListener('pointerup', released, true);
  };
  doc.addEventListener('click', stop, true);
  doc.addEventListener('mousedown', clear, true);
  doc.addEventListener('keydown', clear, true);
  doc.addEventListener('mouseup', released, true);
  doc.addEventListener('pointerup', released, true);
}

/**
 * Call `focus` after the release of this press in the document of `from`, when the focus is
 * then on the page body. A press on a sheet tab that ends a written rename leaves the focus to
 * its click, so the grid says only the sheet that the click chooses. A release away from the
 * tab has no click, and the rename field that had the focus is gone. One release only: the
 * release removes the listeners, and so do the next press and a key press, so a release that
 * the page does not get moves no later focus. @internal
 */
export function focusAfterRelease(from: Node, focus: () => void): void {
  const doc = from.ownerDocument ?? (from as Document);
  const view = doc.defaultView;
  const check = () => {
    const at = doc.activeElement;
    if (!at || at === doc.body) focus();
  };
  // The click of a release comes in the same task as the release, before this timer.
  const released = () => {
    clear();
    if (view) view.setTimeout(check, 0);
    else check();
  };
  const clear = () => {
    doc.removeEventListener('mousedown', clear, true);
    doc.removeEventListener('keydown', clear, true);
    doc.removeEventListener('mouseup', released, true);
  };
  doc.addEventListener('mousedown', clear, true);
  doc.addEventListener('keydown', clear, true);
  doc.addEventListener('mouseup', released, true);
}
