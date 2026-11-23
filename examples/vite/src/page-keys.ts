// The page keys of the demo, read in the capture phase before the editor reads them:
// - Escape closes the message from anywhere, when no cell edit is open: the close button comes
//   after the grid. The closing Escape is used, so it does not also arm the grid's exit or end
//   a pending copy or cut.
// - F6 and Shift+F6 move the focus among the regions of the page: the top bar, then the
//   editor's toolbar, grid, and footer. The editor viewport holds the top bar and cycles its own
//   regions (packages/react/src/landmarks.ts): from the top bar, F6 goes to the toolbar and
//   Shift+F6 to the footer. The demo adds the top bar to the cycle: F6 from the footer and
//   Shift+F6 from the toolbar go to the workbook list.
// - A held Control+S or Command+S saves once: a repeat stops before the toolbar, which claims
//   the key in the viewport from its `onSave`, reads it.
// - During an open, from the download or the file read to the end of the load, and during a
//   headings switch, Control+S and Command+S, and a press of the toolbar Save, stay in the page
//   and go to Save, which says why it waits. The toolbar would only refuse with the general
//   text of its `opening` prop. The focus can be on the body, outside the editor: there a key
//   that would type says that it did nothing.
// - While the bytes of a save wait for a press of Save, Control+S and Command+S, and a press of
//   the toolbar Save, also go to Save, which writes those bytes. The toolbar would prepare the
//   bytes again.

import { useEffect, useRef, type RefObject } from 'react';

export interface PageKeys {
  /** The workbook list of the top bar: the top bar's stop in the F6 cycle. */
  readonly workbookList: RefObject<HTMLSelectElement | null>;
  /** Whether a message shows. */
  readonly messageOpen: boolean;
  /** Whether a cell edit is open: Escape then belongs to the cell editor. */
  readonly editing: () => boolean;
  /** Closes the message. */
  readonly closeMessage: () => void;
  /** Whether a workbook opens (its download, file read, or load) or the headings switch. */
  readonly loading: () => boolean;
  /** Whether the bytes of a save wait for a press of Save. */
  readonly readyToSave: () => boolean;
  /** Presses Save. */
  readonly onSaveKey: () => void;
  /** Says that a key did nothing during a load. */
  readonly onKeyWhileLoading: () => void;
}

/** Reads the page keys on the window of the workbook list, while the component is mounted. */
export function usePageKeys(keys: PageKeys): void {
  const now = useRef(keys);
  now.current = keys;
  useEffect(() => {
    const list = now.current.workbookList.current;
    const win = list?.ownerDocument.defaultView;
    if (!list || !win) return undefined;
    const onKey = (e: KeyboardEvent) => {
      if (e.isComposing) return;
      const { workbookList, messageOpen, editing, closeMessage } = now.current;
      const { loading, readyToSave, onSaveKey, onKeyWhileLoading } = now.current;
      if (e.key === 'Escape' && messageOpen) {
        if (!editing()) {
          e.preventDefault();
          closeMessage();
        }
        return;
      }
      const at = e.target instanceof win.Element ? e.target : null;
      if (e.key === 'F6' && !e.metaKey && !e.ctrlKey && !e.altKey) {
        const out = e.shiftKey ? at?.closest('[role="toolbar"]') : at?.closest('.xlsx-footer');
        if (!out) return;
        e.preventDefault();
        e.stopPropagation();
        workbookList.current?.focus();
        return;
      }
      const save = e.key.toLowerCase() === 's' && (e.metaKey || e.ctrlKey);
      if (save && !e.shiftKey && !e.altKey && (loading() || readyToSave())) {
        e.preventDefault();
        e.stopPropagation();
        if (!e.repeat) onSaveKey();
        return;
      }
      if (e.target === list.ownerDocument.body && loading()) {
        if (e.key.length === 1 && !e.metaKey && !e.ctrlKey && !e.altKey) {
          e.preventDefault();
          onKeyWhileLoading();
        }
        return;
      }
      if (save && !e.shiftKey && !e.altKey && e.repeat) {
        e.preventDefault();
        e.stopPropagation();
      }
    };
    const onClick = (e: MouseEvent) => {
      const at = e.target instanceof win.Element ? e.target : null;
      if (!at?.closest('[data-slot="file.save"]')) return;
      if (!now.current.loading() && !now.current.readyToSave()) return;
      e.preventDefault();
      e.stopPropagation();
      now.current.onSaveKey();
    };
    win.addEventListener('keydown', onKey, true);
    win.addEventListener('click', onClick, true);
    return () => {
      win.removeEventListener('keydown', onKey, true);
      win.removeEventListener('click', onClick, true);
    };
  }, []);
}
