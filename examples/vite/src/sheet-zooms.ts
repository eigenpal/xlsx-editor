// The zoom of each sheet of the demo. A load makes new sheet views at the zoom of the file, so
// the headings switch, which reopens the workbook, puts back the zoom that each sheet had.

import { useState, type RefObject } from 'react';
import type { EditorSnapshot, XlsxEditorRef } from '@xlsx-editor.dev/react';

/** A phone-sized window opens workbooks at this zoom, so more of the sheet shows. */
const NARROW_ZOOM = 0.75;

/** The zoom memory of the sheets of the open workbook. */
export interface SheetZooms {
  /** Reads a snapshot from the editor's change callback, and sets the zoom of a sheet that shows. */
  onSnapshot(snapshot: EditorSnapshot): void;
  /** A new workbook opens: it forgets every zoom. */
  forget(): void;
  /** The same workbook opens again: each sheet gets its zoom back when it next shows. */
  reopen(): void;
}

/**
 * The zoom memory of `editor`. When a sheet first shows after a load, it gets its zoom from
 * before the load, or on a phone the phone zoom. Later, it keeps the zoom that the user sets.
 */
export function useSheetZooms(editor: RefObject<XlsxEditorRef | null>): SheetZooms {
  const [state] = useState((): SheetZooms => {
    let zooms = new Map<number, number>();
    let shownSince = new Set<number>();
    return {
      onSnapshot(snapshot) {
        const sheet = snapshot.activeSheet;
        if (snapshot.status !== 'ready' || sheet < 0) return;
        if (shownSince.has(sheet)) {
          zooms.set(sheet, snapshot.zoom);
          return;
        }
        shownSince.add(sheet);
        const narrow = window.matchMedia('(max-width: 600px)').matches;
        const zoom = zooms.get(sheet) ?? (narrow ? NARROW_ZOOM : snapshot.zoom);
        if (zoom !== snapshot.zoom) editor.current?.getEditor()?.setZoom(zoom);
        else zooms.set(sheet, zoom);
      },
      forget() {
        zooms = new Map();
        shownSince = new Set();
      },
      reopen() {
        shownSince = new Set();
      },
    };
  });
  return state;
}
