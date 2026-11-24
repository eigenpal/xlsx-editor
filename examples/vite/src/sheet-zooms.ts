// The zoom of each sheet of the demo on a phone. A load makes new sheet views at the zoom of the
// file. On a phone-sized window, each sheet first shows at the phone zoom, so more of the sheet
// shows. Later, it keeps the zoom that the user sets.

import { useState, type RefObject } from 'react';
import type { EditorSnapshot, XlsxEditorRef } from '@xlsx-editor.dev/react';

/** A phone-sized window opens workbooks at this zoom, so more of the sheet shows. */
const NARROW_ZOOM = 0.75;

/** The zoom memory of the sheets of the open workbook. */
export interface SheetZooms {
  /** Reads a snapshot from the editor's change callback, and sets the zoom of a sheet that shows. */
  onSnapshot(snapshot: EditorSnapshot): void;
  /** A new workbook opens: it forgets which sheets showed. */
  forget(): void;
}

/**
 * The sheets of `editor` that showed since the last load. When a sheet first shows after a load,
 * it gets the phone zoom on a phone. Later, it keeps the zoom that the user sets.
 */
export function useSheetZooms(editor: RefObject<XlsxEditorRef | null>): SheetZooms {
  const [state] = useState((): SheetZooms => {
    let shownSince = new Set<number>();
    return {
      onSnapshot(snapshot) {
        const sheet = snapshot.activeSheet;
        if (snapshot.status !== 'ready' || sheet < 0 || shownSince.has(sheet)) return;
        shownSince.add(sheet);
        const narrow = window.matchMedia('(max-width: 600px)').matches;
        if (narrow && snapshot.zoom !== NARROW_ZOOM)
          editor.current?.getEditor()?.setZoom(NARROW_ZOOM);
      },
      forget() {
        shownSince = new Set();
      },
    };
  });
  return state;
}
