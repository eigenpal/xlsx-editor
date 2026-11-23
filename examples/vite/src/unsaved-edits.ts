// Whether the open workbook has edits that are not saved. The demo asks before another
// workbook or a reload of the page replaces them.

import { useState, type RefObject } from 'react';
import type { EditorSnapshot, XlsxEditorRef } from '@xlsx-editor.dev/react';

/** The fields of an editor snapshot that tell an edit. */
export type EditFacts = Pick<EditorSnapshot, 'status' | 'revision' | 'editing' | 'editChanged'>;

/** The unsaved-edit state of the workbook that `editor` shows. */
export interface UnsavedEdits {
  /** The content revision of the editor snapshot. */
  revision(): number;
  /**
   * True when the open workbook has an edit after the last clean mark, or an open cell edit
   * with changed text, which a new load drops without a commit. With no open workbook, it is
   * true only during a reopen that carries edits (`holdEdits`).
   */
  hasUnsavedEdits(): boolean;
  /** Tells that a reopen runs whose saved bytes alone hold edits that are not saved to a file. */
  holdEdits(held: boolean): void;
  /**
   * Marks `revision` as the state with no unsaved edit. With `carried`, a reopen kept edits
   * that are not saved, so the new revision still has them.
   */
  markClean(revision: number, carried?: boolean): void;
}

/**
 * The unsaved-edit state over `facts`, read from the editor snapshot. It compares revisions:
 * an undo back to the opened state makes a new revision, so it still counts as an edit. The
 * demo can ask when nothing is lost, never the reverse.
 */
export function unsavedEdits(facts: () => EditFacts): UnsavedEdits {
  let clean = { revision: 0, carried: false };
  let held = false;
  return {
    revision: () => facts().revision,
    hasUnsavedEdits: () => {
      const s = facts();
      if (s.status !== 'ready') return held;
      return clean.carried || s.revision !== clean.revision || (s.editing && s.editChanged);
    },
    holdEdits: (on) => {
      held = on;
    },
    markClean: (at, carried = false) => {
      clean = { revision: at, carried };
    },
  };
}

/** The unsaved-edit state of the workbook that `editor` shows, made once for the component. */
export function useUnsavedEdits(editor: RefObject<XlsxEditorRef | null>): UnsavedEdits {
  const [state] = useState(() =>
    unsavedEdits(
      () =>
        editor.current?.snapshot() ?? {
          status: 'empty',
          revision: 0,
          editing: false,
          editChanged: false,
        }
    )
  );
  return state;
}
