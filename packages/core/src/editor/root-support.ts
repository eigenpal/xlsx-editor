// The parts of the adapters' root that do not depend on a framework: the empty snapshot of a
// mode before the editor mounts, the report of font failures, the wait of `onReady` for the
// first open, and the `mode` prop that waits for an open cell edit. Moved from the React
// adapter's `XlsxEditorRoot.tsx` and `context.ts`, so the React and Vue roots act the same.
// Adapter support, not a stable API.

import {
  EMPTY_SNAPSHOT,
  type EditorFailure,
  type EditorMode,
  type EditorSnapshot,
} from './editor-model.ts';
import { substituteFontFailure } from './substitute-fonts.ts';
import type { XlsxEditorInstance } from './xlsx-editor.ts';

/** The empty snapshot of each mode, so a read before mount keeps one identity. */
const EMPTY_BY_MODE: Readonly<Record<EditorMode, EditorSnapshot>> = {
  edit: EMPTY_SNAPSHOT,
  view: Object.freeze({ ...EMPTY_SNAPSHOT, mode: 'view' }),
};

/** The snapshot before an editor mounts, with the `mode` prop of the root applied. @internal */
export function emptySnapshotOf(mode: EditorMode): EditorSnapshot {
  return EMPTY_BY_MODE[mode] ?? EMPTY_SNAPSHOT;
}

/**
 * Report each font failure of an editor once: a failed open for its fonts, and each set of
 * packaged substitute faces that did not all load (`fontReport().substituteFailures`). Returns
 * the stop function.
 *
 * @internal
 */
export function watchFontErrors(
  instance: Pick<XlsxEditorInstance, 'subscribe' | 'snapshot' | 'fontReport'>,
  report: (failure: EditorFailure) => void
): () => void {
  let reported: EditorFailure | null = null;
  let substitutes = instance.fontReport().substituteFailures;
  return instance.subscribe(() => {
    const failure = instance.snapshot().failure;
    if (failure && failure !== reported && failure.code === 'fonts') {
      reported = failure;
      report(failure);
    }
    const now = instance.fontReport().substituteFailures;
    if (now === substitutes) return;
    substitutes = now;
    const partial = substituteFontFailure(now);
    if (partial) report(partial);
  });
}

/**
 * Call `fire` once, when the first open of `editor` settles (`snapshot().status` is not
 * `'loading'`), or at once when no open runs. Returns the stop function. @internal
 */
export function whenOpenSettles(
  editor: Pick<XlsxEditorInstance, 'subscribe' | 'snapshot'>,
  fire: () => void
): () => void {
  if (editor.snapshot().status !== 'loading') {
    fire();
    return () => {};
  }
  let done = false;
  const off = editor.subscribe(() => {
    if (done || editor.snapshot().status === 'loading') return;
    done = true;
    off();
    fire();
  });
  return off;
}

/** The `mode` prop of a root, applied to one editor. @internal */
export interface ModeProp {
  /** Apply the prop's value. A value equal to the last applied one changes nothing. */
  apply(mode: EditorMode): void;
  /** Stop the wait and the subscription. */
  dispose(): void;
}

/**
 * Apply a root's `mode` prop to `editor` when it changes, as DOCX Editor's root applies it: an
 * instance call to `setMode` stays until the next prop change. The first value applies only when
 * the editor is in another mode. A change that the open cell edit refuses waits, and applies when
 * the edit closes; `onWaits` tells the root whether a change waits, so it can show a notice. A
 * newer `setMode` call (`modeRequests`) during the wait ends the wait: that call wins.
 * @internal
 */
export function followModeProp(
  editor: Pick<XlsxEditorInstance, 'subscribe' | 'snapshot' | 'setMode'>,
  onWaits: (waits: boolean) => void
): ModeProp {
  let applied: EditorMode | null = null;
  let pending: { mode: EditorMode; requests: number } | null = null;
  const off = editor.subscribe(() => {
    const want = pending;
    if (want === null) return;
    const snap = editor.snapshot();
    if (snap.modeRequests === want.requests && snap.editing) return;
    // Clear the wait before the call: `setMode` publishes, which calls this again. A newer
    // applied `setMode` call (`modeRequests`) ends the wait without a call.
    pending = null;
    if (snap.modeRequests === want.requests && !editor.setMode(want.mode))
      pending = { mode: want.mode, requests: editor.snapshot().modeRequests };
    onWaits(pending !== null);
  });
  return {
    apply(mode) {
      if (applied === mode) return;
      const first = applied === null;
      applied = mode;
      // A new editor starts in the mode it was created with.
      if (first && editor.snapshot().mode === mode) {
        pending = null;
        onWaits(false);
        return;
      }
      const ok = editor.setMode(mode);
      pending = ok ? null : { mode, requests: editor.snapshot().modeRequests };
      onWaits(!ok);
    },
    dispose() {
      pending = null;
      off();
    },
  };
}
