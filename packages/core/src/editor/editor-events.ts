// Typed editor events for host chrome.
//
// The public shape is DOCX Editor's (packages/core/src/contracts/editor-events.ts at 5793d0234):
// the map is `EditorEvents`, a host subscribes with `editor.on(event, handler)`, which returns
// the unsubscribe function, `change` passes a `DocumentChange` alone, and `selectionChange`
// passes the snapshot. The way the events are made is PPTX Editor's
// (packages/core/src/editor/editor-events.ts at 4d976427): the editor has one notification
// path, `subscribe`, and each event compares the snapshot before and after a notification. The
// events are not a second path, so the React and Vue `useEditorEvent` hooks, which call
// `editor.on`, report the same events as `subscribe` sees.
//
// Intended differences from DOCX Editor:
// - The map has the five events that this editor makes: `change`, `selectionChange`,
//   `rejection`, `openFailure`, and `contextMenu`. DOCX Editor's drag, drop, paragraph, review, history, and
//   host events have no feature here yet.
// - `rejection` and `openFailure` are PPTX Editor's events, in place of DOCX Editor's
//   `diagnostic` and `error`. DOCX Editor's `diagnostic` is an anonymous channel with a
//   versioned schema of codes and bucketed numbers, for an error tracker. This editor has no
//   such schema, and its `EditRefusal` and `EditorFailure` can carry text from the file (a
//   sheet name, a cell reference, a part name), so they must not be sent as anonymous events.
//   DOCX Editor's `error` reports thrown errors; this editor refuses with a value and does not
//   throw. PPTX Editor solved the same problem, a refusal or a failure that the snapshot holds,
//   with `rejection` and `openFailure`, each with the value and the snapshot.
// - `DocumentChange.source` is `'load'` when a load settles and `'close'` when the workbook
//   closes, in place of `'load' | 'refresh' | 'recovery'`: this editor has no document refresh
//   and no refresh recovery, and it has `close()`. It has no `created`, `deleted`, or `dirty`
//   block ids: a workbook has no blocks, and the store does not report the cells of a change.
// - `rejection` reads only `refusal`, the refusal of the last command. `formatRefusal`,
//   `structureRefusal`, `clearRefusal`, and `sheetEditRefusal` say why a command would be
//   refused for the current selection: they are command states, not refused commands.
// - `DocumentChange.revision` starts again at 0 on each open, where DOCX Editor's revision only
//   grows: it is the store revision of the open workbook (`snapshot().revision`). A host that
//   keeps a revision, for example of its last save, starts again when `source` is `'load'`.
// - `selectionChange` fires only when the shown sheet or the selection changes, where DOCX
//   Editor also fires it for state derived from the selection. A host reads derived state, such
//   as `activeFormat`, with `subscribe` or the adapters' `useEditorState`. It also fires on each
//   load and close, because the sheet and the selection change then.
// - The selection compares by value, so a new selection object that selects the same cells
//   does not fire `selectionChange`.
// - `rejection` fires once for each refusal object that the snapshot shows. The editor stores a
//   new object for each refused command, so a repeat fires again. The exception is a `loading`
//   refusal that equals the one the snapshot shows (the same refused command, with the same
//   text): the editor does not store it again. A settled load shows the `edit-dropped` refusal
//   of its drop again, as the same object, so it does not fire twice, also for a subscription
//   that started during the load (`EditorListeners` passes the stored drop as known).
// - A handler that throws does not stop the other handlers or the editor's own updates: the
//   editor runs every listener and reports each error (`runListeners`), as DOCX Editor's
//   dispatch runs every queued notification.

import type { GridSelection } from '../contracts/index.ts';
import { runListeners } from '../store/store/listener-error.ts';
import type { EditRefusal } from './cell-input.ts';
import type { EditorFailure, EditorSnapshot } from './editor-model.ts';
import type { GridContextMenuRequest } from './grid-context-menu.ts';

/**
 * What caused a `change` event.
 *
 * @public
 */
export interface DocumentChange {
  /**
   * `'load'` when a load settles (the workbook opened or the open failed), and `'close'` when
   * the workbook closes. Absent for an edit, an undo, and a redo. Do not start external
   * processing, such as a save to a server, for a change with a `source`. A file picked in the
   * menu bar's File > Open that does not open keeps the workbook and fires no change.
   */
  readonly source?: 'load' | 'close';
  /**
   * The content revision after the change, the same number as `snapshot().revision`. It starts
   * again at 0 on each open, so compare it with a revision that you kept only while no change
   * with `source` `'load'` came between them.
   */
  readonly revision: number;
}

/**
 * What `editor.on(...)` and `useEditorEvent` can subscribe to, and what each handler receives.
 *
 * @public
 */
export interface EditorEvents {
  /**
   * The workbook content changed: an edit, an undo, or a redo (`snapshot().revision` changed),
   * or a load that settled, or a close (`change.source`). A selection change, a sheet change, a
   * zoom change, and the text of an open cell edit do not fire it. A change to view mode that
   * commits an open changed cell edit fires it for that commit. The root's `onChange` prop
   * differs: it fires after every state change. Read the new state with `snapshot()`.
   */
  change: (change: DocumentChange) => void;
  /**
   * The shown sheet (`snapshot.activeSheet`) or the selection (`snapshot.selection`) changed,
   * also when a load settles or the workbook closes. State derived from the selection, such as
   * `activeFormat`, does not fire it on its own.
   */
  selectionChange: (snapshot: EditorSnapshot) => void;
  /**
   * The editor refused a command, an edit, or a save (`snapshot.refusal`). It fires once for
   * each refused command, also when a command is refused again for the same reason. While a
   * load runs, every command is refused with `loading`; such a refusal fires once for each run
   * of the same refused command, and a repeat of it, with the same text, does not fire until
   * another refusal comes between. A save, or another command, fires again. It does not fire
   * again while the same refusal stays in the snapshot, and a load that dropped an open edit
   * fires `edit-dropped` once.
   */
  rejection: (refusal: EditRefusal, snapshot: EditorSnapshot) => void;
  /** A workbook did not open (`snapshot.failure`, with `snapshot.status` `'failed'`). */
  openFailure: (failure: EditorFailure, snapshot: EditorSnapshot) => void;
  /**
   * A secondary press on the grid, Shift+F10, or the Menu key asks for the grid context menu.
   * The press selected the cells under the pointer first when the selection did not hold them.
   * While at least one handler listens, the editor cancels the browser's own menu; with none,
   * the browser shows its menu. A secondary press in the open cell editor keeps the browser's
   * menu, for its text. The adapters' `ContextMenu` part listens to this event.
   */
  contextMenu: (request: GridContextMenuRequest) => void;
}

const samePoint = (a: GridSelection['active'], b: GridSelection['active']): boolean =>
  a.row === b.row && a.col === b.col;

function sameSelection(a: GridSelection | null, b: GridSelection | null): boolean {
  if (a === b) return true;
  if (!a || !b) return false;
  return (
    a.sheetIndex === b.sheetIndex &&
    samePoint(a.active, b.active) &&
    samePoint(a.anchor, b.anchor) &&
    samePoint(a.range.start, b.range.start) &&
    samePoint(a.range.end, b.range.end)
  );
}

// The snapshots that put back the state from before a load that did not replace it: a picked file
// that did not open (`openPickedWorkbook`). The workbook did not change, so such a snapshot fires
// no `change` and no `openFailure`.
const restored = new WeakSet<EditorSnapshot>();

/** The source of a load or a close that settled between `prev` and `next`, or undefined. */
function settled(prev: EditorSnapshot, next: EditorSnapshot): 'load' | 'close' | undefined {
  // The start of a load is no change: the workbook changes when the load settles.
  if (prev.status === next.status || next.status === 'loading') return undefined;
  return next.status === 'empty' ? 'close' : 'load';
}

/** The part of the editor that `onEditorEvent` reads. */
interface EventSource {
  snapshot(): EditorSnapshot;
  subscribe(listener: () => void): () => void;
}

/** Whether the event `event` happened between snapshots `prev` and `next`. */
function happened(event: keyof EditorEvents, prev: EditorSnapshot, next: EditorSnapshot): boolean {
  if ((event === 'change' || event === 'openFailure') && restored.has(next)) return false;
  switch (event) {
    case 'change':
      return prev.revision !== next.revision || settled(prev, next) !== undefined;
    case 'selectionChange':
      return (
        prev.activeSheet !== next.activeSheet || !sameSelection(prev.selection, next.selection)
      );
    case 'rejection':
      // `onEditorEvent` also skips a refusal object that it reported before.
      return !!next.refusal && next.refusal !== prev.refusal;
    case 'openFailure':
      return !!next.failure && next.status === 'failed' && next.failure !== prev.failure;
    case 'contextMenu':
      // A request is no state change: `EditorListeners.requestContextMenu` calls its handlers.
      return false;
  }
}

/**
 * Call `handler` each time `event` happens on `editor`. Returns the unsubscribe function. The
 * state at the call is the baseline: nothing fires for it. The editor's `on` method calls it.
 */
export function onEditorEvent<E extends keyof EditorEvents>(
  editor: EventSource,
  event: E,
  handler: EditorEvents[E],
  known: EditRefusal | null = null
): () => void {
  let prev = editor.snapshot();
  // The refusal objects seen so far: one that the snapshot shows again (the `edit-dropped`
  // refusal after a load that a `loading` refusal replaced for a while) is not a new refusal.
  // `known` is a refusal from before the call that the snapshot can show later: the stored
  // `edit-dropped` refusal of a drop, while a `loading` refusal hides it.
  const seen = new WeakSet<object>();
  if (prev.refusal) seen.add(prev.refusal);
  if (known) seen.add(known);
  return editor.subscribe(() => {
    const next = editor.snapshot();
    const before = prev;
    prev = next;
    if (next === before || !happened(event, before, next)) return;
    const call = handler as (...args: unknown[]) => void;
    if (event === 'rejection') {
      if (seen.has(next.refusal!)) return;
      seen.add(next.refusal!);
      call(next.refusal, next);
    } else if (event === 'openFailure') call(next.failure, next);
    else if (event === 'change') {
      const source = settled(before, next);
      const change: DocumentChange = source
        ? { source, revision: next.revision }
        : { revision: next.revision };
      call(change);
    } else call(next);
  });
}

/**
 * The listeners of one editor: `subscribe`, `on`, and the notification after each state change.
 * Every listener runs, also when another throws (`runListeners`), so the editor's own updates
 * after a notification, such as the repaint, still run.
 */
export class EditorListeners {
  readonly #listeners = new Set<() => void>();
  readonly #menu = new Set<EditorEvents['contextMenu']>();
  readonly #editor: Pick<EventSource, 'snapshot'>;
  readonly #known: () => EditRefusal | null;

  /**
   * `editor` gives the state that the events compare; `known` gives the refusal that a new `on`
   * subscription must not report later (`onEditorEvent`). Neither is read before the first call.
   */
  constructor(editor: Pick<EventSource, 'snapshot'>, known: () => EditRefusal | null) {
    this.#editor = editor;
    this.#known = known;
  }

  readonly subscribe = (listener: () => void): (() => void) => {
    this.#listeners.add(listener);
    return () => void this.#listeners.delete(listener);
  };

  readonly on = <E extends keyof EditorEvents>(
    event: E,
    handler: EditorEvents[E]
  ): (() => void) => {
    if (event === 'contextMenu') {
      // A fresh function per subscription, so the same handler twice unsubscribes once each.
      const call: EditorEvents['contextMenu'] = (r) => (handler as EditorEvents['contextMenu'])(r);
      this.#menu.add(call);
      return () => void this.#menu.delete(call);
    }
    return onEditorEvent(
      { snapshot: () => this.#editor.snapshot(), subscribe: this.subscribe },
      event,
      handler,
      this.#known()
    );
  };

  /** Whether a handler listens to `contextMenu`: the editor then cancels the browser's menu. */
  hasContextMenu(): boolean {
    return this.#menu.size > 0;
  }

  /** Call every `contextMenu` handler with `request` (`runListeners`). */
  requestContextMenu(request: GridContextMenuRequest): void {
    runListeners([...this.#menu].map((handler) => () => handler(request)));
  }

  /**
   * Call every listener. `restore`: the snapshot puts back the state from before a load that
   * did not replace the workbook, so it fires no `change` and no `openFailure`.
   */
  notify(restore = false): void {
    if (restore) restored.add(this.#editor.snapshot());
    runListeners(this.#listeners);
  }

  /** Drop every listener: the editor is destroyed. */
  clear(): void {
    this.#listeners.clear();
    this.#menu.clear();
  }
}
