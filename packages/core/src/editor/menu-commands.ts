// The state and the run of the slots that only the menu bar shows: `file.open`, the clipboard
// and select all (`edit`), the zoom (`view`), the new sheet (`sheet.add`), and the merge
// commands (`merge`). `toolbarCommandState` and `runToolbarCommand` call this file for these
// slots, so `toolbarCommandState` stays the one source of their enabled state. After the state
// of the same slots in pptx-editor/packages/core/src/editor/chrome-adapter-state.ts and their run
// in pptx-editor/packages/core/src/editor/chrome-actions.ts (`runClipboardSlot`) at fef01c93
// (EigenPal). Differences:
// - The state reads the snapshot and the host only, as every toolbar slot does here, and its
//   reasons are catalog keys.
// - Cut and Copy go through the grid's own clipboard events, which write the workbook's token and
//   the text (`clipboard-events.ts`), so a later Control or Command and V pastes what the menu
//   copied. Where the grid handles no event (`gridClipboardHandled`), they copy through the
//   editor and write the text with the clipboard of the host window.
// - Paste reads the text of the system clipboard through `ToolbarHost.readClipboard`, which the
//   menu bar gives when the host window has a clipboard that a page can read
//   (`systemClipboardReader`). Without it, Paste is disabled with `disabledReason.noClipboard`.
//   A page cannot read the workbook's private clipboard type, so text that is the text of the
//   pending copy or cut, while its outline shows, takes that copy's token (`tokenOfText`): the
//   paste then takes the lane of the paste key, and a cut moves its cells. Other text, and the
//   text of a copy whose outline ended or that expired, pastes in the plain lane. A read that
//   the browser refuses shows the refusal `clipboard-read`.
// - In an open cell edit, Cut, Copy, and Paste act on the text of the cell editor, as Control or
//   Command and X, C, or V do there: Cut and Copy fire their event at the focused input, which
//   the grid lets pass, and Paste puts the read text into the input at its selection.
// - Open is the host's (`ToolbarHost.open`): the engine never fetches a workbook.
// - The merge commands read `mergeCommandState`. The engine leaves a merge over values outside
//   the top-left cell enabled, with `dropsValuesAt`, so that the host asks the user first. The
//   library has no dialog that asks, so the row is disabled with
//   `disabledReason.mergeDropsValues`, which names the cell and tells the user to clear those
//   cells first: the menu bar never drops values without a question.

import type { ChromeSlotId } from './chrome-controls.ts';
import type { EditResult } from './cell-input.ts';
import type { EditorSnapshot } from './editor-model.ts';
import { MAX_ZOOM, MIN_ZOOM } from './editor-model.ts';
import type { PasteData } from './clipboard.ts';
import { mergeCommandState, type MergeCommand } from './merge-state.ts';
import { sheetCommandState } from './sheet-command-state.ts';
import {
  editorHostDocument,
  editorHostWindow,
  hostElementOfEditor,
  hostWindowOf,
} from './host-realm.ts';
import { editingOf } from './cell-editing.ts';
import { clipboardOf } from './clipboard.ts';
import { gridClipboardHandled, gridClipboardRefusal } from './clipboard-events.ts';
import type { XlsxEditorInstance } from './xlsx-editor.ts';
import type { ChromeDisabledReason, ToolbarCommandState, ToolbarHost } from './toolbar-commands.ts';

/** The slots of this file. @internal */
export const MENU_ONLY_SLOTS: ReadonlySet<ChromeSlotId> = new Set<ChromeSlotId>([
  'file.open',
  'edit.cut',
  'edit.copy',
  'edit.paste',
  'edit.selectAll',
  'view.zoomIn',
  'view.zoomOut',
  'view.zoomReset',
  'sheet.add',
  'merge.merge',
  'merge.mergeCenter',
  'merge.unmerge',
]);

/** The slots that view mode leaves enabled: they read the workbook or change only the view. */
const VIEW_SAFE = new Set<ChromeSlotId>([
  'file.open',
  'edit.copy',
  'edit.selectAll',
  'view.zoomIn',
  'view.zoomOut',
  'view.zoomReset',
]);

const MERGE: Readonly<Record<string, MergeCommand>> = {
  'merge.merge': 'merge',
  'merge.mergeCenter': 'mergeCenter',
  'merge.unmerge': 'unmerge',
};

const on = (id: ChromeSlotId): ToolbarCommandState => ({
  id,
  enabled: true,
  disabledReason: null,
  active: false,
});
const off = (id: ChromeSlotId, reason: ChromeDisabledReason): ToolbarCommandState => ({
  id,
  enabled: false,
  disabledReason: reason,
  active: false,
});

/**
 * The state of a slot of `MENU_ONLY_SLOTS`, from the snapshot and the host. `file.open` needs
 * only `host.open`, also with no workbook open; the others need a ready workbook. View mode
 * leaves the slots that do not edit enabled: Open, Copy, Select all, and the zoom. @internal
 */
export function menuSlotState(
  snapshot: EditorSnapshot,
  id: ChromeSlotId,
  host: ToolbarHost
): ToolbarCommandState {
  if (snapshot.status === 'loading') return off(id, { key: 'disabledReason.loading' });
  if (id === 'file.open') {
    if (!host.open) return off(id, { key: 'disabledReason.noOpen' });
    return host.opening?.() ? off(id, { key: 'disabledReason.loading' }) : on(id);
  }
  if (snapshot.status !== 'ready' || !snapshot.selection)
    return off(id, { key: 'disabledReason.editorNotReady' });
  if (snapshot.mode === 'view' && !VIEW_SAFE.has(id))
    return off(id, { key: 'disabledReason.viewMode' });
  switch (id) {
    case 'view.zoomIn':
      return snapshot.zoom >= MAX_ZOOM ? off(id, { key: 'disabledReason.zoomMax' }) : on(id);
    case 'view.zoomOut':
      return snapshot.zoom <= MIN_ZOOM ? off(id, { key: 'disabledReason.zoomMin' }) : on(id);
    case 'view.zoomReset':
      return snapshot.zoom === 1 ? off(id, { key: 'disabledReason.zoomActual' }) : on(id);
    default:
      break;
  }
  // The composed text belongs to the input method until it ends, as for every toolbar slot.
  if (snapshot.editing && snapshot.composing) return off(id, { key: 'disabledReason.composing' });
  // A copy, a cut, and a paste act on one range, and so does a merge (`mergeCommandState`): a
  // selection of several ranges disables them, as the commands refuse it (`multiple-ranges`).
  // In an open cell edit, the clipboard slots act on the text of the cell editor.
  if (
    (id === 'edit.cut' || id === 'edit.copy' || id === 'edit.paste') &&
    !snapshot.editing &&
    (snapshot.selection.ranges?.length ?? 0) > 1
  )
    return off(id, { key: 'disabledReason.multipleRanges' });
  const merge = MERGE[id];
  if (merge) {
    const state = mergeCommandState(snapshot, merge);
    if (state.disabledReason) return off(id, state.disabledReason);
    return state.dropsValuesAt
      ? off(id, { key: 'disabledReason.mergeDropsValues', params: { cell: state.dropsValuesAt } })
      : on(id);
  }
  if (id === 'sheet.add') {
    const state = sheetCommandState(snapshot, 'add', 0);
    return state.disabledReason ? off(id, state.disabledReason) : on(id);
  }
  // A cut clears the values of the copied cells, and a paste writes over the selected cells: a
  // locked cell with content refuses them, as it refuses Clear contents
  // (`selectionContentLocked`). In an open cell edit, both act on the text of the cell editor.
  if (
    (id === 'edit.cut' || id === 'edit.paste') &&
    !snapshot.editing &&
    snapshot.protection &&
    snapshot.selectionContentLocked
  )
    return off(id, { key: 'disabledReason.locked' });
  if (id === 'edit.paste' && !host.readClipboard)
    return off(id, { key: 'disabledReason.noClipboard' });
  return on(id);
}

/**
 * Run a slot of `MENU_ONLY_SLOTS`, or return null for another slot. Paste starts the read of the
 * clipboard and returns `ok`: the paste runs when the text arrives. @internal
 */
export function runMenuSlot(
  editor: XlsxEditorInstance,
  id: ChromeSlotId,
  host: ToolbarHost
): EditResult | null {
  if (!MENU_ONLY_SLOTS.has(id)) return null;
  const state = menuSlotState(editor.snapshot(), id, host);
  if (!state.enabled)
    return { ok: false, refusal: { code: 'not-editable', detail: `${id} is disabled` } };
  switch (id) {
    case 'file.open':
      host.open?.();
      return { ok: true };
    case 'edit.copy':
    case 'edit.cut':
      return copyThroughGrid(editor, id === 'edit.cut' ? 'cut' : 'copy');
    case 'edit.paste': {
      const read = host.readClipboard;
      if (!read) return { ok: false, refusal: { code: 'not-editable', detail: 'no clipboard' } };
      // The read can wait for the user's permission. A refused read pastes nothing, and shows
      // the refusal `clipboard-read`, whose message names the key that pastes in the grid.
      void read().then(
        (data) => {
          if (editor.snapshot().editing) {
            if (data.text) insertEditText(editor, data.text);
            return;
          }
          // As the paste key: no text with HTML is refused with `paste-html`, and an empty
          // clipboard changes nothing. Text that is the text of the pending copy or cut, while
          // its outline shows, takes its token, so the paste takes the lane of the paste key
          // with the private type: a cut moves its cells, and a copy keeps its formulas and
          // cell formats.
          const token =
            data.token ?? (data.text ? clipboardOf(editor)?.tokenOfText(data.text) : undefined);
          editor.paste(token ? { ...data, token } : data);
        },
        // The refusal shows in the snapshot, the status strip, and the live region, as the
        // refusal of an editor command does.
        () => void editingOf(editor)?.refuse({ code: 'clipboard-read' })
      );
      return { ok: true };
    }
    case 'edit.selectAll':
      editor.selectAll();
      return { ok: true };
    case 'view.zoomIn':
      editor.zoomStep(1);
      return { ok: true };
    case 'view.zoomOut':
      editor.zoomStep(-1);
      return { ok: true };
    case 'view.zoomReset':
      editor.setZoom(1);
      return { ok: true };
    case 'sheet.add':
      return editor.addSheet();
    case 'merge.merge':
      return editor.mergeCells();
    case 'merge.mergeCenter':
      return editor.mergeCells({ center: true });
    case 'merge.unmerge':
      return editor.unmergeCells();
    default:
      return null;
  }
}

/**
 * Copy or cut through the grid's own clipboard event: the grid takes the focus, and the
 * document's `copy` or `cut` command fires the event at it. Where the browser fires no event,
 * copy through the editor, and write the text with the host window's clipboard. In an open cell
 * edit, the cell editor takes the focus and the event, as for the key: the selected text of the
 * edit is copied or cut, and the cells are not.
 */
function copyThroughGrid(editor: XlsxEditorInstance, kind: 'copy' | 'cut'): EditResult {
  editor.focus();
  const doc = editorHostDocument(editor);
  const editing = editor.snapshot().editing;
  // The event targets the page selection: on a grid that already had the focus, no focus event
  // moved the selection into it, so the command moves it into the editor's own grid, as a press
  // on the grid does.
  // The editor's own grid: the scroller of the surface that the editor puts in its host
  // (`grid-surface.ts`), not the grid of another editor that the host page nests deeper.
  const grid =
    hostElementOfEditor(editor)?.querySelector(':scope > .xlsx-surface > .xlsx-scroller') ?? null;
  const selection = doc.getSelection?.();
  if (!editing && grid && selection && !grid.contains(selection.anchorNode))
    selection.collapse(grid, 0);
  // The grid counts each event that it handles, written or refused, so a host listener that
  // stops the event or cancels it cannot make the command copy twice or not at all, and a
  // refusal that the grid showed is not shown again by a second copy.
  const before = grid ? gridClipboardHandled(grid) : 0;
  let fired = false;
  try {
    fired = typeof doc.execCommand === 'function' && doc.execCommand(kind);
  } catch {
    fired = false;
  }
  if (grid !== null && gridClipboardHandled(grid) > before) {
    // The result of the grid's copy: a refusal that it showed is the command's result too.
    const refusal = gridClipboardRefusal(grid);
    return refusal ? { ok: false, refusal } : { ok: true };
  }
  if (fired && editing) return { ok: true };
  if (editing) return copyEditText(editor, doc, kind);
  const result = kind === 'cut' ? editor.cut() : editor.copy();
  if (!result.ok) return { ok: false, refusal: result.refusal };
  const clipboard = editorHostWindow(editor).navigator?.clipboard;
  void clipboard?.writeText?.(result.content.text).catch(() => undefined);
  return { ok: true };
}

/**
 * Copy or cut the selected text of the open cell editor where the browser fires no clipboard
 * event: write it with the host window's clipboard, and for a cut, remove it from the input as
 * the key does, with an `input` event that the cell editor reads.
 */
function copyEditText(editor: XlsxEditorInstance, doc: Document, kind: 'copy' | 'cut'): EditResult {
  const input = doc.activeElement as HTMLInputElement | HTMLTextAreaElement | null;
  if (!input || typeof input.setRangeText !== 'function') return { ok: true };
  const start = input.selectionStart ?? 0;
  const end = input.selectionEnd ?? start;
  if (end <= start) return { ok: true };
  const clipboard = editorHostWindow(editor).navigator?.clipboard;
  void clipboard?.writeText?.(input.value.slice(start, end)).catch(() => undefined);
  if (kind === 'cut') {
    input.setRangeText('', start, end, 'end');
    const InputEvent = doc.defaultView?.Event ?? Event;
    input.dispatchEvent(new InputEvent('input', { bubbles: true }));
  }
  return { ok: true };
}

/**
 * Put `text` into the open cell editor in place of its selected text, as a paste key does in the
 * input, with an `input` event that the cell editor reads.
 */
function insertEditText(editor: XlsxEditorInstance, text: string): void {
  editor.focus();
  const doc = editorHostDocument(editor);
  const input = doc.activeElement as HTMLInputElement | HTMLTextAreaElement | null;
  if (!input || typeof input.setRangeText !== 'function') return;
  const start = input.selectionStart ?? input.value.length;
  const end = input.selectionEnd ?? start;
  input.setRangeText(text, start, end, 'end');
  const InputEvent = doc.defaultView?.Event ?? Event;
  input.dispatchEvent(new InputEvent('input', { bubbles: true }));
}

/**
 * The reader of the system clipboard text of the window that holds `editor`, for a toolbar's
 * `ToolbarHost.readClipboard`, or undefined when that window has no clipboard that a page can
 * read. The window is resolved again at each read (`editorHostWindow`), so a slot of the menu
 * bar, such as Paste, placed in a toolbar group works there too. @internal
 */
export function editorClipboardReader(
  editor: XlsxEditorInstance | null | undefined
): (() => Promise<PasteData>) | undefined {
  if (!editor) return undefined;
  if (typeof editorHostWindow(editor).navigator?.clipboard?.readText !== 'function')
    return undefined;
  return async () => {
    return readClipboardData(editorHostWindow(editor).navigator?.clipboard);
  };
}

/**
 * The reader of the system clipboard text of the window that holds `node`, for
 * `ToolbarHost.readClipboard`, or undefined when that window has no clipboard that a page can
 * read. @internal
 */
export function systemClipboardReader(
  node: Node | null | undefined
): (() => Promise<PasteData>) | undefined {
  if (typeof hostWindowOf(node).navigator?.clipboard?.readText !== 'function') return undefined;
  // The window is resolved again at the read: the node can move to another window.
  return async () => {
    return readClipboardData(hostWindowOf(node).navigator?.clipboard);
  };
}

/**
 * The text of a system clipboard, and its HTML when it holds no text, so that a paste routes as
 * the paste key does (`WorkbookClipboard.paste`): text to the plain lane, HTML alone to the
 * refusal `paste-html`. The HTML itself is never read. A clipboard with no reader throws, which the
 * Paste row shows as `clipboard-read`.
 */
async function readClipboardData(clipboard: Clipboard | undefined): Promise<PasteData> {
  if (typeof clipboard?.readText !== 'function') throw new Error('no clipboard');
  const text = await clipboard.readText();
  if (text || typeof clipboard.read !== 'function') return { text };
  try {
    // Only the presence of HTML counts: the HTML is never read.
    for (const item of await clipboard.read())
      if (item.types.includes('text/html')) return { text, html: true };
  } catch {
    // A refused rich read leaves the plain text, as a paste of plain text.
  }
  return { text };
}
