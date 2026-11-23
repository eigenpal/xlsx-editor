// The save of the open workbook in short tasks. The workbook is captured when the save starts:
// the package is immutable, so the bytes describe it as it was then, and an edit made while
// the save runs belongs to the next save.

import { saveXlsxSteps, type WorkbookDocumentStore } from '../store/index.ts';
import type { CellEditing } from './cell-editing.ts';
import type { SavedWorkbook } from './editor-instance.ts';
import { runStepsInSlices } from './sliced-steps.ts';

/** The bytes that a running save wrote, when it ends. No caller holds this array. */
interface WrittenBytes {
  bytes?: Uint8Array<ArrayBuffer>;
}

/** The running save of each opened workbook. Each open has its own store. */
const runningSaves = new WeakMap<
  WorkbookDocumentStore,
  {
    readonly revision: number;
    readonly saving: Promise<SavedWorkbook | null>;
    readonly written: WrittenBytes;
  }
>();

/**
 * The bytes of the last save of each opened workbook. The package is immutable, so another save
 * of the same revision gives the same bytes: repeated presses of Save do no work.
 */
const finishedSaves = new WeakMap<
  WorkbookDocumentStore,
  { readonly revision: number; readonly bytes: Uint8Array<ArrayBuffer> }
>();

/** The running save that each joined call waits for. */
const joinedSaves = new WeakMap<Promise<SavedWorkbook | null>, Promise<SavedWorkbook | null>>();

/** The one running save behind `saving`, so that callers of one save can tell that they joined. */
export const sharedSaveOf = (saving: Promise<SavedWorkbook | null>) =>
  joinedSaves.get(saving) ?? saving;

/**
 * `saved` with bytes of its own, copied from the `written` bytes that no caller holds when the
 * caller first reads them.
 */
function ownBytes(saved: SavedWorkbook | null, written: WrittenBytes): SavedWorkbook | null {
  if (!saved) return null;
  let copy: Uint8Array<ArrayBuffer> | undefined;
  return {
    get bytes() {
      copy ??= (written.bytes ?? saved.bytes).slice();
      return copy;
    },
    revision: saved.revision,
    stillOpen: saved.stillOpen,
  };
}

/**
 * Saves `store`, or refuses as `save` refuses: no workbook is ready, or an open cell edit does
 * not commit. A workbook that cannot serialize is a refused save too, with the cause in the
 * `detail` of the refusal, unless another workbook replaced it. `isCurrent` tells whether the
 * workbook is still open when the bytes are ready. A call during a running save of the same
 * revision joins that save, and gets its own copy of the bytes. A call after a
 * finished save of the same revision gets its own copy of those bytes.
 */
export function saveInSlices(
  store: WorkbookDocumentStore | null,
  edits: CellEditing,
  host: HTMLElement | null,
  isCurrent: () => boolean
): Promise<SavedWorkbook | null> {
  if (!store) return Promise.resolve(edits.refuseSave(null));
  const open = edits.commitOpenEdit();
  if (!open.ok) return Promise.resolve(edits.refuseSave(open.refusal));
  edits.saved(null);
  const revision = store.revision;
  const running = runningSaves.get(store);
  if (running?.revision === revision) {
    const joined = running.saving.then((saved) => ownBytes(saved, running.written));
    joinedSaves.set(joined, running.saving);
    return joined;
  }
  const finished = finishedSaves.get(store);
  if (finished?.revision === revision)
    return Promise.resolve({ bytes: finished.bytes.slice(), revision, stillOpen: isCurrent() });
  const ended = () => {
    if (runningSaves.get(store)?.revision === revision) runningSaves.delete(store);
  };
  // The save rules run when the steps are made, so a workbook can already fail here.
  let steps: ReturnType<typeof saveXlsxSteps>;
  try {
    steps = saveXlsxSteps(store.package);
  } catch (cause) {
    return Promise.resolve(edits.refuseSave({ code: 'save-failed', detail: String(cause) }));
  }
  const written: WrittenBytes = {};
  const saving = runStepsInSlices(steps, host).then(
    (bytes) => {
      written.bytes = bytes;
      ended();
      finishedSaves.set(store, { revision, bytes });
      return { bytes: bytes.slice(), revision, stillOpen: isCurrent() };
    },
    (cause: unknown) => {
      ended();
      if (!isCurrent()) return null;
      return edits.refuseSave({ code: 'save-failed', detail: String(cause) });
    }
  );
  runningSaves.set(store, { revision, saving, written });
  return saving;
}
