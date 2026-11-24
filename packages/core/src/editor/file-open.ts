// The packaged File > Open of the menu bar: it opens a picked file in place of the open workbook
// only when the editor can open it. `XlsxEditorInstance.load` closes the open workbook when a
// load fails, so that no state shows that the editor would refuse to edit or save. A file that
// the user picks by mistake must not cost the open workbook, its edits, or its open cell edit,
// so this takes the engine's picked-file load (`registerPickedLoad`): the same load, with one
// read of the file, the loading status from the start of the file read, and the same rule that
// a later load wins. A file that the browser cannot read or the editor cannot open leaves the
// open workbook in the state it had and shows the refusal `open-failed`. The public `load`
// keeps its contract for every other caller.

import { pickedLoadOf } from './load-steps.ts';
import type { EditorSnapshot } from './editor-model.ts';
import type { XlsxEditorInstance } from './xlsx-editor.ts';

/**
 * Open a picked file (or its bytes) in `editor`. With a workbook open, a file that the browser
 * cannot read or the editor cannot open leaves that workbook, its edits, and its open cell edit
 * as they are, and shows the refusal `open-failed`. With no workbook open, the load shows its own
 * failure. A later pick or load wins over this one. With an editor that the engine did not make,
 * a file that the browser cannot read changes nothing. The promise never rejects. @internal
 */
export async function openPickedWorkbook(
  editor: XlsxEditorInstance,
  file: Blob | Uint8Array
): Promise<EditorSnapshot> {
  const bytes =
    file instanceof Uint8Array ? file : file.arrayBuffer().then((b) => new Uint8Array(b));
  const load = pickedLoadOf(editor);
  if (load) return load(bytes);
  // An editor that the engine did not make has only `load`, and no refusal channel: a file that
  // the browser cannot read leaves it as it is. The call only never rejects.
  try {
    return await editor.load(await bytes);
  } catch {
    return editor.snapshot();
  }
}
