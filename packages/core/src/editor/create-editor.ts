// The entry of the workbook editor: `createXlsxEditor`, and the paint hook of the tests. Kept
// apart from the editor facade (`xlsx-editor.ts`), so the facade stays under the line cap.

import type { XlsxEditorOptions } from './editor-model.ts';
import type { XlsxEditorInstance } from './editor-instance.ts';
import { Editor } from './xlsx-editor.ts';

/**
 * Create a workbook editor. Mount it with `attach`, open a file with `load`.
 *
 * @public
 */
export function createXlsxEditor(options: XlsxEditorOptions = {}): XlsxEditorInstance {
  return new Editor(options);
}

/**
 * Paint an editor's window now, without waiting for an animation frame. For tests.
 *
 * @internal
 */
export function flushEditor(editor: XlsxEditorInstance): void {
  if (editor instanceof Editor) editor.render();
}
