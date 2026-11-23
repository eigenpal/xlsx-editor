/** @jsxImportSource react */
// The host field of `packages/react/docs/editing-cells.md`, written as the guide writes it, so
// typecheck proves it and the test shows that it edits, commits, and cancels the active cell.

import { afterEach, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { act, cleanup, fireEvent, render, waitFor } from '@testing-library/react';
import { buildXlsx, c, row } from '../../../../scripts/fixtures/builder.ts';
import {
  composesKey,
  editMoveOf,
  fieldHistoryKey,
  refusesEditText,
  useEditorSnapshot,
  useXlsxEditor,
  XlsxEditor,
  type FontSource,
  type XlsxEditorInstance,
} from '../index.ts';

function CellField() {
  const editor = useXlsxEditor();
  const snap = useEditorSnapshot();
  const text = snap.editText ?? snap.activeCell?.content ?? '';
  // Mark the field while the engine refuses the text of the open edit.
  const refusal = snap.editing ? snap.refusal : null;
  const invalid = refusal !== null && refusesEditText(refusal.code);
  return (
    <textarea
      rows={1}
      aria-label="Cell"
      aria-invalid={invalid || undefined}
      value={text}
      onChange={(e) => editor?.setEditText(e.currentTarget.value)}
      onKeyDown={(e) => {
        if (!editor || composesKey(e.nativeEvent, false)) return;
        const history = fieldHistoryKey(e, snap.editing);
        const move = editMoveOf(e);
        if (history) {
          e.preventDefault();
          if (history === 'undo') editor.undo();
          else editor.redo();
        } else if (move) {
          e.preventDefault();
          editor.commitEdit(move);
        } else if (e.key === 'Escape') {
          e.preventDefault();
          editor.cancelEdit();
        }
      }}
    />
  );
}

const fontsDir = join(
  dirname(
    Bun.resolveSync('@docx-editor.dev/fonts/package.json', join(import.meta.dir, '../../../core'))
  ),
  'assets'
);
const FONTS: FontSource[] = [
  {
    family: 'Carlito',
    bold: false,
    italic: false,
    bytes: new Uint8Array(readFileSync(join(fontsDir, 'Carlito-Regular.ttf'))),
    origin: 'substitute',
  },
];

afterEach(cleanup);

test('the host field of the editing guide edits, commits, cancels, and marks a refused text', async () => {
  let viewer: XlsxEditorInstance | null = null;
  const bytes = buildXlsx({ sheets: [{ name: 'First', rows: row(1, [c('A1', 'one')]) }] });
  const view = render(
    <div style={{ height: 400 }}>
      <XlsxEditor.Root document={bytes} fonts={FONTS} onReady={(v) => (viewer = v)}>
        <CellField />
        <XlsxEditor.Content />
        <XlsxEditor.StatusMessage />
      </XlsxEditor.Root>
    </div>
  );
  await waitFor(() => expect(viewer?.snapshot().status).toBe('ready'), { timeout: 5000 });
  const field = view.getByLabelText('Cell') as HTMLTextAreaElement;
  expect(field.value).toBe('one');
  act(() => void fireEvent.change(field, { target: { value: '=1+' } }));
  fireEvent.keyDown(field, { key: 'Enter' });
  expect(field.getAttribute('aria-invalid')).toBe('true');
  fireEvent.keyDown(field, { key: 'Escape' });
  expect(viewer!.snapshot().editing).toBe(false);
  expect(field.value).toBe('one');
  act(() => void fireEvent.change(field, { target: { value: '7' } }));
  fireEvent.keyDown(field, { key: 'Tab' });
  expect(viewer!.snapshot().selectionRef).toBe('B1');
  act(() => void viewer!.select({ row: 0, col: 0 }));
  expect(field.value).toBe('7');
});
