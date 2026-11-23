/** @jsxImportSource react */
// The public event hook in happy-dom (`bun run test:dom`): `useEditorEvent` subscribes to the
// editor of the enclosing root, calls the latest handler without a new subscription, and stops
// at unmount. `change` follows content only, while the root's `onChange` follows every state
// change.

import { afterEach, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { act, cleanup, render, waitFor } from '@testing-library/react';
import { buildXlsx, c, row } from '../../../../scripts/fixtures/builder.ts';
import {
  XlsxEditor,
  useEditorEvent,
  type FontSource,
  type EditorEvents,
  type XlsxEditorInstance,
} from '../index.ts';

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

const workbook = buildXlsx({
  sheets: [{ name: 'S', rows: row(1, [c('A1', 1), c('B1', 'two')]) }],
});

describe('useEditorEvent', () => {
  async function probe<E extends keyof EditorEvents>(event: E) {
    const log: string[] = [];
    let rootChanges = 0;
    function Probe({ tag }: { tag: string }) {
      useEditorEvent(event, (() => log.push(tag)) as EditorEvents[E]);
      return null;
    }
    let editor: XlsxEditorInstance | null = null;
    const tree = (tag: string, withProbe = true) => (
      <XlsxEditor.Root
        document={workbook}
        fonts={FONTS}
        onReady={(e) => (editor = e)}
        onChange={() => (rootChanges += 1)}
      >
        <XlsxEditor.Viewport>
          <XlsxEditor.SheetTabs />
          {withProbe ? <Probe tag={tag} /> : null}
          <XlsxEditor.StatusMessage />
        </XlsxEditor.Viewport>
      </XlsxEditor.Root>
    );
    const view = render(tree('first'));
    await waitFor(() => expect(view.getAllByRole('tab').length).toBe(1), { timeout: 5000 });
    return {
      log,
      editor: editor!,
      rootChanges: () => rootChanges,
      rerender: (tag: string, withProbe = true) => view.rerender(tree(tag, withProbe)),
    };
  }

  test('change fires for an edit, not for a selection or a zoom change', async () => {
    const { log, editor, rootChanges } = await probe('change');
    log.splice(0);
    const before = rootChanges();
    act(() => void editor.select({ row: 0, col: 1 }));
    act(() => editor.setZoom(1.5));
    expect(log).toEqual([]);
    // The root's `onChange` follows every state change.
    expect(rootChanges()).toBeGreaterThan(before);
    act(() => void editor.enterValue('3', { row: 0, col: 0 }));
    expect(log).toEqual(['first']);
  });

  test('a new handler is called without a new subscription; unmount stops it', async () => {
    const { log, editor, rerender } = await probe('selectionChange');
    log.splice(0);
    act(() => void editor.select({ row: 1, col: 1 }));
    expect(log).toEqual(['first']);
    let subscribed = 0;
    const on = editor.on.bind(editor);
    editor.on = ((event, handler) => {
      subscribed += 1;
      return on(event, handler);
    }) as typeof editor.on;
    rerender('second');
    expect(subscribed).toBe(0);
    act(() => void editor.select({ row: 2, col: 2 }));
    expect(log).toEqual(['first', 'second']);
    rerender('third', false);
    act(() => void editor.select({ row: 3, col: 3 }));
    expect(log).toEqual(['first', 'second']);
  });

  test('rejection fires once for each refused command', async () => {
    const { log, editor } = await probe('rejection');
    act(() => void editor.setMode('view'));
    act(() => void editor.enterValue('9'));
    act(() => editor.setZoom(2));
    expect(log).toEqual(['first']);
  });
});
