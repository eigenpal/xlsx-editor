// The public event composable in happy-dom, the twin of the React adapter's
// `editor-event.test.tsx`: `useEditorEvent` subscribes to the editor of the enclosing root and
// stops at unmount. `change` follows content only, while the root's `change` emit follows every
// state change.
//
// Differences from React: `setup` runs once, so the handler does not change between renders and
// React's case of a new handler has no twin; the unmount case removes the probe with a ref.

import { afterEach, describe, expect, test } from 'bun:test';
import { defineComponent, h, nextTick, shallowRef } from 'vue';
import { buildXlsx, c, row } from '../../../scripts/fixtures/builder.ts';
import {
  SheetTabs,
  StatusMessage,
  XlsxEditorRoot,
  XlsxEditorViewport,
  useEditorEvent,
  type EditorEvents,
  type XlsxEditorInstance,
} from '../src/index.ts';
import { FONTS, mount, tabsShow, unmountAll } from './mount.ts';

afterEach(unmountAll);

const workbook = buildXlsx({
  sheets: [{ name: 'S', rows: row(1, [c('A1', 1), c('B1', 'two')]) }],
});

describe('useEditorEvent', () => {
  async function probe<E extends keyof EditorEvents>(event: E) {
    const log: string[] = [];
    let rootChanges = 0;
    const shown = shallowRef(true);
    const Probe = defineComponent({
      setup() {
        useEditorEvent(event, (() => log.push(event)) as EditorEvents[E]);
        return () => null;
      },
    });
    let editor: XlsxEditorInstance | null = null;
    const view = mount(() =>
      h(
        XlsxEditorRoot,
        {
          document: workbook,
          fonts: FONTS,
          onReady: (e: XlsxEditorInstance) => (editor = e),
          onChange: () => (rootChanges += 1),
        },
        {
          default: () =>
            h(XlsxEditorViewport, null, () => [
              h(SheetTabs),
              shown.value ? h(Probe) : null,
              h(StatusMessage),
            ]),
        }
      )
    );
    await tabsShow(view, 1);
    return { log, editor: editor!, rootChanges: () => rootChanges, shown };
  }

  test('change fires for an edit, not for a selection or a zoom change', async () => {
    const { log, editor, rootChanges } = await probe('change');
    log.splice(0);
    const before = rootChanges();
    editor.select({ row: 0, col: 1 });
    editor.setZoom(1.5);
    expect(log).toEqual([]);
    // The root's `change` emit follows every state change.
    expect(rootChanges()).toBeGreaterThan(before);
    editor.enterValue('3', { row: 0, col: 0 });
    expect(log).toEqual(['change']);
  });

  test('selectionChange fires for a selection; unmount stops it', async () => {
    const { log, editor, shown } = await probe('selectionChange');
    log.splice(0);
    editor.select({ row: 1, col: 1 });
    expect(log).toEqual(['selectionChange']);
    shown.value = false;
    await nextTick();
    editor.select({ row: 2, col: 2 });
    expect(log).toEqual(['selectionChange']);
  });

  test('rejection fires once for each refused command', async () => {
    const { log, editor } = await probe('rejection');
    editor.setMode('view');
    editor.enterValue('9');
    editor.setZoom(2);
    expect(log).toEqual(['rejection']);
  });
});
