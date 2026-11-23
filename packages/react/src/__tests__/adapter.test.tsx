/** @jsxImportSource react */
// The React adapter in happy-dom (`bun run test:dom` preloads it): the root opens a workbook, the
// tabs show the visible sheets and switch the grid, the formula bar follows the selection, the
// zoom control zooms, and chrome presses keep the keyboard on the grid.

import { afterEach, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { act, cleanup, fireEvent, render, waitFor } from '@testing-library/react';
import { Profiler, useEffect, useRef, useState } from 'react';
import { createPortal } from 'react-dom';
import { buildXlsx, c, row } from '../../../../scripts/fixtures/builder.ts';
import {
  XlsxEditor,
  XlsxEditorContext,
  useEditorColorMode,
  useEditorState,
  useScopeClassName,
  type EditorSnapshot,
  type FontSource,
  type XlsxEditorInstance,
  type XlsxEditorRef,
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
  sheets: [
    { name: 'First', rows: row(1, [c('A1', 'one')]) + row(5, [c('A5', 84, { f: 'A2*2' })]) },
    { name: 'Hidden', state: 'hidden', rows: row(1, [c('A1', 'secret')]) },
    { name: 'Second', rows: row(1, [c('A1', 'two')]) },
  ],
});

/** A field of the formula bar by its label. */
const field = (view: ReturnType<typeof render>, name: string) =>
  view.getByLabelText(name) as HTMLInputElement | HTMLTextAreaElement;

async function open() {
  let viewer: XlsxEditorInstance | null = null;
  const view = render(
    <div style={{ height: 400 }}>
      <XlsxEditor document={workbook} fonts={FONTS} onReady={(v) => (viewer = v)} />
    </div>
  );
  await waitFor(() => expect(view.getAllByRole('tab').length).toBe(2), { timeout: 5000 });
  return { view, viewer: viewer! };
}

describe('XlsxEditor', () => {
  test('shows one tab for each visible sheet, in workbook order', async () => {
    const { view } = await open();
    expect(view.getAllByRole('tab').map((t) => t.textContent)).toEqual(['First', 'Second']);
    expect(view.getByRole('region', { name: 'Workbook editor' })).toBeTruthy();
    expect(view.getByRole('tablist', { name: 'Sheets' })).toBeTruthy();
  });

  test('opens a workbook from an ArrayBuffer', async () => {
    const buffer = workbook.buffer.slice(
      workbook.byteOffset,
      workbook.byteOffset + workbook.byteLength
    ) as ArrayBuffer;
    const view = render(<XlsxEditor document={buffer} fonts={FONTS} />);
    await waitFor(() => expect(view.getAllByRole('tab').length).toBe(2), { timeout: 5000 });
  });

  test('a tab switches the sheet; arrow keys move between tabs', async () => {
    const { view, viewer } = await open();
    const [first, second] = view.getAllByRole('tab');
    expect(first!.getAttribute('aria-selected')).toBe('true');
    fireEvent.click(second!);
    expect(viewer.snapshot().activeSheet).toBe(2);
    fireEvent.keyDown(view.getAllByRole('tab')[1]!, { key: 'ArrowLeft' });
    expect(viewer.snapshot().activeSheet).toBe(0);
  });

  test('the formula bar shows the active reference and the formula text', async () => {
    const { view, viewer } = await open();
    act(() => viewer.select({ row: 4, col: 0 }));
    expect(field(view, 'Cell reference').value).toBe('A5');
    expect(field(view, 'Cell content').value).toBe('=A2*2');
  });

  test('the formula bar follows the text of an open edit', async () => {
    const { view, viewer } = await open();
    const grid = view.container.querySelector('.xlsx-scroller') as HTMLElement;
    act(() => viewer.select({ row: 0, col: 0 }));
    expect(field(view, 'Cell content').value).toBe('one');
    fireEvent.keyDown(grid, { key: '7' });
    const input = view.getByLabelText('Cell editor') as HTMLTextAreaElement;
    expect(field(view, 'Cell content').value).toBe('7');
    input.value = '72';
    fireEvent.input(input);
    expect(field(view, 'Cell content').value).toBe('72');
    fireEvent.keyDown(input, { key: 'Escape' });
    expect(field(view, 'Cell content').value).toBe('one');
  });

  test('the formula bar is not a live region: the grid announces each move once', async () => {
    const { view } = await open();
    for (const name of ['Cell reference', 'Cell content']) {
      expect(field(view, name).closest('[aria-live]:not([aria-live="off"])')).toBeNull();
      expect(field(view, name).getAttribute('role')).toBeNull();
    }
  });

  test('a value edit in a workbook with a formula shows the stale notice', async () => {
    const { view, viewer } = await open();
    const notice = () => view.container.querySelector('.xlsx-notice');
    // The status node is mounted before the notice, so its text is announced when it arrives.
    expect(notice()?.getAttribute('role')).toBe('status');
    expect(notice()?.textContent).toBe('');
    act(() => void viewer.enterValue('7'));
    await waitFor(() =>
      expect(notice()?.textContent).toContain('Formula results can be out of date')
    );
    expect(notice()?.getAttribute('role')).toBe('status');
    expect(viewer.snapshot().stale).toBe(true);
  });

  test('the status strip keeps its place above the grid, and a refusal wraps to two lines', async () => {
    const { view, viewer } = await open();
    const grid = view.container.querySelector('.xlsx-canvas') as HTMLElement;
    const strip = view.container.querySelector('.xlsx-status-strip') as HTMLElement;
    // Empty, the strip already sits above the grid, with only the empty status node.
    expect(strip.textContent).toBe('');
    expect(grid.previousElementSibling).toBe(strip);
    act(() => void viewer.enterValue('7'));
    act(() => void viewer.beginEdit('=SUM(', 'enter'));
    act(() => void viewer.commitEdit());
    await waitFor(() => expect(strip.querySelector('.xlsx-refusal')).not.toBeNull());
    // Both messages share the one line: the notice shows its short text, the full one in title.
    // The short text is visual only: the live text of the notice stays the full text.
    const notice = strip.querySelector('.xlsx-notice') as HTMLElement;
    expect(notice.hasAttribute('data-short')).toBe(true);
    expect(notice.querySelector('.xlsx-notice-brief')!.textContent).toBe(
      'Formula results can be out of date.'
    );
    expect(notice.querySelector('.xlsx-notice-brief')!.getAttribute('aria-hidden')).toBe('true');
    expect(notice.querySelector('.xlsx-notice-full')!.textContent).toContain(
      'the editor shows the stored results'
    );
    expect(notice.title).toContain('the editor shows the stored results');
    expect(view.container.querySelector('.xlsx-status-strip')).toBe(strip);
    expect(grid.previousElementSibling).toBe(strip);
    // The strip is one line, and grows while a refusal shows; it hides nothing of a refusal,
    // and no message is positioned over the grid.
    const css = readFileSync(join(import.meta.dir, '../../../core/src/styles/editor.css'), 'utf8');
    const rule = /\.xlsx-status-strip \{([^}]*)\}/.exec(css)![1]!;
    expect(rule).toMatch(/\bmin-height: \d+px;/);
    expect(rule).not.toContain('max-height');
    expect(rule).toContain('flex-wrap: wrap;');
    expect(rule).toContain('flex: none;');
    // A refusal shows in full on its own row in every layout: no line clamp and no height
    // limit, and it takes the whole row, so the notices go to the next row.
    const refusalRule = /\n\.xlsx-refusal \{([^}]*)\}/.exec(css)![1]!;
    expect(refusalRule).toContain('flex: 1 1 100%;');
    expect(refusalRule).toContain('white-space: normal;');
    expect(css).not.toMatch(/line-clamp/);
    expect(css).not.toMatch(/\.xlsx-refusal[^{]*\{[^}]*\s(max-height|height):/);
    expect(css).toMatch(/\n\.xlsx-refusal ~ \.xlsx-notice:not\(:empty\) \{[^}]*border-top:/);
    // In a narrow editor, a later notice takes its own row, with or without a refusal.
    const narrowBlock = /@container xlsx-editor \(max-width: 599px\) \{[\s\S]*?\n\}/.exec(css)![0];
    expect(narrowBlock).toMatch(
      /\n {2}\.xlsx-notice:not\(:empty\) ~ \.xlsx-notice:not\(:empty\) \{[^}]*flex: 1 1 100%;[^}]*border-top:/
    );
    // Notices that show side by side have a separator, also with an empty notice between them.
    expect(css).toMatch(
      /\n\.xlsx-notice:not\(:empty\) ~ \.xlsx-notice:not\(:empty\) \{[^}]*border-inline-start:/
    );
    // A notice that shows alone wraps instead of being cut.
    expect(css).toMatch(/\n\.xlsx-notice:not\(\[data-short\]\) \{[^}]*white-space: normal;/);
    // The one exception is the visually hidden full text inside a notice: 1px, clipped to
    // nothing, and placed in its notice, which is the positioned parent.
    const exception = /(?:^|\n)\.xlsx-notice\[data-short\] \.xlsx-notice-full \{[^}]*\}/;
    const hidden = exception.exec(css)![0];
    for (const decl of [
      'position: absolute;',
      'width: 1px;',
      'height: 1px;',
      'overflow: hidden;',
      'clip-path: inset(50%);',
    ])
      expect(hidden).toContain(decl);
    expect(css).toMatch(/\n\.xlsx-refusal,\n\.xlsx-notice \{[^}]*position: relative;/);
    expect(css.replace(exception, '')).not.toMatch(
      /\.xlsx-(refusal|notice)[^{]*\{[^}]*position: absolute/
    );
    // In a narrow editor each notice under a refusal takes its own row.
    const narrow = /@container xlsx-editor \(max-width: \d+px\) \{([\s\S]*?)\n\}/.exec(css)![1]!;
    expect(narrow).toMatch(/\.xlsx-refusal ~ \.xlsx-notice:not\(:empty\) \{[^}]*flex: 1 1 100%;/);
    expect(css).toMatch(/\.xlsx-editor-viewport \{[^}]*container: xlsx-editor \/ inline-size;/);
  });

  test('F6 and Shift+F6 move the keyboard among the toolbar, the formula bar, the grid, and the footer', async () => {
    const { view } = await open();
    const nameBox = view.getByLabelText('Cell reference');
    const grid = view.container.querySelector('.xlsx-scroller') as HTMLElement;
    const toolbar = view.getByRole('toolbar');
    const footer = view.container.querySelector('.xlsx-footer') as HTMLElement;
    act(() => grid.focus());
    const f6 = (shiftKey = false) =>
      fireEvent.keyDown(view.container.ownerDocument.activeElement!, { key: 'F6', shiftKey });
    f6();
    expect(footer.contains(document.activeElement)).toBe(true);
    // The footer takes the keyboard on the selected sheet tab.
    expect(document.activeElement?.getAttribute('aria-selected')).toBe('true');
    f6();
    expect(toolbar.contains(document.activeElement)).toBe(true);
    // The formula bar takes the keyboard on the name box.
    f6();
    expect(document.activeElement).toBe(nameBox);
    f6();
    expect(document.activeElement).toBe(grid);
    f6(true);
    expect(document.activeElement).toBe(nameBox);
    f6(true);
    expect(toolbar.contains(document.activeElement)).toBe(true);
    f6(true);
    expect(footer.contains(document.activeElement)).toBe(true);
  });

  test('F6 in the cell editor commits the edit and moves to the next region', async () => {
    const { view, viewer } = await open();
    const footer = view.container.querySelector('.xlsx-footer') as HTMLElement;
    act(() => viewer.select({ row: 1, col: 1 }));
    act(() => void viewer.beginEdit('7'));
    const input = view.getByLabelText('Cell content') as HTMLTextAreaElement;
    const editor = view.container.querySelector('.xlsx-scroller textarea, .xlsx-scroller input');
    expect(editor).not.toBeNull();
    act(() => (editor as HTMLElement).focus());
    fireEvent.keyDown(editor!, { key: 'F6' });
    expect(viewer.snapshot().editing).toBe(false);
    expect(footer.contains(document.activeElement)).toBe(true);
    act(() => viewer.select({ row: 1, col: 1 }));
    expect(input.value).toBe('7');
  });

  test('the zoom control zooms and resets', async () => {
    const { view, viewer } = await open();
    expect(view.getByRole('group', { name: 'Zoom level' })).toBeTruthy();
    fireEvent.click(view.getByRole('button', { name: 'Zoom in' }));
    expect(viewer.snapshot().zoom).toBe(1.1);
    view.getByRole('button', { name: 'Reset the zoom to 100%. The zoom is 110%.' });
    fireEvent.click(view.getByRole('button', { name: /Reset the zoom/ }));
    expect(viewer.snapshot().zoom).toBe(1);
  });

  test('a zoom button at its limit keeps the focus and does nothing', async () => {
    const { view, viewer } = await open();
    act(() => viewer.setZoom(4));
    const zoomIn = view.getByRole('button', { name: 'Zoom in' });
    act(() => zoomIn.focus());
    expect(zoomIn.getAttribute('aria-disabled')).toBe('true');
    expect(zoomIn.hasAttribute('disabled')).toBe(false);
    fireEvent.click(zoomIn);
    expect(viewer.snapshot().zoom).toBe(4);
    expect(document.activeElement).toBe(zoomIn);
    fireEvent.click(view.getByRole('button', { name: 'Zoom out' }));
    expect(zoomIn.getAttribute('aria-disabled')).toBe('false');
  });

  test('pressing chrome does not take the focus from the grid', async () => {
    const { view } = await open();
    const press = fireEvent.mouseDown(view.getAllByRole('tab')[0]!);
    expect(press).toBe(false);
  });

  test('an edit from the grid commits through the engine, and a refusal shows once', async () => {
    const { view, viewer } = await open();
    const grid = view.container.querySelector('.xlsx-scroller') as HTMLElement;
    act(() => viewer.select({ row: 1, col: 1 }));
    fireEvent.keyDown(grid, { key: '5' });
    const input = view.getByLabelText('Cell editor') as HTMLTextAreaElement;
    fireEvent.keyDown(input, { key: 'Enter' });
    act(() => viewer.select({ row: 1, col: 1 }));
    expect(field(view, 'Cell content').value).toBe('5');
    fireEvent.keyDown(grid, { key: '=' });
    input.value = '=SUM(';
    fireEvent.keyDown(input, { key: 'Enter' });
    await waitFor(() =>
      expect(view.container.querySelector('.xlsx-refusal')?.textContent).toBe(
        'The formula was not applied: it cannot be read. Check the quotes, parentheses, and references.'
      )
    );
    // One announcement: the engine's live region. The shown refusal has no live role.
    expect(view.queryByRole('alert')).toBeNull();
    expect(view.container.querySelector('.xlsx-live')?.textContent).toBe(
      'The formula was not applied: it cannot be read. Check the quotes, parentheses, and references.'
    );
    fireEvent.keyDown(input, { key: 'Escape' });
    expect(viewer.snapshot().refusal).toBeNull();
  });

  test('the handle is small: it saves, reads the state, and hands out the editor', async () => {
    let handle: XlsxEditorRef | null = null;
    let viewer: XlsxEditorInstance | null = null;
    const view = render(
      <div style={{ height: 400 }}>
        <XlsxEditor
          ref={(r) => {
            handle = r;
          }}
          document={workbook}
          fonts={FONTS}
          onReady={(v) => (viewer = v)}
        />
      </div>
    );
    await waitFor(() => expect(view.getAllByRole('tab').length).toBe(2), { timeout: 5000 });
    act(() => {
      viewer!.enterValue('edited');
    });
    expect(handle!.save()!.byteLength).toBeGreaterThan(0);
    expect(handle!.getEditor()).toBe(viewer);
    expect(handle!.snapshot()).toBe(viewer!.snapshot());
    expect(Object.keys(handle!).sort()).toEqual([
      'focus',
      'getEditor',
      'load',
      'save',
      'saveInTasks',
      'snapshot',
    ]);
    act(() => {
      expect(handle!.getEditor()!.undo().ok).toBe(true);
    });
    expect(viewer!.snapshot().canUndo).toBe(false);
  });

  test('before the editor mounts, the state and the handle read the mode prop', () => {
    const modes: string[] = [];
    const handleModes: string[] = [];
    let handle: XlsxEditorRef | null = null;
    function Probe() {
      modes.push(useEditorState((s) => s.mode));
      const first = useRef(true);
      // A child's effect runs after the root attaches its handle, before the root's own effect
      // creates the editor.
      useEffect(() => {
        if (first.current) handleModes.push(handle!.snapshot().mode);
        first.current = false;
      });
      return null;
    }
    render(
      <XlsxEditor.Root
        ref={(r) => {
          handle = r;
        }}
        fonts={FONTS}
        mode="view"
      >
        <Probe />
      </XlsxEditor.Root>
    );
    expect(modes[0]).toBe('view');
    expect(handleModes).toEqual(['view']);
    expect(handle!.getEditor()!.snapshot().mode).toBe('view');
  });

  test('in development, a layout with the grid and no StatusMessage gets one warning', async () => {
    const warnings: string[] = [];
    const warn = console.warn;
    console.warn = (message: string) => warnings.push(message);
    try {
      const bare = render(
        <XlsxEditor.Root fonts={FONTS}>
          <XlsxEditor.Content />
        </XlsxEditor.Root>
      );
      await waitFor(() => expect(warnings.length).toBe(1));
      expect(warnings[0]).toContain('XlsxEditor.StatusMessage');
      bare.unmount();
      warnings.length = 0;
      render(
        <XlsxEditor.Root fonts={FONTS}>
          <XlsxEditor.Content />
          <XlsxEditor.StatusMessage />
        </XlsxEditor.Root>
      );
      await act(async () => {});
      expect(warnings).toEqual([]);
    } finally {
      console.warn = warn;
    }
  });

  test('useEditorState re-renders only when its slice changes', async () => {
    let renders = 0;
    let seen = -2;
    const pick = (s: EditorSnapshot) => s.activeSheet;
    function Probe() {
      seen = useEditorState(pick);
      renders += 1;
      return null;
    }
    let viewer: XlsxEditorInstance | null = null;
    const view = render(
      <div style={{ height: 400 }}>
        <XlsxEditor.Root document={workbook} fonts={FONTS} onReady={(v) => (viewer = v)}>
          <XlsxEditor.Viewport>
            <XlsxEditor.SheetTabs />
            <Probe />
          </XlsxEditor.Viewport>
        </XlsxEditor.Root>
      </div>
    );
    await waitFor(() => expect(view.getAllByRole('tab').length).toBe(2), { timeout: 5000 });
    expect(seen).toBe(0);
    const before = renders;
    act(() => viewer!.setZoom(1.5));
    act(() => viewer!.select({ row: 3, col: 3 }));
    expect(renders).toBe(before);
    act(() => viewer!.setActiveSheet(2));
    expect(seen).toBe(2);
    expect(renders).toBe(before + 1);
  });

  test('useEditorState reads no stale snapshot and survives key repeat with no update-depth error', async () => {
    // Each notification comes after `publish()` builds the new snapshot, so a read inside it is
    // never stale; 500 synchronous key presses in one act stay under the update-depth limit.
    const seen: number[] = [];
    const pick = (s: EditorSnapshot) => s.selection?.active.row ?? -1;
    function Probe() {
      seen.push(useEditorState(pick));
      return null;
    }
    let viewer: XlsxEditorInstance | null = null;
    const view = render(
      <div style={{ height: 400 }}>
        <XlsxEditor.Root document={workbook} fonts={FONTS} onReady={(v) => (viewer = v)}>
          <XlsxEditor.Viewport>
            <XlsxEditor.Content />
            <XlsxEditor.SheetTabs />
            <Probe />
          </XlsxEditor.Viewport>
        </XlsxEditor.Root>
      </div>
    );
    await waitFor(() => expect(view.getAllByRole('tab').length).toBe(2), { timeout: 5000 });
    const rows: number[] = [];
    const off = viewer!.subscribe(() => rows.push(viewer!.snapshot().selection!.active.row));
    const grid = view.container.querySelector('.xlsx-scroller') as HTMLElement;
    const errors: unknown[] = [];
    const error = console.error;
    console.error = (...args: unknown[]) => void errors.push(args);
    try {
      act(() => {
        for (let i = 0; i < 500; i += 1) fireEvent.keyDown(grid, { key: 'ArrowDown' });
      });
    } finally {
      console.error = error;
      off();
    }
    expect(errors).toEqual([]);
    expect(rows.length).toBeGreaterThanOrEqual(500);
    expect(rows.every((r, i) => i === 0 || r >= rows[i - 1]!)).toBe(true);
    expect(seen.at(-1)).toBe(viewer!.snapshot().selection!.active.row);
    expect(seen.at(-1)).toBe(500);
  });

  test('the chrome parts render only when their own slice changes', async () => {
    const renders: Record<string, number> = {};
    const count = (id: string) => (renders[id] = (renders[id] ?? 0) + 1);
    let viewer: XlsxEditorInstance | null = null;
    const view = render(
      <div style={{ height: 400 }}>
        <XlsxEditor.Root document={workbook} fonts={FONTS} onReady={(v) => (viewer = v)}>
          <XlsxEditor.Viewport>
            <Profiler id="zoom" onRender={count}>
              <XlsxEditor.ZoomControl />
            </Profiler>
            <Profiler id="tabs" onRender={count}>
              <XlsxEditor.SheetTabs />
            </Profiler>
            <Profiler id="status" onRender={count}>
              <XlsxEditor.StatusMessage />
            </Profiler>
          </XlsxEditor.Viewport>
        </XlsxEditor.Root>
      </div>
    );
    await waitFor(() => expect(view.getAllByRole('tab').length).toBe(2), { timeout: 5000 });
    const before = { ...renders };
    act(() => viewer!.select({ row: 3, col: 3 }));
    expect(renders).toEqual(before);
    act(() => viewer!.setZoom(1.5));
    expect(renders).toEqual({ ...before, zoom: before.zoom! + 1 });
  });

  test('useEditorState reads the new slice when the selector changes', async () => {
    const seen: unknown[] = [];
    const pickers: Record<string, (s: EditorSnapshot) => unknown> = {
      status: (s) => s.status,
      sheet: (s) => s.activeSheet,
    };
    function Probe({ pick }: { pick: string }) {
      seen.push(useEditorState(pickers[pick]!));
      return null;
    }
    const tree = (pick: string) => (
      <XlsxEditor.Root document={workbook} fonts={FONTS}>
        <XlsxEditor.Viewport>
          <XlsxEditor.SheetTabs />
          <Probe pick={pick} />
        </XlsxEditor.Viewport>
      </XlsxEditor.Root>
    );
    const view = render(tree('status'));
    await waitFor(() => expect(view.getAllByRole('tab').length).toBe(2), { timeout: 5000 });
    expect(seen.at(-1)).toBe('ready');
    view.rerender(tree('sheet'));
    expect(seen.at(-1)).toBe(0);
  });

  test('the arrow keys pass over a chart sheet tab', async () => {
    const { viewer } = await open();
    const sheets = [
      { index: 0, name: 'One', visible: true, worksheet: true },
      { index: 1, name: 'Chart', visible: true, worksheet: false },
      { index: 2, name: 'Two', visible: true, worksheet: true },
    ];
    let active = 0;
    const listeners = new Set<() => void>();
    let snap = { ...viewer.snapshot(), sheets, activeSheet: active };
    let refuse = false;
    const fake = {
      ...viewer,
      snapshot: () => snap,
      subscribe: (f: () => void) => (listeners.add(f), () => void listeners.delete(f)),
      setActiveSheet: (i: number) => {
        if (!sheets[i]!.worksheet || refuse) return false;
        active = i;
        snap = { ...snap, activeSheet: active };
        for (const f of listeners) f();
        return true;
      },
    } as unknown as XlsxEditorInstance;
    const view = render(
      <XlsxEditorContext.Provider value={fake}>
        <div data-probe="tabs">
          <XlsxEditor.SheetTabs />
        </div>
      </XlsxEditorContext.Provider>
    );
    const tabs = () =>
      view.container.querySelectorAll<HTMLButtonElement>('[data-probe="tabs"] [role="tab"]');
    expect(tabs().length).toBe(3);
    act(() => void fireEvent.keyDown(tabs()[0]!, { key: 'ArrowRight' }));
    expect(active).toBe(2);
    act(() => void fireEvent.keyDown(tabs()[2]!, { key: 'ArrowLeft' }));
    expect(active).toBe(0);
    // A refused sheet change keeps the focus on the selected tab.
    act(() => tabs()[0]!.focus());
    refuse = true;
    act(() => void fireEvent.keyDown(tabs()[0]!, { key: 'ArrowRight' }));
    expect(active).toBe(0);
    expect(document.activeElement).toBe(tabs()[0]!);
  });

  test('a file that is not a workbook shows an alert', async () => {
    const view = render(<XlsxEditor document={new Uint8Array([1, 2, 3])} fonts={FONTS} />);
    await waitFor(() => expect(view.getByRole('alert').textContent).toContain('not a workbook'), {
      timeout: 5000,
    });
  });

  test('a font load failure fires onFontError once and shows the fonts alert', async () => {
    const errors: string[] = [];
    const view = render(
      <XlsxEditor
        document={workbook}
        fonts={() => Promise.reject(new Error('offline'))}
        onFontError={(f) => errors.push(f.code)}
      />
    );
    await waitFor(() => expect(view.getByRole('alert').textContent).toContain('fonts'), {
      timeout: 5000,
    });
    expect(errors).toEqual(['fonts']);
  });

  test('a new fonts or numberLocale value applies in place, and the workbook stays', async () => {
    const bytes = buildXlsx({ sheets: [{ name: 'S', rows: row(1, [c('A1', 1.5)]) }] });
    let viewer: XlsxEditorInstance | null = null;
    const calls: string[] = [];
    const ui = (fonts: FontSource[], numberLocale?: { decimal: string; group: string }) => (
      <div style={{ height: 400 }}>
        <XlsxEditor
          document={bytes}
          fonts={fonts}
          {...(numberLocale ? { numberLocale } : {})}
          onReady={(v) => {
            viewer = v;
            const setFonts = v.setFonts.bind(v);
            v.setFonts = (f) => (calls.push('setFonts'), setFonts(f));
          }}
        />
      </div>
    );
    const view = render(ui(FONTS));
    await waitFor(() => expect(viewer!.snapshot().status).toBe('ready'), { timeout: 5000 });
    act(() => viewer!.select({ row: 0, col: 0 }));
    expect(field(view, 'Cell content').value).toBe('1.5');
    // The same faces again, in a new list, load nothing.
    view.rerender(ui([...FONTS]));
    expect(calls).toEqual([]);
    const copies = FONTS.map((f) => ({ ...f, bytes: new Uint8Array(f.bytes) }));
    view.rerender(ui(copies, { decimal: ',', group: '.' }));
    expect(calls).toEqual(['setFonts']);
    expect(field(view, 'Cell content').value).toBe('1,5');
    expect(viewer!.snapshot().status).toBe('ready');
    expect(viewer!.snapshot().selectionRef).toBe('A1');
  });

  test('onReady fires after the first open settles, and a select from it sticks', async () => {
    const bytes = buildXlsx({ sheets: [{ name: 'S', rows: row(1, [c('A1', 1), c('B2', 2)]) }] });
    const seen: string[] = [];
    let viewer: XlsxEditorInstance | null = null;
    const ui = (document?: Uint8Array) => (
      <div style={{ height: 300 }}>
        <XlsxEditor
          {...(document ? { document } : {})}
          fonts={FONTS}
          onReady={(v) => {
            viewer = v;
            seen.push(v.snapshot().status);
            v.select({ row: 1, col: 1 });
          }}
        />
      </div>
    );
    const view = render(ui(bytes));
    await waitFor(() => expect(seen).toEqual(['ready']), { timeout: 5000 });
    expect(viewer!.snapshot().selectionRef).toBe('B2');
    // A change of `document` to undefined closes the workbook; onReady does not fire again.
    view.rerender(ui());
    expect(viewer!.snapshot().status).toBe('empty');
    expect(viewer!.snapshot().sheets).toEqual([]);
    expect(viewer!.select({ row: 0, col: 0 })).toBe(false);
    expect(seen).toEqual(['ready']);
    view.unmount();
    // Without a document, onReady fires right after the attach.
    render(ui());
    await waitFor(() => expect(seen).toEqual(['ready', 'empty']));
  });

  test('an inline numberLocale with a state-setting onChange does not loop', async () => {
    const bytes = buildXlsx({ sheets: [{ name: 'S', rows: row(1, [c('A1', 1.5)]) }] });
    let renders = 0;
    let viewer: XlsxEditorInstance | null = null;
    function Host() {
      const [, setSnap] = useState<EditorSnapshot | null>(null);
      renders += 1;
      return (
        <div style={{ height: 300 }}>
          <XlsxEditor
            document={bytes}
            fonts={[...FONTS]}
            numberLocale={{ decimal: ',', group: '.', builtins: { 2: '0,00' } }}
            onReady={(v) => (viewer = v)}
            onChange={(s) => setSnap(s)}
          />
        </div>
      );
    }
    const errors: unknown[] = [];
    const original = console.error;
    console.error = (...a: unknown[]) => void errors.push(a[0]);
    try {
      render(<Host />);
      await waitFor(() => expect(viewer?.snapshot().status).toBe('ready'), { timeout: 5000 });
      const settled = renders;
      await act(() => new Promise((r) => setTimeout(r, 200)));
      expect(renders - settled).toBeLessThan(3);
      expect(errors).toEqual([]);
    } finally {
      console.error = original;
    }
  });

  test('colorMode puts the dark class on the viewport, and system follows the OS', () => {
    const viewport = (view: ReturnType<typeof render>) =>
      view.container.querySelector('.xlsx-editor') as HTMLElement;
    const light = render(<XlsxEditor fonts={FONTS} />);
    expect(viewport(light).classList.contains('dark')).toBe(false);
    light.unmount();
    const dark = render(<XlsxEditor fonts={FONTS} colorMode="dark" />);
    expect(viewport(dark).classList.contains('dark')).toBe(true);
    dark.rerender(<XlsxEditor fonts={FONTS} colorMode="light" />);
    expect(viewport(dark).classList.contains('dark')).toBe(false);
    dark.unmount();
    const original = window.matchMedia;
    window.matchMedia = ((query: string) => ({
      matches: query === '(prefers-color-scheme: dark)',
      media: query,
      addEventListener: () => {},
      removeEventListener: () => {},
    })) as unknown as typeof window.matchMedia;
    try {
      const seen: string[] = [];
      function Probe() {
        seen.push(useEditorColorMode().colorMode);
        return null;
      }
      const system = render(
        <XlsxEditor.Root fonts={FONTS} colorMode="system">
          <XlsxEditor.Viewport>
            <Probe />
          </XlsxEditor.Viewport>
        </XlsxEditor.Root>
      );
      expect(viewport(system).classList.contains('dark')).toBe(true);
      expect(seen.at(-1)).toBe('dark');
      system.rerender(
        <XlsxEditor.Root fonts={FONTS} colorMode="light">
          <XlsxEditor.Viewport>
            <Probe />
          </XlsxEditor.Viewport>
        </XlsxEditor.Root>
      );
      expect(seen.at(-1)).toBe('light');
    } finally {
      window.matchMedia = original;
    }
  });

  test('useScopeClassName follows colorMode, and a portaled host popup resolves the tokens', () => {
    // happy-dom computes a custom property that a rule declares on the element itself, so the
    // test reads `--doc-bg` from the computed style of the popup and of the viewport.
    const sheet = document.createElement('style');
    sheet.textContent = readFileSync(
      join(import.meta.dir, '../../../core/src/styles/editor.css'),
      'utf8'
    );
    document.head.append(sheet);
    try {
      const inside: string[] = [];
      function Inside() {
        inside.push(useScopeClassName());
        return null;
      }
      function HostPopup() {
        const scope = useScopeClassName();
        return createPortal(<div className={`${scope}acme-popup`} data-popup="" />, document.body);
      }
      const tree = (colorMode: 'light' | 'dark') => (
        <XlsxEditor.Root fonts={FONTS} colorMode={colorMode}>
          <XlsxEditor.Viewport>
            <Inside />
          </XlsxEditor.Viewport>
          <HostPopup />
        </XlsxEditor.Root>
      );
      const view = render(tree('light'));
      const popup = () => document.body.querySelector('[data-popup]') as HTMLElement;
      const viewport = () => view.container.querySelector('.xlsx-editor') as HTMLElement;
      const token = (el: HTMLElement) => getComputedStyle(el).getPropertyValue('--doc-bg');
      expect(inside.at(-1)).toBe('');
      expect(popup().parentElement).toBe(document.body);
      expect(popup().className).toBe('xlsx-editor acme-popup');
      expect(token(popup())).not.toBe('');
      expect(token(popup())).toBe(token(viewport()));
      const lightBg = token(popup());

      view.rerender(tree('dark'));
      expect(inside.at(-1)).toBe('');
      expect(popup().className).toBe('xlsx-editor dark acme-popup');
      expect(token(popup())).toBe(token(viewport()));
      expect(token(popup())).not.toBe(lightBg);
      expect(getComputedStyle(popup()).getPropertyValue('color-scheme')).toBe('dark');
    } finally {
      sheet.remove();
    }
  });

  test('the inline-size container is on the viewport class, not on the scope class', () => {
    // The viewport keeps its width from following its content, so a content-sized host cannot
    // grow it without limit. A host popup with the scope class keeps its content width.
    // happy-dom does not compute `container-type`, so the test reads the stylesheet.
    const css = readFileSync(
      join(import.meta.dir, '../../../core/src/styles/editor.css'),
      'utf8'
    ).replace(/\/\*[\s\S]*?\*\//g, '');
    const rules = (end: RegExp) =>
      [...css.matchAll(/([^{}]+)\{([^{}]*)\}/g)].filter(([, selector]) =>
        selector!.split(',').some((part) => end.test(part.trim()))
      );
    const scopeRules = rules(/\.xlsx-editor(\.dark)?$/);
    expect(scopeRules.length).toBeGreaterThan(0);
    for (const [, , body] of scopeRules) expect(body).not.toMatch(/(^|[;\s])container(-type)?\s*:/);
    const viewportRules = rules(/^\.xlsx-editor-viewport$/);
    expect(viewportRules.map(([, , body]) => body!.trim())).toEqual([
      'container: xlsx-editor / inline-size;',
    ]);
    const view = render(
      <XlsxEditor.Root fonts={FONTS}>
        <XlsxEditor.Viewport className="host" />
      </XlsxEditor.Root>
    );
    expect(view.container.querySelector('.xlsx-editor')!.className).toBe(
      'xlsx-editor xlsx-editor-viewport host'
    );
  });

  test('outside a root, useScopeClassName is the light scope', () => {
    const seen: string[] = [];
    function Probe() {
      seen.push(useScopeClassName());
      return null;
    }
    render(<Probe />);
    expect(seen).toEqual(['xlsx-editor ']);
  });

  test('a viewport with label false renders no region; a label names it', () => {
    const tree = (label?: string | false) => (
      <XlsxEditor.Root fonts={FONTS}>
        <XlsxEditor.Viewport label={label} />
      </XlsxEditor.Root>
    );
    const view = render(tree());
    expect(view.getByRole('region', { name: 'Workbook editor' })).toBeTruthy();
    view.rerender(tree(false));
    expect(view.container.querySelector('.xlsx-editor')!.hasAttribute('role')).toBe(false);
    expect(view.queryByRole('region')).toBeNull();
    view.rerender(tree('Budget'));
    expect(view.getByRole('region', { name: 'Budget' })).toBeTruthy();
  });

  test('a locale catalog replaces strings and falls back to English', async () => {
    const view = render(
      <XlsxEditor fonts={FONTS} locale={{ status: { empty: 'Leer', loading: null } }} />
    );
    await waitFor(() => expect(view.getByRole('status').textContent).toBe('Leer'));
  });
});
