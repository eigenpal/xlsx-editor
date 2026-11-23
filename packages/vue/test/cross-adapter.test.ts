// The React and the Vue editors render the same chrome for one workbook, after PPTX Editor's
// cross-adapter test (pptx-editor/packages/vue/test/cross-adapter.test.ts at d4502552), which
// follows DOCX Editor's cross-adapter differential test. Each case opens the same bytes in the
// composed editor of each adapter, brings both editors to the same state through the engine,
// and compares the canonical markup of everything but the painted grid: the toolbar, the
// formula bar, the sheet tabs, the zoom control, and the status line. A control, a label, a
// state, or a reason that one adapter renders and the other does not fails here.
//
// Differences from PPTX Editor: the case renders the whole composed editor, not one toolbar,
// because the XLSX chrome is small enough to compare as one tree.

import { afterEach, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { createElement } from 'react';
import { act, cleanup, fireEvent, render, waitFor, within } from '@testing-library/react';
import { createApp, h, nextTick, type App } from 'vue';
import type { Translations } from '@xlsx-editor.dev/i18n';
import { buildXlsx, c, row } from '../../../scripts/fixtures/builder.ts';
import { XlsxEditor as ReactXlsxEditor } from '../../react/src/index.ts';
import {
  XlsxEditor,
  type EditorMode,
  type FontSource,
  type XlsxEditorInstance,
} from '../src/index.ts';
import { canonicalDom } from './canonical-dom.ts';

const fontsDir = join(
  dirname(
    Bun.resolveSync('@docx-editor.dev/fonts/package.json', join(import.meta.dir, '../../core'))
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

const workbook = buildXlsx({
  sheets: [
    { name: 'First', rows: row(1, [c('A1', 'one')]) + row(5, [c('A5', 84, { f: 'A2*2' })]) },
    { name: 'Hidden', state: 'hidden', rows: row(1, [c('A1', 'secret')]) },
    { name: 'Second', rows: row(1, [c('A1', 'two')]) },
  ],
});

let vueApp: App | null = null;
let vueHost: HTMLElement | null = null;
afterEach(() => {
  cleanup();
  vueApp?.unmount();
  vueHost?.remove();
  vueApp = null;
  vueHost = null;
});

interface Options {
  readonly mode?: EditorMode;
  readonly locale?: Translations;
}

const tabsIn = (host: Element): number => host.querySelectorAll('[role="tab"]').length;

/** The composed React editor on `workbook`, once its tabs show. */
async function openReact(options: Options = {}) {
  let editor: XlsxEditorInstance | null = null;
  const view = render(
    createElement(ReactXlsxEditor, {
      document: workbook,
      fonts: FONTS,
      onReady: (ready: XlsxEditorInstance) => (editor = ready),
      ...options,
    })
  );
  await waitFor(() => expect(tabsIn(view.container)).toBe(2), { timeout: 5000 });
  return { host: view.container, editor: editor! };
}

/** The composed Vue editor on `workbook`, once its tabs show. */
async function openVue(options: Options = {}) {
  let editor: XlsxEditorInstance | null = null;
  const host = document.createElement('div');
  document.body.append(host);
  vueHost = host;
  vueApp = createApp({
    render: () =>
      h(XlsxEditor, {
        document: workbook,
        fonts: FONTS,
        onReady: (ready: XlsxEditorInstance) => (editor = ready),
        ...options,
      }),
  });
  vueApp.mount(host);
  await waitFor(() => expect(tabsIn(host)).toBe(2), { timeout: 5000 });
  return { host, editor: editor! };
}

/** The chrome markup, one element on each line, with the painted grid left out. */
const markup = (host: Element): string =>
  canonicalDom(host, {
    trimText: true,
    formState: true,
    splitIdReferences: true,
    normalizeStyle: true,
    dropEmptyClass: true,
    dropSelectValue: true,
    // React writes a controlled field's value as an attribute too; `formState` reads the value.
    // An option keeps its `value`.
    dropFieldValue: true,
    // React also writes a controlled textarea's first value as its text; `content` reads it.
    skipChildren: (el) => el.classList.contains('xlsx-scroller') || el.tagName === 'TEXTAREA',
  })
    // React writes a zero length with no unit; `0` and `0px` are the same CSS length.
    .replace(/([:;])0px(?=[;"])/g, '$10')
    .replace(/></g, '>\n<');

/** The text of the formula bar's content field, which `formState` does not read. */
const content = (host: Element): string | undefined =>
  host.querySelector<HTMLTextAreaElement>('textarea[aria-label="Cell content"]')?.value;

/** Bring both editors to one state through the engine, and let each adapter render it. */
async function both(
  react: XlsxEditorInstance,
  vue: XlsxEditorInstance,
  step: (editor: XlsxEditorInstance) => void
): Promise<void> {
  act(() => step(react));
  step(vue);
  await nextTick();
}

/** What has the keyboard: its tag, role, and accessible name or text. */
const focused = (): string => {
  const el = document.activeElement;
  if (!el || el === document.body) return 'body';
  const name = el.getAttribute('aria-label') ?? el.textContent?.trim() ?? '';
  return `${el.tagName.toLowerCase()} ${el.getAttribute('role') ?? ''} ${name}`;
};

/**
 * Run each user action in turn, and give the markup and the focused element after each. The
 * two editors run their actions one after the other, never side by side, because an open modal
 * dialog of one editor would take the keyboard from the other.
 */
async function trace(
  host: HTMLElement,
  actions: readonly ((host: HTMLElement) => void)[],
  settle: () => Promise<void>
): Promise<{ markup: string; focus: string }[]> {
  const out: { markup: string; focus: string }[] = [];
  for (const action of actions) {
    action(host);
    await settle();
    out.push({ markup: markup(host), focus: focused() });
  }
  return out;
}

describe('the React and Vue editors render the same chrome', () => {
  test('after the open', async () => {
    const react = await openReact();
    const vue = await openVue();
    expect(markup(vue.host)).toBe(markup(react.host));
  });

  test('for a formula cell, a sheet switch, and a bold cell', async () => {
    const react = await openReact();
    const vue = await openVue();
    await both(react.editor, vue.editor, (e) => void e.select({ row: 4, col: 0 }));
    expect(content(vue.host)).toBe('=A2*2');
    expect(content(vue.host)).toBe(content(react.host));
    expect(markup(vue.host)).toBe(markup(react.host));
    await both(react.editor, vue.editor, (e) => void e.setActiveSheet(2));
    expect(markup(vue.host)).toBe(markup(react.host));
    await both(react.editor, vue.editor, (e) => void e.toggleFormat('bold'));
    expect(vue.host.querySelector('[data-slot="font.bold"][data-active]')).not.toBeNull();
    expect(markup(vue.host)).toBe(markup(react.host));
  });

  test('in view mode', async () => {
    const react = await openReact({ mode: 'view' });
    const vue = await openVue({ mode: 'view' });
    expect(markup(vue.host)).toBe(markup(react.host));
  });

  test('with a partial locale', async () => {
    const locale = { formattingBar: { label: 'Formatierung' } } as Translations;
    const react = await openReact({ locale });
    const vue = await openVue({ locale });
    expect(vue.host.querySelector('[role="toolbar"]')?.getAttribute('aria-label')).toBe(
      'Formatierung'
    );
    expect(markup(vue.host)).toBe(markup(react.host));
  });

  test('with each popup of the toolbar open', async () => {
    const button = (host: HTMLElement, name: string | RegExp) =>
      fireEvent.click(within(host).getByRole('button', { name }));
    // A press on the grid closes the open popup, as a press outside does.
    const outside = (host: HTMLElement) =>
      fireEvent.pointerDown(host.querySelector('.xlsx-scroller')!);
    const actions: ((host: HTMLElement) => void)[] = [
      (host) => button(host, /^Select font family/),
      outside,
      (host) => button(host, /^Font size/),
      outside,
      (host) => button(host, 'Font color'),
      outside,
      (host) => button(host, 'Fill color'),
      outside,
      (host) => button(host, 'Insert and delete'),
      outside,
    ];
    const react = await trace((await openReact()).host, actions, async () => undefined);
    cleanup();
    const vue = await trace((await openVue()).host, actions, () => nextTick());
    // Each even step opens its popup, and each odd step closes it. A boolean keeps a failure
    // message short.
    const opened = (i: number, label: string, role: string) =>
      new RegExp(` aria-label="${label}"[^>]* role="${role}"`).test(vue[i]!.markup);
    expect(opened(0, 'Select font family', 'listbox')).toBe(true);
    expect(vue[0]!.markup.includes('role="option"')).toBe(true);
    expect(opened(2, 'Font size', 'listbox')).toBe(true);
    expect(opened(4, 'Font color', 'listbox')).toBe(true);
    expect(opened(6, 'Fill color', 'listbox')).toBe(true);
    expect(opened(8, 'Insert and delete', 'menu')).toBe(true);
    for (const i of [1, 3, 5, 7, 9])
      expect([i, /role="(listbox|menu)"/.test(vue[i]!.markup)]).toEqual([i, false]);
    for (const [i, step] of vue.entries()) {
      expect(step.markup).toBe(react[i]!.markup);
      expect(step.focus).toBe(react[i]!.focus);
    }
  });

  test('with the number format list, its dialog, the tab menu, and the rename field open', async () => {
    const actions: ((host: HTMLElement) => void)[] = [
      (host) => fireEvent.click(within(host).getByRole('button', { name: /^Number format/ })),
      (host) => fireEvent.click(within(host).getByRole('button', { name: 'More number formats…' })),
      (host) => fireEvent.click(within(host).getByRole('button', { name: 'Cancel' })),
      (host) => fireEvent.contextMenu(within(host).getAllByRole('tab')[1]!),
      (host) => fireEvent.keyDown(within(host).getByRole('menu'), { key: 'Escape' }),
      (host) => fireEvent.doubleClick(within(host).getAllByRole('tab')[0]!),
    ];
    const react = await trace((await openReact()).host, actions, async () => undefined);
    cleanup();
    const vue = await trace((await openVue()).host, actions, () => nextTick());
    expect(vue[1]!.markup).toContain('<dialog ');
    expect(vue[3]!.markup).toContain('role="menu"');
    expect(vue[5]!.focus).toBe('input  Sheet name');
    for (const [i, step] of vue.entries()) {
      expect(step.markup).toBe(react[i]!.markup);
      expect(step.focus).toBe(react[i]!.focus);
    }
  });
});
