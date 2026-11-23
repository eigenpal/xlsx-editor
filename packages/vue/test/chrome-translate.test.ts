// The language of the chrome: the Vue twin of the React adapter's `chrome-translate.test.tsx`.
// `useChromeTranslate` with overrides, the `t` prop of each chrome part, `LocaleProvider` in a
// region and around the root, a runtime change of the strings, and the direction of a
// right-to-left catalog. Vue renders in the next microtask, so a step waits for `nextTick`.

import { afterEach, beforeEach, describe, expect, test } from 'bun:test';
import { fireEvent } from '@testing-library/react';
import { defineComponent, h, nextTick, shallowRef, type VNodeChild } from 'vue';
import { buildXlsx, c, row } from '../../../scripts/fixtures/builder.ts';
import {
  FormulaBar,
  LocaleProvider,
  SheetTabs,
  StatusMessage,
  XlsxEditor,
  XlsxEditorContent,
  XlsxEditorRoot,
  XlsxEditorToolbar,
  XlsxEditorViewport,
  ZoomControl,
  useChromeDir,
  useChromeTranslate,
  useTranslation,
  type ChromeTranslate,
  type Translations,
  type XlsxEditorInstance,
} from '../src/index.ts';
import { FONTS, mount, tabsShow, unmountAll, type Mounted } from './mount.ts';
import { resetPartTranslateWarnings } from '../src/locale.ts';

afterEach(unmountAll);

const workbook = buildXlsx({
  sheets: [
    { name: 'First', rows: row(1, [c('A1', 1234.5)]) },
    { name: 'Second', rows: row(1, [c('A1', 'two')]) },
  ],
});

/** A resolver for the keys it holds; every other key is left to the inherited translator. */
const only =
  (labels: Record<string, string>): ChromeTranslate =>
  (key, params) => {
    const text = Object.hasOwn(labels, key) ? labels[key]! : undefined;
    return text?.replace(/\{(\w+)\}/g, (_, name: string) => String(params?.[name]));
  };

interface PartsOptions {
  readonly locale?: Translations;
  readonly t?: ChromeTranslate;
  readonly children?: () => VNodeChild;
  readonly onReady?: (editor: XlsxEditorInstance) => void;
}

/** A composed layout with the same `t` on every part. */
function parts({ locale, t, children, onReady }: PartsOptions = {}): VNodeChild {
  const part = t ? { t } : {};
  return h(
    XlsxEditorRoot,
    {
      document: workbook,
      fonts: FONTS,
      ...(locale ? { locale } : {}),
      ...(onReady ? { onReady } : {}),
    },
    () => [
      h(XlsxEditorViewport, null, () => [
        h(XlsxEditorToolbar, part),
        h(FormulaBar, part),
        h(StatusMessage, part),
        h(XlsxEditorContent),
        h('div', { class: 'xlsx-footer' }, [h(SheetTabs, part), h(ZoomControl, part)]),
        children?.(),
      ]),
    ]
  );
}

/** A component that runs `read` in its setup and renders nothing. */
const probe = (read: () => void) =>
  h(
    defineComponent({
      setup() {
        read();
        return () => null;
      },
    })
  );

describe('useChromeTranslate', () => {
  test('an override wins with its placeholders filled, and other keys read the root locale', () => {
    const seen: string[] = [];
    const overrides = new Map([['zoom.resetAt', 'Zoom {level}, press to reset']]);
    mount(() =>
      h(XlsxEditorRoot, { locale: { zoom: { zoomIn: 'Vergrößern' } } }, () => [
        probe(() => {
          const t = useChromeTranslate(overrides).value;
          seen.push(t('zoom.resetAt', { level: '80%' }), t('zoom.zoomIn'), t('x.y'));
        }),
      ])
    );
    expect(seen).toEqual(['Zoom 80%, press to reset', 'Vergrößern', 'x.y']);
  });

  test('an override picks its plural branch with the plural rules of the root locale', () => {
    const seen: string[] = [];
    const overrides = new Map([['formulaBar.lines', '{count, plural, few {# (few)} other {#}}']]);
    mount(() =>
      h(XlsxEditorRoot, { locale: { _lang: 'pl' } }, () => [
        probe(() => seen.push(useChromeTranslate(overrides).value('formulaBar.lines', { count: 3 }))),
      ])
    );
    expect(seen).toEqual(['3 (few)']);
  });

  test('a ref of overrides is followed', async () => {
    const overrides = shallowRef<ReadonlyMap<string, string> | undefined>(undefined);
    let t: ReturnType<typeof useChromeTranslate> | null = null;
    mount(() => h(XlsxEditorRoot, null, () => [probe(() => (t = useChromeTranslate(overrides)))]));
    expect(t!.value('zoom.zoomIn')).toBe('Zoom in');
    overrides.value = new Map([['zoom.zoomIn', 'Larger']]);
    await nextTick();
    expect(t!.value('zoom.zoomIn')).toBe('Larger');
  });

  test('passed as a part t, it relabels that part and leaves the others', async () => {
    const overrides = new Map([['zoom.zoomIn', 'Larger']]);
    const Footer = defineComponent({
      setup() {
        const t = useChromeTranslate(overrides);
        return () => h(ZoomControl, { t: t.value });
      },
    });
    const view = mount(() =>
      h(XlsxEditorRoot, { document: workbook, fonts: FONTS }, () => [
        h(XlsxEditorViewport, null, () => [
          h(XlsxEditorToolbar),
          h(XlsxEditorContent),
          h(SheetTabs),
          h(Footer),
        ]),
      ])
    );
    await tabsShow(view);
    expect(view.getByRole('button', { name: 'Larger' })).toBeTruthy();
    expect(view.getByRole('button', { name: 'Zoom out' })).toBeTruthy();
    expect(view.getByRole('toolbar').getAttribute('aria-label')).toBe('Formatting controls');
  });

  test('inside a part with t, it reads that t first', () => {
    const seen: string[] = [];
    mount(() =>
      h(XlsxEditorRoot, null, () => [
        h(XlsxEditorViewport, null, () => [
          h(
            XlsxEditorToolbar,
            { preset: false, t: only({ 'formattingBar.bold': 'Heavy' }) },
            () => [probe(() => seen.push(useChromeTranslate().value('formattingBar.bold')))]
          ),
        ]),
      ])
    );
    expect(seen).toEqual(['Heavy']);
  });

  test('outside a root, it reads English', () => {
    const seen: string[] = [];
    mount(() => probe(() => seen.push(useChromeTranslate().value('zoom.zoomIn'))));
    expect(seen).toEqual(['Zoom in']);
  });
});

describe('the t prop of each part', () => {
  test('relabels the toolbar, the formula bar, the status message, the sheet tabs, and the zoom', async () => {
    const t = only({
      'formattingBar.label': 'Bar',
      'formulaBar.nameBox': 'Address',
      'sheetTabs.add': 'New sheet',
      'zoom.resetAt': 'Reset from {level}',
    });
    const view = mount(() => parts({ t }));
    await tabsShow(view);
    expect(view.getByRole('toolbar').getAttribute('aria-label')).toBe('Bar');
    expect(view.getByRole('textbox', { name: 'Address' })).toBeTruthy();
    expect(view.getByRole('button', { name: 'New sheet' })).toBeTruthy();
    expect(view.getByRole('button', { name: 'Reset from 100%' })).toBeTruthy();
    // A key that t leaves unresolved reads the root's locale, never the key.
    expect(view.getByRole('button', { name: 'Zoom in' })).toBeTruthy();
    expect(view.getByRole('textbox', { name: 'Cell content' })).toBeTruthy();
  });

  test('relabels the status message', () => {
    const view = mount(() =>
      h(XlsxEditorRoot, null, () => [
        h(XlsxEditorViewport, null, () => [
          h(StatusMessage, { t: only({ 'status.empty': 'Nothing open' }) }),
        ]),
      ])
    );
    expect(view.container.textContent).toContain('Nothing open');
  });

  test('reaches the parts inside the toolbar and its number format dialog', async () => {
    const t = only({
      'formattingBar.bold': 'Heavy',
      'numberFormat.more': 'Other formats',
      'dialogs.numberFormat.title': 'Format of numbers',
    });
    const view = mount(() =>
      h(XlsxEditorRoot, { document: workbook, fonts: FONTS }, () => [
        h(XlsxEditorViewport, null, () => [
          h(XlsxEditorToolbar, { t }),
          h(XlsxEditorContent),
          h(SheetTabs),
        ]),
      ])
    );
    await tabsShow(view);
    expect(
      view.container.querySelector('[data-slot="font.bold"]')?.getAttribute('aria-label')
    ).toBe('Heavy');
    fireEvent.click(view.getByRole('button', { name: /^Number format/ }));
    await nextTick();
    fireEvent.click(view.getByRole('button', { name: 'Other formats' }));
    await nextTick();
    expect(view.getByRole('dialog', { name: 'Format of numbers' })).toBeTruthy();
  });

  test('a change of t relabels the part and keeps it mounted', async () => {
    const t = shallowRef<ChromeTranslate | undefined>(undefined);
    const view = mount(() => parts(t.value ? { t: t.value } : {}));
    await tabsShow(view);
    const bar = view.getByRole('toolbar');
    expect(bar.getAttribute('aria-label')).toBe('Formatting controls');
    t.value = only({ 'formattingBar.label': 'Bar' });
    await nextTick();
    expect(view.getByRole('toolbar')).toBe(bar);
    expect(bar.getAttribute('aria-label')).toBe('Bar');
    t.value = undefined;
    await nextTick();
    expect(view.getByRole('toolbar')).toBe(bar);
    expect(bar.getAttribute('aria-label')).toBe('Formatting controls');
  });
});

describe('LocaleProvider', () => {
  test('relabels one region and inherits the rest', async () => {
    const view = mount(() =>
      parts({
        locale: { zoom: { zoomOut: 'Kleiner' } },
        children: () =>
          h(LocaleProvider, { i18n: { zoom: { zoomIn: 'Größer' } } }, () => [
            h('div', { class: 'region' }, [h(ZoomControl)]),
          ]),
      })
    );
    await tabsShow(view);
    const region = view.container.querySelector('.region')!;
    expect(region.querySelector('[aria-label="Größer"]')).toBeTruthy();
    expect(region.querySelector('[aria-label="Kleiner"]')).toBeTruthy();
    const footer = view.container.querySelector('.xlsx-footer')!;
    expect(footer.querySelector('[aria-label="Zoom in"]')).toBeTruthy();
  });

  test('around the root, the root locale merges onto it, and host controls outside read it', async () => {
    const seen: string[] = [];
    const view = mount(() =>
      h(LocaleProvider, { i18n: { _lang: 'de', zoom: { zoomIn: 'Größer', zoomOut: 'Kleiner' } } }, () => [
        probe(() => seen.push(useTranslation().t('zoom.zoomIn'))),
        parts({ locale: { zoom: { zoomOut: 'Verkleinern' } } }),
      ])
    );
    await tabsShow(view);
    expect(seen).toEqual(['Größer']);
    expect(view.getByRole('button', { name: 'Größer' })).toBeTruthy();
    expect(view.getByRole('button', { name: 'Verkleinern' })).toBeTruthy();
  });

  test('with no i18n, it keeps the t of the part above it', () => {
    const view = mount(() =>
      h(XlsxEditorRoot, null, () => [
        h(XlsxEditorViewport, null, () => [
          h(
            XlsxEditorToolbar,
            { preset: false, t: only({ 'formattingBar.bold': 'Heavy' }) },
            () => [h(LocaleProvider, null, () => [h(XlsxEditorToolbar.Bold)])]
          ),
        ]),
      ])
    );
    expect(
      view.container.querySelector('[data-slot="font.bold"]')?.getAttribute('aria-label')
    ).toBe('Heavy');
  });
});

describe('useTranslation', () => {
  test('returns the catalogue, the language, and the direction of the nearest catalog', () => {
    const seen: { lang: string; dir: string; zoomIn: string; t: string }[] = [];
    const read = () => {
      const { t, catalogue, lang, dir } = useTranslation();
      seen.push({ lang, dir, zoomIn: catalogue.value.zoom.zoomIn, t: t('zoom.zoomIn') });
    };
    mount(() => [
      probe(read),
      h(XlsxEditorRoot, { locale: { _lang: 'he', zoom: { zoomIn: 'הגדל' } } }, () => [
        probe(read),
        h(LocaleProvider, { i18n: { _lang: 'de', zoom: { zoomIn: 'Größer' } }, dir: 'ltr' }, () => [
          probe(read),
        ]),
      ]),
    ]);
    expect(seen).toEqual([
      { lang: 'en', dir: 'ltr', zoomIn: 'Zoom in', t: 'Zoom in' },
      { lang: 'he', dir: 'rtl', zoomIn: 'הגדל', t: 'הגדל' },
      { lang: 'de', dir: 'ltr', zoomIn: 'Größer', t: 'Größer' },
    ]);
  });
});

describe('a t that throws', () => {
  beforeEach(resetPartTranslateWarnings);

  test('keeps the editor mounted, shows the catalog text, and warns once in development', async () => {
    const warnings: string[] = [];
    const warn = console.warn;
    console.warn = (line: string) => void warnings.push(line);
    try {
      const t: ChromeTranslate = (key) => {
        if (key === 'zoom.zoomIn') throw new Error('broken');
        return undefined;
      };
      const current = shallowRef<ChromeTranslate>(t);
      const view = mount(() => parts({ t: current.value }));
      await tabsShow(view);
      expect(view.getByRole('button', { name: 'Zoom in' })).toBeTruthy();
      expect(view.getByRole('grid')).toBeTruthy();
      // A new throwing t, as an inline function gives on each render, warns no second time.
      current.value = (key) => t(key);
      await nextTick();
      expect(view.getByRole('button', { name: 'Zoom in' })).toBeTruthy();
      const lines = warnings.filter((line) => line.includes('The t prop of ZoomControl'));
      expect(lines.length).toBe(1);
      expect(lines[0]).toContain('"zoom.zoomIn" (broken)');
    } finally {
      console.warn = warn;
    }
  });
});

describe('a runtime change of the root locale', () => {
  test('relabels every part and keeps the workbook, the selection, and the undo history', async () => {
    const ready: { editor: XlsxEditorInstance | null } = { editor: null };
    const locale = shallowRef<Translations | undefined>(undefined);
    const view = mount(() =>
      parts({ ...(locale.value ? { locale: locale.value } : {}), onReady: (e) => (ready.editor = e) })
    );
    await tabsShow(view);
    const first = ready.editor!;
    // An edit at B2, then a selection that is not the default one.
    first.select({ row: 1, col: 1 });
    expect(first.setEditText('edited').ok).toBe(true);
    expect(first.commitEdit().ok).toBe(true);
    first.select({ row: 2, col: 3 });
    const before = first.snapshot();
    expect(before.canUndo).toBe(true);
    locale.value = { zoom: { zoomIn: 'Größer' } };
    await nextTick();
    expect(view.getByRole('button', { name: 'Größer' })).toBeTruthy();
    expect(ready.editor).toBe(first);
    const after = first.snapshot();
    expect(after.revision).toBe(before.revision);
    expect(after.selection).toEqual(before.selection);
    expect(after.selectionRef).toBe('D3');
    expect(after.canUndo).toBe(true);
    expect(after.sheets.map((s) => s.name)).toEqual(['First', 'Second']);
  });
});

describe('direction', () => {
  const chrome = (view: Mounted) =>
    ['.xlsx-toolbar', '.xlsx-formula-bar', '.xlsx-status-strip', '.xlsx-sheet-bar', '.xlsx-zoom'].map(
      (selector) => view.container.querySelector(selector)!
    );

  test('a right-to-left catalog mirrors each chrome part, and the grid keeps its direction', async () => {
    const view = mount(() => parts({ locale: { _lang: 'ar' } }));
    await tabsShow(view);
    for (const el of chrome(view)) expect(el.getAttribute('dir')).toBe('rtl');
    expect(view.container.querySelector('.xlsx-canvas')?.closest('[dir]')).toBeNull();
  });

  test('an English catalog sets no dir, so the parts keep the direction of the host page', async () => {
    const view = mount(() => parts());
    await tabsShow(view);
    for (const el of chrome(view)) expect(el.hasAttribute('dir')).toBe(false);
  });

  test('the dir of a LocaleProvider decides for its region, and useChromeDir reads it', async () => {
    const seen: (string | undefined)[] = [];
    const view = mount(() =>
      parts({
        locale: { _lang: 'he' },
        children: () =>
          h(LocaleProvider, { dir: 'ltr' }, () => [
            h('div', { class: 'region' }, [h(ZoomControl)]),
            probe(() => seen.push(useChromeDir().value)),
          ]),
      })
    );
    await tabsShow(view);
    expect(seen).toEqual(['ltr']);
    expect(view.container.querySelector('.region .xlsx-zoom')?.getAttribute('dir')).toBe('ltr');
    for (const el of chrome(view)) expect(el.getAttribute('dir')).toBe('rtl');
  });

  test('a dir around the root reaches a root whose left-to-right catalog decides no direction', async () => {
    const view = mount(() =>
      h(LocaleProvider, { dir: 'rtl' }, () => [parts({ locale: { _lang: 'en' } })])
    );
    await tabsShow(view);
    for (const el of chrome(view)) expect(el.getAttribute('dir')).toBe('rtl');
  });

  test('the packaged editor mirrors its footer and its dialog', async () => {
    const view = mount(() =>
      h(XlsxEditor, { document: workbook, fonts: FONTS, locale: { _dir: 'rtl' } })
    );
    await tabsShow(view);
    expect(view.container.querySelector('.xlsx-footer')?.getAttribute('dir')).toBe('rtl');
    fireEvent.click(view.getByRole('button', { name: /^Number format/ }));
    await nextTick();
    fireEvent.click(view.getByRole('button', { name: 'More number formats…' }));
    await nextTick();
    expect(view.getByRole('dialog', { name: 'Number format' }).getAttribute('dir')).toBe('rtl');
  });

  test('in a right-to-left bar, ArrowLeft moves to the next sheet tab', async () => {
    const ready: { editor: XlsxEditorInstance | null } = { editor: null };
    const view = mount(() => parts({ locale: { _lang: 'ar' }, onReady: (e) => (ready.editor = e) }));
    await tabsShow(view);
    const [first] = view.getAllByRole('tab');
    first!.focus();
    fireEvent.keyDown(first!, { key: 'ArrowLeft' });
    expect(ready.editor!.snapshot().activeSheet).toBe(1);
    await nextTick();
    fireEvent.keyDown(view.getAllByRole('tab')[1]!, { key: 'ArrowRight' });
    expect(ready.editor!.snapshot().activeSheet).toBe(0);
  });

  test('in a right-to-left toolbar, ArrowLeft moves to the next control', async () => {
    const view = mount(() => parts({ locale: { _lang: 'ar' } }));
    await tabsShow(view);
    const bold = view.container.querySelector<HTMLElement>('[data-slot="font.bold"]')!;
    bold.focus();
    fireEvent.keyDown(bold, { key: 'ArrowLeft' });
    expect(document.activeElement?.getAttribute('data-slot')).toBe('font.italic');
    fireEvent.keyDown(document.activeElement!, { key: 'ArrowRight' });
    expect(document.activeElement).toBe(bold);
  });

  test('a popup that runs past the left edge of a right-to-left editor aligns to its end', async () => {
    const view = mount(() => parts({ locale: { _lang: 'ar' } }));
    await tabsShow(view);
    const rect = (left: number, right: number) =>
      ({ left, right, top: 0, bottom: 20, width: right - left, height: 20, x: left, y: 0 }) as DOMRect;
    const original = Element.prototype.getBoundingClientRect;
    Element.prototype.getBoundingClientRect = function (this: Element) {
      if (this.classList.contains('xlsx-editor')) return rect(0, 800);
      if (this.hasAttribute('data-popup')) return rect(-40, 160);
      return rect(100, 140);
    };
    try {
      fireEvent.click(view.getByRole('button', { name: /^Number format/ }));
      await nextTick();
      await nextTick();
      const popup = view.container.querySelector<HTMLElement>('[data-popup]')!;
      expect(popup.dataset['align']).toBe('end');
    } finally {
      Element.prototype.getBoundingClientRect = original;
    }
  });

  test('the toolbar keeps a host dir attribute when the chrome decides no direction', async () => {
    const view = mount(() =>
      h(XlsxEditorRoot, null, () => [
        h(XlsxEditorViewport, null, () => [h(XlsxEditorToolbar, { dir: 'rtl', preset: false })]),
      ])
    );
    await nextTick();
    expect(view.getByRole('toolbar').getAttribute('dir')).toBe('rtl');
  });

  test('in a right-to-left palette, ArrowLeft moves to the next swatch', async () => {
    const view = mount(() => parts({ locale: { _lang: 'ar' } }));
    await tabsShow(view);
    fireEvent.click(view.getByRole('button', { name: /^Fill color/ }));
    await nextTick();
    await nextTick();
    const swatches = [...view.container.querySelectorAll<HTMLElement>('.xlsx-toolbar-swatch')];
    expect(swatches.length).toBeGreaterThan(2);
    swatches[0]!.focus();
    fireEvent.keyDown(swatches[0]!, { key: 'ArrowLeft' });
    expect(document.activeElement).toBe(swatches[1]!);
    fireEvent.keyDown(swatches[1]!, { key: 'ArrowRight' });
    expect(document.activeElement).toBe(swatches[0]!);
  });

  test('in a right-to-left bar, the tab menu names the side that each move goes to', async () => {
    const ready: { editor: XlsxEditorInstance | null } = { editor: null };
    const view = mount(() => parts({ locale: { _lang: 'ar' }, onReady: (e) => (ready.editor = e) }));
    await tabsShow(view);
    fireEvent.keyDown(view.getAllByRole('tab')[0]!, { key: 'F10', shiftKey: true });
    await nextTick();
    const items = view.getAllByRole('menuitem').map((el) => el.textContent);
    // The first sheet cannot move earlier: in a right-to-left bar, earlier is to the right.
    expect(items.slice(-2)).toEqual(['Move right', 'Move left']);
    fireEvent.click(view.getByRole('menuitem', { name: 'Move left' }));
    expect(ready.editor!.snapshot().sheets.map((s) => s.name)).toEqual(['Second', 'First']);
  });

  test('in a left-to-right bar, the tab menu keeps its words', async () => {
    const view = mount(() => parts());
    await tabsShow(view);
    fireEvent.keyDown(view.getAllByRole('tab')[0]!, { key: 'F10', shiftKey: true });
    await nextTick();
    expect(view.getAllByRole('menuitem').map((el) => el.textContent).slice(-2)).toEqual([
      'Move left',
      'Move right',
    ]);
  });

  test('in a right-to-left bar, the rename field starts at the right edge of its tab', async () => {
    const view = mount(() => parts({ locale: { _lang: 'ar' } }));
    await tabsShow(view);
    const rect = (left: number, right: number) =>
      ({ left, right, top: 0, bottom: 20, width: right - left, height: 20, x: left, y: 0 }) as DOMRect;
    // The bar spans 0 to 600, the list 100 to 600, the add button 60 to 100, the first tab 500
    // to 600 (the right end, where a right-to-left bar starts). Each element gets its own box,
    // so no stub on a prototype from another test file changes it.
    const boxes: [string, DOMRect][] = [
      ['.xlsx-sheet-bar', rect(0, 600)],
      ['.xlsx-sheet-tabs', rect(100, 600)],
      ['.xlsx-sheet-add', rect(60, 100)],
      ['[data-sheet-index="0"]', rect(500, 600)],
    ];
    const stubbed = boxes.map(([selector, box]) => {
      const el = view.container.querySelector<HTMLElement>(selector)!;
      Object.defineProperty(el, 'getBoundingClientRect', { configurable: true, value: () => box });
      return el;
    });
    try {
      fireEvent.keyDown(view.getAllByRole('tab')[0]!, { key: 'F2' });
      await nextTick();
      const field = view.getByRole('textbox', { name: 'Sheet name' }) as HTMLInputElement;
      expect(field.style.right).toBe('0px');
      expect(field.style.left).toBe('auto');
      // The field ends before the add button: 600 - 100 - 2 pixels.
      expect(field.style.maxWidth).toBe('498px');
    } finally {
      for (const el of stubbed) delete (el as { getBoundingClientRect?: unknown }).getBoundingClientRect;
    }
  });
});

describe('the in-cell editor', () => {
  test('the cell editor sets its direction from its text', async () => {
    const view = mount(() => parts({ locale: { _lang: 'ar' } }));
    await tabsShow(view);
    const input = view.container.querySelector('textarea.xlsx-cell-editor')!;
    expect(input.getAttribute('dir')).toBe('auto');
    expect(input.closest('.xlsx-surface')?.closest('[dir]')).toBeNull();
  });
});

describe('a field in a right-to-left toolbar', () => {
  /** A bar of Bold, the field, and Italic, in a right-to-left catalog. */
  function fieldBar(attrs: Record<string, unknown>) {
    const view = mount(() =>
      h(XlsxEditorRoot, { locale: { _lang: 'ar' } }, () => [
        h(XlsxEditorViewport, null, () => [
          h(XlsxEditorToolbar, { preset: false }, () => [
            h(XlsxEditorToolbar.Bold),
            h('input', { 'aria-label': 'Field', ...attrs }),
            h(XlsxEditorToolbar.Italic),
          ]),
        ]),
      ])
    );
    const input = view.getByRole('textbox', { name: 'Field' }) as HTMLInputElement;
    const press = (key: string) => {
      input.focus();
      input.setSelectionRange(input.value.length, input.value.length);
      fireEvent.keyDown(input, { key });
      return document.activeElement;
    };
    const slot = (id: string) => view.container.querySelector(`[data-slot="${id}"]`);
    return { input, press, slot };
  }

  test('a left-to-right field gives ArrowRight to the bar at its end, and keeps ArrowLeft', () => {
    const { input, press, slot } = fieldBar({ style: { direction: 'ltr' }, value: 'abc' });
    // In a right-to-left bar, ArrowRight moves to the previous control.
    expect(press('ArrowRight')).toBe(slot('font.bold'));
    expect(press('ArrowLeft')).toBe(input);
  });

  test('a right-to-left field gives ArrowLeft to the bar at its end, and keeps ArrowRight', () => {
    const { input, press, slot } = fieldBar({ dir: 'rtl', value: 'שלום' });
    expect(press('ArrowLeft')).toBe(slot('font.italic'));
    expect(press('ArrowRight')).toBe(input);
  });

  test('a right-to-left field of digits follows the field, as a browser caret does', () => {
    // At the end of "11" in a right-to-left field, a browser caret does not move on ArrowLeft
    // and moves back on ArrowRight: the field's direction decides, not the digits.
    const { input, press, slot } = fieldBar({ dir: 'rtl', value: '11' });
    expect(press('ArrowLeft')).toBe(slot('font.italic'));
    expect(press('ArrowRight')).toBe(input);
  });
});
