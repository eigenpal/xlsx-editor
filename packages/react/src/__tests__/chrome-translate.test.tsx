/** @jsxImportSource react */
// The language of the chrome: `useChromeTranslate` with overrides, the `t` prop of each chrome
// part, `LocaleProvider` in a region and around the root, a runtime change of the strings, and
// the direction of a right-to-left catalog. The Vue twin is `chrome-translate.test.ts`.

import { afterEach, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { act, cleanup, fireEvent, render, waitFor } from '@testing-library/react';
import { useState, type ReactNode } from 'react';
import { buildXlsx, c, row } from '../../../../scripts/fixtures/builder.ts';
import {
  LocaleProvider,
  XlsxEditor,
  useChromeDir,
  useChromeTranslate,
  useTranslation,
  type ChromeTranslate,
  type FontSource,
  type Translations,
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

interface PartsProps {
  readonly locale?: Translations;
  readonly t?: ChromeTranslate;
  readonly children?: ReactNode;
  readonly onReady?: (editor: XlsxEditorInstance) => void;
}

/** A composed layout with the same `t` on every part. */
function Parts({ locale, t, children, onReady }: PartsProps) {
  const part = t ? { t } : {};
  return (
    <XlsxEditor.Root
      document={workbook}
      fonts={FONTS}
      {...(locale ? { locale } : {})}
      {...(onReady ? { onReady } : {})}
    >
      <XlsxEditor.Viewport>
        <XlsxEditor.Toolbar {...part} />
        <XlsxEditor.FormulaBar {...part} />
        <XlsxEditor.StatusMessage {...part} />
        <XlsxEditor.Content />
        <div className="xlsx-footer">
          <XlsxEditor.SheetTabs {...part} />
          <XlsxEditor.ZoomControl {...part} />
        </div>
        {children}
      </XlsxEditor.Viewport>
    </XlsxEditor.Root>
  );
}

const ready = async (view: ReturnType<typeof render>) =>
  waitFor(() => expect(view.getAllByRole('tab').length).toBe(2), { timeout: 5000 });

describe('useChromeTranslate', () => {
  test('an override wins with its placeholders filled, and other keys read the root locale', async () => {
    const seen: string[] = [];
    const overrides = new Map([['zoom.resetAt', 'Zoom {level}, press to reset']]);
    function Probe() {
      const t = useChromeTranslate(overrides);
      seen.splice(0, seen.length, t('zoom.resetAt', { level: '80%' }), t('zoom.zoomIn'), t('x.y'));
      return null;
    }
    render(
      <XlsxEditor.Root locale={{ zoom: { zoomIn: 'Vergrößern' } }}>
        <Probe />
      </XlsxEditor.Root>
    );
    expect(seen).toEqual(['Zoom 80%, press to reset', 'Vergrößern', 'x.y']);
  });

  test('an override picks its plural branch with the plural rules of the root locale', () => {
    const overrides = new Map([['formulaBar.lines', '{count, plural, few {# (few)} other {#}}']]);
    const seen: string[] = [];
    function Probe() {
      const t = useChromeTranslate(overrides);
      seen.push(t('formulaBar.lines', { count: 3 }));
      return null;
    }
    render(
      <XlsxEditor.Root locale={{ _lang: 'pl' }}>
        <Probe />
      </XlsxEditor.Root>
    );
    expect(seen.at(-1)).toBe('3 (few)');
  });

  test('passed as a part t, it relabels that part and leaves the others', async () => {
    const overrides = new Map([['zoom.zoomIn', 'Larger']]);
    function Footer() {
      const t = useChromeTranslate(overrides);
      return <XlsxEditor.ZoomControl t={t} />;
    }
    const view = render(
      <XlsxEditor.Root document={workbook} fonts={FONTS}>
        <XlsxEditor.Viewport>
          <XlsxEditor.Toolbar />
          <XlsxEditor.Content />
          <XlsxEditor.SheetTabs />
          <Footer />
        </XlsxEditor.Viewport>
      </XlsxEditor.Root>
    );
    await ready(view);
    expect(view.getByRole('button', { name: 'Larger' })).toBeTruthy();
    expect(view.getByRole('button', { name: 'Zoom out' })).toBeTruthy();
    expect(view.getByRole('toolbar').getAttribute('aria-label')).toBe('Formatting controls');
  });

  test('inside a part with t, it reads that t first', () => {
    const seen: string[] = [];
    function Probe() {
      const t = useChromeTranslate();
      seen.push(t('formattingBar.bold'));
      return null;
    }
    render(
      <XlsxEditor.Root>
        <XlsxEditor.Viewport>
          <XlsxEditor.Toolbar preset={false} t={only({ 'formattingBar.bold': 'Heavy' })}>
            <Probe />
          </XlsxEditor.Toolbar>
        </XlsxEditor.Viewport>
      </XlsxEditor.Root>
    );
    expect(seen.at(-1)).toBe('Heavy');
  });

  test('outside a root, it reads English', () => {
    const seen: string[] = [];
    function Probe() {
      seen.push(useChromeTranslate()('zoom.zoomIn'));
      return null;
    }
    render(<Probe />);
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
      'status.stale': 'Stale',
    });
    const view = render(<Parts t={t} />);
    await ready(view);
    expect(view.getByRole('toolbar').getAttribute('aria-label')).toBe('Bar');
    expect(view.getByRole('textbox', { name: 'Address' })).toBeTruthy();
    expect(view.getByRole('button', { name: 'New sheet' })).toBeTruthy();
    expect(view.getByRole('button', { name: 'Reset from 100%' })).toBeTruthy();
    // A key that t leaves unresolved reads the root's locale, never the key.
    expect(view.getByRole('button', { name: 'Zoom in' })).toBeTruthy();
    expect(view.getByRole('textbox', { name: 'Cell content' })).toBeTruthy();
  });

  test('reaches the parts inside the toolbar and its number format dialog', async () => {
    const t = only({
      'formattingBar.bold': 'Heavy',
      'numberFormat.more': 'Other formats',
      'dialogs.numberFormat.title': 'Format of numbers',
    });
    const view = render(
      <XlsxEditor.Root document={workbook} fonts={FONTS}>
        <XlsxEditor.Viewport>
          <XlsxEditor.Toolbar t={t} />
          <XlsxEditor.Content />
          <XlsxEditor.SheetTabs />
        </XlsxEditor.Viewport>
      </XlsxEditor.Root>
    );
    await ready(view);
    expect(
      view.container.querySelector('[data-slot="font.bold"]')?.getAttribute('aria-label')
    ).toBe('Heavy');
    fireEvent.click(view.getByRole('button', { name: /^Number format/ }));
    fireEvent.click(view.getByRole('button', { name: 'Other formats' }));
    expect(view.getByRole('dialog', { name: 'Format of numbers' })).toBeTruthy();
  });

  test('a change of t relabels the part and keeps it mounted', async () => {
    function Switch() {
      const [t, setT] = useState<ChromeTranslate | undefined>(undefined);
      return (
        <>
          <button type="button" onClick={() => setT(() => only({ 'formattingBar.label': 'Bar' }))}>
            set
          </button>
          <button type="button" onClick={() => setT(undefined)}>
            unset
          </button>
          <Parts {...(t ? { t } : {})} />
        </>
      );
    }
    const view = render(<Switch />);
    await ready(view);
    const bar = view.getByRole('toolbar');
    expect(bar.getAttribute('aria-label')).toBe('Formatting controls');
    fireEvent.click(view.getByRole('button', { name: 'set' }));
    expect(view.getByRole('toolbar')).toBe(bar);
    expect(bar.getAttribute('aria-label')).toBe('Bar');
    fireEvent.click(view.getByRole('button', { name: 'unset' }));
    expect(view.getByRole('toolbar')).toBe(bar);
    expect(bar.getAttribute('aria-label')).toBe('Formatting controls');
  });
});

describe('LocaleProvider', () => {
  test('relabels one region and inherits the rest', async () => {
    const view = render(
      <Parts locale={{ zoom: { zoomOut: 'Kleiner' } }}>
        <LocaleProvider i18n={{ zoom: { zoomIn: 'Größer' } }}>
          <div className="region">
            <XlsxEditor.ZoomControl />
          </div>
        </LocaleProvider>
      </Parts>
    );
    await ready(view);
    const region = view.container.querySelector('.region')!;
    expect(region.querySelector('[aria-label="Größer"]')).toBeTruthy();
    expect(region.querySelector('[aria-label="Kleiner"]')).toBeTruthy();
    // The footer outside the provider keeps the root's catalog.
    const footer = view.container.querySelector('.xlsx-footer')!;
    expect(footer.querySelector('[aria-label="Zoom in"]')).toBeTruthy();
  });

  test('around the root, the root locale merges onto it, and host controls outside read it', async () => {
    const seen: string[] = [];
    function HostLabel() {
      const { t } = useTranslation();
      seen.push(t('zoom.zoomIn'));
      return null;
    }
    const view = render(
      <LocaleProvider i18n={{ _lang: 'de', zoom: { zoomIn: 'Größer', zoomOut: 'Kleiner' } }}>
        <HostLabel />
        <Parts locale={{ zoom: { zoomOut: 'Verkleinern' } }} />
      </LocaleProvider>
    );
    await ready(view);
    expect(seen.at(-1)).toBe('Größer');
    expect(view.getByRole('button', { name: 'Größer' })).toBeTruthy();
    expect(view.getByRole('button', { name: 'Verkleinern' })).toBeTruthy();
  });

  test('with no i18n, it keeps the t of the part above it', () => {
    const view = render(
      <XlsxEditor.Root>
        <XlsxEditor.Viewport>
          <XlsxEditor.Toolbar preset={false} t={only({ 'formattingBar.bold': 'Heavy' })}>
            <LocaleProvider>
              <XlsxEditor.Toolbar.Bold />
            </LocaleProvider>
          </XlsxEditor.Toolbar>
        </XlsxEditor.Viewport>
      </XlsxEditor.Root>
    );
    expect(
      view.container.querySelector('[data-slot="font.bold"]')?.getAttribute('aria-label')
    ).toBe('Heavy');
  });
});

describe('useTranslation', () => {
  test('returns the catalogue, the language, and the direction of the nearest catalog', () => {
    const seen: { lang: string; dir: string; zoomIn: string; t: string }[] = [];
    function Probe() {
      const { t, catalogue, lang, dir } = useTranslation();
      seen.push({ lang, dir, zoomIn: catalogue.zoom.zoomIn, t: t('zoom.zoomIn') });
      return null;
    }
    render(
      <>
        <Probe />
        <XlsxEditor.Root locale={{ _lang: 'he', zoom: { zoomIn: 'הגדל' } }}>
          <Probe />
          <LocaleProvider i18n={{ _lang: 'de', zoom: { zoomIn: 'Größer' } }} dir="ltr">
            <Probe />
          </LocaleProvider>
        </XlsxEditor.Root>
      </>
    );
    expect(seen).toEqual([
      { lang: 'en', dir: 'ltr', zoomIn: 'Zoom in', t: 'Zoom in' },
      { lang: 'he', dir: 'rtl', zoomIn: 'הגדל', t: 'הגדל' },
      { lang: 'de', dir: 'ltr', zoomIn: 'Größer', t: 'Größer' },
    ]);
  });
});

describe('a t that throws', () => {
  test('keeps the editor mounted, shows the catalog text, and warns once in development', async () => {
    const warnings: string[] = [];
    const warn = console.warn;
    console.warn = (line: string) => void warnings.push(line);
    try {
      const t: ChromeTranslate = (key) => {
        if (key === 'zoom.zoomIn') throw new Error('broken');
        return undefined;
      };
      const view = render(<Parts t={t} />);
      await ready(view);
      expect(view.getByRole('button', { name: 'Zoom in' })).toBeTruthy();
      expect(view.getByRole('grid')).toBeTruthy();
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
    const opened: { editor: XlsxEditorInstance | null } = { editor: null };
    function Switch() {
      const [locale, setLocale] = useState<Translations | undefined>(undefined);
      return (
        <>
          <button type="button" onClick={() => setLocale({ zoom: { zoomIn: 'Größer' } })}>
            de
          </button>
          <Parts {...(locale ? { locale } : {})} onReady={(e) => (opened.editor = e)} />
        </>
      );
    }
    const view = render(<Switch />);
    await ready(view);
    const first = opened.editor!;
    // An edit at B2, then a selection that is not the default one.
    act(() => {
      first.select({ row: 1, col: 1 });
      expect(first.setEditText('edited').ok).toBe(true);
      expect(first.commitEdit().ok).toBe(true);
      first.select({ row: 2, col: 3 });
    });
    const before = first.snapshot();
    expect(before.canUndo).toBe(true);
    fireEvent.click(view.getByRole('button', { name: 'de' }));
    expect(view.getByRole('button', { name: 'Größer' })).toBeTruthy();
    expect(opened.editor).toBe(first);
    const after = first.snapshot();
    expect(after.revision).toBe(before.revision);
    expect(after.selection).toEqual(before.selection);
    expect(after.selectionRef).toBe('D3');
    expect(after.canUndo).toBe(true);
    expect(after.sheets.map((s) => s.name)).toEqual(['First', 'Second']);
  });
});

describe('direction', () => {
  const parts = (view: ReturnType<typeof render>) =>
    [
      '.xlsx-toolbar',
      '.xlsx-formula-bar',
      '.xlsx-status-strip',
      '.xlsx-sheet-bar',
      '.xlsx-zoom',
    ].map((selector) => view.container.querySelector(selector)!);

  test('a right-to-left catalog mirrors each chrome part, and the grid keeps its direction', async () => {
    const view = render(<Parts locale={{ _lang: 'ar' }} />);
    await ready(view);
    for (const el of parts(view)) expect(el.getAttribute('dir')).toBe('rtl');
    expect(view.container.querySelector('.xlsx-canvas')?.closest('[dir]')).toBeNull();
  });

  test('an English catalog sets no dir, so the parts keep the direction of the host page', async () => {
    const view = render(<Parts />);
    await ready(view);
    for (const el of parts(view)) expect(el.hasAttribute('dir')).toBe(false);
  });

  test('the dir of a LocaleProvider decides for its region, and useChromeDir reads it', async () => {
    const seen: (string | undefined)[] = [];
    function Probe() {
      seen.push(useChromeDir());
      return null;
    }
    const view = render(
      <Parts locale={{ _lang: 'he' }}>
        <LocaleProvider dir="ltr">
          <div className="region">
            <XlsxEditor.ZoomControl />
          </div>
          <Probe />
        </LocaleProvider>
      </Parts>
    );
    await ready(view);
    expect(seen.at(-1)).toBe('ltr');
    expect(view.container.querySelector('.region .xlsx-zoom')?.getAttribute('dir')).toBe('ltr');
    for (const el of parts(view)) expect(el.getAttribute('dir')).toBe('rtl');
  });

  test('a dir around the root reaches a root whose left-to-right catalog decides no direction', async () => {
    const view = render(
      <LocaleProvider dir="rtl">
        <Parts locale={{ _lang: 'en' }} />
      </LocaleProvider>
    );
    await ready(view);
    for (const el of parts(view)) expect(el.getAttribute('dir')).toBe('rtl');
  });

  test('the packaged editor mirrors its footer and its dialog', async () => {
    const view = render(<XlsxEditor document={workbook} fonts={FONTS} locale={{ _dir: 'rtl' }} />);
    await ready(view);
    expect(view.container.querySelector('.xlsx-footer')?.getAttribute('dir')).toBe('rtl');
    fireEvent.click(view.getByRole('button', { name: /^Number format/ }));
    fireEvent.click(view.getByRole('button', { name: 'More number formats…' }));
    expect(view.getByRole('dialog', { name: 'Number format' }).getAttribute('dir')).toBe('rtl');
  });

  test('in a right-to-left bar, ArrowLeft moves to the next sheet tab', async () => {
    const opened: { editor: XlsxEditorInstance | null } = { editor: null };
    const view = render(<Parts locale={{ _lang: 'ar' }} onReady={(e) => (opened.editor = e)} />);
    await ready(view);
    const [first] = view.getAllByRole('tab');
    act(() => first!.focus());
    fireEvent.keyDown(first!, { key: 'ArrowLeft' });
    expect(opened.editor!.snapshot().activeSheet).toBe(1);
    fireEvent.keyDown(view.getAllByRole('tab')[1]!, { key: 'ArrowRight' });
    expect(opened.editor!.snapshot().activeSheet).toBe(0);
  });

  test('in a right-to-left toolbar, ArrowLeft moves to the next control', async () => {
    const view = render(<Parts locale={{ _lang: 'ar' }} />);
    await ready(view);
    const bold = view.container.querySelector<HTMLElement>('[data-slot="font.bold"]')!;
    act(() => bold.focus());
    fireEvent.keyDown(bold, { key: 'ArrowLeft' });
    expect(document.activeElement?.getAttribute('data-slot')).toBe('font.italic');
    fireEvent.keyDown(document.activeElement!, { key: 'ArrowRight' });
    expect(document.activeElement).toBe(bold);
  });

  test('a popup that runs past the left edge of a right-to-left editor aligns to its end', async () => {
    const view = render(<Parts locale={{ _lang: 'ar' }} />);
    await ready(view);
    const rect = (left: number, right: number) =>
      ({
        left,
        right,
        top: 0,
        bottom: 20,
        width: right - left,
        height: 20,
        x: left,
        y: 0,
      }) as DOMRect;
    const original = Element.prototype.getBoundingClientRect;
    Element.prototype.getBoundingClientRect = function (this: Element) {
      if (this.classList.contains('xlsx-editor')) return rect(0, 800);
      if (this.hasAttribute('data-popup')) return rect(-40, 160);
      return rect(100, 140);
    };
    try {
      fireEvent.click(view.getByRole('button', { name: /^Number format/ }));
      const popup = view.container.querySelector<HTMLElement>('[data-popup]')!;
      expect(popup.dataset['align']).toBe('end');
    } finally {
      Element.prototype.getBoundingClientRect = original;
    }
  });

  test('in a right-to-left palette, ArrowLeft moves to the next swatch', async () => {
    const view = render(<Parts locale={{ _lang: 'ar' }} />);
    await ready(view);
    fireEvent.click(view.getByRole('button', { name: /^Fill color/ }));
    const swatches = [...view.container.querySelectorAll<HTMLElement>('.xlsx-toolbar-swatch')];
    expect(swatches.length).toBeGreaterThan(2);
    act(() => swatches[0]!.focus());
    fireEvent.keyDown(swatches[0]!, { key: 'ArrowLeft' });
    expect(document.activeElement).toBe(swatches[1]!);
    fireEvent.keyDown(swatches[1]!, { key: 'ArrowRight' });
    expect(document.activeElement).toBe(swatches[0]!);
  });

  test('in a right-to-left bar, the tab menu names the side that each move goes to', async () => {
    const opened: { editor: XlsxEditorInstance | null } = { editor: null };
    const view = render(<Parts locale={{ _lang: 'ar' }} onReady={(e) => (opened.editor = e)} />);
    await ready(view);
    const [first] = view.getAllByRole('tab');
    fireEvent.keyDown(first!, { key: 'F10', shiftKey: true });
    const items = view.getAllByRole('menuitem').map((el) => el.textContent);
    // The first sheet cannot move earlier: in a right-to-left bar, earlier is to the right.
    expect(items.slice(-2)).toEqual(['Move right', 'Move left']);
    fireEvent.click(view.getByRole('menuitem', { name: 'Move left' }));
    expect(opened.editor!.snapshot().sheets.map((s) => s.name)).toEqual(['Second', 'First']);
  });

  test('in a left-to-right bar, the tab menu keeps its words', async () => {
    const view = render(<Parts />);
    await ready(view);
    fireEvent.keyDown(view.getAllByRole('tab')[0]!, { key: 'F10', shiftKey: true });
    expect(
      view
        .getAllByRole('menuitem')
        .map((el) => el.textContent)
        .slice(-2)
    ).toEqual(['Move left', 'Move right']);
  });

  test('in a right-to-left bar, the rename field starts at the right edge of its tab', async () => {
    const view = render(<Parts locale={{ _lang: 'ar' }} />);
    await ready(view);
    const rect = (left: number, right: number) =>
      ({
        left,
        right,
        top: 0,
        bottom: 20,
        width: right - left,
        height: 20,
        x: left,
        y: 0,
      }) as DOMRect;
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
      const [first] = view.getAllByRole('tab');
      fireEvent.keyDown(first!, { key: 'F2' });
      const field = view.getByRole('textbox', { name: 'Sheet name' }) as HTMLInputElement;
      expect(field.style.right).toBe('0px');
      expect(field.style.left).toBe('auto');
      // The field ends before the add button: 600 - 100 - 2 pixels.
      expect(field.style.maxWidth).toBe('498px');
    } finally {
      for (const el of stubbed)
        delete (el as { getBoundingClientRect?: unknown }).getBoundingClientRect;
    }
  });
});
