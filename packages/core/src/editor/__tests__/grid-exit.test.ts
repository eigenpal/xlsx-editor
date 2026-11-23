// The exit from the grid (`grid-exit.ts`): Escape with nothing to cancel arms it and says how to
// leave, and the next Tab or Shift+Tab goes to the browser, which moves the focus out of the
// grid. Any other key, a press, or a focus that leaves the grid disarms it; a modifier key alone
// does not. While it is armed, a cue over the grid shows the hint. The grid's accessible
// description names the exit. An Escape that forgets a pending cut, or that cancels a cell edit,
// does not arm it.

import { afterEach, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { Window } from 'happy-dom';
import { buildXlsx, c, row } from '../../../../../scripts/fixtures/builder.ts';
import { createXlsxEditor, type XlsxEditorInstance } from '../index.ts';
import { flushEditor } from './flush-editor.ts';
import type { FontSource } from '../../layout/index.ts';

const fontsDir = join(
  dirname(Bun.resolveSync('@docx-editor.dev/fonts/package.json', import.meta.dir)),
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

let viewers: XlsxEditorInstance[] = [];
afterEach(() => {
  for (const v of viewers) v.destroy();
  viewers = [];
});

async function mount(merges?: string[]) {
  const win = new Window();
  const doc = win.document as unknown as Document;
  const host = doc.createElement('div');
  doc.body.appendChild(host);
  const viewer = createXlsxEditor({ fonts: FONTS, translate: (k) => k });
  viewers.push(viewer);
  viewer.attach(host);
  const scroller = host.querySelector('.xlsx-scroller') as HTMLElement;
  Object.defineProperty(scroller, 'clientWidth', { get: () => 800 });
  Object.defineProperty(scroller, 'clientHeight', { get: () => 400 });
  await viewer.load(buildXlsx({ sheets: [{ name: 'S', rows: row(1, [c('A1', 'a')]), merges }] }));
  flushEditor(viewer);
  viewer.select({ row: 4, col: 3 });
  const live = host.querySelector('.xlsx-live') as HTMLElement;
  const input = host.querySelector('.xlsx-cell-editor') as HTMLTextAreaElement;
  return { viewer, scroller, live, input, win };
}

type Win = Awaited<ReturnType<typeof mount>>['win'];
const W = (win: Win) => win as unknown as typeof globalThis;

/** Dispatch a keydown on `el`; returns whether the default stays (the browser acts). */
const key = (win: Win, el: Element, k: string, init: Record<string, unknown> = {}) =>
  el.dispatchEvent(
    new (W(win).KeyboardEvent)('keydown', {
      key: k,
      bubbles: true,
      cancelable: true,
      ...init,
    })
  );

const active = (v: XlsxEditorInstance) => v.snapshot().selection?.active;

describe('the exit from the grid', () => {
  test('Escape, then Tab, leaves the grid from any cell, and says how', async () => {
    const { viewer, scroller, live, win } = await mount();
    expect(key(win, scroller, 'Tab')).toBe(false);
    expect(active(viewer)).toEqual({ row: 4, col: 4 });
    expect(key(win, scroller, 'Escape')).toBe(true);
    expect(live.textContent).toBe('grid.exitHint');
    expect(key(win, scroller, 'Tab')).toBe(true);
    expect(active(viewer)).toEqual({ row: 4, col: 4 });
    // The exit is used: the next Tab moves the active cell again.
    expect(key(win, scroller, 'Tab')).toBe(false);
    expect(active(viewer)).toEqual({ row: 4, col: 5 });
  });

  test('Escape, then Shift+Tab, leaves the grid; the Shift alone keeps the exit', async () => {
    const { viewer, scroller, win } = await mount();
    key(win, scroller, 'Escape');
    expect(key(win, scroller, 'Shift', { shiftKey: true })).toBe(true);
    expect(key(win, scroller, 'Tab', { shiftKey: true })).toBe(true);
    expect(active(viewer)).toEqual({ row: 4, col: 3 });
  });

  test('another key, a press, or a focus that leaves the grid disarms the exit', async () => {
    const { viewer, scroller, win } = await mount();
    key(win, scroller, 'Escape');
    expect(key(win, scroller, 'ArrowDown')).toBe(false);
    expect(key(win, scroller, 'Tab')).toBe(false);
    expect(active(viewer)).toEqual({ row: 5, col: 4 });
    key(win, scroller, 'Escape');
    scroller.dispatchEvent(new (W(win).FocusEvent)('focusout'));
    expect(key(win, scroller, 'Tab')).toBe(false);
    key(win, scroller, 'Escape');
    const stage = scroller.querySelector('.xlsx-stage')!;
    stage.dispatchEvent(new (W(win).Event)('pointerdown'));
    expect(key(win, scroller, 'Tab')).toBe(false);
    // Mod+Tab is no exit: the key table takes it, as before, and it disarms the exit.
    key(win, scroller, 'Escape');
    const before = active(viewer);
    key(win, scroller, 'Tab', { ctrlKey: true });
    expect(active(viewer)).not.toEqual(before);
    expect(key(win, scroller, 'Tab')).toBe(false);
  });

  test('an Escape that forgets a pending cut does not arm the exit', async () => {
    const { viewer, scroller, live, win } = await mount();
    const cut = new (W(win).Event)('cut', {
      bubbles: true,
      cancelable: true,
    });
    Object.defineProperty(cut, 'clipboardData', { value: { setData: () => {} } });
    scroller.focus();
    scroller.dispatchEvent(cut);
    expect(key(win, scroller, 'Escape')).toBe(false);
    expect(live.textContent).not.toBe('grid.exitHint');
    expect(key(win, scroller, 'Tab')).toBe(false);
    // The next Escape, with nothing to forget, arms it.
    key(win, scroller, 'Escape');
    expect(key(win, scroller, 'Tab')).toBe(true);
    expect(viewer.snapshot().selection?.active).toEqual({ row: 4, col: 4 });
  });

  test('from an open cell edit: Escape cancels it, and the next Escape arms the exit', async () => {
    const { viewer, scroller, input, live, win } = await mount();
    expect(viewer.beginEdit('x').ok).toBe(true);
    key(win, input, 'Escape');
    expect(viewer.snapshot().editing).toBe(false);
    expect(live.textContent).not.toBe('grid.exitHint');
    key(win, scroller, 'Escape');
    expect(live.textContent).toBe('grid.exitHint');
    expect(key(win, scroller, 'Tab')).toBe(true);
  });

  test('the grid names the exit in its accessible description', async () => {
    const { scroller } = await mount();
    expect(scroller.getAttribute('aria-description')).toBe('grid.exitDescription');
  });

  test('Escape with a modifier, or while composing, does not arm the exit', async () => {
    const { scroller, live, win } = await mount();
    for (const init of [
      { shiftKey: true },
      { ctrlKey: true },
      { altKey: true },
      { metaKey: true },
      { isComposing: true },
    ]) {
      key(win, scroller, 'Escape', init);
      expect(live.textContent).not.toBe('grid.exitHint');
      expect(key(win, scroller, 'Tab')).toBe(false);
    }
  });

  test('Alt+Tab and Meta+Tab do not use the exit, and AltGraph alone keeps it', async () => {
    const { scroller, win } = await mount();
    for (const init of [{ altKey: true }, { metaKey: true }]) {
      key(win, scroller, 'Escape');
      key(win, scroller, 'Tab', init);
      expect(key(win, scroller, 'Tab')).toBe(false);
    }
    key(win, scroller, 'Escape');
    expect(key(win, scroller, 'AltGraph')).toBe(true);
    expect(key(win, scroller, 'Tab')).toBe(true);
  });

  test('a key from an element inside the grid neither arms nor uses the exit', async () => {
    const { viewer, scroller, live, win } = await mount();
    const stage = scroller.querySelector('.xlsx-stage')!;
    key(win, stage, 'Escape');
    expect(live.textContent).not.toBe('grid.exitHint');
    expect(key(win, scroller, 'Tab')).toBe(false);
    key(win, scroller, 'Escape');
    expect(key(win, stage, 'Tab')).toBe(true);
    // The armed exit stays for the next key on the grid itself.
    expect(active(viewer)).toEqual({ row: 4, col: 4 });
    expect(key(win, scroller, 'Tab')).toBe(true);
  });

  test('opening a cell edit moves the focus, which disarms the exit', async () => {
    const { viewer, scroller, win } = await mount();
    key(win, scroller, 'Escape');
    expect(viewer.beginEdit('x').ok).toBe(true);
    viewer.cancelEdit();
    expect(key(win, scroller, 'Tab')).toBe(false);
  });

  test('a press on the chrome or on the scrollbar, which keep the focus, disarms the exit', async () => {
    const { scroller, win } = await mount();
    const chrome = scroller.ownerDocument.createElement('button');
    scroller.ownerDocument.body.appendChild(chrome);
    for (const target of [chrome, scroller]) {
      key(win, scroller, 'Escape');
      target.dispatchEvent(new (W(win).Event)('pointerdown', { bubbles: true }));
      expect(key(win, scroller, 'Tab')).toBe(false);
    }
  });

  test('a held Escape after a pending cut arms the exit and says the hint', async () => {
    const { scroller, live, win } = await mount();
    const cut = new (W(win).Event)('cut', {
      bubbles: true,
      cancelable: true,
    });
    Object.defineProperty(cut, 'clipboardData', { value: { setData: () => {} } });
    scroller.focus();
    scroller.dispatchEvent(cut);
    expect(key(win, scroller, 'Escape')).toBe(false);
    expect(live.textContent).not.toBe('grid.exitHint');
    key(win, scroller, 'Escape', { repeat: true });
    expect(live.textContent).toBe('grid.exitHint');
    expect(key(win, scroller, 'Tab')).toBe(true);
  });

  test('a held Escape says the hint once', async () => {
    const { scroller, live, win } = await mount();
    key(win, scroller, 'Escape');
    expect(live.textContent).toBe('grid.exitHint');
    live.textContent = '';
    key(win, scroller, 'Escape', { repeat: true });
    expect(live.textContent).toBe('');
    expect(key(win, scroller, 'Tab')).toBe(true);
  });

  test('Home and then Shift+Tab leave the grid from a merged cell in column A', async () => {
    const { viewer, scroller, win } = await mount(['A5:E5']);
    expect(key(win, scroller, 'Home')).toBe(false);
    expect(viewer.snapshot().selectionRef).toBe('A5:E5');
    expect(key(win, scroller, 'Tab', { shiftKey: true })).toBe(true);
    // Tab steps over the merged cell, as from one cell.
    expect(key(win, scroller, 'Tab')).toBe(false);
    expect(active(viewer)).toEqual({ row: 4, col: 5 });
  });

  test('a cue shows the hint while the exit is armed, and hides when it is used or canceled', async () => {
    const { scroller, win } = await mount();
    const cue = scroller.parentElement!.querySelector('.xlsx-exit-cue')!;
    // Outside the painted stage, and hidden from assistive technology: the live region speaks.
    expect(cue.closest('.xlsx-stage')).toBeNull();
    expect(cue.getAttribute('aria-hidden')).toBe('true');
    expect(cue.hasAttribute('data-shown')).toBe(false);
    key(win, scroller, 'Escape');
    expect(cue.hasAttribute('data-shown')).toBe(true);
    expect(cue.textContent).toBe('grid.exitHint');
    key(win, scroller, 'Shift', { shiftKey: true });
    expect(cue.hasAttribute('data-shown')).toBe(true);
    key(win, scroller, 'Tab');
    expect(cue.hasAttribute('data-shown')).toBe(false);
    for (const cancel of [
      () => key(win, scroller, 'ArrowDown'),
      () => scroller.dispatchEvent(new (W(win).FocusEvent)('focusout')),
      () => scroller.dispatchEvent(new (W(win).Event)('pointerdown', { bubbles: true })),
    ]) {
      key(win, scroller, 'Escape');
      expect(cue.hasAttribute('data-shown')).toBe(true);
      cancel();
      expect(cue.hasAttribute('data-shown')).toBe(false);
    }
  });

  test('Enter and Shift+Enter move past a merged cell, as from one cell', async () => {
    const { viewer, scroller, win } = await mount(['B5:C7']);
    viewer.select({ row: 4, col: 1 });
    expect(viewer.snapshot().selectionRef).toBe('B5:C7');
    expect(key(win, scroller, 'Enter')).toBe(false);
    expect(active(viewer)).toEqual({ row: 7, col: 1 });
    viewer.select({ row: 4, col: 1 });
    expect(key(win, scroller, 'Enter', { shiftKey: true })).toBe(false);
    expect(active(viewer)).toEqual({ row: 3, col: 1 });
  });

  test('a committed edit with Enter or Tab moves past a merged cell, as from one cell', async () => {
    const { viewer, input, win } = await mount(['B5:C7']);
    for (const [k, to] of [
      ['Enter', { row: 7, col: 1 }],
      ['Tab', { row: 4, col: 3 }],
    ] as const) {
      viewer.select({ row: 4, col: 1 });
      expect(viewer.beginEdit('x').ok).toBe(true);
      key(win, input, k);
      expect(viewer.snapshot().editing).toBe(false);
      expect(active(viewer)).toEqual(to);
    }
  });

  test('view mode has the same exit', async () => {
    const { viewer, scroller, win } = await mount();
    viewer.setMode('view');
    key(win, scroller, 'Escape');
    expect(key(win, scroller, 'Tab')).toBe(true);
  });
});
