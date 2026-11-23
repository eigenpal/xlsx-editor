// The frame of a host page around the editor: the part of the window where a chrome part can
// show inside host containers that clip their overflow, and the scale of a host transform. After
// pptx-editor/packages/core/src/editor/__tests__/fixed-frame.test.ts at 3bf89a88 (EigenPal). The
// browser check is the `embedding` step of `scripts/verify-demo.ts`.

import { afterEach, describe, expect, test } from 'bun:test';
import { GlobalWindow } from 'happy-dom';
import { sheetMenuLeft, sheetRenameBox, type EditorSheet } from '../index.ts';
import { visibleArea, windowScale } from '../host-frame.ts';

// A DOM of this file's own: the core run has no global DOM (`bun run test`).
const win = new GlobalWindow({ width: 1000, height: 800 });
const document = win.document as unknown as Document;

const box = (left: number, top: number, width: number, height: number): DOMRect =>
  ({ left, top, width, height, right: left + width, bottom: top + height }) as DOMRect;

/** A `div` under `parent` with `style`, measured at `rect`. */
function el(parent: Element, style: Partial<CSSStyleDeclaration>, rect?: DOMRect): HTMLElement {
  const node = document.createElement('div');
  Object.assign(node.style, style);
  if (rect) node.getBoundingClientRect = () => rect;
  parent.appendChild(node);
  return node;
}

/** A layout box of `width` by `height` that shows at `rect`. */
function laidOut(node: HTMLElement, width: number, height: number, rect: DOMRect): HTMLElement {
  Object.defineProperty(node, 'offsetWidth', { configurable: true, get: () => width });
  Object.defineProperty(node, 'offsetHeight', { configurable: true, get: () => height });
  node.getBoundingClientRect = () => rect;
  return node;
}

const WINDOW = { left: 0, top: 0, right: 1000, bottom: 800 };

afterEach(() => {
  document.body.replaceChildren();
  document.body.removeAttribute('style');
  document.body.getBoundingClientRect = () => box(0, 0, 0, 0);
});

describe('visibleArea', () => {
  test('with no host container that clips, the area is the window', () => {
    const popup = el(el(document.body, { position: 'relative' }), { position: 'absolute' });
    expect(visibleArea(popup)).toEqual(WINDOW);
  });

  test('a host container that clips its overflow limits the area to its box', () => {
    // A drawer that slides in with a transform and clips its overflow.
    const drawer = el(
      document.body,
      { overflow: 'hidden', transform: 'translateX(0)' },
      box(600, 100, 300, 400)
    );
    const popup = el(el(drawer, { position: 'relative' }), { position: 'absolute' });
    expect(visibleArea(popup)).toEqual({ left: 600, top: 100, right: 900, bottom: 500 });
  });

  test('the element’s own overflow does not count, and nested clips intersect', () => {
    const outer = el(document.body, { overflow: 'auto' }, box(0, 50, 700, 600));
    const inner = el(outer, { overflow: 'hidden' }, box(100, 0, 900, 400));
    const popup = el(inner, { overflow: 'auto' }, box(150, 60, 200, 900));
    expect(visibleArea(popup)).toEqual({ left: 100, top: 50, right: 700, bottom: 400 });
  });

  test('an absolute element escapes a static ancestor that clips, not a positioned one', () => {
    const page = el(
      document.body,
      { position: 'relative', overflow: 'hidden' },
      box(0, 0, 900, 700)
    );
    const card = el(page, { position: 'static', overflow: 'hidden' }, box(10, 10, 100, 100));
    const popup = el(card, { position: 'absolute' });
    // The card is static: it is not the popup's containing block, so it does not clip it.
    expect(visibleArea(popup)).toEqual({ left: 0, top: 0, right: 900, bottom: 700 });
  });

  test('a fixed sheet is not clipped by the page around it or by a body scroll lock', () => {
    document.body.style.overflow = 'hidden';
    document.body.getBoundingClientRect = () => box(0, 0, 1000, 200);
    const main = el(document.body, { overflow: 'auto' }, box(0, 0, 500, 300));
    const sheet = el(
      main,
      { position: 'fixed', transform: 'translateY(0)', overflow: 'hidden' },
      box(400, 100, 600, 700)
    );
    const popup = el(el(sheet, { position: 'relative' }), { position: 'absolute' });
    // Only the sheet's own box: not `main`'s, and not the scroll-locked body's.
    expect(visibleArea(popup)).toEqual({ left: 400, top: 100, right: 1000, bottom: 800 });
  });

  test('a body scroll lock clips nothing while the root does not clip', () => {
    // A host locks the page scroll with `overflow: hidden` on the body: the body's overflow
    // moves to the window, so it does not cut the area to the body's box.
    document.body.style.overflow = 'hidden';
    document.body.getBoundingClientRect = () => box(0, 0, 1000, 200);
    const popup = el(el(document.body, { position: 'relative' }), { position: 'absolute' });
    expect(visibleArea(popup)).toEqual(WINDOW);
  });

  test('a modal dialog stops the walk: the page around it does not clip it', () => {
    const drawer = el(document.body, { overflow: 'hidden' }, box(600, 100, 300, 400));
    const dialog = el(drawer, { overflow: 'auto' }, box(100, 50, 800, 700));
    // happy-dom has no top layer: the dialog matches `:modal` as an open modal dialog does.
    const matches = dialog.matches.bind(dialog);
    Object.defineProperty(dialog, 'matches', {
      value: (selector: string) => selector.includes(':modal') || matches(selector),
    });
    const popup = el(el(dialog, { position: 'relative' }), { position: 'absolute' });
    // Only the dialog's own box, not the drawer's.
    expect(visibleArea(popup)).toEqual({ left: 100, top: 50, right: 900, bottom: 750 });
  });

  test('an element in the top layer is not clipped by the page around it', () => {
    const drawer = el(document.body, { overflow: 'hidden' }, box(600, 100, 300, 400));
    const dialog = el(drawer, { position: 'relative' }, box(100, 50, 800, 700));
    const matches = dialog.matches.bind(dialog);
    Object.defineProperty(dialog, 'matches', {
      value: (selector: string) => selector.includes(':modal') || matches(selector),
    });
    expect(visibleArea(dialog)).toEqual(WINDOW);
  });

  test('clips that do not meet give an empty area, not a negative one', () => {
    const outer = el(document.body, { overflow: 'hidden' }, box(0, 0, 100, 100));
    const inner = el(outer, { overflow: 'hidden' }, box(200, 200, 100, 100));
    const popup = el(el(inner, { position: 'relative' }), { position: 'absolute' });
    expect(visibleArea(popup)).toEqual({ left: 200, top: 200, right: 200, bottom: 200 });
  });
});

describe('windowScale', () => {
  test('a host that scales the editor gives its scale on each axis', () => {
    const node = laidOut(el(document.body, {}), 400, 100, box(0, 0, 500, 200));
    expect(windowScale(node)).toEqual({ x: 1.25, y: 2 });
  });

  test('an element with no layout box has a scale of 1', () => {
    expect(windowScale(el(document.body, {}))).toEqual({ x: 1, y: 1 });
  });

  test('a fractional layout size with no transform has a scale of 1', () => {
    // The layout size is rounded to whole pixels: 400.4 by 99.6 shows as 400 by 100.
    const node = laidOut(el(document.body, {}), 400, 100, box(0, 0, 400.4, 99.6));
    expect(windowScale(node)).toEqual({ x: 1, y: 1 });
  });
});

describe('the sheet tabs inside a host that scales the editor by 2', () => {
  const sheets = [{ name: 'One', index: 0, visible: true }] as unknown as readonly EditorSheet[];

  test('the tab menu opens at the pointer, in the bar’s pixels', () => {
    const bar = laidOut(el(document.body, {}), 300, 30, box(100, 0, 600, 60));
    const tab = el(bar, {}, box(140, 0, 80, 60));
    expect(sheetMenuLeft(bar, tab, 300)).toBe(100);
    // A key opens it at the tab's left edge.
    expect(sheetMenuLeft(bar, tab)).toBe(20);
  });

  test('the rename field lies over its tab, in the bar’s pixels', () => {
    const bar = laidOut(el(document.body, {}), 300, 30, box(100, 0, 600, 60));
    const list = el(bar, {}, box(100, 0, 400, 60));
    const tab = el(list, {}, box(140, 0, 80, 60));
    tab.dataset.sheetIndex = '0';
    const add = el(bar, {}, box(520, 0, 40, 60));
    expect(sheetRenameBox({ bar, list, add }, sheets, 'One')).toEqual({
      name: 'One',
      left: 20,
      width: 40,
      maxWidth: 188,
    });
  });

  test('the rename field of a tab that scrolled out stays over the tab list', () => {
    const bar = laidOut(el(document.body, {}), 300, 30, box(100, 0, 600, 60));
    const list = el(bar, {}, box(140, 0, 360, 60));
    const tab = el(list, {}, box(560, 0, 80, 60));
    tab.dataset.sheetIndex = '0';
    const parts = { bar, list, add: null };
    // Past the list's right edge (500): the field ends there, 160 of the bar's pixels in.
    expect(sheetRenameBox(parts, sheets, 'One')?.left).toBe(160);
    // Before the list's left edge (140): the field starts there, 20 of the bar's pixels in.
    tab.getBoundingClientRect = () => box(60, 0, 80, 60);
    expect(sheetRenameBox(parts, sheets, 'One')?.left).toBe(20);
  });
});
