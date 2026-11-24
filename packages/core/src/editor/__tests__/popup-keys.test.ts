// The keys of a chrome popup that a pointer press opened (`popup-keys.ts`): which keys belong to
// the popup, which key origins it hears, how it moves the keyboard into a menu panel or a toolbar
// popup, and which rows a submenu reads its state from (`submenuRowSlots`).

import { afterEach, describe, expect, test } from 'bun:test';
import { Window } from 'happy-dom';
import {
  customHexDraft,
  enterMenuPanel,
  enterToolbarPopup,
  hexFromText,
  listenForPointerPopupKeys,
  popupEntryOf,
  submenuRowSlots,
  type PopupEntry,
  type PopupEntryResult,
} from '../index.ts';

const stops: (() => void)[] = [];
afterEach(() => stops.splice(0).forEach((stop) => stop()));

/** An editor instance with a grid and a menu bar popup, and an input outside the editor. */
function page() {
  const win = new Window();
  const doc = win.document as unknown as Document;
  doc.body.innerHTML = `
    <div class="xlsx-editor">
      <div class="xlsx-menubar"><div class="popup"><button class="row">Copy</button></div></div>
      <div class="xlsx-surface"><div class="xlsx-scroller" tabindex="0"></div></div>
      <div aria-modal="true"><input class="modal-field"></div>
    </div>
    <input class="host-field">`;
  const $ = (s: string) => doc.querySelector<HTMLElement>(s)!;
  return { win, doc, $, popup: $('.popup'), grid: $('.xlsx-scroller') };
}

function press(el: Element, key: string, init: Record<string, unknown> = {}) {
  const View = el.ownerDocument!.defaultView as unknown as typeof globalThis;
  const event = new View.KeyboardEvent('keydown', {
    key,
    bubbles: true,
    cancelable: true,
    ...init,
  });
  el.dispatchEvent(event);
  return event;
}

function listen(popup: HTMLElement, result: PopupEntryResult = 'done') {
  const heard: PopupEntry[] = [];
  stops.push(
    listenForPointerPopupKeys({
      popup,
      contains: (node) => popup.contains(node),
      enter: (entry) => {
        heard.push(entry);
        return result;
      },
    })
  );
  return heard;
}

describe('popupEntryOf', () => {
  const of = (key: string, mods: Partial<Record<'altKey' | 'ctrlKey' | 'metaKey', boolean>> = {}) =>
    popupEntryOf({ key, altKey: false, ctrlKey: false, metaKey: false, ...mods })?.kind ?? null;

  test('the arrows, Home, End, Enter, Space, and a character belong to the popup', () => {
    expect(['ArrowDown', 'Enter', ' ', 'Home'].map((k) => of(k))).toEqual(Array(4).fill('first'));
    expect(['ArrowUp', 'End'].map((k) => of(k))).toEqual(['last', 'last']);
    expect(['ArrowLeft', 'ArrowRight'].map((k) => of(k))).toEqual(['side', 'side']);
    expect(['a', '1', 'Ä'].map((k) => of(k))).toEqual(['char', 'char', 'char']);
  });

  test('Escape, Tab, F6, and keys with a modifier pass', () => {
    expect(['Escape', 'Tab', 'F6', 'Shift'].map((k) => of(k))).toEqual([null, null, null, null]);
    expect(of('ArrowDown', { ctrlKey: true })).toBeNull();
    expect(of('c', { metaKey: true })).toBeNull();
    expect(of('a', { altKey: true })).toBeNull();
  });
});

describe('listenForPointerPopupKeys', () => {
  test('a key from the grid goes to the popup, and the grid never hears it', () => {
    const { popup, grid } = page();
    const heard = listen(popup);
    let gridHeard = false;
    grid.addEventListener('keydown', () => (gridHeard = true));
    const event = press(grid, 'ArrowDown');
    expect(heard.map((e) => e.kind)).toEqual(['first']);
    expect(event.defaultPrevented).toBe(true);
    expect(gridHeard).toBe(false);
  });

  test('a typed character that a field takes keeps its default', () => {
    const { popup, grid } = page();
    listen(popup, 'type');
    const event = press(grid, '1');
    expect(event.defaultPrevented).toBe(false);
  });

  test('a key from the popup, the host page, or a modal dialog, and a composed key, pass', () => {
    const { popup, grid, $ } = page();
    const heard = listen(popup);
    press($('.row'), 'ArrowDown');
    press($('.host-field'), 'ArrowDown');
    press($('.modal-field'), 'ArrowDown');
    press(grid, 'a', { isComposing: true });
    expect(heard).toEqual([]);
  });

  test('a key that the page already used passes', () => {
    const { popup, grid, doc } = page();
    // A capture listener on the document added first runs first.
    doc.addEventListener('keydown', (e) => e.preventDefault(), { capture: true });
    const heard = listen(popup);
    press(grid, 'ArrowDown');
    expect(heard).toEqual([]);
  });

  test('the disposer stops the listener', () => {
    const { popup, grid } = page();
    const heard = listen(popup);
    stops.splice(0).forEach((stop) => stop());
    press(grid, 'ArrowDown');
    expect(heard).toEqual([]);
  });
});

describe('entering a popup', () => {
  function panel() {
    const { doc } = page();
    const root = doc.createElement('div');
    root.innerHTML = ['Undo', 'Cut', 'Copy', 'Select all']
      .map((t) => `<button role="menuitem"><span class="xlsx-menu-label">${t}</span></button>`)
      .join('');
    doc.body.append(root);
    return { doc, root, rows: [...root.querySelectorAll<HTMLElement>('button')] };
  }
  const entry = (kind: PopupEntry['kind'], key: string) => ({ kind, key }) as PopupEntry;

  test('a menu panel takes its first row, its last row, or the row that starts with a letter', () => {
    const { doc, root, rows } = panel();
    const step: number[] = [];
    const options = { rtl: false, step: (by: -1 | 1) => step.push(by) };
    expect(enterMenuPanel(entry('first', 'ArrowDown'), root, options)).toBe('done');
    expect(doc.activeElement).toBe(rows[0]);
    enterMenuPanel(entry('last', 'End'), root, options);
    expect(doc.activeElement).toBe(rows[3]);
    enterMenuPanel(entry('char', 'c'), root, options);
    expect(doc.activeElement).toBe(rows[1]);
    enterMenuPanel(entry('char', 'z'), root, options);
    expect(doc.activeElement).toBe(rows[0]);
    enterMenuPanel(entry('side', 'ArrowRight'), root, options);
    enterMenuPanel(entry('side', 'ArrowRight'), root, { ...options, rtl: true });
    expect(step).toEqual([1, -1]);
  });

  test('a toolbar popup gives a character to its field, else moves to the item that starts with it', () => {
    const { doc, rows } = panel();
    const field = doc.createElement('input');
    doc.body.append(field);
    let focusedIn = 0;
    const focusIn = () => (focusedIn += 1);
    expect(enterToolbarPopup(entry('char', '1'), rows, field, focusIn)).toBe('type');
    expect(doc.activeElement).toBe(field);
    expect(enterToolbarPopup(entry('char', 's'), rows, null, focusIn)).toBe('done');
    expect(doc.activeElement).toBe(rows[3]);
    enterToolbarPopup(entry('last', 'ArrowUp'), rows, null, focusIn);
    expect(doc.activeElement).toBe(rows[3]);
    enterToolbarPopup(entry('first', 'ArrowDown'), rows, null, focusIn);
    expect(focusedIn).toBe(1);
    expect(enterToolbarPopup(entry('first', 'Enter'), [], null, focusIn)).toBeNull();
  });
});

describe('submenuRowSlots', () => {
  type Row = { slotId?: string; hidden?: boolean; separator?: boolean };
  const slots = (rows: Row[]) =>
    submenuRowSlots(
      rows,
      (r) => (r.slotId ? { slotId: r.slotId, hidden: r.hidden === true } : null),
      (r) => r.separator === true
    );

  test('lists the shown slot rows, past separators', () => {
    expect(
      slots([
        { slotId: 'clear.contents' },
        { separator: true },
        { slotId: 'clear.formats', hidden: true },
        { slotId: 'clear.all' },
      ])
    ).toEqual(['clear.contents', 'clear.all']);
  });

  test('is empty when a row is not a slot row', () => {
    expect(slots([{ slotId: 'clear.contents' }, {}])).toEqual([]);
  });
});

describe('the custom hex field', () => {
  test('a pasted #RRGGBB keeps its six digits', () => {
    expect(customHexDraft('#00FF00')).toBe('00FF00');
    expect(customHexDraft(' #00ff00 ')).toBe('00ff00');
    expect(hexFromText(customHexDraft('#00FF00'))).toBe('00FF00');
    expect(hexFromText(customHexDraft('00F0'))).toBeNull();
  });
});
