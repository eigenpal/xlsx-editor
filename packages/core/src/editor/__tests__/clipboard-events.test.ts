// The clipboard listeners of a grid element (`clipboard-events.ts`) in a happy-dom window of
// their own: which events the grid takes, and how it keeps the focus.

import { describe, expect, test } from 'bun:test';
import { Window } from 'happy-dom';
import type { CopyResult } from '../clipboard.ts';
import { clipboardHandlers, type ClipboardEventHost } from '../clipboard-events.ts';

describe('clipboard listeners', () => {
  test('a grid takes the focus back when the page selection moves it to an editing host', () => {
    const win = new Window();
    const doc = win.document as unknown as Document;
    const host = doc.createElement('div');
    host.setAttribute('contenteditable', 'true');
    host.tabIndex = 0;
    const grid = doc.createElement('div');
    grid.tabIndex = 0;
    host.append(grid);
    doc.body.append(host);
    const { onFocus } = clipboardHandlers({} as ClipboardEventHost, grid);
    let focuses = 0;
    grid.addEventListener('focus', (e) => {
      focuses += 1;
      onFocus(e as FocusEvent);
    });
    // An engine moves the focus to the editing host that holds the new selection.
    const selection = doc.getSelection()!;
    const collapse = selection.collapse.bind(selection);
    selection.collapse = ((node: Node, offset?: number) => {
      collapse(node, offset);
      host.focus();
    }) as typeof selection.collapse;
    grid.focus();
    expect(doc.activeElement).toBe(grid);
    expect(selection.anchorNode).toBe(grid);
    // The second focus finds the selection in the grid, so it does not move the focus again.
    expect(focuses).toBe(2);
    win.close();
  });

  test('a press on the focused grid moves the page selection from host text into the grid', () => {
    const win = new Window();
    const doc = win.document as unknown as Document;
    const text = doc.createElement('p');
    text.textContent = 'Host text';
    const grid = doc.createElement('div');
    grid.tabIndex = 0;
    doc.body.append(text, grid);
    const { onPress } = clipboardHandlers({} as ClipboardEventHost, grid);
    const selection = doc.getSelection()!;
    // With no focus, the press leaves the selection: the focus event moves it.
    selection.selectAllChildren(text);
    onPress();
    expect(selection.anchorNode).toBe(text);
    grid.focus();
    selection.selectAllChildren(text);
    onPress();
    expect(selection.anchorNode).toBe(grid);
    expect(doc.activeElement).toBe(grid);
    win.close();
  });

  test('the listeners act for the focused grid only, once', () => {
    const win = new Window();
    const doc = win.document as unknown as Document;
    const calls: string[] = [];
    const copied = (name: string): CopyResult =>
      ({ ok: true, content: { text: name, token: name } }) as unknown as CopyResult;
    // Two grids on one page, one in a shadow root, each with the listeners of the editor.
    const plain = doc.createElement('div');
    // An ancestor of the host page that stops each clipboard event of its content.
    const wrapper = doc.createElement('div');
    for (const type of ['copy', 'cut', 'paste'])
      wrapper.addEventListener(type, (e) => e.stopPropagation());
    wrapper.append(plain);
    const host = doc.createElement('div');
    const shadow = host.attachShadow({ mode: 'open' });
    const inner = doc.createElement('div');
    shadow.append(inner);
    const editable = doc.createElement('div');
    editable.contentEditable = 'true';
    const field = doc.createElement('textarea');
    doc.body.append(wrapper, host, editable, field);
    for (const [name, grid] of [
      ['plain', plain],
      ['shadow', inner],
    ] as const) {
      grid.tabIndex = 0;
      const h = clipboardHandlers(
        {
          copy: () => (calls.push(`${name} copy`), copied(name)),
          cut: () => (calls.push(`${name} cut`), copied(name)),
          paste: () => (calls.push(`${name} paste`), { ok: true }) as never,
        },
        grid
      );
      grid.addEventListener('copy', h.onCopy as EventListener);
      grid.addEventListener('cut', h.onCut as EventListener);
      grid.addEventListener('paste', h.onPaste as EventListener);
    }
    const fire = (type: string, target: EventTarget) => {
      const e = new (win as unknown as typeof globalThis).Event(type, {
        bubbles: true,
        cancelable: true,
        composed: true,
      });
      const data = new Map<string, string>();
      Object.defineProperty(e, 'clipboardData', {
        value: {
          setData: (t: string, v: string) => void data.set(t, v),
          getData: (t: string) => data.get(t) ?? '',
          types: ['text/plain'],
        },
      });
      target.dispatchEvent(e);
      return e.defaultPrevented;
    };
    // A node inside each grid, and a field inside the plain grid.
    const cell = doc.createElement('div');
    const input = doc.createElement('textarea');
    plain.append(cell, input);
    const innerCell = doc.createElement('div');
    inner.append(innerCell);
    // The plain grid has the focus: a cut at a node inside it is its own, once, and not the
    // other grid's; an event on the rest of the page is no grid's.
    plain.focus();
    expect(fire('cut', cell)).toBe(true);
    expect(fire('cut', doc.body)).toBe(false);
    expect(fire('copy', innerCell)).toBe(false);
    expect(calls).toEqual(['plain cut']);
    // An event at the grid acts, also under an ancestor that stops it.
    expect(fire('copy', plain)).toBe(true);
    expect(calls).toEqual(['plain cut', 'plain copy']);
    // An event at a node inside the grid acts once; one at a field inside it passes through.
    expect(fire('paste', cell)).toBe(true);
    expect(fire('paste', input)).toBe(false);
    expect(calls).toEqual(['plain cut', 'plain copy', 'plain paste']);
    calls.length = 0;
    // Editable content inside the grid keeps its events, also at a text node inside it.
    const own = doc.createElement('div');
    own.setAttribute('contenteditable', 'true');
    const ownText = doc.createTextNode('Own text');
    own.append(ownText);
    plain.append(own);
    expect(fire('paste', own)).toBe(false);
    expect(fire('copy', ownText)).toBe(false);
    expect(calls).toEqual([]);
    // A grid inside editable content of the host page keeps its events, and so does its cell.
    wrapper.setAttribute('contenteditable', 'true');
    expect(fire('copy', plain)).toBe(true);
    expect(fire('paste', cell)).toBe(true);
    expect(fire('paste', own)).toBe(false);
    expect(calls).toEqual(['plain copy', 'plain paste']);
    wrapper.removeAttribute('contenteditable');
    calls.length = 0;
    // Editable content and a field keep their own events.
    expect(fire('paste', editable)).toBe(false);
    expect(fire('copy', field)).toBe(false);
    expect(calls).toEqual([]);
    // The grid in the shadow root has the focus: an event at it, and one at a node inside it.
    calls.length = 0;
    inner.focus();
    expect(fire('copy', inner)).toBe(true);
    expect(fire('paste', innerCell)).toBe(true);
    expect(calls).toEqual(['shadow copy', 'shadow paste']);
    win.close();
  });
});
