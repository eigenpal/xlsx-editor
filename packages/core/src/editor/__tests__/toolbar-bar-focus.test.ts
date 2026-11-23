// The arrow keys of the bar (`moveBarFocus`) where the adapter tests do not reach: a bar in a
// shadow root, and a control that hands its focus on. The core run has no global DOM
// (`bun run test`), so the file uses a DOM of its own.

import { describe, expect, test } from 'bun:test';
import { GlobalWindow } from 'happy-dom';
import { moveBarFocus } from '../toolbar-layout.ts';

const win = new GlobalWindow({ width: 1000, height: 800 });
const document = win.document as unknown as Document;

/** A bar with one button for each name, in `parent`. */
function bar(parent: ParentNode, names: readonly string[]): HTMLElement {
  const el = document.createElement('div');
  for (const name of names) {
    const button = document.createElement('button');
    button.textContent = name;
    el.append(button);
  }
  parent.append(el);
  return el;
}

const control = (el: HTMLElement, name: string) =>
  [...el.querySelectorAll<HTMLElement>('button')].find((b) => b.textContent === name)!;

/** Press `key` on `target` in `el`; true when the bar took the key. */
function press(el: HTMLElement, target: HTMLElement, key: string): boolean {
  let prevented = false;
  moveBarFocus({ key, target, currentTarget: el, preventDefault: () => (prevented = true) });
  return prevented;
}

describe('the arrow keys of the bar', () => {
  test('in a shadow root, pass over a control that refuses the focus', () => {
    const host = document.createElement('div');
    document.body.append(host);
    const root = host.attachShadow({ mode: 'open' });
    const el = bar(root, ['A', 'Ghost', 'B']);
    control(el, 'Ghost').focus = () => undefined;
    control(el, 'A').focus();
    expect(root.activeElement).toBe(control(el, 'A'));
    expect(press(el, control(el, 'A'), 'ArrowRight')).toBe(true);
    expect(root.activeElement).toBe(control(el, 'B'));
    press(el, control(el, 'B'), 'ArrowLeft');
    expect(root.activeElement).toBe(control(el, 'A'));
    host.remove();
  });

  test('stop at a control that hands its focus on, inside it or out of the bar', () => {
    const el = bar(document.body, ['A', 'Widget', 'Picker', 'B']);
    // A host widget focuses its own field, which is no stop of the bar.
    const inner = document.createElement('input');
    inner.tabIndex = -1;
    const widget = control(el, 'Widget');
    widget.append(inner);
    widget.addEventListener('focus', () => inner.focus());
    // A host picker opens a field outside the bar.
    const outside = document.createElement('input');
    document.body.append(outside);
    const picker = control(el, 'Picker');
    picker.addEventListener('focus', () => outside.focus());

    control(el, 'A').focus();
    press(el, control(el, 'A'), 'ArrowRight');
    expect(document.activeElement).toBe(inner);
    control(el, 'B').focus();
    press(el, control(el, 'B'), 'ArrowLeft');
    expect(document.activeElement).toBe(outside);
    el.remove();
    outside.remove();
  });
});
