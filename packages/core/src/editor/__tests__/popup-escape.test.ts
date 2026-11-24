// Escape for a toolbar popup that a pointer press opened (`popup-escape.ts`): one Escape closes
// the popup, wherever the keyboard is. A key from this editor stops there, so the grid does not
// spend it too. A key from the host page goes on, with its default prevented, so a host layer
// that closes on Escape (a side panel that holds the editor) can see that the key closed the
// popup and stay open.

import { afterEach, describe, expect, test } from 'bun:test';
import { Window } from 'happy-dom';
import { listenForPopupEscape } from '../index.ts';

const stops: (() => void)[] = [];
afterEach(() => stops.splice(0).forEach((stop) => stop()));

/** An editor instance with a grid and a toolbar popup, and a host button outside the editor. */
function page() {
  const win = new Window();
  const doc = win.document as unknown as Document;
  doc.body.innerHTML = `
    <div class="host-panel">
      <div class="xlsx-editor">
        <div class="xlsx-toolbar"><div class="popup"><button class="option">11</button></div></div>
        <div class="xlsx-surface"><div class="xlsx-scroller" tabindex="0"></div></div>
      </div>
    </div>
    <button class="host-card">Open</button>`;
  const $ = (s: string) => doc.querySelector<HTMLElement>(s)!;
  return { doc, $, popup: $('.popup') };
}

/** Listens for the popup, and records each close and what a later host listener sees. */
function listen(doc: Document, popup: HTMLElement) {
  const closes: boolean[] = [];
  const host: { prevented: boolean }[] = [];
  stops.push(
    listenForPopupEscape({
      popup,
      contains: (node) => popup.contains(node),
      close: (fromInside) => closes.push(fromInside),
    })
  );
  const onHost = (e: Event) => host.push({ prevented: e.defaultPrevented });
  doc.defaultView!.addEventListener('keydown', onHost);
  stops.push(() => doc.defaultView!.removeEventListener('keydown', onHost));
  return { closes, host };
}

function escape(el: Element) {
  const View = el.ownerDocument!.defaultView as unknown as typeof globalThis;
  const event = new View.KeyboardEvent('keydown', {
    key: 'Escape',
    bubbles: true,
    cancelable: true,
  });
  el.dispatchEvent(event);
  return event;
}

describe('listenForPopupEscape', () => {
  test('an Escape from the grid closes the popup and stops there', () => {
    const { doc, $, popup } = page();
    const { closes, host } = listen(doc, popup);
    const event = escape($('.xlsx-scroller'));
    expect(closes).toEqual([false]);
    expect(event.defaultPrevented).toBe(true);
    expect(host).toEqual([]);
  });

  test('an Escape from inside the popup closes it and stops there', () => {
    const { doc, $, popup } = page();
    const { closes, host } = listen(doc, popup);
    escape($('.option'));
    expect(closes).toEqual([true]);
    expect(host).toEqual([]);
  });

  test('an Escape from the host closes the popup, and the host sees it as handled', () => {
    const { doc, $, popup } = page();
    const { closes, host } = listen(doc, popup);
    const event = escape($('.host-card'));
    expect(closes).toEqual([false]);
    expect(event.defaultPrevented).toBe(true);
    // The key goes on to the host, which can keep its own layer open.
    expect(host).toEqual([{ prevented: true }]);
  });

  test('a popup that skips the key, for a nested popup, leaves it to the next listener', () => {
    const { $, popup } = page();
    const closes: boolean[] = [];
    stops.push(
      listenForPopupEscape({
        popup,
        contains: (node) => popup.contains(node),
        skip: () => true,
        close: (fromInside) => closes.push(fromInside),
      })
    );
    const event = escape($('.host-card'));
    expect(closes).toEqual([]);
    expect(event.defaultPrevented).toBe(false);
  });

  test('an Escape during an input method composition does nothing', () => {
    const { doc, $, popup } = page();
    const { closes } = listen(doc, popup);
    const View = doc.defaultView as unknown as typeof globalThis;
    $('.host-card').dispatchEvent(
      new View.KeyboardEvent('keydown', { key: 'Escape', bubbles: true, isComposing: true })
    );
    expect(closes).toEqual([]);
  });
});
