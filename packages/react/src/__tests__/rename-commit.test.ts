// `blockNextClick` (design D4 of the change `workbook-structure-editing`): it blocks the click
// of one press, and no later click: a release with no click, a key press, or the next press
// ends it. `focusAfterRelease` gives the focus after one release that leaves it on the body.

import { afterEach, describe, expect, test } from 'bun:test';
import { blockNextClick, focusAfterRelease } from '@xlsx-editor.dev/core/editor';

let mounted: HTMLElement[] = [];
afterEach(() => {
  for (const el of mounted) el.remove();
  mounted = [];
});

function control(): { el: HTMLButtonElement; clicks: () => number } {
  const el = document.createElement('button');
  document.body.appendChild(el);
  mounted.push(el);
  let n = 0;
  el.addEventListener('click', () => (n += 1));
  return { el, clicks: () => n };
}

const tick = () => new Promise((resolve) => setTimeout(resolve, 5));

describe('blockNextClick', () => {
  test('blocks the click of the press', () => {
    const { el, clicks } = control();
    blockNextClick(el);
    el.dispatchEvent(new MouseEvent('mouseup', { bubbles: true }));
    el.click();
    expect(clicks()).toBe(0);
    el.click();
    expect(clicks()).toBe(1);
  });

  test('a release away with no click blocks no later keyboard click', async () => {
    const { el, clicks } = control();
    const away = control();
    blockNextClick(el);
    away.el.dispatchEvent(new MouseEvent('mouseup', { bubbles: true }));
    await tick();
    el.click();
    expect(clicks()).toBe(1);
  });

  test('a key press ends the block', () => {
    const { el, clicks } = control();
    blockNextClick(el);
    el.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
    el.click();
    expect(clicks()).toBe(1);
  });
});

describe('focusAfterRelease', () => {
  /** The calls of `focus` after a release on `at`, when the task of the release ends. */
  async function release(at: HTMLElement, calls: () => number): Promise<number> {
    at.dispatchEvent(new MouseEvent('mouseup', { bubbles: true }));
    await new Promise((done) => setTimeout(done, 0));
    return calls();
  }

  function armed() {
    const { el } = control();
    (document.activeElement as HTMLElement | null)?.blur();
    let n = 0;
    focusAfterRelease(el, () => (n += 1));
    return { el, calls: () => n };
  }

  test('gives the focus after a release that leaves it on the body, once', async () => {
    const { el, calls } = armed();
    expect(await release(document.body, calls)).toBe(1);
    expect(await release(el, calls)).toBe(1);
  });

  test('gives no focus when the click of the release gave it', async () => {
    const { el, calls } = armed();
    el.addEventListener('mouseup', () => el.focus());
    expect(await release(el, calls)).toBe(0);
  });

  test('the next press or a key press ends it', async () => {
    for (const end of [
      new MouseEvent('mousedown', { bubbles: true }),
      new KeyboardEvent('keydown', { key: 'Tab', bubbles: true }),
    ]) {
      const { el, calls } = armed();
      el.dispatchEvent(end);
      expect(await release(document.body, calls)).toBe(0);
    }
  });
});
