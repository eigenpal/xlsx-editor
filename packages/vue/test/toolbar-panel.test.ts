// The overflow panel of the toolbar in happy-dom (`bun run test:dom`): the groups that do not
// fit open in the More panel, its lists render in flow, and a list taller than the visible
// panel, or than a host container that clips it, scrolls inside itself. The Vue twin of the
// React adapter's `toolbar-panel.test.tsx`.

import { afterEach, describe, expect, test } from 'bun:test';
import { fireEvent, waitFor } from '@testing-library/react';
import { nextTick } from 'vue';
import type { Translations } from '@xlsx-editor.dev/i18n';
import { unmountAll } from './mount.ts';
import {
  type Clip,
  clipBox,
  groupWidth,
  openNarrow,
  type Rect,
  stubLayout,
  stubScrollHeight,
  workbook,
} from './toolbar-stubs.ts';

afterEach(unmountAll);

describe('the overflow panel', () => {
  test('a short bar moves the font family field before the font group', async () => {
    // History is 60 pixels, the font group 380 with its 150 pixel family field, and every other
    // group too wide to stay. In 340 pixels the font group fits without the field.
    const rects = (el: HTMLElement): Rect | null => {
      if (el.dataset.slot === 'font.family') return { width: 150 };
      const group = el.dataset.group;
      if (!group) return null;
      if (group === 'history') return { width: 60 };
      if (group === 'font')
        return { width: el.querySelector('[data-slot="font.family"]') ? 380 : 230 };
      return { width: 1000 };
    };
    const restore = stubLayout(rects, 340);
    try {
      const { view } = await openNarrow(340);
      const bar = view.getByRole('toolbar');
      await waitFor(() =>
        expect(
          [...bar.querySelectorAll<HTMLElement>('[data-group]')].map((g) => g.dataset.group)
        ).toEqual(['history', 'font'])
      );
      const font = bar.querySelector('[data-group="font"]')!;
      expect(font.querySelector('[data-slot="font.family"]')).toBeNull();
      expect(font.querySelector('[data-slot="font.bold"]')).not.toBeNull();
      fireEvent.click(view.getByRole('button', { name: 'More' }));
      const panel = await waitFor(() => view.getByRole('dialog', { name: 'More' }));
      const sections = [...panel.querySelectorAll<HTMLElement>('[role="group"]')];
      expect(sections[0]!.getAttribute('aria-label')).toBe('Font');
      expect(sections[0]!.querySelector('[data-slot="font.family"]')).not.toBeNull();
      expect(sections[0]!.querySelector('[data-slot="font.bold"]')).toBeNull();
    } finally {
      restore();
    }
  });

  test('a list in the panel renders in flow, and the keyboard reaches it', async () => {
    const restore = stubLayout(groupWidth, 400);
    try {
      const { view } = await openNarrow(400);
      const bar = view.getByRole('toolbar');
      await waitFor(() => expect(bar.querySelectorAll('[data-group]').length).toBe(5));
      const more = view.getByRole('button', { name: 'More' });
      more.focus();
      fireEvent.keyDown(more, { key: 'Enter' });
      await nextTick();
      const panel = view.getByRole('dialog', { name: 'More' });
      await waitFor(() => expect(panel.contains(document.activeElement)).toBe(true));
      const trigger = view.getByRole('button', { name: /^Number format/ });
      expect(document.activeElement).toBe(trigger);
      fireEvent.keyDown(trigger, { key: 'ArrowDown' });
      await nextTick();
      const list = view.getByRole('listbox', { name: 'Number format' });
      expect(panel.contains(list)).toBe(true);
      expect(list.closest('[data-popup]')?.hasAttribute('data-inline')).toBe(true);
      await waitFor(() => expect(list.contains(document.activeElement)).toBe(true));
    } finally {
      restore();
    }
  });

  test('a list that opens below the visible panel scrolls the panel, and keeps focus', async () => {
    // The panel shows 300 pixels; the list opens in flow 400 pixels down, 200 pixels tall.
    const rects = (el: HTMLElement): Rect | null =>
      el.classList.contains('xlsx-toolbar-panel')
        ? { top: 0, bottom: 300, height: 300 }
        : el.hasAttribute('data-inline')
          ? { top: 400, bottom: 600, height: 200 }
          : groupWidth(el);
    const restore = stubLayout(rects, 400);
    try {
      const { view } = await openNarrow(400);
      const bar = view.getByRole('toolbar');
      await waitFor(() => expect(bar.querySelectorAll('[data-group]').length).toBe(5));
      fireEvent.click(view.getByRole('button', { name: 'More' }));
      await nextTick();
      const panel = view.getByRole('dialog', { name: 'More' });
      panel.style.overflowY = 'auto';
      Object.defineProperty(panel, 'clientHeight', { configurable: true, value: 300 });
      Object.defineProperty(panel, 'scrollHeight', { configurable: true, value: 900 });
      const scrolls: { top?: number }[] = [];
      panel.scrollBy = ((o: { top?: number }) => void scrolls.push(o)) as typeof panel.scrollBy;
      const trigger = view.getByRole('button', { name: /^Number format/ });
      trigger.focus();
      fireEvent.click(trigger);
      await nextTick();
      expect(view.getByRole('listbox', { name: 'Number format' })).toBeTruthy();
      // The least scroll that shows the whole list: its end, 600, to the panel's end, 300.
      expect(scrolls.map((o) => o.top)).toEqual([300]);
      expect(document.activeElement).toBe(trigger);
    } finally {
      restore();
    }
  });

  /**
   * Open the popup of `control` in the overflow panel of a bar of `barWidth` pixels that shows
   * `groups` groups. The panel shows `panelBottom` pixels, popups are 500 pixels, and items
   * (options and swatches) are 24 pixels, and a size entry is 28 pixels. A host container
   * around the editor shows `host.clip` pixels, the editor shows at `host.scale` times its
   * layout size, and the panel has a padding of `host.padding` pixels. Gives the popup.
   */
  async function openInPanel(
    control: string | RegExp,
    listName: string,
    panelBottom: number,
    [barWidth, groups] = [400, 5],
    host: { clip?: Clip; scale?: number; padding?: number } = {}
  ) {
    const scale = host.scale ?? 1;
    const rects = (el: HTMLElement): Rect | null =>
      scale !== 1 && el.classList.contains('xlsx-editor')
        ? { width: 400 * scale, height: 400 * scale }
        : host.clip !== undefined && el.hasAttribute('data-test-clip')
          ? clipBox(host.clip)
          : el.classList.contains('xlsx-toolbar-panel')
            ? { top: 0, bottom: panelBottom, height: panelBottom }
            : el.getAttribute('role') === 'option' || el.hasAttribute('data-color')
              ? { height: 24 }
              : el.classList.contains('xlsx-toolbar-entry')
                ? { height: 28 }
                : groupWidth(el);
    const restore = stubLayout(rects, barWidth);
    const unstub = stubScrollHeight((el) => (el.hasAttribute('data-popup') ? 500 : null));
    try {
      const { view } = await openNarrow(barWidth);
      const bar = view.getByRole('toolbar');
      await waitFor(() => expect(bar.querySelectorAll('[data-group]').length).toBe(groups));
      if (host.clip !== undefined) {
        const clip = view.container.firstElementChild as HTMLElement;
        clip.style.overflow = 'hidden';
        clip.setAttribute('data-test-clip', '');
      }
      const editor = view.container.querySelector<HTMLElement>('.xlsx-editor')!;
      for (const size of ['offsetWidth', 'offsetHeight'])
        Object.defineProperty(editor, size, { configurable: true, value: 400 });
      fireEvent.click(view.getByRole('button', { name: 'More' }));
      await nextTick();
      if (host.padding !== undefined) view.getByRole('dialog').style.padding = `${host.padding}px`;
      fireEvent.click(view.getByRole('button', { name: control }));
      await nextTick();
      // A list is a listbox, and the color picker is a dialog.
      const popup =
        view.queryByRole('listbox', { name: listName }) ??
        view.getByRole('dialog', { name: listName });
      return popup.closest<HTMLElement>('[data-popup]')!;
    } finally {
      unstub();
      restore();
    }
  }

  test('a list taller than the visible panel scrolls inside itself', async () => {
    const popup = await openInPanel(/^Number format/, 'Number format', 300);
    expect(popup.style.maxHeight).toBe('300px');
    expect(popup.hasAttribute('data-capped')).toBe(true);
  });

  test('in a host container that clips its overflow, a list in the panel is capped to it', async () => {
    // A panel of 300 pixels in a container of 200: the list shows 200 pixels.
    const popup = await openInPanel(/^Number format/, 'Number format', 300, undefined, {
      clip: 200,
    });
    expect(popup.style.maxHeight).toBe('200px');
  });

  test('in a host container that clips the top of the panel, a list in the panel is capped to it', async () => {
    // A panel from 0 to 300 in a container from 100 to 260: the list shows 160 pixels.
    const popup = await openInPanel(/^Number format/, 'Number format', 300, undefined, {
      clip: [100, 260],
    });
    expect(popup.style.maxHeight).toBe('160px');
  });

  test('inside a host that scales the editor by 2, a list in the panel is capped in its own pixels', async () => {
    // A panel of 300 window pixels, less its padding of 8 twice at scale 2: 268 window pixels, or
    // 134 of the list's pixels.
    const popup = await openInPanel(/^Number format/, 'Number format', 300, undefined, {
      scale: 2,
      padding: 8,
    });
    expect(popup.style.maxHeight).toBe('134px');
  });

  test('a palette taller than the visible panel scrolls inside itself too', async () => {
    // In a bar of 100 pixels, no group fits: the fill group is in the panel.
    const popup = await openInPanel('Fill color options', 'Fill color', 300, [100, 0]);
    expect(popup.closest('[role="dialog"]')).not.toBeNull();
    expect(popup.style.maxHeight).toBe('300px');
  });

  test('a panel that shows less than one item caps nothing', async () => {
    const popup = await openInPanel(/^Number format/, 'Number format', 20);
    expect(popup.style.maxHeight).toBe('');
    expect(popup.hasAttribute('data-capped')).toBe(false);
  });

  test('a room for one item but not for the entry above it caps nothing', async () => {
    // In a bar of 100 pixels, the font size list is in the panel. Its entry (28) and one item
    // (24) need 52 pixels: a room of 40 shows no item, and a room of 60 shows one.
    const short = await openInPanel(/^Font size/, 'Font size', 40, [100, 0]);
    expect(short.style.maxHeight).toBe('');
    unmountAll();
    const room = await openInPanel(/^Font size/, 'Font size', 60, [100, 0]);
    expect(room.style.maxHeight).toBe('60px');
  });

  test('a mouse open of the number format dialog closes the panel', async () => {
    const restore = stubLayout(groupWidth, 400);
    try {
      const { view } = await openNarrow(400);
      const bar = view.getByRole('toolbar');
      await waitFor(() => expect(bar.querySelectorAll('[data-group]').length).toBe(5));
      fireEvent.click(view.getByRole('button', { name: 'More' }));
      await nextTick();
      fireEvent.click(view.getByRole('button', { name: /^Number format/ }));
      await nextTick();
      fireEvent.click(view.getByRole('button', { name: 'More number formats…' }));
      await nextTick();
      expect(view.getByRole('dialog', { name: 'Number format' })).toBeTruthy();
      expect(view.queryByRole('dialog', { name: 'More' })).toBeNull();
    } finally {
      restore();
    }
  });

  test('an open of another workbook closes the panel', async () => {
    const restore = stubLayout(groupWidth, 400);
    try {
      const { view, viewer } = await openNarrow(400);
      const bar = view.getByRole('toolbar');
      await waitFor(() => expect(bar.querySelectorAll('[data-group]').length).toBe(5));
      fireEvent.click(view.getByRole('button', { name: 'More' }));
      await nextTick();
      expect(view.getByRole('dialog', { name: 'More' })).toBeTruthy();
      await viewer.load(workbook());
      await nextTick();
      expect(view.queryByRole('dialog', { name: 'More' })).toBeNull();
    } finally {
      restore();
    }
  });

  test('the groups are measured again after a change of locale', async () => {
    const restore = stubLayout(groupWidth, 200);
    try {
      const { view, rerender } = await openNarrow(200);
      const bar = view.getByRole('toolbar');
      await waitFor(() => expect(bar.querySelectorAll('[data-group]').length).toBe(3));
      const short = Object.fromEntries(
        ['history', 'font', 'fill', 'borders', 'alignment', 'number', 'cells', 'clear', 'file'].map(
          (g) => [g, g.slice(0, 1)]
        )
      );
      // A group that the bar never showed is measured again too: its cached width is stale.
      await rerender({ formattingBar: { groups: short } } as Translations);
      await waitFor(() => expect(bar.querySelectorAll('[data-group]').length).toBe(9));
      expect(view.queryByRole('button', { name: 'More' })).toBeNull();
    } finally {
      restore();
    }
  });
  test('inside a host that scales the editor by 2, the bar fits the same groups', async () => {
    // The bar and its groups show at twice their layout size: a group of 60 layout pixels is 120
    // window pixels wide. The bar has 400 layout pixels, as at scale 1, so 5 groups fit.
    const restore = stubLayout((el) => {
      if (el.classList.contains('xlsx-toolbar')) return { width: 800, height: 80 };
      const width = groupWidth(el)?.width;
      return width === undefined ? null : { width: width * 2 };
    }, 400);
    const proto = HTMLElement.prototype as unknown as Record<string, unknown>;
    const sizes = ['offsetWidth', 'offsetHeight'] as const;
    const own = sizes.map((size) => Object.getOwnPropertyDescriptor(HTMLElement.prototype, size)!);
    sizes.forEach((size, i) =>
      Object.defineProperty(proto, size, {
        configurable: true,
        get(this: HTMLElement) {
          if (this.classList.contains('xlsx-toolbar')) return size === 'offsetWidth' ? 400 : 40;
          return own[i]!.get!.call(this);
        },
      })
    );
    try {
      const { view } = await openNarrow(400);
      const bar = view.getByRole('toolbar');
      await waitFor(() => expect(bar.querySelectorAll('[data-group]').length).toBe(5));
    } finally {
      sizes.forEach((size, i) => Object.defineProperty(proto, size, own[i]!));
      restore();
    }
  });
});
