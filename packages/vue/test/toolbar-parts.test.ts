// The parts of the toolbar in happy-dom (`bun run test:dom`). The Vue twin of the React adapter's
// `toolbar-parts.test.tsx`, with the cases that only a Vue host can write: a part's props as a
// template passes them (`slot-id`, a bare `hidden`), a template `class` beside `className`, a
// `v-if` branch that renders nothing, and content as an array, a vnode, or a string. The
// development check itself is in toolbar-dev-warnings.test.ts.

import { afterEach, describe, expect, test } from 'bun:test';
import { fireEvent, waitFor } from '@testing-library/react';
import {
  Comment,
  Fragment,
  createVNode,
  defineComponent,
  h,
  nextTick,
  ref,
  type VNodeChild,
} from 'vue';
import { buildXlsx, c, row } from '../../../scripts/fixtures/builder.ts';
import {
  StatusMessage,
  XlsxEditorContent,
  XlsxEditorRoot,
  XlsxEditorToolbar,
  XlsxEditorToolbarGroup,
  XlsxEditorToolbarSlot,
  XlsxEditorViewport,
  type XlsxEditorInstance,
} from '../src/index.ts';
import { FONTS, mount, unmountAll } from './mount.ts';

afterEach(unmountAll);

const workbook = () => buildXlsx({ sheets: [{ name: 'S', rows: row(1, [c('A1', 1)]) }] });

/** Mount the editor with `toolbar` as its toolbar, and wait for the workbook. */
async function open(toolbar: () => VNodeChild) {
  let viewer: XlsxEditorInstance | null = null;
  const view = mount(() =>
    h(
      XlsxEditorRoot,
      { document: workbook(), fonts: FONTS, onReady: (v: XlsxEditorInstance) => (viewer = v) },
      {
        default: () =>
          h(XlsxEditorViewport, { style: { height: '400px' } }, () => [
            toolbar(),
            h(XlsxEditorContent),
            h(StatusMessage),
          ]),
      }
    )
  );
  await waitFor(() => expect(viewer?.snapshot().status).toBe('ready'), { timeout: 5000 });
  await nextTick();
  return { view, bar: view.getByRole('toolbar'), viewer: viewer! };
}

/** The ids of the groups that the bar shows, in order. */
const barGroups = (bar: HTMLElement) =>
  [...bar.querySelectorAll<HTMLElement>(':scope > [data-group]')].map((g) => g.dataset.group);

/** The slots of one group of the bar, in order. */
const slotsOf = (bar: HTMLElement, group: string) =>
  [...bar.querySelectorAll<HTMLElement>(`[data-group="${group}"] [data-slot]`)].map(
    (s) => s.dataset.slot
  );

const BUILT_IN = [
  'history',
  'font',
  'fill',
  'borders',
  'alignment',
  'number',
  'cells',
  'clear',
  'file',
];

const button = (text: string, slot?: string) =>
  h('button', { type: 'button', ...(slot ? { 'data-slot': slot } : {}) }, text);

/**
 * Props as a compiled template passes them: kebab-case names, and a bare boolean attribute as
 * an empty string. The prop types of `h` do not describe them.
 */
const template = (props: Record<string, unknown>) => props as never;

/** Collect `console.warn` messages while `run` runs. */
async function warnings(run: () => Promise<void>): Promise<string[]> {
  const seen: string[] = [];
  const original = console.warn;
  console.warn = (message: string) => void seen.push(message);
  try {
    await run();
  } finally {
    console.warn = original;
  }
  return seen;
}

/** A toolbar layout for happy-dom, which has none: rectangles by element, the rest unchanged. */
type Rect = Partial<DOMRect>;
function stubLayout(rectOf: (el: HTMLElement) => Rect | null, barWidth: number): () => void {
  const proto = HTMLElement.prototype as unknown as Record<string, unknown>;
  const rect = Element.prototype.getBoundingClientRect;
  const width = Object.getOwnPropertyDescriptor(HTMLElement.prototype, 'clientWidth')!;
  Element.prototype.getBoundingClientRect = function (this: HTMLElement) {
    const r = rectOf(this);
    return (
      r
        ? { x: 0, y: 0, top: 0, left: 0, right: 0, bottom: 0, width: 0, height: 0, ...r }
        : rect.call(this)
    ) as DOMRect;
  };
  Object.defineProperty(proto, 'clientWidth', {
    configurable: true,
    get(this: HTMLElement) {
      return this.classList.contains('xlsx-toolbar') ? barWidth : width.get!.call(this);
    },
  });
  return () => {
    Element.prototype.getBoundingClientRect = rect;
    Object.defineProperty(proto, 'clientWidth', width);
  };
}

/**
 * A group is ten pixels for each letter of its heading, and the fixed block is `fixed` pixels.
 * The built-in headings in English add up to 700 pixels.
 */
const widths =
  (fixed = 0) =>
  (el: HTMLElement): Rect | null =>
    el.dataset.group
      ? { width: (el.getAttribute('aria-label') ?? '').length * 10 }
      : el.hasAttribute('data-toolbar-fixed')
        ? { width: fixed }
        : null;

describe('XlsxEditorToolbar', () => {
  test('takes a host class and style, which its type accepts as for a component', async () => {
    const props: Parameters<typeof XlsxEditorToolbar>[0] = { class: 'mine', style: { order: 1 } };
    const { bar } = await open(() => h(XlsxEditorToolbar, props));
    expect(bar.classList.contains('mine')).toBe(true);
    expect(bar.style.order).toBe('1');
  });
});

describe('XlsxEditorToolbar.Slot', () => {
  test('replaces its slot where it stands, in the bar', async () => {
    const { bar } = await open(() =>
      h(XlsxEditorToolbar, null, () => [
        h(XlsxEditorToolbar.Slot, { slotId: 'font.bold' }, () => button('Strong', 'host.bold')),
      ])
    );
    expect(slotsOf(bar, 'font')).toEqual([
      'font.family',
      'font.size',
      'host.bold',
      'font.italic',
      'font.underline',
      'font.color',
    ]);
    expect(barGroups(bar)).toEqual(BUILT_IN);
  });

  test('takes its props as a template passes them: slot-id and a bare hidden', async () => {
    const { bar } = await open(() =>
      h(XlsxEditorToolbar, null, () => [
        h(XlsxEditorToolbarSlot, template({ 'slot-id': 'font.underline', hidden: '' })),
        h(XlsxEditorToolbarSlot, template({ 'slot-id': 'file.save', hidden: '' })),
      ])
    );
    expect(slotsOf(bar, 'font')).not.toContain('font.underline');
    expect(barGroups(bar)).toEqual(BUILT_IN.filter((id) => id !== 'file'));
  });

  test('with a v-if branch that renders nothing, keeps the built-in control', async () => {
    const { bar } = await open(() =>
      h(XlsxEditorToolbar, null, () => [
        h(XlsxEditorToolbarSlot, { slotId: 'font.bold' }, () => [
          createVNode(Comment, null, 'v-if'),
        ]),
      ])
    );
    expect(slotsOf(bar, 'font')).toContain('font.bold');
  });

  test('the parts inside a Fragment are read as parts, and the last part for a slot wins', async () => {
    const { bar } = await open(() =>
      h(XlsxEditorToolbar, null, () => [
        h(XlsxEditorToolbarSlot, { slotId: 'font.bold', hidden: true }),
        h(Fragment, null, [
          h(XlsxEditorToolbarSlot, { slotId: 'font.bold' }, () => button('Strong', 'host.bold')),
          h(Fragment, null, [
            h(XlsxEditorToolbarSlot, { slotId: 'font.italic', hidden: true }),
            h(XlsxEditorToolbarGroup, { id: 'borders', hidden: true }),
          ]),
        ]),
      ])
    );
    expect(slotsOf(bar, 'font')).toContain('host.bold');
    expect(slotsOf(bar, 'font')).not.toContain('font.bold');
    expect(slotsOf(bar, 'font')).not.toContain('font.italic');
    expect(barGroups(bar)).not.toContain('borders');
    expect(bar.querySelector('[data-toolbar-fixed]')).toBeNull();
  });

  test('takes content as an array, a vnode, or a string, as a render function passes it', async () => {
    const { bar } = await open(() =>
      h(XlsxEditorToolbar, null, () => [
        h(XlsxEditorToolbarSlot, { slotId: 'font.bold' }, [button('Strong', 'host.bold')]),
        h(XlsxEditorToolbarSlot, { slotId: 'font.italic' }, button('Slanted', 'host.italic')),
        h(XlsxEditorToolbarGroup, { id: 'note', label: 'Note' }, 'Draft'),
        h(XlsxEditorToolbarGroup, { id: 'empty', label: 'Empty' }, ''),
      ])
    );
    expect(slotsOf(bar, 'font')).toEqual([
      'font.family',
      'font.size',
      'host.bold',
      'host.italic',
      'font.underline',
      'font.color',
    ]);
    expect(barGroups(bar)).toEqual([...BUILT_IN, 'note']);
    expect(bar.querySelector('[data-group="note"]')?.textContent).toBe('Draft');
  });

  test('the hidden list wins over its content, with one warning for the page', async () => {
    const toolbar = () =>
      h(XlsxEditorToolbar, { hidden: ['font.italic'] }, () => [
        h(XlsxEditorToolbarSlot, { slotId: 'font.italic' }, () => button('Slanted')),
      ]);
    const seen = await warnings(async () => {
      const { bar, view } = await open(toolbar);
      expect(view.queryByRole('button', { name: 'Slanted' })).toBeNull();
      expect(slotsOf(bar, 'font')).not.toContain('font.italic');
      // A new toolbar with the same problem prints nothing new.
      unmountAll();
      await open(toolbar);
    });
    expect(seen).toEqual([
      '[xlsx-editor] XlsxEditorToolbar.Slot "font.italic" has content, but the hidden list ' +
        'names the slot, so the slot is left out. Remove it from the hidden list to show the ' +
        'content.',
    ]);
  });

  test('an id outside the registry shows its content after the groups, with a warning', async () => {
    let bar!: HTMLElement;
    const seen = await warnings(async () => {
      ({ bar } = await open(() =>
        h(XlsxEditorToolbar, null, () => [
          h(XlsxEditorToolbarSlot, { slotId: 'font.strike' as 'font.bold' }, () =>
            button('Strike')
          ),
        ])
      ));
    });
    const fixed = bar.querySelector('[data-toolbar-fixed]')!;
    expect(fixed.textContent).toBe('Strike');
    expect(bar.lastElementChild).toBe(fixed);
    expect(seen.length).toBe(1);
    expect(seen[0]).toStartWith('[xlsx-editor] "font.strike" is not a toolbar slot id');
  });

  test('outside the preset, an id outside the registry renders nothing, with a warning', async () => {
    let bar!: HTMLElement;
    const seen = await warnings(async () => {
      ({ bar } = await open(() =>
        h(XlsxEditorToolbar, { preset: false }, () => [
          h(XlsxEditorToolbarSlot, { slotId: 'font.shadow' as 'font.bold' }),
        ])
      ));
    });
    expect(bar.childElementCount).toBe(0);
    expect(seen).toEqual([
      '[xlsx-editor] "font.shadow" is not a toolbar slot id, so its XlsxEditorToolbar.Slot ' +
        'renders nothing. Use a ChromeSlotId, such as "font.bold".',
    ]);
  });

  test('a replaced number format list leaves the dialog opener as its own button', async () => {
    const { bar, view } = await open(() =>
      h(XlsxEditorToolbar, null, () => [
        h(XlsxEditorToolbarSlot, { slotId: 'number.format' }, () => button('Formats')),
      ])
    );
    expect(slotsOf(bar, 'number')).toContain('number.custom');
    expect(view.getByRole('button', { name: 'More number formats…' })).toBeTruthy();
  });

  test('in a menu group, the content follows the menu, which keeps the other items', async () => {
    const { bar, view } = await open(() =>
      h(XlsxEditorToolbar, null, () => [
        h(XlsxEditorToolbarSlot, { slotId: 'clear.all' }, () => button('Wipe')),
      ])
    );
    const group = bar.querySelector('[data-group="clear"]')!;
    expect(group.lastElementChild?.textContent).toBe('Wipe');
    fireEvent.click(group.querySelector('button')!);
    await nextTick();
    const items = view.getAllByRole('menuitem').map((m) => m.textContent);
    expect(items).toEqual(['Clear contents', 'Clear formats']);
  });

  test('a change of the parts or of the hidden list rearranges the bar', async () => {
    const replace = ref(false);
    const hidden = ref<'font.italic'[]>([]);
    const share = ref(false);
    const { bar } = await open(() =>
      h(XlsxEditorToolbar, { hidden: hidden.value }, () => [
        replace.value
          ? h(XlsxEditorToolbarSlot, { slotId: 'font.bold' }, () => button('Strong', 'host.bold'))
          : null,
        share.value
          ? h(XlsxEditorToolbarGroup, { id: 'share', label: 'Share' }, () => button('Share'))
          : null,
      ])
    );
    expect(slotsOf(bar, 'font')).toContain('font.bold');
    replace.value = true;
    await nextTick();
    await nextTick();
    expect(slotsOf(bar, 'font')).toContain('host.bold');
    expect(slotsOf(bar, 'font')).not.toContain('font.bold');
    hidden.value = ['font.italic'];
    share.value = true;
    await waitFor(() => expect(barGroups(bar)).toEqual([...BUILT_IN, 'share']));
    expect(slotsOf(bar, 'font')).toEqual([
      'font.family',
      'font.size',
      'host.bold',
      'font.underline',
      'font.color',
    ]);
  });

  test('loose children keep their own state when a sibling before them goes', async () => {
    const show = ref(true);
    const { view } = await open(() =>
      h(XlsxEditorToolbar, null, () => [
        // As a render function writes a branch, and as a template compiles `<template v-if>`.
        show.value ? h('input', { 'aria-label': 'A' }) : null,
        show.value
          ? h(Fragment, { key: 0 }, [h('input', { 'aria-label': 'C' })])
          : createVNode(Comment, null, 'v-if'),
        h('input', { 'aria-label': 'B' }),
      ])
    );
    const b = view.getByRole('textbox', { name: 'B' }) as HTMLInputElement;
    b.value = 'typed';
    show.value = false;
    await nextTick();
    await nextTick();
    expect(view.queryByRole('textbox', { name: 'A' })).toBeNull();
    expect(view.getByRole('textbox', { name: 'B' })).toBe(b);
    expect(b.value).toBe('typed');
  });

  test('a loose child and the content of a slot outside the registry keep their elements', async () => {
    const show = ref(true);
    let view!: Awaited<ReturnType<typeof open>>['view'];
    await warnings(async () => {
      ({ view } = await open(() =>
        h(XlsxEditorToolbar, null, () => [
          show.value ? h('input', { 'aria-label': 'L' }) : null,
          show.value
            ? h(XlsxEditorToolbarSlot, { slotId: 'x.w' as 'font.bold' }, () => [
                h('input', { 'aria-label': 'W' }),
              ])
            : null,
          h(XlsxEditorToolbarSlot, { slotId: 'x.y' as 'font.bold' }, () => [
            h('input', { 'aria-label': 'S' }),
          ]),
          show.value
            ? h(XlsxEditorToolbarSlot, { slotId: 'x.z' as 'font.bold' }, () => [
                h('input', { 'aria-label': 'Z' }),
              ])
            : null,
        ])
      ));
    });
    const field = view.getByRole('textbox', { name: 'S' }) as HTMLInputElement;
    field.value = 'typed';
    show.value = false;
    await nextTick();
    await nextTick();
    expect(view.queryByRole('textbox', { name: 'L' })).toBeNull();
    expect(view.queryByRole('textbox', { name: 'W' })).toBeNull();
    expect(view.queryByRole('textbox', { name: 'Z' })).toBeNull();
    expect(view.getByRole('textbox', { name: 'S' })).toBe(field);
    expect(field.value).toBe('typed');
  });

  test('two lists with the same keys stay apart when they change', async () => {
    const order = ref([0, 1]);
    const list = (name: string) =>
      h(
        Fragment,
        order.value.map((i) => h('button', { key: i, type: 'button' }, `${name}${i}`))
      );
    const seen = await warnings(async () => {
      const { bar, view } = await open(() =>
        h(XlsxEditorToolbar, null, () => [
          h(XlsxEditorToolbarGroup, { id: 'mine', label: 'Mine' }, () => [list('a'), list('b')]),
          list('c'),
          list('d'),
        ])
      );
      const a0 = view.getByRole('button', { name: 'a0' });
      const c0 = view.getByRole('button', { name: 'c0' });
      order.value = [1, 0];
      await nextTick();
      await nextTick();
      // A keyed child moves with its element.
      expect(view.getByRole('button', { name: 'a0' })).toBe(a0);
      expect(view.getByRole('button', { name: 'c0' })).toBe(c0);
      const texts = (el: Element | null) =>
        [...(el?.querySelectorAll('button') ?? [])].map((x) => x.textContent);
      expect(texts(bar.querySelector('[data-group="mine"]'))).toEqual(['a1', 'a0', 'b1', 'b0']);
      expect(texts(bar.querySelector('[data-toolbar-fixed]'))).toEqual(['c1', 'c0', 'd1', 'd0']);
    });
    expect(seen.filter((w) => String(w).includes('Duplicate keys'))).toEqual([]);
  });
  test('the content of two slots of one group keeps its own element when both change', async () => {
    const which = ref<'bold' | 'underline'>('bold');
    const field = (name: string) => h('input', { 'aria-label': name });
    const { view } = await open(() =>
      h(XlsxEditorToolbar, null, () => [
        h(XlsxEditorToolbarSlot, { slotId: `font.${which.value}` }, () => field(which.value)),
      ])
    );
    const bold = view.getByRole('textbox', { name: 'bold' });
    which.value = 'underline';
    await nextTick();
    await nextTick();
    // The bold control comes back, and the underline content is a new element.
    expect(view.getByRole('textbox', { name: 'underline' })).not.toBe(bold);
    expect(bold.isConnected).toBe(false);
  });
});

describe('XlsxEditorToolbar.Group', () => {
  test('a host group follows every built-in group, or the group it names', async () => {
    const { bar } = await open(() =>
      h(XlsxEditorToolbar, null, () => [
        h(XlsxEditorToolbar.Group, { id: 'share', label: 'Share' }, () => button('Share')),
        h(
          XlsxEditorToolbar.Group,
          { id: 'review', label: 'Review', after: 'history', class: 'one', className: 'two' },
          () => button('Review')
        ),
      ])
    );
    expect(barGroups(bar)).toEqual(['history', 'review', ...BUILT_IN.slice(1), 'share']);
    const review = bar.querySelector('[data-group="review"]')!;
    expect(review.getAttribute('aria-label')).toBe('Review');
    expect(review.className).toBe('xlsx-toolbar-group one two');
  });

  test('a host group with no content or with hidden is left out', async () => {
    const { bar } = await open(() =>
      h(XlsxEditorToolbar, null, () => [
        h(XlsxEditorToolbarGroup, { id: 'empty', label: 'Empty' }),
        h(XlsxEditorToolbarGroup, template({ id: 'gone', label: 'Gone', hidden: '' }), () =>
          button('Gone')
        ),
      ])
    );
    expect(barGroups(bar)).toEqual(BUILT_IN);
  });

  test('a built-in id adds its content after the group controls, or hides the group', async () => {
    const { bar } = await open(() =>
      h(XlsxEditorToolbar, null, () => [
        h(XlsxEditorToolbarGroup, { id: 'file' }, () => button('Share', 'host.share')),
        h(XlsxEditorToolbarGroup, { id: 'borders', hidden: true }),
      ])
    );
    expect(slotsOf(bar, 'file')).toEqual(['file.save', 'host.share']);
    expect(barGroups(bar)).toEqual(BUILT_IN.filter((id) => id !== 'borders'));
  });

  test('a Slot with no content in a host group renders the built-in control', async () => {
    const { bar } = await open(() =>
      h(XlsxEditorToolbar, null, () => [
        h(XlsxEditorToolbarSlot, { slotId: 'font.bold', hidden: true }),
        h(XlsxEditorToolbarGroup, { id: 'mine', label: 'Mine' }, () => [
          h(XlsxEditorToolbarSlot, { slotId: 'font.bold' }),
        ]),
      ])
    );
    expect(slotsOf(bar, 'font')).not.toContain('font.bold');
    expect(slotsOf(bar, 'mine')).toEqual(['font.bold']);
  });

  test('the bar measures a host group and moves it into the More panel with its heading', async () => {
    // The built-in groups take 700 pixels and "Share" 50; the More button takes 40.
    const restore = stubLayout(widths(), 749);
    try {
      const { bar, view } = await open(() =>
        h(XlsxEditorToolbar, null, () => [
          h(XlsxEditorToolbarGroup, { id: 'share', label: 'Share' }, () => button('Share it')),
          h(XlsxEditorToolbarGroup, { id: 'team', labelKey: 'formattingBar.groups.file' }, () =>
            button('Team')
          ),
          h(XlsxEditorToolbarGroup, { id: 'plain' }, () => [
            h(XlsxEditorToolbarSlot, { slotId: 'font.bold' }),
          ]),
        ])
      );
      await waitFor(() => expect(barGroups(bar)).toEqual(BUILT_IN));
      fireEvent.click(view.getByRole('button', { name: 'More' }));
      await nextTick();
      const panel = view.getByRole('dialog', { name: 'More' });
      const sections = [...panel.querySelectorAll('.xlsx-toolbar-section')];
      expect(sections.map((s) => s.getAttribute('aria-label'))).toEqual(['Share', 'File', 'plain']);
      expect(sections[0]!.textContent).toBe('ShareShare it');
      // A Slot in the panel renders its control as a row.
      expect(sections[2]!.querySelector('[data-slot="font.bold"]')?.textContent).toContain('Bold');
    } finally {
      restore();
    }
  });

  test('with room for all, the bar shows a host group and no More button', async () => {
    const restore = stubLayout(widths(), 750);
    try {
      const { bar, view } = await open(() =>
        h(XlsxEditorToolbar, null, () => [
          h(XlsxEditorToolbarGroup, { id: 'share', label: 'Share' }, () => button('Share it')),
        ])
      );
      await waitFor(() => expect(barGroups(bar)).toEqual([...BUILT_IN, 'share']));
      expect(view.queryByRole('button', { name: 'More' })).toBeNull();
    } finally {
      restore();
    }
  });

  test('in the More panel, slot content and group content show in the section of the group', async () => {
    // 690 less the More button leaves 650: the first seven groups fit (610), and the clear and
    // file groups move into the panel.
    const restore = stubLayout(widths(), 690);
    try {
      const { bar, view } = await open(() =>
        h(XlsxEditorToolbar, null, () => [
          h(XlsxEditorToolbarSlot, { slotId: 'file.save' }, () => button('Upload')),
          h(XlsxEditorToolbarGroup, { id: 'file' }, () => button('Export')),
        ])
      );
      await waitFor(() => expect(barGroups(bar)).toEqual(BUILT_IN.slice(0, 7)));
      fireEvent.click(view.getByRole('button', { name: 'More' }));
      await nextTick();
      const panel = view.getByRole('dialog', { name: 'More' });
      const file = panel.querySelector('.xlsx-toolbar-section[aria-label="File"]')!;
      expect([...file.querySelectorAll('button')].map((b) => b.textContent)).toEqual([
        'Upload',
        'Export',
      ]);
      expect(file.querySelector('[data-slot="file.save"]')).toBeNull();
    } finally {
      restore();
    }
  });

  test('a built-in group ignores a label and an anchor, with a warning each', async () => {
    let bar!: HTMLElement;
    const seen = await warnings(async () => {
      ({ bar } = await open(() =>
        h(XlsxEditorToolbar, null, () => [
          h(XlsxEditorToolbarGroup, { id: 'font', label: 'Type', after: 'file' }, () =>
            button('Extra')
          ),
          h(XlsxEditorToolbarGroup, { id: 'late', after: 'nowhere' }, () => button('Late')),
        ])
      ));
    });
    expect(barGroups(bar)).toEqual([...BUILT_IN, 'late']);
    expect(bar.querySelector('[data-group="font"]')?.getAttribute('aria-label')).toBe('Font');
    expect(seen.map((m) => m.split(',')[0])).toEqual([
      '[xlsx-editor] XlsxEditorToolbar.Group "font" is a built-in group',
      '[xlsx-editor] XlsxEditorToolbar.Group "font" is a built-in group',
      '[xlsx-editor] XlsxEditorToolbar.Group "late" has after="nowhere"',
    ]);
  });
});

describe('loose children', () => {
  test('show in one block after the More button, which the bar charges as fixed width', async () => {
    // 849 less the block's 100 leaves 749: the host group "Share" moves into the panel.
    const restore = stubLayout(widths(100), 849);
    try {
      const { bar, view } = await open(() =>
        h(XlsxEditorToolbar, null, () => [
          h('span', null, 'Saved'),
          h(XlsxEditorToolbarGroup, { id: 'share', label: 'Share' }, () => button('Share it')),
        ])
      );
      await waitFor(() => expect(barGroups(bar)).toEqual(BUILT_IN));
      const more = view.getByRole('button', { name: 'More' }).parentElement!;
      const fixed = bar.querySelector('[data-toolbar-fixed]')!;
      expect(more.nextElementSibling).toBe(fixed);
      expect(fixed.textContent).toBe('Saved');
    } finally {
      restore();
    }
  });

  test('a block that comes or goes measures the bar again', async () => {
    const restore = stubLayout(widths(100), 849);
    try {
      const loose = ref(false);
      const { bar } = await open(() =>
        h(XlsxEditorToolbar, null, () => [
          h(XlsxEditorToolbarGroup, { id: 'share', label: 'Share' }, () => button('Share it')),
          loose.value ? h('span', null, 'Saved') : null,
        ])
      );
      await waitFor(() => expect(barGroups(bar)).toEqual([...BUILT_IN, 'share']));
      loose.value = true;
      await waitFor(() => expect(barGroups(bar)).toEqual(BUILT_IN));
      loose.value = false;
      await waitFor(() => expect(barGroups(bar)).toEqual([...BUILT_IN, 'share']));
    } finally {
      restore();
    }
  });
});

describe('the arrow keys in the preset bar', () => {
  test('reach a control at any depth in a host group and in the fixed block', async () => {
    // A component of the host renders the part, so the toolbar shows it as a loose child.
    const Color = defineComponent(() => () => h(XlsxEditorToolbarSlot, { slotId: 'font.color' }));
    const { bar, view } = await open(() =>
      h(XlsxEditorToolbar, null, () => [
        h(XlsxEditorToolbarGroup, { id: 'share', label: 'Share' }, () => [
          h('span', { title: 'Share the workbook' }, [button('Share it')]),
          h('div', { hidden: true }, [button('Hidden')]),
          h('div', { 'aria-hidden': 'true' }, [button('Unseen')]),
        ]),
        h(Color),
        h('a', { href: '#help' }, 'Help'),
        h('button', { type: 'button', disabled: true }, 'Off'),
        h('button', { type: 'button', tabindex: -1 }, 'Out'),
        h('select', { 'aria-label': 'Pick' }, [h('option', 'One')]),
        h('span', { role: 'button', tabindex: 0 }, 'Chip'),
        h('textarea', { 'aria-label': 'Memo' }),
        h('input', { 'aria-label': 'Note', value: 'ab' }),
      ])
    );
    const fixed = bar.querySelector<HTMLElement>('[data-toolbar-fixed]')!;
    const stops = [
      bar.querySelector<HTMLElement>('[data-slot="file.save"]')!,
      view.getByRole('button', { name: 'Share it' }),
      fixed.querySelector<HTMLElement>('button[data-slot="font.color"]')!,
      fixed.querySelector<HTMLElement>('.xlsx-toolbar-colorsplit-caret')!,
      view.getByRole('link', { name: 'Help' }),
      view.getByRole('combobox', { name: 'Pick' }),
      view.getByRole('button', { name: 'Chip' }),
      view.getByRole('textbox', { name: 'Memo' }),
    ];
    stops[0]!.focus();
    for (const next of stops.slice(1)) {
      fireEvent.keyDown(document.activeElement!, { key: 'ArrowRight' });
      expect(document.activeElement).toBe(next);
    }
    // A text area keeps every key, with its own action: `keyDown` gives true for a key that is
    // not cancelled.
    expect(fireEvent.keyDown(stops[7]!, { key: 'ArrowLeft' })).toBe(true);
    expect(fireEvent.keyDown(stops[7]!, { key: 'End' })).toBe(true);
    expect(document.activeElement).toBe(stops[7]!);
    // A field keeps the keys that move its caret, and gives an arrow back at the end of its text.
    // The bar cancels a key that moves the focus. The keys stop at the ends of the bar.
    const note = view.getByRole('textbox', { name: 'Note' }) as HTMLInputElement;
    note.focus();
    note.setSelectionRange(2, 2);
    expect(fireEvent.keyDown(note, { key: 'ArrowLeft' })).toBe(true);
    expect(fireEvent.keyDown(note, { key: 'Home' })).toBe(true);
    expect(document.activeElement).toBe(note);
    // At the end of its text, the field gives ArrowRight to the bar, which stops at its end.
    expect(fireEvent.keyDown(note, { key: 'ArrowRight' })).toBe(false);
    expect(document.activeElement).toBe(note);
    note.setSelectionRange(0, 0);
    expect(fireEvent.keyDown(note, { key: 'ArrowLeft' })).toBe(false);
    expect(document.activeElement).toBe(stops[7]!);
    const undo = view.getByRole('button', { name: 'Undo' });
    undo.focus();
    fireEvent.keyDown(undo, { key: 'ArrowLeft' });
    expect(document.activeElement).toBe(undo);
    fireEvent.keyDown(undo, { key: 'End' });
    expect(document.activeElement).toBe(note);
  });

  test('skip the controls of an open list, and leave a key that a host control handled', async () => {
    const { bar, view } = await open(() =>
      h(XlsxEditorToolbar, null, () => [
        h(XlsxEditorToolbarGroup, { id: 'mine', label: 'Mine' }, () => [
          h(
            'button',
            { type: 'button', onKeydown: (e: KeyboardEvent) => e.preventDefault() },
            'Own'
          ),
          button('Next'),
          h('div', { role: 'menu' }, [button('One'), button('Two')]),
          h('div', { role: 'dialog' }, [button('Three')]),
        ]),
      ])
    );
    const size = bar.querySelector<HTMLElement>('button[data-slot="font.size"]')!;
    fireEvent.click(size);
    await waitFor(() => expect(bar.querySelector('[data-popup]')).not.toBeNull());
    // An item of the open list keeps the arrow keys.
    const item = bar.querySelector<HTMLElement>('[data-popup] [role="option"]')!;
    item.focus();
    fireEvent.keyDown(item, { key: 'ArrowLeft' });
    expect(document.activeElement).toBe(item);
    size.focus();
    fireEvent.keyDown(size, { key: 'ArrowRight' });
    expect(document.activeElement).toBe(view.getByRole('button', { name: 'Bold' }));
    // A menu of the host moves its own focus: Home stays in it.
    const two = view.getByRole('button', { name: 'Two' });
    two.focus();
    fireEvent.keyDown(two, { key: 'Home' });
    expect(document.activeElement).toBe(two);
    // So does a dialog of the host.
    const three = view.getByRole('button', { name: 'Three' });
    three.focus();
    fireEvent.keyDown(three, { key: 'Home' });
    expect(document.activeElement).toBe(three);
    const own = view.getByRole('button', { name: 'Own' });
    own.focus();
    fireEvent.keyDown(own, { key: 'ArrowRight' });
    expect(document.activeElement).toBe(own);
  });

  test('pass over a control that cannot take the focus, also in a bar in a host dialog', async () => {
    const { bar, view } = await open(() =>
      h(XlsxEditorToolbar, null, () => [
        h(XlsxEditorToolbarGroup, { id: 'mine', label: 'Mine' }, () => [
          button('A'),
          h('input', { type: 'hidden' }),
          h(
            'button',
            {
              type: 'button',
              // A control that is not rendered refuses the focus.
              ref: (el: unknown) => {
                if (el) (el as HTMLElement).focus = () => undefined;
              },
            },
            'Ghost'
          ),
          button('B'),
        ]),
      ])
    );
    const a = view.getByRole('button', { name: 'A' });
    a.focus();
    fireEvent.keyDown(a, { key: 'ArrowRight' });
    expect(document.activeElement).toBe(view.getByRole('button', { name: 'B' }));
    fireEvent.keyDown(document.activeElement!, { key: 'ArrowLeft' });
    expect(document.activeElement).toBe(a);
    // A dialog or a menu of the host around the bar does not take the keys of the bar.
    for (const role of ['dialog', 'menu']) {
      bar.parentElement!.setAttribute('role', role);
      const undo = bar.querySelector<HTMLElement>('[data-slot="history.undo"]')!;
      undo.focus();
      fireEvent.keyDown(undo, { key: 'ArrowRight' });
      expect(document.activeElement).toBe(bar.querySelector('[data-slot="history.redo"]'));
    }
  });
});

describe(':preset="false"', () => {
  test('renders only the children, in order, with no measure and no More panel', async () => {
    const restore = stubLayout(widths(), 100);
    try {
      const { bar, view } = await open(() =>
        h(XlsxEditorToolbar, { preset: false }, () => [
          h(XlsxEditorToolbarGroup, { id: 'mine', label: 'Mine', class: 'host' }, () => [
            h(XlsxEditorToolbarSlot, { slotId: 'font.bold' }),
            h(XlsxEditorToolbarSlot, { slotId: 'font.italic', hidden: true }),
            h(XlsxEditorToolbarSlot, { slotId: 'history.undo' }, () => button('Back')),
          ]),
          h(XlsxEditorToolbarSlot, { slotId: 'file.save' }),
        ])
      );
      expect(bar.hasAttribute('data-overflow')).toBe(false);
      expect(bar.querySelector('[data-group]')).toBeNull();
      expect(view.queryByRole('button', { name: 'More' })).toBeNull();
      const mine = bar.querySelector('[data-host-group="mine"]')!;
      expect(mine.getAttribute('role')).toBe('group');
      expect(mine.className).toBe('xlsx-toolbar-group host');
      expect(mine.textContent).toContain('Back');
      expect(
        [...bar.querySelectorAll<HTMLElement>('[data-slot]')].map((s) => s.dataset.slot)
      ).toEqual(['font.bold', 'file.save']);
    } finally {
      restore();
    }
  });

  test('the arrow keys move between the controls of host groups and loose controls', async () => {
    const { bar, view } = await open(() =>
      h(XlsxEditorToolbar, { preset: false }, () => [
        h(XlsxEditorToolbarGroup, { id: 'mine', label: 'Mine' }, () => [
          h(XlsxEditorToolbarSlot, { slotId: 'font.bold' }),
          button('Own'),
          h(XlsxEditorToolbarSlot, { slotId: 'font.size' }),
        ]),
        h(XlsxEditorToolbarSlot, { slotId: 'font.color' }),
        button('Loose'),
      ])
    );
    // A list is a block that holds its trigger, a stop. A color split button is two stops: its
    // main half and its caret.
    const stops = [
      view.getByRole('button', { name: 'Bold' }),
      view.getByRole('button', { name: 'Own' }),
      bar.querySelector<HTMLElement>('button[data-slot="font.size"]')!,
      bar.querySelector<HTMLElement>('button[data-slot="font.color"]')!,
      view.getByRole('button', { name: 'Font color options' }),
      view.getByRole('button', { name: 'Loose' }),
    ];
    stops[0]!.focus();
    for (const next of stops.slice(1)) {
      fireEvent.keyDown(document.activeElement!, { key: 'ArrowRight' });
      expect(document.activeElement).toBe(next);
    }
    fireEvent.keyDown(document.activeElement!, { key: 'Home' });
    expect(document.activeElement).toBe(stops[0]!);
    fireEvent.keyDown(document.activeElement!, { key: 'End' });
    expect(document.activeElement).toBe(stops[5]!);
  });

  test('a group with no content is left out, and a built-in group has its own heading', async () => {
    const { bar } = await open(() =>
      h(XlsxEditorToolbar, { preset: false }, () => [
        h(XlsxEditorToolbarGroup, { id: 'share', label: 'Share' }, () => [null]),
        h(XlsxEditorToolbarGroup, { id: 'font', label: 'Text' }, () => [
          h(XlsxEditorToolbarSlot, { slotId: 'font.bold' }),
        ]),
      ])
    );
    expect(bar.querySelector('[data-host-group="share"]')).toBeNull();
    expect(bar.querySelector('[data-host-group="font"]')!.getAttribute('aria-label')).toBe('Font');
  });

  test('a hidden Slot and a hidden Group render nothing, also with content', async () => {
    const { bar } = await open(() =>
      h(XlsxEditorToolbar, { preset: false }, () => [
        h(XlsxEditorToolbarSlot, { slotId: 'font.bold', hidden: true }, () => button('Strong')),
        h(XlsxEditorToolbarGroup, { id: 'mine', label: 'Mine', hidden: true }, () =>
          button('Share')
        ),
      ])
    );
    expect(bar.childElementCount).toBe(0);
  });

  test('a built-in control there runs its command', async () => {
    const { view, viewer } = await open(() =>
      h(XlsxEditorToolbar, { preset: false }, () => [
        h(XlsxEditorToolbarSlot, { slotId: 'font.bold' }),
      ])
    );
    const before = viewer.snapshot().revision;
    fireEvent.click(view.getByRole('button', { name: 'Bold' }));
    await waitFor(() => expect(viewer.snapshot().revision).toBeGreaterThan(before));
  });
});
