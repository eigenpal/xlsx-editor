// The customization parts of the toolbar in happy-dom (`bun run test:dom`). The Vue twin of the
// React adapter's `toolbar-customize.test.tsx`, with the cases that only a Vue host can write:
// the `select` event of an action, a part's props as a template passes them (`class-name`, a
// bare `hidden`, a `priority` as text), and a template `class` beside `className`.

import { afterEach, describe, expect, test } from 'bun:test';
import { fireEvent, waitFor } from '@testing-library/react';
import { createVNode, h, nextTick, ref, type VNodeChild } from 'vue';
import { buildXlsx, c, row } from '../../../scripts/fixtures/builder.ts';
import { TOOLBAR_NAMED_PARTS } from '@xlsx-editor.dev/core/editor';
import {
  StatusMessage,
  XlsxEditorContent,
  XlsxEditorRoot,
  XlsxEditorToolbar,
  XlsxEditorViewport,
  type XlsxEditorInstance,
} from '../src/index.ts';
import { FONTS, mount, unmountAll, type Mounted } from './mount.ts';
import { groupWidth, stubLayout } from './toolbar-stubs.ts';

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

/** The compiler patch flag of a template element whose only dynamic binding is its class. */
const CLASS_FLAG = 2;

/** The ids of the groups that the bar shows, in order. */
const barGroups = (bar: HTMLElement) =>
  [...bar.querySelectorAll<HTMLElement>(':scope > [data-group]')].map((g) => g.dataset.group);

const slot = (root: HTMLElement, id: string) =>
  root.querySelector<HTMLElement>(`[data-slot="${id}"]`);

/**
 * Props as a compiled template passes them: kebab-case names, and a bare boolean attribute as
 * an empty string. The prop types of `h` do not describe them.
 */
const template = (props: Record<string, unknown>) => props as never;

/** Collect `console.warn` messages while `run` runs. */
async function warnings(run: () => Promise<void>): Promise<string[]> {
  const seen: string[] = [];
  const original = console.warn;
  console.warn = (message: string) => void seen.push(String(message));
  try {
    await run();
  } finally {
    console.warn = original;
  }
  return seen;
}

/** Open the "More" panel and return it. */
async function openPanel(view: Mounted): Promise<HTMLElement> {
  fireEvent.click(view.getByRole('button', { name: 'More' }));
  await nextTick();
  return view.getByRole('dialog', { name: 'More' });
}

const T = XlsxEditorToolbar;

describe('the named parts', () => {
  test('one part for each slot of the registry, from the core table', () => {
    const parts = T as unknown as Record<string, unknown>;
    for (const name of Object.keys(TOOLBAR_NAMED_PARTS)) expect(parts[name]).toBeTruthy();
  });

  test('hidden removes the slot where it stands, also as a bare template attribute', async () => {
    const { bar } = await open(() =>
      h(T, null, () => [h(T.Bold, template({ hidden: '' })), h(T.FontSize, { hidden: true })])
    );
    expect(slot(bar, 'font.bold')).toBeNull();
    expect(slot(bar, 'font.size')).toBeNull();
    expect(slot(bar, 'font.italic')).toBeTruthy();
  });

  test('icon and className change the control, which keeps its label, state, and command', async () => {
    const { bar } = await open(() =>
      h(T, null, () => [
        h(
          T.Bold,
          template({ icon: h('b', { 'data-host-icon': '' }, 'B'), 'class-name': 'host-bold' })
        ),
      ])
    );
    const bold = slot(bar, 'font.bold')!;
    expect(bold.closest('[data-group]')?.getAttribute('data-group')).toBe('font');
    expect(bold.className).toBe('xlsx-toolbar-button host-bold');
    expect(bold.querySelector('[data-host-icon]')).toBeTruthy();
    expect(bold.querySelector('svg')).toBeNull();
    expect(bold.getAttribute('aria-label')).toBe('Bold');
    fireEvent.click(bold);
    await waitFor(() => expect(slot(bar, 'font.bold')!.getAttribute('aria-pressed')).toBe('true'));
  });

  test('the default slot of a button part is its icon, not a replacement', async () => {
    const { bar } = await open(() =>
      h(T, null, () => [h(T.Italic, null, () => [h('i', { 'data-host-icon': '' }, 'I')])])
    );
    const italic = slot(bar, 'font.italic')!;
    expect(italic.tagName).toBe('BUTTON');
    expect(italic.querySelector('[data-host-icon]')).toBeTruthy();
  });

  test('className and a template class reach the trigger of a value part', async () => {
    const { bar } = await open(() =>
      h(T, null, () => [
        h(T.FontSize, { className: 'host-size' }),
        h(T.FillColor, { class: 'host-fill' }),
      ])
    );
    expect(slot(bar, 'font.size')!.classList.contains('host-size')).toBe(true);
    expect(slot(bar, 'fill.color')!.classList.contains('host-fill')).toBe(true);
  });

  test('a changed slot of a menu group leaves the menu and shows after the menu button', async () => {
    const { bar, view } = await open(() =>
      h(T, null, () => [h(T.InsertRows, { icon: h('b', null, 'R') })])
    );
    const cells = bar.querySelector<HTMLElement>('[data-group="cells"]')!;
    expect(slot(cells, 'cells.insertRows')!.textContent).toBe('R');
    fireEvent.click(cells.querySelector('[data-group-menu="cells"]')!);
    await nextTick();
    const menu = view.getByRole('menu');
    expect(
      [...menu.querySelectorAll<HTMLElement>('[data-slot]')].map((s) => s.dataset.slot)
    ).toEqual(['cells.insertColumns', 'cells.deleteRows', 'cells.deleteColumns']);
  });
});

describe('asChild', () => {
  test('a host element renders the control, with its label, state, and press', async () => {
    const { bar } = await open(() =>
      h(T, null, () => [
        h(T.Slot, { slotId: 'font.bold', asChild: true }, () => [
          h('span', { class: 'host', 'data-own': '' }, 'Strong'),
        ]),
      ])
    );
    const bold = slot(bar, 'font.bold')!;
    expect(bold.tagName).toBe('SPAN');
    expect(bold.className).toBe('xlsx-toolbar-button host');
    expect(bold.hasAttribute('data-own')).toBe(true);
    expect(bold.getAttribute('role')).toBe('button');
    expect(bold.tabIndex).toBe(0);
    expect(bold.getAttribute('aria-label')).toBe('Bold');
    fireEvent.click(bold);
    await waitFor(() => expect(slot(bar, 'font.bold')!.getAttribute('aria-pressed')).toBe('true'));
    fireEvent.keyDown(slot(bar, 'font.bold')!, { key: 'Enter' });
    await waitFor(() => expect(slot(bar, 'font.bold')!.getAttribute('aria-pressed')).toBe('false'));
  });

  test('a named part with asChild renders its single child', async () => {
    const { bar } = await open(() =>
      h(T, null, () => [
        h(T.Undo, { asChild: true }, () => [
          h('button', { type: 'button', class: 'host-undo' }, 'Back'),
        ]),
      ])
    );
    const undo = slot(bar, 'history.undo')!;
    expect(undo.className).toBe('xlsx-toolbar-button host-undo');
    expect(undo.textContent).toBe('Back');
    expect(undo.getAttribute('aria-disabled')).toBe('true');
  });

  test('children that are not one element fall back to the built-in button, with a warning', async () => {
    let bar!: HTMLElement;
    const seen = await warnings(async () => {
      ({ bar } = await open(() =>
        h(T, null, () => [
          h(T.Slot, { slotId: 'font.underline', asChild: true }, () => [
            h('span', null, 'one'),
            h('span', null, 'two'),
          ]),
        ])
      ));
    });
    expect(slot(bar, 'font.underline')!.tagName).toBe('BUTTON');
    expect(seen.some((w) => w.includes('"font.underline" has asChild'))).toBe(true);
  });

  test('icon on a value control is ignored, with a warning', async () => {
    let bar!: HTMLElement;
    const seen = await warnings(async () => {
      ({ bar } = await open(() =>
        h(T, null, () => [
          h(T.Slot, { slotId: 'font.family', icon: h('b', { 'data-host-icon': '' }, 'F') }),
        ])
      ));
    });
    expect(bar.querySelector('[data-host-icon]')).toBeNull();
    expect(
      seen.some((w) => w.includes('"font.family" has icon, but the slot is not a button'))
    ).toBe(true);
  });
});

describe('XlsxEditorToolbar.Action', () => {
  test('a host button with the toolbar class, its label, the guard, and the select event', async () => {
    let pressed = 0;
    const { view } = await open(() =>
      h(T, null, () => [
        h(T.Group, { id: 'review', label: 'Review' }, () => [
          h(T.Action, {
            label: 'Send',
            icon: h('b', { 'data-host-icon': '' }, 'S'),
            className: 'host-send',
            onSelect: () => (pressed += 1),
          }),
        ]),
      ])
    );
    const send = view.getByRole('button', { name: 'Send' });
    expect(send.closest('[data-group]')?.getAttribute('data-group')).toBe('review');
    expect(send.className).toBe('xlsx-toolbar-button host-send');
    expect(send.querySelector('[data-host-icon]')).toBeTruthy();
    expect(fireEvent.mouseDown(send)).toBe(false);
    fireEvent.click(send);
    expect(pressed).toBe(1);
  });

  test('without an icon or children, the label shows as text; active sets aria-pressed', async () => {
    const { view } = await open(() =>
      h(T, { preset: false }, () => [
        h(T.Action, { label: 'Publish' }),
        h(T.Action, { label: 'Track', active: true }),
      ])
    );
    expect(view.getByRole('button', { name: 'Publish' }).textContent).toBe('Publish');
    const track = view.getByRole('button', { name: 'Track' });
    expect(track.getAttribute('aria-pressed')).toBe('true');
    expect(track.hasAttribute('data-active')).toBe(true);
  });

  test('a disabled action stays focusable, says why, and emits no select', async () => {
    let pressed = 0;
    const { bar, view } = await open(() =>
      h(T, { preset: false }, () => [
        h(T.Slot, { slotId: 'font.bold' }),
        h(T.Action, {
          label: 'Send',
          disabled: true,
          disabledReason: 'Save the workbook first.',
          onSelect: () => (pressed += 1),
        }),
      ])
    );
    const send = view.getByRole('button', { name: 'Send' });
    expect(send.hasAttribute('disabled')).toBe(false);
    expect(send.getAttribute('aria-disabled')).toBe('true');
    expect(send.getAttribute('title')).toBe('Save the workbook first.');
    const reason = bar.querySelector(`#${CSS.escape(send.getAttribute('aria-describedby')!)}`);
    expect(reason?.textContent).toBe('Save the workbook first.');
    fireEvent.click(send);
    expect(pressed).toBe(0);
    const bold = slot(bar, 'font.bold')!;
    bold.focus();
    fireEvent.keyDown(bold, { key: 'ArrowRight' });
    expect(document.activeElement).toBe(send);
  });

  test('a held Enter runs the action once', async () => {
    const { view } = await open(() =>
      h(T, { preset: false }, () => [h(T.Action, { label: 'Step' })])
    );
    const step = view.getByRole('button', { name: 'Step' });
    expect(fireEvent.keyDown(step, { key: 'Enter', repeat: true })).toBe(false);
    expect(fireEvent.keyDown(step, { key: 'Enter' })).toBe(true);
  });

  test('asChild renders the host element with the action merged in', async () => {
    let pressed = 0;
    const { view } = await open(() =>
      h(T, { preset: false }, () => [
        h(T.Action, { label: 'Open', asChild: true, onSelect: () => (pressed += 1) }, () => [
          h('a', { href: '#open', class: 'host-link' }, 'Open'),
        ]),
      ])
    );
    const link = view.getByRole('button', { name: 'Open' });
    expect(link.tagName).toBe('A');
    expect(link.className).toBe('xlsx-toolbar-button host-link');
    fireEvent.click(link);
    expect(pressed).toBe(1);
  });

  test('in a collapsed group, the action is a row of the More panel with its label', async () => {
    let pressed = 0;
    const restore = stubLayout(groupWidth, 759);
    try {
      const { bar, view } = await open(() =>
        h(T, null, () => [
          h(T.Group, { id: 'review', label: 'Review' }, () => [
            h(T.Action, { label: 'Send', icon: h('b', null, 'S'), onSelect: () => pressed++ }),
            h(T.Separator),
          ]),
        ])
      );
      await waitFor(() => expect(barGroups(bar)).not.toContain('review'));
      const panel = await openPanel(view);
      const section = panel.querySelector<HTMLElement>('[aria-label="Review"]')!;
      const row = section.querySelector<HTMLElement>('button[aria-label="Send"]')!;
      expect(row.className).toBe('xlsx-toolbar-row');
      expect(row.textContent).toBe('SSend');
      fireEvent.click(row);
      expect(pressed).toBe(1);
      expect(section.querySelector('[role="separator"]')?.getAttribute('aria-orientation')).toBe(
        'horizontal'
      );
    } finally {
      restore();
    }
  });
});

describe('asChild with a template child', () => {
  // A compiled template marks an element with a patch flag, here CLASS only. The merged state
  // attributes and handlers must still update on each render.
  test('a named part keeps its state up to date after a press and after a mode change', async () => {
    const mode = ref<'edit' | 'view'>('edit');
    let viewer: XlsxEditorInstance | null = null;
    const view = mount(() =>
      h(
        XlsxEditorRoot,
        {
          document: workbook(),
          fonts: FONTS,
          mode: mode.value,
          onReady: (v: XlsxEditorInstance) => (viewer = v),
        },
        {
          default: () =>
            h(XlsxEditorViewport, { style: { height: '400px' } }, () => [
              h(T, null, () => [
                h(T.Bold, { asChild: true }, () => [
                  createVNode('button', { type: 'button', class: 'host' }, 'B', CLASS_FLAG),
                ]),
              ]),
              h(XlsxEditorContent),
              h(StatusMessage),
            ]),
        }
      )
    );
    await waitFor(() => expect(viewer?.snapshot().status).toBe('ready'), { timeout: 5000 });
    await nextTick();
    const bar = view.getByRole('toolbar');
    expect(slot(bar, 'font.bold')!.getAttribute('aria-pressed')).toBe('false');
    fireEvent.click(slot(bar, 'font.bold')!);
    await waitFor(() => expect(slot(bar, 'font.bold')!.getAttribute('aria-pressed')).toBe('true'));
    mode.value = 'view';
    await waitFor(() => expect(slot(bar, 'font.bold')!.getAttribute('aria-disabled')).toBe('true'));
    expect(slot(bar, 'font.bold')!.hasAttribute('data-disabled')).toBe(true);
  });

  test('an action keeps its state up to date when it turns disabled', async () => {
    const off = ref(false);
    let pressed = 0;
    const { view } = await open(() =>
      h(T, { preset: false }, () => [
        h(
          T.Action,
          {
            label: 'Go',
            asChild: true,
            disabled: off.value,
            disabledReason: 'Not now.',
            onSelect: () => (pressed += 1),
          },
          () => [createVNode('button', { type: 'button', class: 'host' }, 'Go', CLASS_FLAG)]
        ),
      ])
    );
    fireEvent.click(view.getByRole('button', { name: 'Go' }));
    expect(pressed).toBe(1);
    off.value = true;
    await nextTick();
    const go = view.getByRole('button', { name: 'Go' });
    expect(go.getAttribute('aria-disabled')).toBe('true');
    expect(go.getAttribute('title')).toBe('Not now.');
    fireEvent.click(go);
    expect(pressed).toBe(1);
  });

  test('a disabled action on a link stops the navigation', async () => {
    const { view } = await open(() =>
      h(T, { preset: false }, () => [
        h(T.Action, { label: 'Open', asChild: true, disabled: true }, () => [
          h('a', { href: '#open' }, 'Open'),
        ]),
      ])
    );
    expect(fireEvent.click(view.getByRole('button', { name: 'Open' }))).toBe(false);
  });

  test('the More panel opened from the keyboard focuses a host element of asChild', async () => {
    const restore = stubLayout(groupWidth, 759);
    try {
      const { bar, view } = await open(() =>
        h(T, null, () => [
          h(T.Group, { id: 'review', label: 'Review' }, () => [
            h(T.Action, { label: 'Go', asChild: true }, () => [h('span', null, 'Go')]),
          ]),
        ])
      );
      await waitFor(() => expect(barGroups(bar)).not.toContain('review'));
      fireEvent.keyDown(view.getByRole('button', { name: 'More' }), { key: 'Enter' });
      await waitFor(() =>
        expect((document.activeElement as HTMLElement | null)?.textContent).toBe('Go')
      );
      expect(document.activeElement?.tagName).toBe('SPAN');
    } finally {
      restore();
    }
  });
});

describe('XlsxEditorToolbar.Separator', () => {
  test('a vertical rule with its class', async () => {
    const { bar } = await open(() =>
      h(T, { preset: false }, () => [h(T.Separator, { className: 'host-rule' })])
    );
    const rule = bar.querySelector('[role="separator"]')!;
    expect(rule.getAttribute('aria-orientation')).toBe('vertical');
    expect(rule.className).toBe('xlsx-toolbar-separator host-rule');
  });
});

describe('priority and pinned', () => {
  test('a host group with no priority leaves the bar first', async () => {
    const restore = stubLayout(groupWidth, 400);
    try {
      const { bar } = await open(() =>
        h(T, null, () => [
          h(T.Group, { id: 'share', label: 'Share', after: 'history' }, () => [
            h('button', { type: 'button' }, 'Share it'),
          ]),
        ])
      );
      await waitFor(() =>
        expect(barGroups(bar)).toEqual(['history', 'font', 'fill', 'borders', 'alignment'])
      );
    } finally {
      restore();
    }
  });

  test('a priority, also as template text, keeps a host group longer', async () => {
    const restore = stubLayout(groupWidth, 400);
    try {
      const { bar, view } = await open(() =>
        h(T, null, () => [
          h(T.Group, template({ id: 'share', label: 'Share', priority: '95' }), () => [
            h('button', { type: 'button' }, 'Share it'),
          ]),
        ])
      );
      await waitFor(() =>
        expect(barGroups(bar)).toEqual(['history', 'font', 'fill', 'borders', 'alignment', 'share'])
      );
      const panel = await openPanel(view);
      expect(
        [...panel.querySelectorAll('.xlsx-toolbar-section')].map((s) =>
          s.getAttribute('aria-label')
        )
      ).toEqual(['Number format', 'Insert and delete', 'Clear', 'File']);
    } finally {
      restore();
    }
  });

  test('a pinned group stays at every width', async () => {
    const restore = stubLayout(groupWidth, 100);
    try {
      const { bar } = await open(() =>
        h(T, null, () => [
          h(T.Group, { id: 'share', label: 'Share', pinned: true }, () => [
            h('button', { type: 'button' }, 'Share it'),
          ]),
          h(T.Group, template({ id: 'history', pinned: '' })),
        ])
      );
      await waitFor(() => expect(barGroups(bar)).toEqual(['history', 'share']));
    } finally {
      restore();
    }
  });
});

describe('overflowContent', () => {
  test('of a group: in place of the default slot of a host group, after a built-in group', async () => {
    const restore = stubLayout(groupWidth, 100);
    try {
      const { view } = await open(() =>
        h(T, null, () => [
          h(
            T.Group,
            {
              id: 'share',
              label: 'Share',
              overflowContent: () => h('button', { type: 'button' }, 'Share as a row'),
            },
            () => [h('button', { type: 'button' }, 'Share it')]
          ),
          h(T.Group, {
            id: 'file',
            overflowContent: () => h('button', { type: 'button' }, 'Export'),
          }),
        ])
      );
      await waitFor(() => expect(view.queryByRole('button', { name: 'More' })).toBeTruthy());
      const panel = await openPanel(view);
      const share = panel.querySelector<HTMLElement>('[aria-label="Share"]')!;
      expect([...share.querySelectorAll('button')].map((b) => b.textContent)).toEqual([
        'Share as a row',
      ]);
      const file = panel.querySelector<HTMLElement>('[aria-label="File"]')!;
      expect(
        [...file.querySelectorAll('button')].map(
          (b) => b.getAttribute('aria-label') ?? b.textContent
        )
      ).toEqual(['Save', 'Export']);
    } finally {
      restore();
    }
  });

  test('of a slot: in place of the control in its row of the More panel', async () => {
    const restore = stubLayout(groupWidth, 100);
    try {
      const { bar, view } = await open(() =>
        h(T, null, () => [
          h(T.Save, { overflowContent: () => h('button', { type: 'button' }, 'Upload') }),
          h(T.Group, { id: 'history', pinned: true }),
        ])
      );
      await waitFor(() => expect(barGroups(bar)).toEqual(['history']));
      const panel = await openPanel(view);
      const file = panel.querySelector<HTMLElement>('[aria-label="File"]')!;
      expect(slot(file, 'file.save')).toBeNull();
      expect(file.textContent).toContain('Upload');
    } finally {
      restore();
    }
  });
});

describe(':preset="false"', () => {
  test('named parts, groups, actions, slots, and separators render in their order', async () => {
    const { bar } = await open(() =>
      h(T, { preset: false }, () => [
        h(T.Action, { label: 'First' }),
        h(T.Separator),
        h(T.Bold, { className: 'host-bold' }),
        h(T.Group, { id: 'mine', label: 'Mine' }, () => [
          h(T.Italic),
          h(T.Action, { label: 'Inner' }),
        ]),
        h(T.Slot, { slotId: 'history.undo' }),
        h(T.DeleteRows),
      ])
    );
    const order = [...bar.querySelectorAll<HTMLElement>('button, [role="separator"]')].map(
      (el) => el.dataset.slot ?? el.getAttribute('aria-label') ?? el.getAttribute('role')
    );
    expect(order).toEqual([
      'First',
      'separator',
      'font.bold',
      'font.italic',
      'Inner',
      'history.undo',
      'cells.deleteRows',
    ]);
    expect(slot(bar, 'font.bold')!.classList.contains('host-bold')).toBe(true);
    expect(slot(bar, 'cells.deleteRows')!.textContent).toBe('Delete rows');
  });
});
