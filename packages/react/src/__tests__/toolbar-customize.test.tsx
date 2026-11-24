/** @jsxImportSource react */
// The customization parts of the toolbar in happy-dom (`bun run test:dom`): the named parts
// (`XlsxEditor.Toolbar.Bold`), `icon`, `asChild`, and `className` on a slot, `Toolbar.Action`,
// `Toolbar.Separator`, `priority`, `pinned`, and `overflowContent` on a group and a slot, and a
// `preset={false}` bar of every kind of part. The Vue twin is
// packages/vue/test/toolbar-customize.test.ts.

import { afterEach, describe, expect, test } from 'bun:test';
import { cleanup, fireEvent, render, waitFor } from '@testing-library/react';
import type { ReactNode } from 'react';
import { TOOLBAR_NAMED_PARTS } from '@xlsx-editor.dev/core/editor';
import { XlsxEditor, type XlsxEditorInstance } from '../index.ts';
import { FONTS, groupWidth, stubLayout, workbook } from './toolbar-stubs.tsx';

afterEach(cleanup);

/** Mount the editor with `toolbar` as its toolbar, and wait for the workbook. */
async function open(toolbar: ReactNode) {
  let viewer: XlsxEditorInstance | null = null;
  const view = render(
    <XlsxEditor.Root document={workbook()} fonts={FONTS} onReady={(v) => (viewer = v)}>
      <XlsxEditor.Viewport style={{ height: 400 }}>
        {toolbar}
        <XlsxEditor.Content />
        <XlsxEditor.StatusMessage />
      </XlsxEditor.Viewport>
    </XlsxEditor.Root>
  );
  await waitFor(() => expect(viewer?.snapshot().status).toBe('ready'), { timeout: 5000 });
  return { view, bar: view.getByRole('toolbar'), viewer: viewer! };
}

/** The ids of the groups that the bar shows, in order. */
const barGroups = (bar: HTMLElement) =>
  [...bar.querySelectorAll<HTMLElement>(':scope > [data-group]')].map((g) => g.dataset.group);

const slot = (root: HTMLElement, id: string) =>
  root.querySelector<HTMLElement>(`[data-slot="${id}"]`);

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

/** Open the "More" panel and return it. */
function openPanel(view: Awaited<ReturnType<typeof open>>['view']): HTMLElement {
  fireEvent.click(view.getByRole('button', { name: 'More' }));
  return view.getByRole('dialog', { name: 'More' });
}

describe('the named parts', () => {
  test('one part for each slot of the registry, from the core table', () => {
    const parts = XlsxEditor.Toolbar as unknown as Record<string, unknown>;
    for (const name of Object.keys(TOOLBAR_NAMED_PARTS))
      expect(typeof parts[name]).toBe('function');
  });

  test('hidden removes the slot where it stands, as Toolbar.Slot does', async () => {
    const { bar } = await open(
      <XlsxEditor.Toolbar>
        <XlsxEditor.Toolbar.Bold hidden />
        <XlsxEditor.Toolbar.FontSize hidden />
      </XlsxEditor.Toolbar>
    );
    expect(slot(bar, 'font.bold')).toBeNull();
    expect(slot(bar, 'font.size')).toBeNull();
    expect(slot(bar, 'font.italic')).toBeTruthy();
  });

  test('icon and className change the control, which keeps its label, state, and command', async () => {
    const { bar } = await open(
      <XlsxEditor.Toolbar>
        <XlsxEditor.Toolbar.Bold icon={<b data-host-icon="">B</b>} className="host-bold" />
      </XlsxEditor.Toolbar>
    );
    const bold = slot(bar, 'font.bold')!;
    expect(bold.closest('[data-group]')?.getAttribute('data-group')).toBe('font');
    expect(bold.className).toBe('xlsx-toolbar-button host-bold');
    expect(bold.querySelector('[data-host-icon]')).toBeTruthy();
    expect(bold.querySelector('svg')).toBeNull();
    expect(bold.getAttribute('aria-label')).toBe('Bold');
    expect(bold.getAttribute('aria-pressed')).toBe('false');
    fireEvent.click(bold);
    await waitFor(() => expect(slot(bar, 'font.bold')!.getAttribute('aria-pressed')).toBe('true'));
  });

  test('the children of a button part are its icon, not a replacement', async () => {
    const { bar } = await open(
      <XlsxEditor.Toolbar>
        <XlsxEditor.Toolbar.Italic>
          <i data-host-icon="">I</i>
        </XlsxEditor.Toolbar.Italic>
      </XlsxEditor.Toolbar>
    );
    const italic = slot(bar, 'font.italic')!;
    expect(italic.tagName).toBe('BUTTON');
    expect(italic.querySelector('[data-host-icon]')).toBeTruthy();
    expect(italic.getAttribute('aria-label')).toBe('Italic');
  });

  test('className reaches the trigger of a value part', async () => {
    const { bar } = await open(
      <XlsxEditor.Toolbar>
        <XlsxEditor.Toolbar.FontSize className="host-size" />
        <XlsxEditor.Toolbar.FillColor className="host-fill" />
      </XlsxEditor.Toolbar>
    );
    expect(slot(bar, 'font.size')!.classList.contains('host-size')).toBe(true);
    expect(slot(bar, 'fill.color')!.classList.contains('host-fill')).toBe(true);
  });

  test('a changed slot of a menu group leaves the menu and shows after the menu button', async () => {
    const { bar, view } = await open(
      <XlsxEditor.Toolbar>
        <XlsxEditor.Toolbar.InsertRows icon={<b>R</b>} />
      </XlsxEditor.Toolbar>
    );
    const cells = bar.querySelector<HTMLElement>('[data-group="cells"]')!;
    const own = slot(cells, 'cells.insertRows')!;
    expect(own.tagName).toBe('BUTTON');
    expect(own.textContent).toBe('R');
    fireEvent.click(cells.querySelector('[data-group-menu="cells"]')!);
    const menu = view.getByRole('menu');
    expect(
      [...menu.querySelectorAll<HTMLElement>('[data-slot]')].map((s) => s.dataset.slot)
    ).toEqual(['cells.insertColumns', 'cells.deleteRows', 'cells.deleteColumns']);
  });
});

describe('asChild', () => {
  test('a host element renders the control, with its label, state, and press', async () => {
    const { bar } = await open(
      <XlsxEditor.Toolbar>
        <XlsxEditor.Toolbar.Slot slotId="font.bold" asChild>
          <span className="host" data-own="">
            Strong
          </span>
        </XlsxEditor.Toolbar.Slot>
      </XlsxEditor.Toolbar>
    );
    const bold = slot(bar, 'font.bold')!;
    expect(bold.tagName).toBe('SPAN');
    expect(bold.className).toBe('xlsx-toolbar-button host');
    expect(bold.hasAttribute('data-own')).toBe(true);
    expect(bold.getAttribute('role')).toBe('button');
    expect(bold.tabIndex).toBe(0);
    expect(bold.getAttribute('aria-label')).toBe('Bold');
    expect(bold.textContent).toBe('Strong');
    fireEvent.click(bold);
    await waitFor(() => expect(slot(bar, 'font.bold')!.getAttribute('aria-pressed')).toBe('true'));
    // Enter presses an element that is not a button.
    fireEvent.keyDown(slot(bar, 'font.bold')!, { key: 'Enter' });
    await waitFor(() => expect(slot(bar, 'font.bold')!.getAttribute('aria-pressed')).toBe('false'));
  });

  test('a named part with asChild renders its single child', async () => {
    const { bar } = await open(
      <XlsxEditor.Toolbar>
        <XlsxEditor.Toolbar.Undo asChild>
          <button type="button" className="host-undo">
            Back
          </button>
        </XlsxEditor.Toolbar.Undo>
      </XlsxEditor.Toolbar>
    );
    const undo = slot(bar, 'history.undo')!;
    expect(undo.className).toBe('xlsx-toolbar-button host-undo');
    expect(undo.textContent).toBe('Back');
    expect(undo.getAttribute('aria-disabled')).toBe('true');
    expect(undo.hasAttribute('role')).toBe(false);
  });

  test('children that are not one element fall back to the built-in button, with a warning', async () => {
    let bar!: HTMLElement;
    const seen = await warnings(async () => {
      ({ bar } = await open(
        <XlsxEditor.Toolbar>
          <XlsxEditor.Toolbar.Slot slotId="font.underline" asChild>
            <span>one</span>
            <span>two</span>
          </XlsxEditor.Toolbar.Slot>
        </XlsxEditor.Toolbar>
      ));
    });
    expect(slot(bar, 'font.underline')!.tagName).toBe('BUTTON');
    expect(seen.some((w) => w.includes('"font.underline" has asChild'))).toBe(true);
  });

  test('icon and asChild on a value control are ignored, with a warning', async () => {
    let bar!: HTMLElement;
    const seen = await warnings(async () => {
      ({ bar } = await open(
        <XlsxEditor.Toolbar>
          <XlsxEditor.Toolbar.Slot slotId="font.family" icon={<b data-host-icon="">F</b>} />
        </XlsxEditor.Toolbar>
      ));
    });
    expect(bar.querySelector('[data-host-icon]')).toBeNull();
    expect(
      seen.some((w) => w.includes('"font.family" has icon, but the slot is not a button'))
    ).toBe(true);
  });
  test('icon on a color part replaces the glyph above the color bar, with no warning', async () => {
    let bar!: HTMLElement;
    const seen = await warnings(async () => {
      ({ bar } = await open(
        <XlsxEditor.Toolbar>
          <XlsxEditor.Toolbar.FontColor icon={<b data-host-icon="">A</b>} />
        </XlsxEditor.Toolbar>
      ));
    });
    const main = slot(bar, 'font.color')!;
    expect(main.querySelector('[data-host-icon]')).not.toBeNull();
    expect(main.querySelector('.xlsx-toolbar-color-bar')).not.toBeNull();
    expect(seen).toEqual([]);
  });
});

describe('XlsxEditor.Toolbar.Action', () => {
  test('a host button with the toolbar class, its label, and the mousedown guard', async () => {
    let pressed = 0;
    const { bar, view } = await open(
      <XlsxEditor.Toolbar>
        <XlsxEditor.Toolbar.Group id="review" label="Review">
          <XlsxEditor.Toolbar.Action
            label="Send"
            icon={<b data-host-icon="">S</b>}
            className="host-send"
            onSelect={() => (pressed += 1)}
          />
        </XlsxEditor.Toolbar.Group>
      </XlsxEditor.Toolbar>
    );
    const send = view.getByRole('button', { name: 'Send' });
    expect(send.closest('[data-group]')?.getAttribute('data-group')).toBe('review');
    expect(send.className).toBe('xlsx-toolbar-button host-send');
    expect(send.getAttribute('title')).toBe('Send');
    expect(send.querySelector('[data-host-icon]')).toBeTruthy();
    // The guard keeps the keyboard on the grid: the mousedown default is prevented.
    expect(fireEvent.mouseDown(send)).toBe(false);
    fireEvent.click(send);
    expect(pressed).toBe(1);
    expect(bar.contains(send)).toBe(true);
  });

  test('without an icon or children, the label shows as text', async () => {
    const { view } = await open(
      <XlsxEditor.Toolbar preset={false}>
        <XlsxEditor.Toolbar.Action label="Publish" />
      </XlsxEditor.Toolbar>
    );
    expect(view.getByRole('button', { name: 'Publish' }).textContent).toBe('Publish');
  });

  test('pressed state sets aria-pressed and data-active', async () => {
    const { view } = await open(
      <XlsxEditor.Toolbar preset={false}>
        <XlsxEditor.Toolbar.Action label="Track" active />
        <XlsxEditor.Toolbar.Action label="Plain" />
      </XlsxEditor.Toolbar>
    );
    const track = view.getByRole('button', { name: 'Track' });
    expect(track.getAttribute('aria-pressed')).toBe('true');
    expect(track.hasAttribute('data-active')).toBe(true);
    expect(view.getByRole('button', { name: 'Plain' }).hasAttribute('aria-pressed')).toBe(false);
  });

  test('a disabled action stays focusable, says why, and does not run', async () => {
    let pressed = 0;
    const { bar, view } = await open(
      <XlsxEditor.Toolbar preset={false}>
        <XlsxEditor.Toolbar.Slot slotId="font.bold" />
        <XlsxEditor.Toolbar.Action
          label="Send"
          disabled
          disabledReason="Save the workbook first."
          onSelect={() => (pressed += 1)}
        />
      </XlsxEditor.Toolbar>
    );
    const send = view.getByRole('button', { name: 'Send' });
    expect(send.hasAttribute('disabled')).toBe(false);
    expect(send.getAttribute('aria-disabled')).toBe('true');
    expect(send.hasAttribute('data-disabled')).toBe(true);
    expect(send.getAttribute('title')).toBe('Save the workbook first.');
    const reason = bar.querySelector(`#${CSS.escape(send.getAttribute('aria-describedby')!)}`);
    expect(reason?.textContent).toBe('Save the workbook first.');
    fireEvent.click(send);
    expect(pressed).toBe(0);
    // The arrow keys of the bar reach it, so the keyboard reaches its reason.
    const bold = slot(bar, 'font.bold')!;
    bold.focus();
    fireEvent.keyDown(bold, { key: 'ArrowRight' });
    expect(document.activeElement).toBe(send);
  });

  test('a held Enter runs the action once', async () => {
    const { view } = await open(
      <XlsxEditor.Toolbar preset={false}>
        <XlsxEditor.Toolbar.Action label="Step" />
      </XlsxEditor.Toolbar>
    );
    const step = view.getByRole('button', { name: 'Step' });
    expect(fireEvent.keyDown(step, { key: 'Enter', repeat: true })).toBe(false);
    expect(fireEvent.keyDown(step, { key: 'Enter' })).toBe(true);
  });

  test('asChild renders the host element with the action merged in', async () => {
    let pressed = 0;
    const { view } = await open(
      <XlsxEditor.Toolbar preset={false}>
        <XlsxEditor.Toolbar.Action label="Open" asChild onSelect={() => (pressed += 1)}>
          <a href="#open" className="host-link">
            Open
          </a>
        </XlsxEditor.Toolbar.Action>
      </XlsxEditor.Toolbar>
    );
    // A link keeps its own role, Tab stop, and Enter.
    const link = view.getByRole('link', { name: 'Open' });
    expect(link.tagName).toBe('A');
    expect(link.hasAttribute('role')).toBe(false);
    expect(link.hasAttribute('tabindex')).toBe(false);
    expect(link.className).toBe('xlsx-toolbar-button host-link');
    fireEvent.click(link);
    expect(pressed).toBe(1);
  });

  test('in a collapsed group, the action is a row of the More panel with its label', async () => {
    let pressed = 0;
    // The built-in groups take 700 pixels and "Review" 60; the More button takes 40.
    const restore = stubLayout(groupWidth, 759);
    try {
      const { bar, view } = await open(
        <XlsxEditor.Toolbar>
          <XlsxEditor.Toolbar.Group id="review" label="Review">
            <XlsxEditor.Toolbar.Action label="Send" icon={<b>S</b>} onSelect={() => pressed++} />
            <XlsxEditor.Toolbar.Separator />
          </XlsxEditor.Toolbar.Group>
        </XlsxEditor.Toolbar>
      );
      await waitFor(() => expect(barGroups(bar)).not.toContain('review'));
      const panel = openPanel(view);
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

describe('asChild and the More panel', () => {
  test('a disabled action on a link stops the navigation', async () => {
    const { view } = await open(
      <XlsxEditor.Toolbar preset={false}>
        <XlsxEditor.Toolbar.Action label="Open" asChild disabled>
          <a href="#open">Open</a>
        </XlsxEditor.Toolbar.Action>
      </XlsxEditor.Toolbar>
    );
    expect(fireEvent.click(view.getByRole('link', { name: 'Open' }))).toBe(false);
  });

  test('a disabled control on a link stops the navigation', async () => {
    const { bar } = await open(
      <XlsxEditor.Toolbar>
        <XlsxEditor.Toolbar.Undo asChild>
          <a href="#undo">Back</a>
        </XlsxEditor.Toolbar.Undo>
      </XlsxEditor.Toolbar>
    );
    const undo = slot(bar, 'history.undo')!;
    expect(undo.hasAttribute('role')).toBe(false);
    expect(fireEvent.click(undo)).toBe(false);
  });

  test('the More panel opened from the keyboard focuses a host element of asChild', async () => {
    const restore = stubLayout(groupWidth, 759);
    try {
      const { bar, view } = await open(
        <XlsxEditor.Toolbar>
          <XlsxEditor.Toolbar.Group id="review" label="Review">
            {/* A host button with `disabled` takes no focus, so the panel skips it. */}
            <button type="button" disabled>
              Off
            </button>
            <XlsxEditor.Toolbar.Action label="Go" asChild>
              <span>Go</span>
            </XlsxEditor.Toolbar.Action>
          </XlsxEditor.Toolbar.Group>
        </XlsxEditor.Toolbar>
      );
      await waitFor(() => expect(barGroups(bar)).not.toContain('review'));
      fireEvent.keyDown(view.getByRole('button', { name: 'More' }), { key: 'Enter' });
      await waitFor(() => expect(document.activeElement?.textContent).toBe('Go'));
      expect(document.activeElement?.tagName).toBe('SPAN');
    } finally {
      restore();
    }
  });
});

describe('XlsxEditor.Toolbar.Separator', () => {
  test('a vertical rule with its class', async () => {
    const { bar } = await open(
      <XlsxEditor.Toolbar preset={false}>
        <XlsxEditor.Toolbar.Separator className="host-rule" />
      </XlsxEditor.Toolbar>
    );
    const rule = bar.querySelector('[role="separator"]')!;
    expect(rule.getAttribute('aria-orientation')).toBe('vertical');
    expect(rule.className).toBe('xlsx-toolbar-separator host-rule');
  });
});

describe('priority and pinned', () => {
  // Each group is ten pixels for each letter of its heading: the built-in groups take 700, and
  // "Share" 50. The More button takes 40.
  test('a host group with no priority leaves the bar first', async () => {
    const restore = stubLayout(groupWidth, 400);
    try {
      const { bar } = await open(
        <XlsxEditor.Toolbar>
          <XlsxEditor.Toolbar.Group id="share" label="Share" after="history">
            <button type="button">Share it</button>
          </XlsxEditor.Toolbar.Group>
        </XlsxEditor.Toolbar>
      );
      await waitFor(() =>
        expect(barGroups(bar)).toEqual(['history', 'font', 'fill', 'borders', 'alignment'])
      );
    } finally {
      restore();
    }
  });

  test('a priority keeps a host group while built-in groups with a lower one leave', async () => {
    const restore = stubLayout(groupWidth, 400);
    try {
      const { bar, view } = await open(
        <XlsxEditor.Toolbar>
          <XlsxEditor.Toolbar.Group id="share" label="Share" priority={95}>
            <button type="button">Share it</button>
          </XlsxEditor.Toolbar.Group>
        </XlsxEditor.Toolbar>
      );
      // file (10), clear (20), cells (30), and number (40) leave: 790 less 390 is 400.
      await waitFor(() =>
        expect(barGroups(bar)).toEqual(['history', 'font', 'fill', 'borders', 'alignment', 'share'])
      );
      const panel = openPanel(view);
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
      const { bar } = await open(
        <XlsxEditor.Toolbar>
          <XlsxEditor.Toolbar.Group id="share" label="Share" pinned>
            <button type="button">Share it</button>
          </XlsxEditor.Toolbar.Group>
          <XlsxEditor.Toolbar.Group id="history" pinned />
        </XlsxEditor.Toolbar>
      );
      await waitFor(() => expect(barGroups(bar)).toEqual(['history', 'share']));
    } finally {
      restore();
    }
  });
});

describe('overflowContent', () => {
  test('of a group: in place of the children of a host group, after a built-in group', async () => {
    const restore = stubLayout(groupWidth, 100);
    try {
      const { view } = await open(
        <XlsxEditor.Toolbar>
          <XlsxEditor.Toolbar.Group
            id="share"
            label="Share"
            overflowContent={() => <button type="button">Share as a row</button>}
          >
            <button type="button">Share it</button>
          </XlsxEditor.Toolbar.Group>
          <XlsxEditor.Toolbar.Group
            id="file"
            overflowContent={() => <button type="button">Export</button>}
          />
        </XlsxEditor.Toolbar>
      );
      await waitFor(() => expect(view.queryByRole('button', { name: 'More' })).toBeTruthy());
      const panel = openPanel(view);
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
      const { bar, view } = await open(
        <XlsxEditor.Toolbar>
          <XlsxEditor.Toolbar.Save overflowContent={() => <button type="button">Upload</button>} />
          <XlsxEditor.Toolbar.Group id="history" pinned />
        </XlsxEditor.Toolbar>
      );
      await waitFor(() => expect(barGroups(bar)).toEqual(['history']));
      const panel = openPanel(view);
      const file = panel.querySelector<HTMLElement>('[aria-label="File"]')!;
      expect(slot(file, 'file.save')).toBeNull();
      expect(file.textContent).toContain('Upload');
    } finally {
      restore();
    }
  });
});

describe('preset={false}', () => {
  test('named parts, groups, actions, slots, and separators render in their order', async () => {
    const { bar } = await open(
      <XlsxEditor.Toolbar preset={false}>
        <XlsxEditor.Toolbar.Action label="First" />
        <XlsxEditor.Toolbar.Separator />
        <XlsxEditor.Toolbar.Bold className="host-bold" />
        <XlsxEditor.Toolbar.Group id="mine" label="Mine">
          <XlsxEditor.Toolbar.Italic />
          <XlsxEditor.Toolbar.Action label="Inner" />
        </XlsxEditor.Toolbar.Group>
        <XlsxEditor.Toolbar.Slot slotId="history.undo" />
        <XlsxEditor.Toolbar.DeleteRows />
      </XlsxEditor.Toolbar>
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
    // A slot with no registry icon shows its label in the bar.
    expect(slot(bar, 'cells.deleteRows')!.textContent).toBe('Delete rows');
    expect(bar.querySelector('[data-group="font"]')).toBeNull();
  });
});
