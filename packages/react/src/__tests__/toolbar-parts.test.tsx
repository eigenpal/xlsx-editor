/** @jsxImportSource react */
// The parts of the toolbar in happy-dom (`bun run test:dom`): `XlsxEditor.Toolbar.Slot` replaces
// or hides a slot where it stands, `XlsxEditor.Toolbar.Group` adds a host group that the bar
// measures and collapses with its own groups, `preset={false}` renders only the children, and
// each composition that cannot do what it says prints one warning. The Vue twin is
// packages/vue/test/toolbar-parts.test.ts. The development check itself is in
// toolbar-dev-warnings.test.ts.

import { afterEach, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { cleanup, fireEvent, render, waitFor } from '@testing-library/react';
import { Fragment, type ReactNode } from 'react';
import { buildXlsx, c, row } from '../../../../scripts/fixtures/builder.ts';
import { XlsxEditor, type FontSource, type XlsxEditorInstance } from '../index.ts';

const fontsDir = join(
  dirname(
    Bun.resolveSync('@docx-editor.dev/fonts/package.json', join(import.meta.dir, '../../../core'))
  ),
  'assets'
);
const FONTS: FontSource[] = [
  {
    family: 'Carlito',
    bold: false,
    italic: false,
    bytes: new Uint8Array(readFileSync(join(fontsDir, 'Carlito-Regular.ttf'))),
    origin: 'substitute',
  },
];

afterEach(cleanup);

const workbook = () => buildXlsx({ sheets: [{ name: 'S', rows: row(1, [c('A1', 1)]) }] });

/**
 * Mount the editor with `toolbar` as its toolbar, and wait for the workbook. `rerender` renders
 * the same editor with another toolbar.
 */
async function open(toolbar: ReactNode) {
  let viewer: XlsxEditorInstance | null = null;
  const document = workbook();
  const tree = (bar: ReactNode) => (
    <XlsxEditor.Root document={document} fonts={FONTS} onReady={(v) => (viewer = v)}>
      <XlsxEditor.Viewport style={{ height: 400 }}>
        {bar}
        <XlsxEditor.Content />
        <XlsxEditor.StatusMessage />
      </XlsxEditor.Viewport>
    </XlsxEditor.Root>
  );
  const view = render(tree(toolbar));
  await waitFor(() => expect(viewer?.snapshot().status).toBe('ready'), { timeout: 5000 });
  const rerender = (next: ReactNode) => view.rerender(tree(next));
  return { view, bar: view.getByRole('toolbar'), viewer: viewer!, rerender };
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

describe('XlsxEditor.Toolbar.Slot', () => {
  test('replaces its slot where it stands, in the bar', async () => {
    const { bar } = await open(
      <XlsxEditor.Toolbar>
        <XlsxEditor.Toolbar.Slot slotId="font.bold">
          <button type="button" data-slot="host.bold">
            Strong
          </button>
        </XlsxEditor.Toolbar.Slot>
      </XlsxEditor.Toolbar>
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

  test('with hidden, removes its slot, and a group with no slot left goes', async () => {
    const { bar } = await open(
      <XlsxEditor.Toolbar>
        <XlsxEditor.Toolbar.Slot slotId="font.underline" hidden />
        <XlsxEditor.Toolbar.Slot slotId="file.save" hidden />
      </XlsxEditor.Toolbar>
    );
    expect(slotsOf(bar, 'font')).not.toContain('font.underline');
    expect(barGroups(bar)).toEqual(BUILT_IN.filter((id) => id !== 'file'));
  });

  test('with no children or an empty string, keeps the built-in control', async () => {
    const { bar } = await open(
      <XlsxEditor.Toolbar>
        <XlsxEditor.Toolbar.Slot slotId="font.bold">{null}</XlsxEditor.Toolbar.Slot>
        <XlsxEditor.Toolbar.Slot slotId="font.italic">{''}</XlsxEditor.Toolbar.Slot>
        {''}
      </XlsxEditor.Toolbar>
    );
    expect(slotsOf(bar, 'font')).toContain('font.bold');
    expect(slotsOf(bar, 'font')).toContain('font.italic');
    expect(bar.querySelector('[data-toolbar-fixed]')).toBeNull();
  });

  test('a loose child and the content of a slot outside the registry keep their elements', async () => {
    const original = console.warn;
    console.warn = () => {};
    try {
      const toolbar = (show: boolean) => (
        <XlsxEditor.Toolbar>
          {show ? <input aria-label="L" /> : null}
          {show ? (
            <XlsxEditor.Toolbar.Slot slotId={'x.w' as 'font.bold'}>
              <input aria-label="W" />
            </XlsxEditor.Toolbar.Slot>
          ) : null}
          <XlsxEditor.Toolbar.Slot slotId={'x.y' as 'font.bold'}>
            <input aria-label="S" />
          </XlsxEditor.Toolbar.Slot>
          {show ? (
            <XlsxEditor.Toolbar.Slot slotId={'x.z' as 'font.bold'}>
              <input aria-label="Z" />
            </XlsxEditor.Toolbar.Slot>
          ) : null}
        </XlsxEditor.Toolbar>
      );
      const { view, rerender } = await open(toolbar(true));
      const field = view.getByRole('textbox', { name: 'S' }) as HTMLInputElement;
      field.value = 'typed';
      rerender(toolbar(false));
      expect(view.queryByRole('textbox', { name: 'L' })).toBeNull();
      expect(view.queryByRole('textbox', { name: 'W' })).toBeNull();
      expect(view.queryByRole('textbox', { name: 'Z' })).toBeNull();
      expect(view.getByRole('textbox', { name: 'S' })).toBe(field);
      expect(field.value).toBe('typed');
    } finally {
      console.warn = original;
    }
  });

  test('the content of two slots of one group keeps its own element when both change', async () => {
    const toolbar = (slot: 'bold' | 'underline') => (
      <XlsxEditor.Toolbar>
        <XlsxEditor.Toolbar.Slot slotId={`font.${slot}`}>
          <input aria-label={slot} />
        </XlsxEditor.Toolbar.Slot>
      </XlsxEditor.Toolbar>
    );
    const { view, rerender } = await open(toolbar('bold'));
    const bold = view.getByRole('textbox', { name: 'bold' });
    rerender(toolbar('underline'));
    // The bold control comes back, and the underline content is a new element.
    expect(view.getByRole('textbox', { name: 'underline' })).not.toBe(bold);
    expect(bold.isConnected).toBe(false);
  });

  test('two lists with the same keys stay apart when they change', async () => {
    const seen: string[] = [];
    const original = console.error;
    console.error = (...args: unknown[]) => void seen.push(String(args[0]));
    try {
      // Each list is a Fragment of keyed children, as a template loop renders one in Vue.
      const list = (name: string, order: number[]) => (
        <Fragment>
          {order.map((i) => (
            <button key={i} type="button">
              {`${name}${i}`}
            </button>
          ))}
        </Fragment>
      );
      const toolbar = (order: number[]) => (
        <XlsxEditor.Toolbar>
          <XlsxEditor.Toolbar.Group id="mine" label="Mine">
            {list('a', order)}
            {list('b', order)}
          </XlsxEditor.Toolbar.Group>
          {list('c', order)}
          {list('d', order)}
        </XlsxEditor.Toolbar>
      );
      const { bar, view, rerender } = await open(toolbar([0, 1]));
      const a0 = view.getByRole('button', { name: 'a0' });
      const c0 = view.getByRole('button', { name: 'c0' });
      rerender(toolbar([1, 0]));
      const texts = (el: Element | null) =>
        [...(el?.querySelectorAll('button') ?? [])].map((x) => x.textContent);
      expect(texts(bar.querySelector('[data-group="mine"]'))).toEqual(['a1', 'a0', 'b1', 'b0']);
      expect(texts(bar.querySelector('[data-toolbar-fixed]'))).toEqual(['c1', 'c0', 'd1', 'd0']);
      // A keyed child moves with its element.
      expect(view.getByRole('button', { name: 'a0' })).toBe(a0);
      expect(view.getByRole('button', { name: 'c0' })).toBe(c0);
      expect(seen.filter((m) => m.includes('same key'))).toEqual([]);
    } finally {
      console.error = original;
    }
  });

  test('loose children in a Fragment keep their own state when a sibling goes', async () => {
    const seen: string[] = [];
    const original = console.error;
    console.error = (...args: unknown[]) => void seen.push(String(args[0]));
    try {
      const toolbar = (show: boolean) => (
        <XlsxEditor.Toolbar>
          {show ? <input aria-label="A" defaultValue="a" /> : null}
          <>
            <input aria-label="B" defaultValue="b" />
          </>
        </XlsxEditor.Toolbar>
      );
      const { view, rerender } = await open(toolbar(true));
      const b = view.getByRole('textbox', { name: 'B' }) as HTMLInputElement;
      b.value = 'typed';
      rerender(toolbar(false));
      expect((view.getByRole('textbox', { name: 'B' }) as HTMLInputElement).value).toBe('typed');
      expect(seen.filter((m) => m.includes('same key'))).toEqual([]);
    } finally {
      console.error = original;
    }
  });

  test('the parts inside a Fragment are read as parts, and the last part for a slot wins', async () => {
    const { bar } = await open(
      <XlsxEditor.Toolbar>
        <XlsxEditor.Toolbar.Slot slotId="font.bold" hidden />
        <>
          <XlsxEditor.Toolbar.Slot slotId="font.bold">
            <button type="button" data-slot="host.bold">
              Strong
            </button>
          </XlsxEditor.Toolbar.Slot>
          <>
            <XlsxEditor.Toolbar.Slot slotId="font.italic" hidden />
            <XlsxEditor.Toolbar.Group id="borders" hidden />
          </>
        </>
      </XlsxEditor.Toolbar>
    );
    expect(slotsOf(bar, 'font')).toContain('host.bold');
    expect(slotsOf(bar, 'font')).not.toContain('font.bold');
    expect(slotsOf(bar, 'font')).not.toContain('font.italic');
    expect(barGroups(bar)).not.toContain('borders');
    expect(bar.querySelector('[data-toolbar-fixed]')).toBeNull();
  });

  test('the hidden list wins over its content, with one warning for the page', async () => {
    const toolbar = (
      <XlsxEditor.Toolbar hidden={['font.italic']}>
        <XlsxEditor.Toolbar.Slot slotId="font.italic">
          <button type="button">Slanted</button>
        </XlsxEditor.Toolbar.Slot>
      </XlsxEditor.Toolbar>
    );
    let bar!: HTMLElement;
    let view!: Awaited<ReturnType<typeof open>>['view'];
    const seen = await warnings(async () => {
      ({ bar, view } = await open(toolbar));
      expect(view.queryByRole('button', { name: 'Slanted' })).toBeNull();
      expect(slotsOf(bar, 'font')).not.toContain('font.italic');
      // A new toolbar with the same problem prints nothing new.
      cleanup();
      await open(toolbar);
    });
    expect(seen).toEqual([
      '[xlsx-editor] XlsxEditor.Toolbar.Slot "font.italic" has content, but the hidden list ' +
        'names the slot, so the slot is left out. Remove it from the hidden list to show the ' +
        'content.',
    ]);
  });

  test('an id outside the registry shows its content after the groups, with a warning', async () => {
    let bar!: HTMLElement;
    const seen = await warnings(async () => {
      ({ bar } = await open(
        <XlsxEditor.Toolbar>
          <XlsxEditor.Toolbar.Slot slotId={'font.strike' as 'font.bold'}>
            <button type="button">Strike</button>
          </XlsxEditor.Toolbar.Slot>
        </XlsxEditor.Toolbar>
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
      ({ bar } = await open(
        <XlsxEditor.Toolbar preset={false}>
          <XlsxEditor.Toolbar.Slot slotId={'font.shadow' as 'font.bold'} />
        </XlsxEditor.Toolbar>
      ));
    });
    expect(bar.childElementCount).toBe(0);
    expect(seen).toEqual([
      '[xlsx-editor] "font.shadow" is not a toolbar slot id, so its XlsxEditor.Toolbar.Slot ' +
        'renders nothing. Use a ChromeSlotId, such as "font.bold".',
    ]);
  });

  test('a new render applies a changed part and a changed hidden list', async () => {
    const { bar, rerender } = await open(
      <XlsxEditor.Toolbar>
        <XlsxEditor.Toolbar.Slot slotId="font.bold">
          <button type="button" data-slot="host.bold">
            Strong
          </button>
        </XlsxEditor.Toolbar.Slot>
      </XlsxEditor.Toolbar>
    );
    expect(slotsOf(bar, 'font')).toContain('host.bold');
    rerender(
      <XlsxEditor.Toolbar hidden={['font.italic']}>
        <XlsxEditor.Toolbar.Slot slotId="font.bold" hidden />
        <XlsxEditor.Toolbar.Group id="share" label="Share">
          <button type="button">Share</button>
        </XlsxEditor.Toolbar.Group>
      </XlsxEditor.Toolbar>
    );
    await waitFor(() => expect(barGroups(bar)).toEqual([...BUILT_IN, 'share']));
    expect(slotsOf(bar, 'font')).toEqual([
      'font.family',
      'font.size',
      'font.underline',
      'font.color',
    ]);
  });

  test('a replaced number format list leaves the dialog opener as its own button', async () => {
    const { bar, view } = await open(
      <XlsxEditor.Toolbar>
        <XlsxEditor.Toolbar.Slot slotId="number.format">
          <button type="button">Formats</button>
        </XlsxEditor.Toolbar.Slot>
      </XlsxEditor.Toolbar>
    );
    expect(slotsOf(bar, 'number')).toContain('number.custom');
    expect(view.getByRole('button', { name: 'More number formats…' })).toBeTruthy();
  });

  test('in a menu group, the content follows the menu, which keeps the other items', async () => {
    const { bar, view } = await open(
      <XlsxEditor.Toolbar>
        <XlsxEditor.Toolbar.Slot slotId="clear.all">
          <button type="button">Wipe</button>
        </XlsxEditor.Toolbar.Slot>
      </XlsxEditor.Toolbar>
    );
    const group = bar.querySelector('[data-group="clear"]')!;
    expect(group.lastElementChild?.textContent).toBe('Wipe');
    fireEvent.click(group.querySelector('button')!);
    const items = view.getAllByRole('menuitem').map((m) => m.textContent);
    expect(items).toEqual(['Clear contents', 'Clear formats']);
  });
});

describe('XlsxEditor.Toolbar.Group', () => {
  test('a host group follows every built-in group, or the group it names', async () => {
    const { bar } = await open(
      <XlsxEditor.Toolbar>
        <XlsxEditor.Toolbar.Group id="share" label="Share">
          <button type="button">Share</button>
        </XlsxEditor.Toolbar.Group>
        <XlsxEditor.Toolbar.Group id="review" label="Review" after="history" className="mine">
          <button type="button">Review</button>
        </XlsxEditor.Toolbar.Group>
      </XlsxEditor.Toolbar>
    );
    expect(barGroups(bar)).toEqual(['history', 'review', ...BUILT_IN.slice(1), 'share']);
    const review = bar.querySelector('[data-group="review"]')!;
    expect(review.getAttribute('aria-label')).toBe('Review');
    expect(review.className).toBe('xlsx-toolbar-group mine');
  });

  test('a host group with no content or with hidden is left out', async () => {
    const { bar } = await open(
      <XlsxEditor.Toolbar>
        <XlsxEditor.Toolbar.Group id="empty" label="Empty" />
        <XlsxEditor.Toolbar.Group id="gone" label="Gone" hidden>
          <button type="button">Gone</button>
        </XlsxEditor.Toolbar.Group>
      </XlsxEditor.Toolbar>
    );
    expect(barGroups(bar)).toEqual(BUILT_IN);
  });

  test('a built-in id adds its content after the group controls, or hides the group', async () => {
    const { bar } = await open(
      <XlsxEditor.Toolbar>
        <XlsxEditor.Toolbar.Group id="file">
          <button type="button" data-slot="host.share">
            Share
          </button>
        </XlsxEditor.Toolbar.Group>
        <XlsxEditor.Toolbar.Group id="borders" hidden />
      </XlsxEditor.Toolbar>
    );
    expect(slotsOf(bar, 'file')).toEqual(['file.save', 'host.share']);
    expect(barGroups(bar)).toEqual(BUILT_IN.filter((id) => id !== 'borders'));
  });

  test('a Slot with no children in a host group renders the built-in control', async () => {
    const { bar } = await open(
      <XlsxEditor.Toolbar>
        <XlsxEditor.Toolbar.Slot slotId="font.bold" hidden />
        <XlsxEditor.Toolbar.Group id="mine" label="Mine">
          <XlsxEditor.Toolbar.Slot slotId="font.bold" />
        </XlsxEditor.Toolbar.Group>
      </XlsxEditor.Toolbar>
    );
    expect(slotsOf(bar, 'font')).not.toContain('font.bold');
    expect(slotsOf(bar, 'mine')).toEqual(['font.bold']);
  });

  test('the bar measures a host group and moves it into the More panel with its heading', async () => {
    // The built-in groups take 700 pixels and "Share" 50; the More button takes 40.
    const restore = stubLayout(widths(), 749);
    try {
      const { bar, view } = await open(
        <XlsxEditor.Toolbar>
          <XlsxEditor.Toolbar.Group id="share" label="Share">
            <button type="button">Share it</button>
          </XlsxEditor.Toolbar.Group>
          <XlsxEditor.Toolbar.Group id="team" labelKey="formattingBar.groups.file">
            <button type="button">Team</button>
          </XlsxEditor.Toolbar.Group>
          <XlsxEditor.Toolbar.Group id="plain">
            <XlsxEditor.Toolbar.Slot slotId="font.bold" />
          </XlsxEditor.Toolbar.Group>
        </XlsxEditor.Toolbar>
      );
      await waitFor(() => expect(barGroups(bar)).toEqual(BUILT_IN));
      fireEvent.click(view.getByRole('button', { name: 'More' }));
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
      const { bar, view } = await open(
        <XlsxEditor.Toolbar>
          <XlsxEditor.Toolbar.Group id="share" label="Share">
            <button type="button">Share it</button>
          </XlsxEditor.Toolbar.Group>
        </XlsxEditor.Toolbar>
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
      const { bar, view } = await open(
        <XlsxEditor.Toolbar>
          <XlsxEditor.Toolbar.Slot slotId="file.save">
            <button type="button">Upload</button>
          </XlsxEditor.Toolbar.Slot>
          <XlsxEditor.Toolbar.Group id="file">
            <button type="button">Export</button>
          </XlsxEditor.Toolbar.Group>
        </XlsxEditor.Toolbar>
      );
      await waitFor(() => expect(barGroups(bar)).toEqual(BUILT_IN.slice(0, 7)));
      fireEvent.click(view.getByRole('button', { name: 'More' }));
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
      ({ bar } = await open(
        <XlsxEditor.Toolbar>
          <XlsxEditor.Toolbar.Group id="font" label="Type" after="file">
            <button type="button">Extra</button>
          </XlsxEditor.Toolbar.Group>
          <XlsxEditor.Toolbar.Group id="late" after="nowhere">
            <button type="button">Late</button>
          </XlsxEditor.Toolbar.Group>
        </XlsxEditor.Toolbar>
      ));
    });
    expect(barGroups(bar)).toEqual([...BUILT_IN, 'late']);
    expect(bar.querySelector('[data-group="font"]')?.getAttribute('aria-label')).toBe('Font');
    expect(seen.map((m) => m.split(',')[0])).toEqual([
      '[xlsx-editor] XlsxEditor.Toolbar.Group "font" is a built-in group',
      '[xlsx-editor] XlsxEditor.Toolbar.Group "font" is a built-in group',
      '[xlsx-editor] XlsxEditor.Toolbar.Group "late" has after="nowhere"',
    ]);
  });
});

describe('the arrow keys in the preset bar', () => {
  test('reach a control at any depth in a host group and in the fixed block', async () => {
    // A component of the host renders the part, so the toolbar shows it as a loose child.
    const Color = () => <XlsxEditor.Toolbar.Slot slotId="font.color" />;
    const { bar, view } = await open(
      <XlsxEditor.Toolbar>
        <XlsxEditor.Toolbar.Group id="share" label="Share">
          <span title="Share the workbook">
            <button type="button">Share it</button>
          </span>
          <div hidden>
            <button type="button">Hidden</button>
          </div>
          <div aria-hidden="true">
            <button type="button">Unseen</button>
          </div>
        </XlsxEditor.Toolbar.Group>
        <Color />
        <a href="#help">Help</a>
        <button type="button" disabled>
          Off
        </button>
        <button type="button" tabIndex={-1}>
          Out
        </button>
        <select aria-label="Pick">
          <option>One</option>
        </select>
        <span role="button" tabIndex={0}>
          Chip
        </span>
        <textarea aria-label="Memo" />
        <input aria-label="Note" defaultValue="ab" />
      </XlsxEditor.Toolbar>
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
    const { bar, view } = await open(
      <XlsxEditor.Toolbar>
        <XlsxEditor.Toolbar.Group id="mine" label="Mine">
          <button type="button" onKeyDown={(e) => e.preventDefault()}>
            Own
          </button>
          <button type="button">Next</button>
          <div role="menu">
            <button type="button">One</button>
            <button type="button">Two</button>
          </div>
          <div role="dialog">
            <button type="button">Three</button>
          </div>
        </XlsxEditor.Toolbar.Group>
      </XlsxEditor.Toolbar>
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
    const { bar, view } = await open(
      <XlsxEditor.Toolbar>
        <XlsxEditor.Toolbar.Group id="mine" label="Mine">
          <button type="button">A</button>
          <input type="hidden" />
          <button
            type="button"
            ref={(el) => {
              // A control that is not rendered refuses the focus.
              if (el) el.focus = () => undefined;
            }}
          >
            Ghost
          </button>
          <button type="button">B</button>
        </XlsxEditor.Toolbar.Group>
      </XlsxEditor.Toolbar>
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

describe('loose children', () => {
  test('show in one block after the More button, which the bar charges as fixed width', async () => {
    // 849 less the block's 100 leaves 749: the host group "Share" moves into the panel.
    const restore = stubLayout(widths(100), 849);
    try {
      const { bar, view } = await open(
        <XlsxEditor.Toolbar>
          <span>Saved</span>
          <XlsxEditor.Toolbar.Group id="share" label="Share">
            <button type="button">Share it</button>
          </XlsxEditor.Toolbar.Group>
        </XlsxEditor.Toolbar>
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
      const share = (
        <XlsxEditor.Toolbar.Group id="share" label="Share">
          <button type="button">Share it</button>
        </XlsxEditor.Toolbar.Group>
      );
      const { bar, rerender } = await open(<XlsxEditor.Toolbar>{share}</XlsxEditor.Toolbar>);
      await waitFor(() => expect(barGroups(bar)).toEqual([...BUILT_IN, 'share']));
      rerender(
        <XlsxEditor.Toolbar>
          {share}
          <span>Saved</span>
        </XlsxEditor.Toolbar>
      );
      await waitFor(() => expect(barGroups(bar)).toEqual(BUILT_IN));
      rerender(<XlsxEditor.Toolbar>{share}</XlsxEditor.Toolbar>);
      await waitFor(() => expect(barGroups(bar)).toEqual([...BUILT_IN, 'share']));
    } finally {
      restore();
    }
  });
});

describe('preset={false}', () => {
  test('renders only the children, in order, with no measure and no More panel', async () => {
    const restore = stubLayout(widths(), 100);
    try {
      const { bar, view } = await open(
        <XlsxEditor.Toolbar preset={false}>
          <XlsxEditor.Toolbar.Group id="mine" label="Mine" className="host">
            <XlsxEditor.Toolbar.Slot slotId="font.bold" />
            <XlsxEditor.Toolbar.Slot slotId="font.italic" hidden />
            <XlsxEditor.Toolbar.Slot slotId="history.undo">
              <button type="button">Back</button>
            </XlsxEditor.Toolbar.Slot>
          </XlsxEditor.Toolbar.Group>
          <XlsxEditor.Toolbar.Slot slotId="file.save" />
        </XlsxEditor.Toolbar>
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
    const { bar, view } = await open(
      <XlsxEditor.Toolbar preset={false}>
        <XlsxEditor.Toolbar.Group id="mine" label="Mine">
          <XlsxEditor.Toolbar.Slot slotId="font.bold" />
          <button type="button">Own</button>
          <XlsxEditor.Toolbar.Slot slotId="font.size" />
        </XlsxEditor.Toolbar.Group>
        <XlsxEditor.Toolbar.Slot slotId="font.color" />
        <button type="button">Loose</button>
      </XlsxEditor.Toolbar>
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

  test('a group with no children is left out, and a built-in group has its own heading', async () => {
    const { bar } = await open(
      <XlsxEditor.Toolbar preset={false}>
        <XlsxEditor.Toolbar.Group id="share" label="Share">
          {false}
        </XlsxEditor.Toolbar.Group>
        <XlsxEditor.Toolbar.Group id="font" label="Text">
          <XlsxEditor.Toolbar.Slot slotId="font.bold" />
        </XlsxEditor.Toolbar.Group>
      </XlsxEditor.Toolbar>
    );
    expect(bar.querySelector('[data-host-group="share"]')).toBeNull();
    expect(bar.querySelector('[data-host-group="font"]')!.getAttribute('aria-label')).toBe('Font');
  });

  test('a hidden Slot and a hidden Group render nothing, also with children', async () => {
    const { bar } = await open(
      <XlsxEditor.Toolbar preset={false}>
        <XlsxEditor.Toolbar.Slot slotId="font.bold" hidden>
          <button type="button">Strong</button>
        </XlsxEditor.Toolbar.Slot>
        <XlsxEditor.Toolbar.Group id="mine" label="Mine" hidden>
          <button type="button">Share</button>
        </XlsxEditor.Toolbar.Group>
      </XlsxEditor.Toolbar>
    );
    expect(bar.childElementCount).toBe(0);
  });

  test('a built-in control there runs its command', async () => {
    const { view, viewer } = await open(
      <XlsxEditor.Toolbar preset={false}>
        <XlsxEditor.Toolbar.Slot slotId="font.bold" />
      </XlsxEditor.Toolbar>
    );
    const before = viewer.snapshot().revision;
    fireEvent.click(view.getByRole('button', { name: 'Bold' }));
    await waitFor(() => expect(viewer.snapshot().revision).toBeGreaterThan(before));
  });
});
