// The arrangement of the preset toolbar with the host's parts, shared by both adapters. The
// placement case is a copy of docx-editor/packages/core/src/editor/__tests__/toolbar-room.test.ts
// at 10eeaf67 (EigenPal). The measure with a fixed block uses a DOM of this file's own: the core
// run has no global DOM (`bun run test`).

import { describe, expect, test } from 'bun:test';
import { GlobalWindow } from 'happy-dom';
import {
  arrangeToolbar,
  arrangeToolbarGroups,
  toolbarGroupHeading,
  type ToolbarArrangement,
} from '../toolbar-arrange.ts';
import { TOOLBAR_FIXED_ATTRIBUTE, measureToolbar, observeToolbar } from '../toolbar-layout.ts';
import { CHROME_GROUPS } from '../chrome-controls.ts';

const BUILT_IN = CHROME_GROUPS.map((g) => g.id);
const ids = (a: ToolbarArrangement) => a.groups.map((g) => g.id);
const controls = (a: ToolbarArrangement, id: string) =>
  a.groups.find((g) => g.id === id)?.chrome?.controls.map((c) => c.id);

describe('arrangeToolbarGroups', () => {
  test('host groups follow the built-in groups, or the group they name', () => {
    expect(
      arrangeToolbarGroups(
        ['history', 'text', 'review'],
        [
          { id: 'a' },
          { id: 'b', after: 'history' },
          { id: 'c', after: 'history' },
          { id: 'd', after: 'b' },
        ]
      )
    ).toEqual(['history', 'b', 'd', 'c', 'text', 'review', 'a']);
    expect(arrangeToolbarGroups(['history'], [{ id: 'x', after: 'missing' }])).toEqual([
      'history',
      'x',
    ]);
  });

  test('a cycle of anchors places its groups at the end, and a built-in id is not a host', () => {
    expect(
      arrangeToolbarGroups(
        ['history', 'font'],
        [
          { id: 'a', after: 'b' },
          { id: 'b', after: 'a' },
          { id: 'font', after: 'history' },
          { id: 'self', after: 'self' },
        ]
      )
    ).toEqual(['history', 'font', 'self', 'a', 'b']);
  });
});

describe('an anchor that leads into a cycle', () => {
  test('places the group last, with a problem for each group of the cycle and after it', () => {
    const a = arrangeToolbar({
      groups: [
        { id: 'a', after: 'b', content: true },
        { id: 'b', after: 'a', content: true },
        { id: 'self', after: 'self', content: true },
        { id: 'free', after: 'history', content: true },
        { id: 'next', after: 'free', content: true },
        { id: 'tail', after: 'a', content: true },
      ],
    });
    expect(ids(a)).toEqual([
      'history',
      'free',
      'next',
      ...BUILT_IN.slice(1),
      'self',
      'a',
      'b',
      'tail',
    ]);
    const cycle = (id: string, after: string, loop: string) =>
      `XlsxEditor.Toolbar.Group "${id}" has after="${after}", which leads into a cycle of ` +
      `anchors: ${loop}. The group is placed after every other group. Change the after of one ` +
      'group of the cycle.';
    expect(a.problems).toEqual([
      cycle('a', 'b', '"a", "b"'),
      cycle('b', 'a', '"b", "a"'),
      cycle('tail', 'a', '"a", "b"'),
      'XlsxEditor.Toolbar.Group "self" has after="self", its own id, so the group is placed as a ' +
        'group with no after. Name another group.',
    ]);
  });

  test('a group before its anchor in a chain that names no group gets a problem', () => {
    const a = arrangeToolbar({
      groups: [
        { id: 'a', after: 'b', content: true },
        { id: 'b', after: 'c', content: true },
        { id: 'c', after: 'nope', content: true },
      ],
    });
    expect(ids(a)).toEqual([...BUILT_IN, 'a', 'b', 'c']);
    const before = (id: string, after: string) =>
      `XlsxEditor.Toolbar.Group "${id}" has after="${after}", but its chain of anchors ends at ` +
      `"c", whose after="nope" names no group. So "${after}" is placed after every other group, ` +
      'and this group stands before it. Fix the after of "c".';
    expect(a.problems).toEqual([
      `XlsxEditor.Toolbar.Group "c" has after="nope", but no group has that id. The group is ` +
        `placed after every other group. Known ids: ${[...BUILT_IN, 'a', 'b', 'c'].join(', ')}.`,
      before('a', 'b'),
      before('b', 'c'),
    ]);
  });

  test('a group after a group that names itself follows it, with no problem of its own', () => {
    const a = arrangeToolbar({
      groups: [
        { id: 't', after: 's', content: true },
        { id: 's', after: 's', content: true },
        { id: 'x', content: true },
      ],
    });
    expect(ids(a)).toEqual([...BUILT_IN, 's', 't', 'x']);
    expect(a.problems).toEqual([
      'XlsxEditor.Toolbar.Group "s" has after="s", its own id, so the group is placed as a ' +
        'group with no after. Name another group.',
    ]);
  });

  test('a group after a hidden built-in group keeps its place, with no problem', () => {
    const a = arrangeToolbar({
      groups: [
        { id: 'borders', hidden: true, content: false },
        { id: 'mine', after: 'borders', content: true },
      ],
    });
    expect(ids(a)).toEqual(['history', 'font', 'fill', 'mine', ...BUILT_IN.slice(4)]);
    expect(a.problems).toEqual([]);
  });

  test('a chain of anchors with no cycle has no problem', () => {
    const a = arrangeToolbar({
      groups: [
        { id: 'a', content: true },
        { id: 'b', after: 'a', content: true },
      ],
    });
    expect(ids(a)).toEqual([...BUILT_IN, 'a', 'b']);
    expect(a.problems).toEqual([]);
  });

  test('a group after a missing anchor gets no cycle problem, and only the first one a problem', () => {
    const a = arrangeToolbar({
      groups: [
        { id: 'x', after: 'nowhere', content: true },
        { id: 'y', after: 'x', content: true },
      ],
    });
    expect(ids(a)).toEqual([...BUILT_IN, 'x', 'y']);
    expect(a.problems).toEqual([
      `XlsxEditor.Toolbar.Group "x" has after="nowhere", but no group has that id. The group is ` +
        `placed after every other group. Known ids: ${[...BUILT_IN, 'x', 'y'].join(', ')}.`,
    ]);
  });
});

describe('arrangeToolbar', () => {
  test('with no parts, the bar is the registry, and the hidden list leaves out slots', () => {
    expect(ids(arrangeToolbar({}))).toEqual(BUILT_IN);
    const a = arrangeToolbar({ hidden: ['history.undo', 'history.redo', 'font.bold'] });
    expect(ids(a)).toEqual(BUILT_IN.filter((id) => id !== 'history'));
    expect(controls(a, 'font')).toEqual(['family', 'size', 'italic', 'underline', 'color']);
    expect(a.problems).toEqual([]);
  });

  test('a hidden slot part removes its slot, and a group with no slot left goes', () => {
    const a = arrangeToolbar({
      slots: [
        { slotId: 'file.save', hidden: true, content: false },
        { slotId: 'font.underline', hidden: true, content: false },
      ],
    });
    expect(ids(a)).not.toContain('file');
    expect(controls(a, 'font')).toEqual(['family', 'size', 'bold', 'italic', 'color']);
    expect(a.replaced.size).toBe(0);
  });

  test('a slot part with content replaces its slot, and one with no content changes nothing', () => {
    const a = arrangeToolbar({
      slots: [
        { slotId: 'font.bold', hidden: false, content: true },
        { slotId: 'font.italic', hidden: false, content: false },
      ],
    });
    expect([...a.replaced]).toEqual(['font.bold']);
    // The slot keeps its place in the group: the adapter renders the content there.
    expect(controls(a, 'font')).toContain('bold');
    expect(ids(a)).toEqual(BUILT_IN);
  });

  test('the last slot part wins', () => {
    const a = arrangeToolbar({
      slots: [
        { slotId: 'font.bold', hidden: false, content: true },
        { slotId: 'font.bold', hidden: true, content: false },
      ],
    });
    expect(a.replaced.size).toBe(0);
    expect(controls(a, 'font')).not.toContain('bold');
    const b = arrangeToolbar({
      slots: [
        { slotId: 'font.bold', hidden: true, content: false },
        { slotId: 'font.bold', hidden: false, content: true },
      ],
    });
    expect([...b.replaced]).toEqual(['font.bold']);
  });

  test('the hidden list wins over slot content, with a problem', () => {
    const a = arrangeToolbar({
      hidden: ['file.save'],
      slots: [{ slotId: 'file.save', hidden: false, content: true }],
    });
    expect(ids(a)).not.toContain('file');
    expect(a.replaced.size).toBe(0);
    expect(a.problems).toEqual([
      'XlsxEditor.Toolbar.Slot "file.save" has content, but the hidden list names the slot, so ' +
        'the slot is left out. Remove it from the hidden list to show the content.',
    ]);
  });

  test('a slot id outside the registry is appended, with a problem', () => {
    const a = arrangeToolbar({ slots: [{ slotId: 'font.strike', hidden: false, content: true }] });
    expect(a.outside).toEqual(['font.strike']);
    expect(a.problems).toEqual([
      '"font.strike" is not a toolbar slot id, so the content of its XlsxEditor.Toolbar.Slot is ' +
        'appended after the groups. Use a ChromeSlotId, such as "font.bold".',
    ]);
    expect(ids(a)).toEqual(BUILT_IN);
  });

  test('a slot id outside the registry with hidden or no content changes nothing', () => {
    const a = arrangeToolbar({
      slots: [
        { slotId: 'font.shadow', hidden: true, content: true },
        { slotId: 'font.outline', hidden: false, content: false },
      ],
    });
    expect(a.outside).toEqual([]);
    expect(ids(a)).toEqual(BUILT_IN);
    expect(a.problems).toEqual(
      ['font.shadow', 'font.outline'].map(
        (id) =>
          `"${id}" is not a toolbar slot id, so its XlsxEditor.Toolbar.Slot changes nothing. ` +
          'Use a ChromeSlotId, such as "font.bold".'
      )
    );
  });

  test('slot content in a hidden group is left out, with a problem', () => {
    const a = arrangeToolbar({
      slots: [{ slotId: 'file.save', hidden: false, content: true }],
      groups: [{ id: 'file', hidden: true, content: false }],
    });
    expect(ids(a)).not.toContain('file');
    expect(a.problems).toEqual([
      'XlsxEditor.Toolbar.Slot "file.save" has content, but its group "file" is hidden, so the ' +
        'content is left out. Show the group, or place the content in a host group.',
    ]);
  });

  test('the problems name the parts as the adapter names them', () => {
    const a = arrangeToolbar({
      hidden: ['file.save'],
      slots: [{ slotId: 'file.save', hidden: false, content: true }],
      groups: [{ id: 'x', after: 'nowhere', content: true }],
      names: { slot: 'Toolbar.Slot', group: 'Toolbar.Group' },
    });
    expect(a.problems.map((p) => p.split(' "')[0])).toEqual(['Toolbar.Slot', 'Toolbar.Group']);
  });

  test('a host group follows every built-in group, or the group it names', () => {
    const a = arrangeToolbar({
      groups: [
        { id: 'share', content: true },
        { id: 'review', after: 'history', label: 'Review', content: true },
        { id: 'late', after: 'review', content: true },
      ],
    });
    expect(ids(a)).toEqual(['history', 'review', 'late', ...BUILT_IN.slice(1), 'share']);
    const review = a.groups.find((g) => g.id === 'review')!;
    expect(review).toEqual({ id: 'review', chrome: null, label: 'Review', labelKey: undefined });
    expect(a.problems).toEqual([]);
  });

  test('a host group with no content or with hidden goes, but still anchors a group', () => {
    const a = arrangeToolbar({
      groups: [
        { id: 'empty', after: 'font', content: false },
        { id: 'gone', after: 'empty', hidden: true, content: true },
        { id: 'next', after: 'gone', content: true },
      ],
    });
    expect(ids(a)).toEqual(['history', 'font', 'next', ...BUILT_IN.slice(2)]);
  });

  test('a built-in id adds content to its group, or hides it', () => {
    const a = arrangeToolbar({
      hidden: ['file.save'],
      groups: [
        { id: 'file', content: true },
        { id: 'borders', hidden: true, content: false },
      ],
    });
    // The file group has no control left, but its part adds content: it stays, empty of
    // controls, in its place.
    expect(ids(a)).toEqual(BUILT_IN.filter((id) => id !== 'borders'));
    expect(controls(a, 'file')).toEqual([]);
  });

  test('a built-in group ignores a label, a label key, and an anchor, with a problem each', () => {
    const a = arrangeToolbar({
      groups: [{ id: 'font', label: 'Type', labelKey: 'x', after: 'file', content: true }],
    });
    expect(ids(a)).toEqual(BUILT_IN);
    expect(a.problems).toEqual(
      (['label', 'labelKey', 'after'] as const).map(
        (setting) =>
          `XlsxEditor.Toolbar.Group "font" is a built-in group, so its ${setting} is ignored. ` +
          `A built-in group keeps its own label and place in the bar.`
      )
    );
  });

  test('an anchor that names no group places the group last, with a problem', () => {
    const a = arrangeToolbar({ groups: [{ id: 'x', after: 'nowhere', content: true }] });
    expect(ids(a).at(-1)).toBe('x');
    expect(a.problems).toEqual([
      `XlsxEditor.Toolbar.Group "x" has after="nowhere", but no group has that id. The group is ` +
        `placed after every other group. Known ids: ${[...BUILT_IN, 'x'].join(', ')}.`,
    ]);
  });

  test('the last group part wins, at the place of its appearance', () => {
    const a = arrangeToolbar({
      groups: [
        { id: 'a', content: true },
        { id: 'b', content: true },
        { id: 'a', label: 'Again', content: true },
      ],
    });
    expect(ids(a).slice(-2)).toEqual(['b', 'a']);
    expect(a.groups.at(-1)?.label).toBe('Again');
  });

  test('the key changes with the groups, their controls, and the replaced slots only', () => {
    const base = arrangeToolbar({ groups: [{ id: 'a', label: 'One', content: true }] });
    expect(arrangeToolbar({ groups: [{ id: 'a', label: 'Two', content: true }] }).key).toBe(
      base.key
    );
    const keys = [
      arrangeToolbar({}).key,
      base.key,
      arrangeToolbar({ hidden: ['font.bold'] }).key,
      arrangeToolbar({ slots: [{ slotId: 'font.bold', hidden: false, content: true }] }).key,
      // An id with a space or a comma is not two ids.
      arrangeToolbar({ groups: [{ id: 'a +b', content: true }] }).key,
      arrangeToolbar({ groups: [{ id: 'a,b', content: true }] }).key,
      arrangeToolbar({
        groups: [
          { id: 'a', content: true },
          { id: 'b', content: true },
        ],
      }).key,
    ];
    expect(new Set(keys).size).toBe(keys.length);
  });

  test('a host heading is its label, else its translated key, else its id', () => {
    const t = (key: string) => `<${key}>`;
    expect(toolbarGroupHeading({ id: 'a', label: 'A', labelKey: 'k' }, t)).toBe('A');
    expect(toolbarGroupHeading({ id: 'a', labelKey: 'k' }, t)).toBe('<k>');
    expect(toolbarGroupHeading({ id: 'a' }, t)).toBe('a');
  });

  test('a built-in heading is the translated label of the group, also with a label', () => {
    const t = (key: string) => `<${key}>`;
    expect(toolbarGroupHeading({ id: 'font', label: 'Text' }, t)).toBe(
      '<formattingBar.groups.font>'
    );
  });
});

describe('the bar with a fixed block', () => {
  const win = new GlobalWindow({ width: 1000, height: 800 });
  const document = win.document as unknown as Document;

  /**
   * A bar of `width` CSS pixels with groups of 100, an optional fixed block, and `gap` pixels
   * between two items, shown at `scale` window pixels for each CSS pixel.
   */
  function bar(
    width: number,
    groups: readonly string[],
    fixed?: number,
    gap = 0,
    scale = 1
  ): HTMLElement {
    const el = document.createElement('div');
    if (gap > 0) el.style.columnGap = `${gap}px`;
    Object.defineProperty(el, 'clientWidth', { get: () => width });
    Object.defineProperty(el, 'offsetWidth', { get: () => width });
    el.getBoundingClientRect = () => ({ width: width * scale, height: 0 }) as DOMRect;
    for (const id of groups) {
      const g = document.createElement('div');
      g.dataset.group = id;
      g.getBoundingClientRect = () => ({ width: 100 * scale }) as DOMRect;
      el.append(g);
    }
    if (fixed !== undefined) {
      const block = document.createElement('div');
      block.setAttribute(TOOLBAR_FIXED_ATTRIBUTE, '');
      block.getBoundingClientRect = () => ({ width: fixed * scale }) as DOMRect;
      el.append(block);
    }
    document.body.append(el);
    return el;
  }
  const groups = ['a', 'b', 'c', 'd'].map((id) => ({ id }));

  test('the groups fit in the room that the fixed block leaves', () => {
    expect(measureToolbar(bar(400, ['a', 'b', 'c', 'd']), groups, new Map())).toBe(4);
    // 400 less 150 leaves 250: the More button (40) and two groups.
    expect(measureToolbar(bar(400, ['a', 'b', 'c', 'd'], 150), groups, new Map())).toBe(2);
  });

  test('in a host that scales the editor, the fixed block is measured in CSS pixels', () => {
    // At scale 2, the block shows 300 window pixels wide but takes 150 of the 400 CSS pixels.
    expect(measureToolbar(bar(400, ['a', 'b', 'c', 'd'], 150, 0, 2), groups, new Map())).toBe(2);
  });

  test('the fixed block is charged its gap, as each group is', () => {
    // 400 less 140 and one gap of 10 leaves 250. The More button (40), a gap, and one group
    // take 150; a second group with its gap would take 260.
    expect(measureToolbar(bar(400, ['a', 'b', 'c', 'd'], 140, 10), groups, new Map())).toBe(1);
    // With no gap, the same bar holds two groups.
    expect(measureToolbar(bar(400, ['a', 'b', 'c', 'd'], 140), groups, new Map())).toBe(2);
  });

  test('a fixed block that fills the bar leaves no group, and an unmeasured bar shows all', () => {
    expect(measureToolbar(bar(400, ['a', 'b', 'c', 'd'], 500), groups, new Map())).toBe(0);
    expect(measureToolbar(bar(0, ['a', 'b', 'c', 'd'], 150), groups, new Map())).toBe(4);
  });

  test('only the groups and the fixed block of the bar count, not host content with their marks', () => {
    const el = bar(400, ['a', 'b', 'c', 'd']);
    const own = document.createElement('span');
    own.dataset.group = 'a';
    own.getBoundingClientRect = () => ({ width: 10 }) as DOMRect;
    el.querySelector('[data-group="d"]')!.append(own);
    const block = document.createElement('div');
    block.setAttribute(TOOLBAR_FIXED_ATTRIBUTE, '');
    block.getBoundingClientRect = () => ({ width: 1000 }) as DOMRect;
    el.querySelector('[data-group="c"]')!.append(block);
    const widths = new Map<string, number>();
    expect(measureToolbar(el, groups, widths)).toBe(4);
    expect(widths.get('a')).toBe(100);
  });

  test('the bar watches its fixed block', () => {
    const watched: Element[] = [];
    const view = win as unknown as { ResizeObserver: unknown };
    const original = view.ResizeObserver;
    view.ResizeObserver = class {
      observe(el: Element) {
        watched.push(el);
      }
      disconnect() {}
    };
    try {
      const el = bar(400, ['a'], 50);
      observeToolbar(el, () => undefined)();
      expect(watched.map((w) => w.hasAttribute(TOOLBAR_FIXED_ATTRIBUTE))).toEqual([
        false,
        false,
        true,
      ]);
    } finally {
      view.ResizeObserver = original;
    }
  });
});
