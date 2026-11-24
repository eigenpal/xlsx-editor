// The arrangement of the preset toolbar with the host's parts, shared by both adapters. The
// placement case is a copy of docx-editor/packages/core/src/editor/__tests__/toolbar-room.test.ts
// at 10eeaf67 (EigenPal). The measure with a fixed block uses a DOM of this file's own: the core
// run has no global DOM (`bun run test`).

import { describe, expect, test } from 'bun:test';
import { GlobalWindow } from 'happy-dom';
import {
  TOOLBAR_COLLAPSE_ORDER,
  arrangeToolbar,
  arrangeToolbarGroups,
  collapseOrder,
  toolbarGroupHeading,
  toolbarGroupPriority,
  type ToolbarArrangement,
} from '../toolbar-arrange.ts';
import {
  TOOLBAR_FIXED_ATTRIBUTE,
  groupsThatFit,
  measureToolbar,
  observeToolbar,
  toolbarCollapse,
} from '../toolbar-layout.ts';
import { chromeSlotId, defaultChromeGroups } from '../chrome-controls.ts';
import { TOOLBAR_NAMED_PARTS, isToolbarButtonSlot } from '../toolbar-named-parts.ts';

// The preset bar: the menu bar's contextual groups and controls are not in it.
const BUILT_IN = defaultChromeGroups().map((g) => g.id);
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

  test('a slot that only the menu bar shows is named as such, not as an unknown id', () => {
    const a = arrangeToolbar({
      slots: [
        { slotId: 'edit.copy', hidden: false, content: true },
        { slotId: 'view.zoomIn', hidden: false, content: false },
      ],
    });
    expect(a.outside).toEqual(['edit.copy']);
    const advice =
      'Place it in a host XlsxEditor.Toolbar.Group, where it renders its control (Paste reads the ' +
      'clipboard there, and Open stays disabled), or use a slot of the preset toolbar, such as ' +
      '"font.bold".';
    expect(a.problems).toEqual([
      '"edit.copy" is a slot that only the menu bar shows, not a toolbar slot, so the content of ' +
        `its XlsxEditor.Toolbar.Slot is appended after the groups. ${advice}`,
      '"view.zoomIn" is a slot that only the menu bar shows, not a toolbar slot, so its ' +
        `XlsxEditor.Toolbar.Slot changes nothing. ${advice}`,
    ]);
    const vue = arrangeToolbar({
      slots: [{ slotId: 'edit.copy', hidden: false, content: false }],
      names: { slot: 'XlsxEditorToolbar.Slot', group: 'XlsxEditorToolbar.Group' },
    });
    expect(vue.problems[0]).toContain('Place it in a host XlsxEditorToolbar.Group, where');
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
    expect(review).toEqual({
      id: 'review',
      chrome: null,
      label: 'Review',
      labelKey: undefined,
      priority: undefined,
      pinned: false,
    });
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
    expect(measureToolbar(bar(400, ['a', 'b', 'c', 'd']), groups, new Map())).toEqual([]);
    // 400 less 150 leaves 250: the More button (40) and two groups.
    expect(measureToolbar(bar(400, ['a', 'b', 'c', 'd'], 150), groups, new Map())).toEqual([
      'c',
      'd',
    ]);
  });

  test('in a host that scales the editor, the fixed block is measured in CSS pixels', () => {
    // At scale 2, the block shows 300 window pixels wide but takes 150 of the 400 CSS pixels.
    expect(measureToolbar(bar(400, ['a', 'b', 'c', 'd'], 150, 0, 2), groups, new Map())).toEqual([
      'c',
      'd',
    ]);
  });

  test('the fixed block is charged its gap, as each group is', () => {
    // 400 less 140 and one gap of 10 leaves 250. The More button (40), a gap, and one group
    // take 150; a second group with its gap would take 260.
    expect(measureToolbar(bar(400, ['a', 'b', 'c', 'd'], 140, 10), groups, new Map())).toEqual([
      'b',
      'c',
      'd',
    ]);
    // With no gap, the same bar holds two groups.
    expect(measureToolbar(bar(400, ['a', 'b', 'c', 'd'], 140), groups, new Map())).toEqual([
      'c',
      'd',
    ]);
  });

  test('a fixed block that fills the bar leaves no group, and an unmeasured bar shows all', () => {
    expect(measureToolbar(bar(400, ['a', 'b', 'c', 'd'], 500), groups, new Map())).toEqual([
      'a',
      'b',
      'c',
      'd',
    ]);
    expect(measureToolbar(bar(0, ['a', 'b', 'c', 'd'], 150), groups, new Map())).toEqual([]);
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
    expect(measureToolbar(el, groups, widths)).toEqual([]);
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

describe('the collapse order', () => {
  test('the built-in groups collapse from the end of the bar, and every one is declared', () => {
    expect([...TOOLBAR_COLLAPSE_ORDER].sort()).toEqual([...BUILT_IN].sort());
    expect(collapseOrder(BUILT_IN)).toEqual([...BUILT_IN].reverse());
    expect(toolbarGroupPriority('file')).toBe(10);
    expect(toolbarGroupPriority('history')).toBe(90);
    expect(toolbarGroupPriority('share')).toBeUndefined();
  });

  test('a group with no priority leaves first, in reverse bar order', () => {
    expect(collapseOrder(['history', 'a', 'font', 'b'])).toEqual(['b', 'a', 'font', 'history']);
  });

  test('a priority sorts a group among the built-in groups, and overrides a built-in one', () => {
    const priorities = new Map([
      ['share', 85],
      ['file', 95],
    ]);
    expect(collapseOrder(['history', 'font', 'file', 'share'], TOOLBAR_COLLAPSE_ORDER, priorities))
      // font 80, share 85, history 90, file 95.
      .toEqual(['font', 'share', 'history', 'file']);
  });

  test('two groups with the same priority: the one further along the bar leaves first', () => {
    const priorities = new Map([
      ['a', 50],
      ['b', 50],
    ]);
    expect(collapseOrder(['a', 'b'], TOOLBAR_COLLAPSE_ORDER, priorities)).toEqual(['b', 'a']);
  });

  test('NaN counts as no priority', () => {
    expect(collapseOrder(['history', 'a'], TOOLBAR_COLLAPSE_ORDER, new Map([['a', NaN]]))).toEqual([
      'a',
      'history',
    ]);
  });
});

describe('toolbarCollapse', () => {
  const all = (ids: readonly string[], width = 100) => new Map(ids.map((id) => [id, width]));

  test('with the default order, the groups that stay are the start that groupsThatFit counts', () => {
    const groups = BUILT_IN.map((id) => ({ id }));
    for (const width of [0, 150, 300, 499, 700, 900, 1000]) {
      const hidden = toolbarCollapse(groups, all(BUILT_IN), width, 4);
      const count = groupsThatFit(
        BUILT_IN.map(() => 100),
        width,
        4
      );
      expect(hidden).toEqual(BUILT_IN.slice(count));
    }
  });

  test('a host group with no priority leaves before every built-in group', () => {
    // The host group stands after history, at the start, and still leaves first.
    const groups = [{ id: 'history' }, { id: 'share' }, { id: 'font' }, { id: 'file' }];
    // 400 holds three groups beside the More button (40), not four.
    expect(toolbarCollapse(groups, all(groups.map((g) => g.id)), 399)).toEqual(['share']);
  });

  test('a priority keeps a host group longer than built-in groups with a lower one', () => {
    const groups = [
      { id: 'history' },
      { id: 'font' },
      { id: 'share', priority: 95 },
      { id: 'file' },
    ];
    // Two groups and the More button fit in 250: file (10) and font (80) leave, share stays.
    expect(toolbarCollapse(groups, all(groups.map((g) => g.id)), 250)).toEqual(['font', 'file']);
  });

  test('a pinned group never leaves, also when nothing else fits', () => {
    const groups = [{ id: 'history' }, { id: 'share', pinned: true }, { id: 'file' }];
    expect(toolbarCollapse(groups, all(groups.map((g) => g.id)), 150)).toEqual(['history', 'file']);
    expect(toolbarCollapse(groups, all(groups.map((g) => g.id)), 1)).toEqual(['history', 'file']);
  });

  test('a pinned built-in group stays while the groups after it leave', () => {
    const groups = [{ id: 'history' }, { id: 'font' }, { id: 'file', pinned: true }];
    expect(toolbarCollapse(groups, all(groups.map((g) => g.id)), 250)).toEqual(['font']);
  });
});

describe('arrangeToolbar with priorities, pinned groups, and overflow content', () => {
  test('a group carries the priority and the pinned flag of its request', () => {
    const a = arrangeToolbar({
      groups: [
        { id: 'share', content: true, priority: 15, pinned: true },
        { id: 'file', content: false, priority: 95 },
        { id: 'font', content: false, priority: NaN },
      ],
    });
    const share = a.groups.find((g) => g.id === 'share');
    expect(share?.priority).toBe(15);
    expect(share?.pinned).toBe(true);
    expect(a.groups.find((g) => g.id === 'file')?.priority).toBe(95);
    expect(a.groups.find((g) => g.id === 'font')?.priority).toBeUndefined();
    expect(a.groups.find((g) => g.id === 'history')?.pinned).toBe(false);
    expect(a.problems).toEqual([]);
  });

  test('a change of priority or pinned changes the key, so the bar measures again', () => {
    const base = arrangeToolbar({ groups: [{ id: 'share', content: true }] });
    const ranked = arrangeToolbar({ groups: [{ id: 'share', content: true, priority: 15 }] });
    const pinned = arrangeToolbar({ groups: [{ id: 'share', content: true, pinned: true }] });
    expect(new Set([base.key, ranked.key, pinned.key]).size).toBe(3);
  });

  test('overflow content keeps a built-in group with no control left', () => {
    const a = arrangeToolbar({
      hidden: ['file.save'],
      groups: [{ id: 'file', content: false, overflowContent: true }],
    });
    expect(ids(a)).toContain('file');
    expect(controls(a, 'file')).toEqual([]);
  });

  test('a host group with overflow content and no children is left out, with a problem', () => {
    const a = arrangeToolbar({
      groups: [{ id: 'share', content: false, overflowContent: true }],
    });
    expect(ids(a)).not.toContain('share');
    expect(a.problems.join('\n')).toContain('"share" has overflowContent but no children');
  });
});

describe('arrangeToolbar with changed controls', () => {
  test('icon, asChild, and className change a button slot where it stands', () => {
    const a = arrangeToolbar({
      slots: [
        { slotId: 'font.bold', hidden: false, content: false, icon: true },
        { slotId: 'cells.insertRows', hidden: false, content: false, asChild: true },
        { slotId: 'font.size', hidden: false, content: false, className: true },
      ],
    });
    expect([...a.customized].sort()).toEqual(['cells.insertRows', 'font.bold', 'font.size']);
    expect(a.replaced.size).toBe(0);
    expect(a.problems).toEqual([]);
  });

  test('icon and asChild on a value control are ignored, with a problem', () => {
    const a = arrangeToolbar({
      slots: [
        { slotId: 'font.size', hidden: false, content: false, icon: true },
        { slotId: 'fill.color', hidden: false, content: false, asChild: true },
      ],
    });
    expect(a.customized.size).toBe(0);
    expect(a.problems).toHaveLength(2);
    expect(a.problems[0]).toContain('"font.size" has icon, but the slot is not a button');
    expect(a.problems[1]).toContain('"fill.color" has asChild, but the slot is not a button');
  });

  test('icon on a color control replaces its glyph, as on a button', () => {
    const a = arrangeToolbar({
      slots: [{ slotId: 'font.color', hidden: false, content: false, icon: true }],
    });
    expect([...a.customized]).toEqual(['font.color']);
    expect(a.problems).toEqual([]);
  });

  test('asChild with children on a value slot replaces the control, with a problem', () => {
    const a = arrangeToolbar({
      slots: [{ slotId: 'font.size', hidden: false, content: true, asChild: true }],
    });
    expect([...a.replaced]).toEqual(['font.size']);
    expect(a.problems).toHaveLength(1);
    expect(a.problems[0]).toContain('"font.size" has asChild, but the slot is not a button');
  });

  test('content wins over icon and className, with a problem', () => {
    const a = arrangeToolbar({
      slots: [{ slotId: 'font.bold', hidden: false, content: true, icon: true, className: true }],
    });
    expect([...a.replaced]).toEqual(['font.bold']);
    expect(a.customized.size).toBe(0);
    expect(a.problems).toHaveLength(2);
  });

  test('a hidden or listed slot is not customized', () => {
    const a = arrangeToolbar({
      hidden: ['font.italic'],
      slots: [
        { slotId: 'font.bold', hidden: true, content: false, icon: true },
        { slotId: 'font.italic', hidden: false, content: false, icon: true },
      ],
    });
    expect(a.customized.size).toBe(0);
  });

  test('customized slots change the key', () => {
    const a = arrangeToolbar({});
    const b = arrangeToolbar({
      slots: [{ slotId: 'font.bold', hidden: false, content: false, icon: true }],
    });
    expect(a.key).not.toBe(b.key);
  });
});

describe('TOOLBAR_NAMED_PARTS', () => {
  test('names each toolbar slot exactly once, with PascalCase names', () => {
    const slots: string[] = Object.values(TOOLBAR_NAMED_PARTS);
    // The slots of the preset toolbar: the slots that only the menu bar shows have menu row names.
    const preset = defaultChromeGroups().flatMap((g) => g.controls.map((c) => chromeSlotId(g, c)));
    expect([...slots].sort()).toEqual([...preset].sort());
    expect(new Set(slots).size).toBe(slots.length);
    for (const name of Object.keys(TOOLBAR_NAMED_PARTS)) expect(name).toMatch(/^[A-Z][A-Za-z]+$/);
  });

  test('a button part is a slot of the icon shape', () => {
    expect(isToolbarButtonSlot('font.bold')).toBe(true);
    expect(isToolbarButtonSlot('cells.deleteRows')).toBe(true);
    expect(isToolbarButtonSlot('font.family')).toBe(false);
    expect(isToolbarButtonSlot('fill.color')).toBe(false);
    expect(isToolbarButtonSlot('number.custom')).toBe(false);
    expect(isToolbarButtonSlot('nope')).toBe(false);
  });
});
