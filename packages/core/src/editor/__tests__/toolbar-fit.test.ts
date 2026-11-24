// The fit of a short bar (`toolbarFit`, `measureToolbarFit`): the wide fields of the registry
// (`ChromeControl.wide`) leave for the "More" panel before their groups do, so a bar beside a
// chat keeps the font controls. The measure uses a DOM of this file's own: the core run has no
// global DOM (`bun run test`).

import { describe, expect, test } from 'bun:test';
import { GlobalWindow } from 'happy-dom';
import { defaultChromeGroups } from '../chrome-controls.ts';
import {
  TOOLBAR_FIT_ALL,
  measureToolbarFit,
  toolbarCollapse,
  toolbarFit,
  trimToolbarGroup,
  wideSlots,
} from '../toolbar-layout.ts';

/** The built-in groups, as the bar gets them, with the widths that the chat panel measured. */
const GROUPS = defaultChromeGroups()
  .filter((g) => g.id !== 'file')
  .map((g) => ({ id: g.id, chrome: g }));
const WIDTHS = new Map(
  Object.entries({
    history: 62,
    font: 381,
    fill: 57,
    borders: 107,
    alignment: 235,
    number: 229,
    cells: 55,
    clear: 55,
  })
);
/** The font family field (150) and the number format field (130), each with a gap of 2. */
const SAVINGS = new Map([
  ['font', 152],
  ['number', 132],
]);

describe('toolbarFit', () => {
  test('the registry marks the font family and number format fields as wide', () => {
    expect(GROUPS.flatMap((g) => wideSlots(g))).toEqual(['font.family', 'number.format']);
  });

  test('a bar that holds every group moves nothing', () => {
    expect(toolbarFit(GROUPS, WIDTHS, SAVINGS, 2000, 4)).toBe(TOOLBAR_FIT_ALL);
  });

  test('a half panel keeps the font controls, and moves the family field before the group', () => {
    // 476 px: before, only history fit, and the font group left whole.
    expect(toolbarCollapse(GROUPS, WIDTHS, 476, 4)).toEqual([
      'font',
      'fill',
      'borders',
      'alignment',
      'number',
      'cells',
      'clear',
    ]);
    expect(toolbarFit(GROUPS, WIDTHS, SAVINGS, 476, 4)).toEqual({
      collapsed: ['borders', 'alignment', 'number', 'cells', 'clear'],
      trimmed: ['font'],
    });
  });

  test('a small overflow moves the wide field of the group that leaves first', () => {
    // Every group with its gaps is 1209 px. 1150 px is 59 px short: the number format field
    // (132) leaves, and the font family field stays.
    expect(toolbarFit(GROUPS, WIDTHS, SAVINGS, 1150, 4)).toEqual({
      collapsed: [],
      trimmed: ['number'],
    });
  });

  test('a group that stays gets its wide field back when the room allows', () => {
    // 700 px: the font family field fits again once the later groups have left.
    const fit = toolbarFit(GROUPS, WIDTHS, SAVINGS, 700, 4);
    expect(fit.trimmed).toEqual([]);
    expect(fit.collapsed).toEqual(['alignment', 'number', 'cells', 'clear']);
  });

  test('with no saving, the fit is the collapse of whole groups', () => {
    for (const width of [300, 476, 700, 1000])
      expect(toolbarFit(GROUPS, WIDTHS, new Map(), width, 4).collapsed).toEqual(
        toolbarCollapse(GROUPS, WIDTHS, width, 4)
      );
  });

  test('a trimmed group splits into its bar part and its panel part', () => {
    const font = GROUPS.find((g) => g.id === 'font')!;
    const ids = (g: typeof font) => g.chrome.controls.map((c) => c.id);
    expect(ids(trimToolbarGroup(font, 'panel'))).toEqual(['family']);
    expect(ids(trimToolbarGroup(font, 'bar'))).toEqual([
      'size',
      'bold',
      'italic',
      'underline',
      'color',
    ]);
  });
});

describe('measureToolbarFit', () => {
  const win = new GlobalWindow({ width: 1000, height: 800 });
  const document = win.document as unknown as Document;

  /** A bar of `width` px with the font and fill groups; the font group holds a 150 px field. */
  function bar(width: number, trimmed: boolean): HTMLElement {
    const el = document.createElement('div');
    Object.defineProperty(el, 'clientWidth', { get: () => width });
    const font = document.createElement('div');
    font.dataset.group = 'font';
    font.getBoundingClientRect = () => ({ width: trimmed ? 230 : 380 }) as DOMRect;
    if (!trimmed) {
      const field = document.createElement('button');
      field.dataset.slot = 'font.family';
      field.getBoundingClientRect = () => ({ width: 150 }) as DOMRect;
      font.append(field);
    }
    const fill = document.createElement('div');
    fill.dataset.group = 'fill';
    fill.getBoundingClientRect = () => ({ width: 60 }) as DOMRect;
    el.append(font, fill);
    document.body.append(el);
    return el;
  }
  const groups = GROUPS.filter((g) => g.id === 'font' || g.id === 'fill');

  test('measures the saving of a wide field, and keeps the full width of a trimmed group', () => {
    const widths = new Map<string, number>();
    // 380 + 60 does not fit in 400; without the field, 230 + 60 and the More button do.
    expect(measureToolbarFit(bar(400, false), groups, widths)).toEqual({
      collapsed: [],
      trimmed: ['font'],
    });
    expect(widths.get('font')).toBe(380);
    // The next render shows the trimmed group: the full width stays 380, so the fit holds.
    expect(measureToolbarFit(bar(400, true), groups, widths, ['font'])).toEqual({
      collapsed: [],
      trimmed: ['font'],
    });
    expect(widths.get('font')).toBe(380);
    // A wider bar gives the field back.
    expect(measureToolbarFit(bar(600, true), groups, widths, ['font'])).toBe(TOOLBAR_FIT_ALL);
  });

  test('a trimmed group with no measured saving shows whole first', () => {
    expect(measureToolbarFit(bar(400, true), groups, new Map(), ['font'])).toBe(TOOLBAR_FIT_ALL);
  });
});
