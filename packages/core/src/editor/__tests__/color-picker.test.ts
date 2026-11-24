// The headless color picker (`color-picker.ts`): the grid, the names, the typed hex, and the
// keyboard model of an opened grid.

import { describe, expect, test } from 'bun:test';
import { Window } from 'happy-dom';
import {
  COLOR_PICKER_STANDARD_COLORS,
  colorPickHex,
  colorPickLabel,
  colorPickerGrid,
  customHexDraft,
  focusColorSwatch,
  hexFromText,
  moveColorSwatchFocus,
  workbookThemeHexes,
} from '../color-picker.ts';

const THEME = [
  'FFFFFF',
  '000000',
  'E7E6E6',
  '44546A',
  '4472C4',
  'ED7D31',
  'A5A5A5',
  'FFC000',
  '5B9BD5',
  '70AD47',
];

const t = (key: string, params?: Readonly<Record<string, string | number>>) =>
  params ? `${key}(${Object.values(params).join(',')})` : key;

describe('the color grid', () => {
  test('has the theme row, five tint and shade rows, and the standard row, ten wide', () => {
    const grid = colorPickerGrid(THEME);
    expect(grid.theme.map((row) => row.length)).toEqual([10, 10, 10, 10, 10, 10]);
    expect(grid.theme[0]!.map((s) => s.hex)).toEqual(THEME.map((hex) => `#${hex}`));
    expect(grid.standard.map((s) => s.hex)).toEqual(COLOR_PICKER_STANDARD_COLORS.map((c) => c.hex));
    // Every swatch paints six hex digits, the only text that reaches a style.
    for (const swatch of [...grid.theme.flat(), ...grid.standard])
      expect(swatch.hex).toMatch(/^#[0-9A-F]{6}$/);
  });

  test('a white column darkens, a black column lightens, and an accent goes both ways', () => {
    const grid = colorPickerGrid(THEME);
    const column = (i: number) => grid.theme.slice(1).map((row) => row[i]!.pick);
    expect(column(0)).toEqual(['bg1:-5', 'bg1:-15', 'bg1:-25', 'bg1:-35', 'bg1:-50']);
    expect(column(1)).toEqual(['tx1:50', 'tx1:35', 'tx1:25', 'tx1:15', 'tx1:5']);
    expect(column(4)).toEqual([
      'accent1:80',
      'accent1:60',
      'accent1:40',
      'accent1:-25',
      'accent1:-50',
    ]);
    expect(colorPickHex('bg1:-50', THEME)).toBe('#808080');
    expect(colorPickHex('tx1:50', THEME)).toBe('#808080');
  });

  test('a workbook with no theme, or a theme with a missing or bad entry, shows the default', () => {
    expect(workbookThemeHexes([])).toEqual(THEME);
    expect(workbookThemeHexes(undefined)).toEqual(THEME);
    const own = ['112233', ...THEME.slice(1)];
    expect(workbookThemeHexes(own)).toEqual(own);
    expect(workbookThemeHexes([undefined, ...THEME.slice(1)])).toEqual(THEME);
    expect(workbookThemeHexes(['url(x)', ...THEME.slice(1)])).toEqual(THEME);
  });
});

describe('the names and the typed hex', () => {
  test('a swatch is named by its theme slot and shift, or its standard name', () => {
    expect(colorPickLabel('accent1', t)).toBe('colorPicker.theme.accent1');
    expect(colorPickLabel('accent1:40', t)).toBe(
      'colorPicker.lighterPercent(colorPicker.theme.accent1,40)'
    );
    expect(colorPickLabel('tx1:-25', t)).toBe(
      'colorPicker.darkerPercent(colorPicker.theme.text1,25)'
    );
    expect(colorPickLabel('#ff0000', t)).toBe('colorPicker.colors.red');
    expect(colorPickLabel('#123456', t)).toBe('#123456');
  });

  test('the custom field keeps six hex digits, and Apply takes three or six', () => {
    expect(customHexDraft('#1a-2B3c9')).toBe('1a2B3c');
    expect(hexFromText('1a2b3c')).toBe('1A2B3C');
    expect(hexFromText('#abc')).toBe('AABBCC');
    expect(hexFromText('12345')).toBeNull();
    expect(hexFromText('red')).toBeNull();
  });
});

describe('the keyboard model', () => {
  const grid = (dir?: string) => {
    const doc = new Window().document;
    const root = doc.createElement('div');
    if (dir) root.setAttribute('dir', dir);
    const none = doc.createElement('button');
    none.setAttribute('data-color-none', '');
    root.append(none);
    for (let i = 0; i < 20; i++) {
      const swatch = doc.createElement('button');
      swatch.setAttribute('data-color-swatch', '');
      swatch.setAttribute('aria-pressed', String(i === 12));
      swatch.textContent = String(i);
      root.append(swatch);
    }
    doc.body.append(root);
    return { root: root as unknown as HTMLElement, doc, none };
  };
  const focused = (doc: { activeElement: unknown }) =>
    (doc.activeElement as { textContent: string } | null)?.textContent;

  test('the open focuses the current swatch, and only it is in the Tab order', () => {
    const { root, doc } = grid();
    focusColorSwatch(root);
    expect(focused(doc)).toBe('12');
    const stops = [...root.querySelectorAll<HTMLElement>('[data-color-swatch]')].filter(
      (s) => s.tabIndex === 0
    );
    expect(stops.map((s) => s.textContent)).toEqual(['12']);
  });

  test('the arrows move by one and by a row, and Up from the first row reaches the clear button', () => {
    const { root, doc, none } = grid();
    focusColorSwatch(root);
    expect(moveColorSwatchFocus(root, 'ArrowRight')).toBe(true);
    expect(focused(doc)).toBe('13');
    moveColorSwatchFocus(root, 'ArrowUp');
    expect(focused(doc)).toBe('3');
    moveColorSwatchFocus(root, 'ArrowUp');
    expect(doc.activeElement).toBe(none);
    moveColorSwatchFocus(root, 'ArrowDown');
    expect(focused(doc)).toBe('0');
    moveColorSwatchFocus(root, 'End');
    expect(focused(doc)).toBe('19');
    expect(moveColorSwatchFocus(root, 'Tab')).toBe(false);
  });

  test('in a right-to-left grid, ArrowLeft moves to the next swatch', () => {
    const { root, doc } = grid('rtl');
    focusColorSwatch(root);
    moveColorSwatchFocus(root, 'ArrowLeft');
    expect(focused(doc)).toBe('13');
  });
});
