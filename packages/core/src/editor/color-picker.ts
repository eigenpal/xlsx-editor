// The color picker, headless: the grid of theme colors with their tints and shades, the standard
// colors, the name and the paint of each pick, the custom hex text, and the keyboard model of an
// opened grid. The font color and fill color controls of React and Vue render from it, so both
// adapters show the same grid, name each swatch the same way, and move the focus with the same
// keys.
//
// Taken from docx-editor/packages/core/src/editor/color-picker.ts at 592db9fd (EigenPal).
// Intended differences:
// - The theme colors come from the workbook theme (`workbookThemeHexes`), in theme index order:
//   `lt1`, `dk1`, `lt2`, `dk2`, then the accents (`cascade/theme.ts`). That is the order of the
//   first row of the grid, so a slot needs no name map.
// - A pick writes six hex digits: the font color and the fill color of a cell format are
//   `CT_Color` with `@rgb` (ECMA-376 Part 1, 18.8.19), as DOCX Editor's font color writes
//   `RRGGBB`. The automatic font color writes `@auto` (`AUTO_COLOR`).
// - No `ColorPickerTranslate` type: the adapters pass their own translator.
// - The keyboard model reads the `data-color-swatch` and `data-color-none` attributes only, and
//   takes the reading direction from `inlineStep`, as the rest of the chrome does.

import { inlineStep } from './chrome-direction.ts';

/** A theme color slot of the picker, in the order of the grid's first row. @public */
export type ColorPickerThemeSlot =
  | 'bg1'
  | 'tx1'
  | 'bg2'
  | 'tx2'
  | 'accent1'
  | 'accent2'
  | 'accent3'
  | 'accent4'
  | 'accent5'
  | 'accent6';

/**
 * A color picker value: a theme color, a theme tint or shade (`accent1:40` is Lighter 40%,
 * `accent1:-25` is Darker 25%), or `#RRGGBB`. @public
 */
export type ColorPick = ColorPickerThemeSlot | `${ColorPickerThemeSlot}:${number}` | `#${string}`;

/** The theme colors of the grid, in the order of its first row. @public */
export const COLOR_PICKER_THEME_SLOTS: readonly ColorPickerThemeSlot[] = [
  'bg1',
  'tx1',
  'bg2',
  'tx2',
  'accent1',
  'accent2',
  'accent3',
  'accent4',
  'accent5',
  'accent6',
];

/** The i18n key of each theme color's name. */
const THEME_NAME_KEYS: Readonly<Record<ColorPickerThemeSlot, string>> = {
  bg1: 'colorPicker.theme.background1',
  tx1: 'colorPicker.theme.text1',
  bg2: 'colorPicker.theme.background2',
  tx2: 'colorPicker.theme.text2',
  accent1: 'colorPicker.theme.accent1',
  accent2: 'colorPicker.theme.accent2',
  accent3: 'colorPicker.theme.accent3',
  accent4: 'colorPicker.theme.accent4',
  accent5: 'colorPicker.theme.accent5',
  accent6: 'colorPicker.theme.accent6',
};

/** The standard colors of the grid, with the i18n key of each name. @public */
export const COLOR_PICKER_STANDARD_COLORS: readonly {
  readonly hex: string;
  readonly nameKey: string;
}[] = [
  { hex: '#C00000', nameKey: 'colorPicker.colors.darkRed' },
  { hex: '#FF0000', nameKey: 'colorPicker.colors.red' },
  { hex: '#FFC000', nameKey: 'colorPicker.colors.orange' },
  { hex: '#FFFF00', nameKey: 'colorPicker.colors.yellow' },
  { hex: '#92D050', nameKey: 'colorPicker.colors.lightGreen' },
  { hex: '#00B050', nameKey: 'colorPicker.colors.green' },
  { hex: '#00B0F0', nameKey: 'colorPicker.colors.lightBlue' },
  { hex: '#0070C0', nameKey: 'colorPicker.colors.blue' },
  { hex: '#002060', nameKey: 'colorPicker.colors.darkBlue' },
  { hex: '#7030A0', nameKey: 'colorPicker.colors.purple' },
];

/**
 * The value of the font color command that writes the automatic color: `x:color` with
 * `@auto` true (`CT_Color`, ECMA-376 Part 1, 18.8.19), which the grid paints black. @public
 */
export const AUTO_COLOR = 'auto';

/**
 * The chip colors of the button above the grid, `RRGGBB`: the automatic font color paints black
 * (`cascade/color.ts`, measured on fixture colors-fonts), and a cell with no fill shows the
 * white grid, which the chrome theme never changes. @public
 */
export const CLEAR_CHIP_HEX = { auto: '000000', none: 'FFFFFF' } as const;

/** Six hex digits, the only color text that reaches a swatch style. */
const HEX6 = /^[0-9A-F]{6}$/;

/** The default theme colors, for a workbook with no theme, in slot order. */
const DEFAULT_THEME_HEXES: readonly string[] = [
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

/**
 * The theme colors of a workbook, `RRGGBB` in `COLOR_PICKER_THEME_SLOTS` order, from
 * `editor.getWorkbookThemeColors()`; the default theme when the workbook has none or lacks an
 * entry. @public
 */
export function workbookThemeHexes(
  colors: readonly (string | undefined)[] | undefined
): readonly string[] {
  const hexes = (colors ?? []).map((hex) => hex?.replace(/^#/, '').toUpperCase() ?? '');
  return hexes.length >= DEFAULT_THEME_HEXES.length &&
    hexes.slice(0, DEFAULT_THEME_HEXES.length).every((hex) => HEX6.test(hex))
    ? hexes.slice(0, DEFAULT_THEME_HEXES.length)
    : DEFAULT_THEME_HEXES;
}

/**
 * The name of a pick: a theme color with its tint or shade, a standard color by name ("Purple"),
 * or the hex code of any other color. @public
 */
export function colorPickLabel(
  pick: ColorPick,
  t: (key: string, params?: Readonly<Record<string, string | number>>) => string
): string {
  if (pick.startsWith('#')) {
    const standard = COLOR_PICKER_STANDARD_COLORS.find((c) => c.hex === pick.toUpperCase());
    return standard ? t(standard.nameKey) : pick.toUpperCase();
  }
  const [theme, shift] = pick.split(':') as [ColorPickerThemeSlot, string?];
  const name = t(THEME_NAME_KEYS[theme]);
  const n = Number(shift ?? 0);
  if (n > 0) return t('colorPicker.lighterPercent', { color: name, percent: n });
  if (n < 0) return t('colorPicker.darkerPercent', { color: name, percent: -n });
  return name;
}

/**
 * `RRGGBB` of `hex` with the luminance transform of a theme tint or shade: `lumMod` of
 * `100 - |shift|` percent, and `lumOff` of `shift` percent for a tint (ECMA-376 Part 1,
 * 20.1.2.3.20 and 20.1.2.3.21), on the HSL lightness.
 */
function shifted(hex: string, shift: number): string {
  const [r, g, b] = [0, 2, 4].map((at) => parseInt(hex.slice(at, at + 2), 16) / 255) as [
    number,
    number,
    number,
  ];
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const l = (max + min) / 2;
  const d = max - min;
  const s = d === 0 ? 0 : d / (1 - Math.abs(2 * l - 1));
  let h = 0;
  if (d !== 0) {
    if (max === r) h = ((g - b) / d) % 6;
    else if (max === g) h = (b - r) / d + 2;
    else h = (r - g) / d + 4;
  }
  const lum = Math.max(0, Math.min(1, l * (1 - Math.abs(shift) / 100) + Math.max(0, shift) / 100));
  const c = (1 - Math.abs(2 * lum - 1)) * s;
  const x = c * (1 - Math.abs(((((h % 6) + 6) % 6) % 2) - 1));
  const m = lum - c / 2;
  const sector = Math.floor(((h % 6) + 6) % 6);
  const rgb = [
    [c, x, 0],
    [x, c, 0],
    [0, c, x],
    [0, x, c],
    [x, 0, c],
    [c, 0, x],
  ][sector]!;
  return rgb
    .map((v) =>
      Math.max(0, Math.min(255, Math.round((v + m) * 255)))
        .toString(16)
        .padStart(2, '0')
    )
    .join('')
    .toUpperCase();
}

/** The `#RRGGBB` that a pick paints, from the workbook's theme colors (`workbookThemeHexes`). @public */
export function colorPickHex(pick: ColorPick, themeHexes: readonly string[]): string {
  if (pick.startsWith('#')) {
    const hex = pick.slice(1).toUpperCase();
    return HEX6.test(hex) ? `#${hex}` : '#000000';
  }
  const [theme, shift] = pick.split(':') as [ColorPickerThemeSlot, string?];
  const base = themeHexes[COLOR_PICKER_THEME_SLOTS.indexOf(theme)] ?? '000000';
  const n = Math.max(-100, Math.min(100, Number(shift ?? 0)));
  return `#${shifted(base, Number.isFinite(n) ? n : 0)}`;
}

/** The tint and shade steps under one theme color, by the color's lightness. */
function steps(hex: string): readonly number[] {
  const channels = [0, 2, 4].map((at) => parseInt(hex.slice(at, at + 2), 16));
  const lightness = (Math.max(...channels) + Math.min(...channels)) / 2 / 255;
  if (lightness === 1) return [-5, -15, -25, -35, -50];
  if (lightness === 0) return [50, 35, 25, 15, 5];
  if (lightness >= 0.8) return [-10, -25, -50, -75, -90];
  return [80, 60, 40, -25, -50];
}

/** One swatch of the color grid. @public */
export interface ColorSwatch {
  readonly pick: ColorPick;
  /** The color the swatch paints, `#RRGGBB`. */
  readonly hex: string;
}

/** The color grid: theme rows (one column per theme color) and the standard row. @public */
export interface ColorPickerGrid {
  /** The theme colors, then five rows of their tints and shades. */
  readonly theme: readonly (readonly ColorSwatch[])[];
  readonly standard: readonly ColorSwatch[];
}

/**
 * The color grid: one row of theme colors, five rows of their tints and shades (one column per
 * theme color), and the row of standard colors. @public
 */
export function colorPickerGrid(themeHexes: readonly string[]): ColorPickerGrid {
  const base = COLOR_PICKER_THEME_SLOTS.map((slot) => ({
    pick: slot as ColorPick,
    hex: colorPickHex(slot, themeHexes),
  }));
  const ladders = COLOR_PICKER_THEME_SLOTS.map((_, i) => steps(themeHexes[i] ?? '000000'));
  const variants = [0, 1, 2, 3, 4].map((row) =>
    COLOR_PICKER_THEME_SLOTS.map((slot, column) => {
      const pick = `${slot}:${ladders[column]![row]!}` as ColorPick;
      return { pick, hex: colorPickHex(pick, themeHexes) };
    })
  );
  const standard = COLOR_PICKER_STANDARD_COLORS.map((c) => ({
    pick: c.hex as ColorPick,
    hex: c.hex,
  }));
  return { theme: [base, ...variants], standard };
}

/**
 * The custom color field's text with every character that is not a hex digit removed, at most
 * six digits: the draft that the field shows. @public
 */
export const customHexDraft = (text: string): string =>
  text.replace(/[^0-9A-Fa-f]/g, '').slice(0, 6);

/**
 * The command value of typed hex text (`#RRGGBB`, `RRGGBB`, or three digits): six upper-case hex
 * digits, or null when it names no color. @public
 */
export function hexFromText(text: string): string | null {
  let hex = text.trim().replace(/^#/, '').toUpperCase();
  if (/^[0-9A-F]{3}$/.test(hex)) hex = [...hex].map((c) => c + c).join('');
  return HEX6.test(hex) ? hex : null;
}

// --- The keyboard model of an opened grid ----------------------------------------------------

/** The columns of every swatch row of the grid. */
const SWATCH_COLUMNS = 10;

/** A grid swatch. */
const SWATCH_SELECTOR = '[data-color-swatch]';
/** The button above the swatches: Automatic or No fill. */
const NONE_SELECTOR = '[data-color-none]';

const swatchesOf = (grid: HTMLElement): HTMLElement[] => [
  ...grid.querySelectorAll<HTMLElement>(SWATCH_SELECTOR),
];

/** Keep one swatch in the Tab order, so Tab leaves the grid in one step. */
function rovingSwatch(grid: HTMLElement, active: HTMLElement | null): void {
  for (const swatch of swatchesOf(grid)) swatch.tabIndex = swatch === active ? 0 : -1;
}

/**
 * Focus the current swatch of an opened grid (`aria-pressed="true"`), else its first button,
 * and keep one swatch in the Tab order. @public
 */
export function focusColorSwatch(grid: HTMLElement | null): void {
  if (!grid) return;
  const current =
    grid.querySelector<HTMLElement>(`${SWATCH_SELECTOR}[aria-pressed="true"]`) ??
    grid.querySelector<HTMLElement>('button');
  if (current?.matches(SWATCH_SELECTOR)) rovingSwatch(grid, current);
  else rovingSwatch(grid, swatchesOf(grid)[0] ?? null);
  current?.focus({ preventScroll: true });
}

/**
 * Move the focus in an opened grid for an arrow key, Home, or End. The swatches are rows of ten;
 * Up from the first row goes to the Automatic (or No fill) button. Left and Right follow the
 * reading direction. Returns whether the key moved the focus, so the caller can prevent the
 * default action. @public
 */
export function moveColorSwatchFocus(grid: HTMLElement | null, key: string): boolean {
  if (!grid) return false;
  const swatches = swatchesOf(grid);
  const none = grid.querySelector<HTMLElement>(NONE_SELECTOR);
  const active = grid.ownerDocument.activeElement;
  const at = swatches.findIndex((swatch) => swatch === active);
  // Along the reading direction (`inlineStep`), as the bar and the lists move.
  const inline = inlineStep(key, grid);
  const step = inline === 0 ? undefined : inline;
  let next: HTMLElement | null | undefined;
  if (at < 0) {
    if (active !== none || (key !== 'ArrowDown' && key !== 'Home')) return false;
    next = swatches[0];
  } else if (step !== undefined) next = swatches[at + step];
  else if (key === 'ArrowDown') next = swatches[at + SWATCH_COLUMNS];
  else if (key === 'ArrowUp') next = at < SWATCH_COLUMNS ? none : swatches[at - SWATCH_COLUMNS];
  else if (key === 'Home') next = swatches[0];
  else if (key === 'End') next = swatches[swatches.length - 1];
  else return false;
  if (!next) return true;
  if (next !== none) rovingSwatch(grid, next);
  next.focus({ preventScroll: true });
  return true;
}
