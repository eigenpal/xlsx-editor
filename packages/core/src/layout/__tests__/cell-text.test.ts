import { beforeAll, describe, expect, test } from 'bun:test';
import type { ResolvedAlignment, ResolvedFont } from '../../cascade/index.ts';
import { layoutCellText, type CellTextInput, type TextContext } from '../cell-text.ts';
import { PRINT_PROFILE } from '../metrics.ts';
import { fitNumber } from '../number-fit.ts';
import { substituteMeasurer } from './test-fonts.ts';

let ctx: TextContext;
beforeAll(async () => {
  ctx = { measurer: await substituteMeasurer(), profile: PRINT_PROFILE, maxDigitWidth: 6 };
});

const FONT: ResolvedFont = {
  family: 'Calibri',
  sizePt: 11,
  bold: false,
  italic: false,
  underline: 'none',
  strike: false,
  color: '000000',
  vertAlign: 'baseline',
};
const ALIGN: ResolvedAlignment = {
  horizontal: 'general',
  vertical: 'bottom',
  wrapText: false,
  shrinkToFit: false,
  indent: 0,
  textRotation: 0,
};

function lay(
  text: string,
  alignment: Partial<ResolvedAlignment> = {},
  extra: Partial<CellTextInput> = {}
) {
  return layoutCellText(
    {
      runs: [{ text, font: FONT }],
      kind: 'text',
      alignment: { ...ALIGN, ...alignment },
      box: { x: 0, y: 0, width: 96, height: 42 },
      ...extra,
    },
    ctx
  );
}

describe('horizontal alignment', () => {
  test('left, right, and center in a 96 point cell', () => {
    // `TRUE` is 23 points wide (5 + 6 + 7 + 5).
    expect(lay('TRUE', { horizontal: 'left' }).lines[0]!.x).toBe(3);
    expect(lay('TRUE', { horizontal: 'right' }).lines[0]!.x).toBe(94 - 23);
    expect(lay('TRUE', { horizontal: 'center' }).lines[0]!.x).toBe(3 + 34);
  });

  test('an indent moves text by one maximum digit width for each level (text-layout.pdf)', () => {
    expect(lay('x', { horizontal: 'left', indent: 2 }).lines[0]!.x).toBe(3 + 12);
  });

  test('fill repeats the text across the cell', () => {
    const fill = lay('ab', { horizontal: 'fill' });
    const text = fill.lines[0]!.spans.map((s) => s.text).join('');
    expect(text.length).toBeGreaterThan(2);
    expect(fill.lines[0]!.width).toBeLessThanOrEqual(91);
    expect(fill.mayOverflow).toBe(false);
  });

  test('justify spreads the space of every line but the last', () => {
    const j = lay('one two three four five six', { horizontal: 'justify' });
    expect(j.lines.length).toBeGreaterThan(1);
    expect(j.lines[0]!.wordSpacing).toBeGreaterThan(0);
    expect(j.lines.at(-1)!.wordSpacing).toBeUndefined();
  });
});

describe('vertical alignment and wrapping', () => {
  test('bottom, top, and center place the baseline in the box', () => {
    // A 13 pixel line (10 + 3) with a 14 point pitch, in a box of 42 points whose top pixel is
    // the gridline.
    expect(lay('x').lines[0]!.baseline).toBe(42 - 3);
    expect(lay('x', { vertical: 'top' }).lines[0]!.baseline).toBe(1 + 10);
    // (41 - 14) / 2 rounds half up, as `Middle` in text-layout.pdf.
    expect(lay('x', { vertical: 'center' }).lines[0]!.baseline).toBe(1 + 14 + 10);
  });

  test('wrapped lines break at spaces and stack by the line height', () => {
    const w = lay('alpha beta gamma delta epsilon', { wrapText: true, vertical: 'top' });
    expect(w.lines.length).toBeGreaterThan(1);
    expect(w.lines[1]!.baseline - w.lines[0]!.baseline).toBe(14);
    expect(w.mayOverflow).toBe(false);
  });

  test('a line feed breaks a wrapped line and draws nothing when unwrapped', () => {
    expect(lay('a\nb', { wrapText: true }).lines).toHaveLength(2);
    const one = lay('a\nb');
    expect(one.lines).toHaveLength(1);
    expect(one.lines[0]!.spans.map((s) => s.text).join('')).toBe('ab');
  });
});

describe('rotation and stacked text', () => {
  test('90 degrees turns the line counterclockwise around an origin in the cell', () => {
    const r = lay('Up', { textRotation: 90 });
    expect(r.rotation!.angle).toBe(90);
    expect(r.right - r.left).toBe(13);
  });

  test('135 is 45 degrees clockwise', () => {
    expect(lay('Down', { textRotation: 135 }).rotation!.angle).toBe(-45);
  });

  test('general aligns rotated text by its angle, and rotated numbers by the number rule', () => {
    // Text at 45 degrees aligns left (text-layout.pdf); a number keeps `general` right.
    const text = lay('12', { textRotation: 45 });
    const number = lay('12', { textRotation: 45 }, { kind: 'number' });
    expect(text.left).toBe(text.areaLeft);
    expect(number.right).toBeCloseTo(number.areaRight - 1, 9);
    // A number shown through `@` aligns as text.
    const asText = lay('12', { textRotation: 45 }, { kind: 'number', alignKind: 'text' });
    expect(asText.left).toBe(asText.areaLeft);
    // A boolean keeps `general` center.
    const bool = lay('TRUE', { textRotation: 45 }, { kind: 'boolean' });
    expect(bool.left - bool.areaLeft).toBeCloseTo(bool.areaRight - bool.right, 9);
  });

  test('only bottom-aligned text at 90 degrees keeps the margin above the bottom', () => {
    // The box is 42 points high and its top pixel is the gridline: the free height is 1 to 42.
    const bottom = lay('Up', { textRotation: 90 });
    const center = lay('Up', { textRotation: 90, vertical: 'center' });
    const w = bottom.lines[0]!.width;
    // At 90 degrees the line runs up from its origin, so its box is `originY - w` to `originY`.
    expect(bottom.rotation!.originY).toBe(42 - 2);
    expect(center.rotation!.originY - w / 2).toBe((1 + 42) / 2);
  });

  test('255 stacks one character on each line', () => {
    expect(lay('abc', { textRotation: 255 }).lines).toHaveLength(3);
  });
});

describe('fitNumber', () => {
  const width = (s: string) => s.length * 6;
  test('General drops decimals, then goes to scientific form, then shows #', () => {
    expect(fitNumber(1234567.891, '1234567.891', true, '.', 48, width)).toBe('1234568');
    expect(fitNumber(123456789012, '1.23457E+11', true, '.', 48, width)).toBe('1.2E+11');
    expect(fitNumber(0.000123456, '0.000123456', true, '.', 25, width)).toBe('0');
    expect(fitNumber(1234567.891, '1234567.891', true, '.', 20, width)).toBe('###');
  });

  test('a formatted number that does not fit fills the area with #', () => {
    expect(fitNumber(1234.5, '1,234.50', false, '.', 30, width)).toBe('#####');
    expect(fitNumber(1234.5, '1,234.50', false, '.', 48, width)).toBe('1,234.50');
  });
});

describe('only text wraps', () => {
  test('a number with wrapText shows on one line', () => {
    const n = lay(
      '1234567.891',
      { wrapText: true },
      { kind: 'number', box: { x: 0, y: 0, width: 30, height: 42 } }
    );
    expect(n.lines.length).toBe(1);
    expect(n.lines[0]!.spans.map((s) => s.text).join('')).toBe('1234567.891');
    // The same characters as text break into lines.
    const t = lay(
      '1234567.891',
      { wrapText: true },
      { box: { x: 0, y: 0, width: 30, height: 42 } }
    );
    expect(t.lines.length).toBeGreaterThan(1);
  });

  test('a number with justify or distributed alignment shows on one line', () => {
    const box = { x: 0, y: 0, width: 30, height: 42 };
    for (const horizontal of ['justify', 'distributed'] as const)
      expect(lay('1 234 567', { horizontal }, { kind: 'number', box }).lines.length).toBe(1);
  });
});

describe('a line break in text that does not wrap', () => {
  const words = (t: ReturnType<typeof lay>) => t.lines[0]!.spans.flatMap((s) => s.positions);

  test('keeps the words apart by the width of a space', () => {
    const joined = lay('cellAcellB');
    const broken = lay('cellA\ncellB');
    const spaced = lay('cellA cellB');
    expect(broken.lines.length).toBe(1);
    // No glyph for the line break; the second word starts where it starts after a space.
    expect(broken.lines[0]!.spans.map((s) => s.text).join('')).toBe('cellAcellB');
    expect(broken.lines[0]!.width).toBe(spaced.lines[0]!.width);
    expect(broken.lines[0]!.width).toBeGreaterThan(joined.lines[0]!.width);
    expect(words(broken)[5]).toBe(words(spaced)[6]);
  });

  test('takes one space for CR LF and for a run of breaks, and none at the ends', () => {
    const spaced = lay('a b').lines[0]!.width;
    expect(lay('a\r\nb').lines[0]!.width).toBe(spaced);
    expect(lay('a\n\n\nb').lines[0]!.width).toBe(spaced);
    expect(lay('a\rb').lines[0]!.width).toBe(spaced);
    expect(lay('\na\n').lines[0]!.width).toBe(lay('a').lines[0]!.width);
  });

  test('still starts a line in wrapped text', () => {
    expect(lay('cellA\ncellB', { wrapText: true }).lines.length).toBe(2);
  });
});
