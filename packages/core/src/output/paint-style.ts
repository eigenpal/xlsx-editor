// The paint styles of the canvas paint (design D7): CSS colors from resolved colors, the line of
// each `ST_BorderStyle` (ECMA-376 Part 1, 18.18.3), and the paint of a cell fill (`x:fill`,
// `CT_Fill`: `x:patternFill`, `CT_PatternFill`, or `x:gradientFill`, `CT_GradientFill`).
// DOM-free.
//
// Border lines are in grid pixels of the layout's profile, so a thin line is one pixel at every
// zoom. Only `thin` is measured (colors-fonts.pdf, row 31); the other styles follow the line
// weights and dash lengths that the style names state, the approximation "Border line styles".

import { toCss, type Rgba } from '@eigenpal/ooxml-core';
import type { GradientRect, ResolvedFill, Rgb } from '../cascade/index.ts';
import { PATTERN_TILES, patternCoverage } from './pattern-tiles.ts';

const HEX = /^[0-9A-Fa-f]{6}$/;

/** A resolved `RRGGBB` color as CSS, or `fallback` when it is not one. */
export function cssColor(rgb: Rgb | undefined, fallback = '#000000'): string {
  return rgb !== undefined && HEX.test(rgb) ? `#${rgb.toUpperCase()}` : fallback;
}

/** A DrawingML color with alpha as CSS. */
export function cssRgba(color: Rgba): string {
  return toCss(color);
}

/** One border line: its width and its dash pattern, in grid pixels. */
export interface BorderLine {
  /** The width of the line, in grid pixels. */
  readonly width: number;
  /** Dash and gap lengths in grid pixels; empty for a solid line. */
  readonly dash: readonly number[];
  /** Two thin lines with one pixel between them (`double`). */
  readonly double: boolean;
}

const LINES: Readonly<Record<string, BorderLine>> = Object.freeze({
  thin: { width: 1, dash: [], double: false },
  medium: { width: 2, dash: [], double: false },
  thick: { width: 3, dash: [], double: false },
  double: { width: 3, dash: [], double: true },
  hair: { width: 1, dash: [1, 1], double: false },
  dotted: { width: 1, dash: [1, 1], double: false },
  dashed: { width: 1, dash: [3, 1], double: false },
  dashDot: { width: 1, dash: [9, 3, 3, 3], double: false },
  dashDotDot: { width: 1, dash: [9, 3, 3, 3, 3, 3], double: false },
  mediumDashed: { width: 2, dash: [9, 3], double: false },
  mediumDashDot: { width: 2, dash: [9, 3, 3, 3], double: false },
  mediumDashDotDot: { width: 2, dash: [9, 3, 3, 3, 3, 3], double: false },
  slantDashDot: { width: 2, dash: [11, 1, 5, 1], double: false },
});

/**
 * The line of a border style (`x:left/@style` and its siblings, `ST_BorderStyle`), or undefined
 * for `none` and for an unknown style.
 */
export function borderLine(style: string): BorderLine | undefined {
  return Object.hasOwn(LINES, style) ? LINES[style] : undefined;
}

const channel = (rgb: string, i: number): number => parseInt(rgb.slice(i * 2, i * 2 + 2), 16);

function mix(a: Rgb, b: Rgb, share: number): string {
  if (!HEX.test(a) || !HEX.test(b)) return cssColor(a);
  const out = [0, 1, 2].map((i) =>
    Math.round(channel(a, i) * share + channel(b, i) * (1 - share))
      .toString(16)
      .padStart(2, '0')
  );
  return `#${out.join('').toUpperCase()}`;
}

/** One color stop of a gradient fill, with its position from 0 to 1. */
export interface GradientStop {
  readonly offset: number;
  readonly color: string;
}

/** What a cell fill paints: nothing, one color, or a gradient. */
export type FillPaint =
  | { readonly kind: 'none' }
  | { readonly kind: 'color'; readonly color: string }
  | {
      readonly kind: 'pattern';
      /** A key of `PATTERN_TILES`. */
      readonly pattern: string;
      readonly foreground: string;
      readonly background: string;
      /**
       * The mix of the two colors in the share of the set pixels: the paint when one tile is
       * smaller than 8 device pixels or larger than 512 (approximation "Fill pixels").
       */
      readonly mix: string;
    }
  | {
      readonly kind: 'gradient';
      /** `x:gradientFill/@type` (`ST_GradientType`, default `linear`). */
      readonly type: 'linear' | 'path';
      /** `x:gradientFill/@degree`: clockwise degrees from the left edge, for `linear`. */
      readonly degree: number;
      /** The inner rectangle of a `path` gradient, as fractions of the box of the cell. */
      readonly rect: GradientRect;
      readonly stops: readonly GradientStop[];
    };

/** The paint of a resolved fill. */
export function fillPaint(fill: ResolvedFill): FillPaint {
  if (fill.kind === 'none') return { kind: 'none' };
  if (fill.kind === 'solid') return { kind: 'color', color: cssColor(fill.color, '#FFFFFF') };
  if (fill.kind === 'pattern') {
    const background = fill.background ?? 'FFFFFF';
    const mixed = mix(fill.foreground, background, patternCoverage(fill.pattern));
    if (!Object.hasOwn(PATTERN_TILES, fill.pattern)) return { kind: 'color', color: mixed };
    return {
      kind: 'pattern',
      pattern: fill.pattern,
      foreground: cssColor(fill.foreground, '#000000'),
      background: cssColor(background, '#FFFFFF'),
      mix: mixed,
    };
  }
  const stops = fill.stops
    .map((s) => ({ offset: Math.min(1, Math.max(0, s.position)), color: cssColor(s.color) }))
    .sort((a, b) => a.offset - b.offset);
  if (stops.length === 0) return { kind: 'none' };
  if (stops.length === 1) return { kind: 'color', color: stops[0]!.color };
  return { kind: 'gradient', type: fill.type, degree: fill.degree, rect: fill.rect, stops };
}

/**
 * The gridline color on screen, as `RRGGBB`. Not measured: the reference PDFs print no gridlines
 * (approximation "Gridline color"). `x:sheetView` (`CT_SheetView`, sml.xsd) can name the color
 * with `@defaultGridColor` (`xsd:boolean`, default true) and `@colorId` (an indexed color,
 * default 64). The engine does not read these two attributes; it reads only `@showGridLines`.
 */
export const DEFAULT_GRIDLINE_COLOR: Rgb = 'D4D4D4';

/** A border edge band: where a line of `width` grid pixels lies on a gridline pixel at `at`. */
export function lineBand(at: number, widthPx: number, pixel: number): [number, number] {
  const start = at - Math.floor((widthPx - 1) / 2) * pixel;
  return [start, widthPx * pixel];
}
