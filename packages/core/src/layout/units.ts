// Length units. Layout works in points. EMU convert at the fixed ratio of ECMA-376 Part 1,
// 20.1.2.1 (914400 EMU per inch). CSS pixels, which a host uses to place layout records on a
// page, are 96 for each inch (CSS Values and Units Level 4, 6.2).
//
// The pixel grid of column widths and row heights is a different density: the one that the
// width formulas of Part 1, 18.3.1.13, round to. It belongs to the grid profile (`metrics.ts`),
// and the reference PDFs measure it at 72 pixels for each inch (the approximation "Pixel
// density").

/** EMU in one point: 914400 EMU per inch, 72 points per inch. */
export const EMU_PER_POINT = 12_700;
/** CSS pixels in one inch. */
export const CSS_PIXELS_PER_INCH = 96;
/** Points in one CSS pixel. */
export const POINTS_PER_PIXEL = 72 / CSS_PIXELS_PER_INCH;

export const emuToPoints = (emu: number): number => emu / EMU_PER_POINT;
/** CSS pixels to points. */
export const pixelsToPoints = (pixels: number): number => pixels * POINTS_PER_PIXEL;
/** Points to CSS pixels. */
export const pointsToPixels = (points: number): number => points / POINTS_PER_PIXEL;
