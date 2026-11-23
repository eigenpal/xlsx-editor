// The tiles of the pattern fills (`ST_PatternType`, ECMA-376 Part 1, 18.18.55), as the reference
// PDF of fixture `fills-patterns` draws them: each pattern is an 8 x 8 tile
// of 1 point pixels, with the foreground color on the set pixels ('1') and the background color
// on the others. Row 0 is the top row. In the reference PDF, the tile grid starts at the
// bottom-left corner of the page, the origin of PDF pattern space; on screen it starts at the
// sheet origin of each pane. An absent foreground paints black, and an absent background paints
// white.
// `scripts/fidelity/pdf-fills.ts` reads the tiles from the reference PDF, and
// `scripts/fidelity/pdf-fills.test.ts` keeps this table equal to it.

/** The side of a tile, in pattern pixels. */
export const PATTERN_TILE_SIZE = 8;

/** The size of one pattern pixel, in points. */
export const PATTERN_PIXEL_PT = 1;

/** The rows of each tile, top first. */
export const PATTERN_TILES: Readonly<Record<string, readonly string[]>> = Object.freeze({
  darkGray: [
    '11101110',
    '10111011',
    '11101110',
    '10111011',
    '11101110',
    '10111011',
    '11101110',
    '10111011',
  ],
  mediumGray: [
    '10101010',
    '01010101',
    '10101010',
    '01010101',
    '10101010',
    '01010101',
    '10101010',
    '01010101',
  ],
  lightGray: [
    '10001000',
    '00100010',
    '10001000',
    '00100010',
    '10001000',
    '00100010',
    '10001000',
    '00100010',
  ],
  gray125: [
    '10001000',
    '00000000',
    '00100010',
    '00000000',
    '10001000',
    '00000000',
    '00100010',
    '00000000',
  ],
  gray0625: [
    '10000000',
    '00000000',
    '00001000',
    '00000000',
    '10000000',
    '00000000',
    '00001000',
    '00000000',
  ],
  darkHorizontal: [
    '11111111',
    '11111111',
    '00000000',
    '00000000',
    '11111111',
    '11111111',
    '00000000',
    '00000000',
  ],
  darkVertical: [
    '11001100',
    '11001100',
    '11001100',
    '11001100',
    '11001100',
    '11001100',
    '11001100',
    '11001100',
  ],
  darkDown: [
    '11001100',
    '01100110',
    '00110011',
    '10011001',
    '11001100',
    '01100110',
    '00110011',
    '10011001',
  ],
  darkUp: [
    '00110011',
    '01100110',
    '11001100',
    '10011001',
    '00110011',
    '01100110',
    '11001100',
    '10011001',
  ],
  darkGrid: [
    '10011001',
    '01100110',
    '01100110',
    '10011001',
    '10011001',
    '01100110',
    '01100110',
    '10011001',
  ],
  darkTrellis: [
    '11111111',
    '01100110',
    '11111111',
    '10011001',
    '11111111',
    '01100110',
    '11111111',
    '10011001',
  ],
  lightHorizontal: [
    '11111111',
    '00000000',
    '00000000',
    '00000000',
    '11111111',
    '00000000',
    '00000000',
    '00000000',
  ],
  lightVertical: [
    '10001000',
    '10001000',
    '10001000',
    '10001000',
    '10001000',
    '10001000',
    '10001000',
    '10001000',
  ],
  lightDown: [
    '10001000',
    '01000100',
    '00100010',
    '00010001',
    '10001000',
    '01000100',
    '00100010',
    '00010001',
  ],
  lightUp: [
    '00010001',
    '00100010',
    '01000100',
    '10001000',
    '00010001',
    '00100010',
    '01000100',
    '10001000',
  ],
  lightGrid: [
    '11111111',
    '10001000',
    '10001000',
    '10001000',
    '11111111',
    '10001000',
    '10001000',
    '10001000',
  ],
  lightTrellis: [
    '10001000',
    '01010101',
    '00100010',
    '01010101',
    '10001000',
    '01010101',
    '00100010',
    '01010101',
  ],
});

/** The share of the set pixels of a pattern: 0 for an unknown pattern. */
export function patternCoverage(pattern: string): number {
  const tile = Object.hasOwn(PATTERN_TILES, pattern) ? PATTERN_TILES[pattern] : undefined;
  if (!tile) return 0;
  let set = 0;
  for (const row of tile) for (const bit of row) if (bit === '1') set += 1;
  return set / (PATTERN_TILE_SIZE * PATTERN_TILE_SIZE);
}
