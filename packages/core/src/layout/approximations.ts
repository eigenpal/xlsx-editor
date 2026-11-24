// The approximations the engine knows it makes (AGENTS.md, "Fidelity": name every approximation
// in one list in code and in the public known approximations). Each entry's title is a bold item
// under "Known approximations" in `docs/site/content/known-approximations.mdx`; a test keeps the
// two lists equal, in order.
//
// Same shape as pptx-editor/packages/core/src/layout/approximations.ts at 1943d2c9 (EigenPal).
// Difference: `perSheet` replaces `perSlide`. Layout reports the entries that a sheet uses when
// the layout lane lands (section 5 of the active change); until then every entry is listed once.

/** One known approximation. */
export interface Approximation {
  readonly id: string;
  /** The bold title of its item in the public list, without the period. */
  readonly title: string;
  /** Whether layout detects it on a sheet. */
  readonly perSheet: boolean;
  /**
   * The cases that no reference PDF measures, each as its public item names it word for word.
   * A test keeps each one in the text of the item.
   */
  readonly unmeasured?: readonly string[];
}

/** Every known approximation, in the order of the public list. */
export const APPROXIMATIONS: readonly Approximation[] = Object.freeze([
  { id: 'pixel-density', title: 'Pixel density', perSheet: false },
  {
    id: 'screen-geometry',
    title: 'Screen geometry',
    perSheet: false,
    unmeasured: [
      'the whole-pixel snap of rotated text in the screen profile',
      'the minimum descent of bottom-aligned text in the screen profile',
    ],
  },
  { id: 'tint-rounding', title: 'Tint rounding', perSheet: false },
  { id: 'row-auto-height', title: 'Row auto height', perSheet: false },
  { id: 'wrap-measure-budget', title: 'Row height of very long wrapped text', perSheet: false },
  {
    id: 'unmeasured-text-layout',
    title: 'Text layout without a measurement',
    perSheet: false,
    unmeasured: [
      'overflow that reaches a cell with a constant value',
      'overflow that reaches a formula with no cached value',
      'the minimum descent of bottom-aligned text in fonts other than Carlito 6 and 8',
      'stacked text at the top or center of a cell',
      'a number in a cell with `wrapText`',
      'a line break in a cell without `wrapText`',
    ],
  },
  {
    id: 'unmeasured-rotation',
    title: 'Rotated text without a measurement',
    perSheet: false,
    unmeasured: [
      'angles other than 45, 90, and -45 degrees',
      '-90 degrees (`textRotation` 180)',
      'numbers, booleans, and errors with rotation',
      'rotated text at the top or center of a cell',
      'an explicit horizontal alignment of rotated text',
      'indent with rotation',
      'rotated text in a merged range',
      '`justify`, `distributed`, and `centerContinuous` with rotation',
      '`wrapText` with rotation',
      'rotated text taller than its cell',
    ],
  },
  { id: 'measured-text-offsets', title: 'Measured text offsets', perSheet: false },
  { id: 'overflow-reach', title: 'Overflow reach', perSheet: false },
  { id: 'drawing-objects', title: 'Drawing objects', perSheet: false },
  { id: 'border-line-styles', title: 'Border line styles', perSheet: false },
  { id: 'fill-pixels', title: 'Fill pixels', perSheet: false },
  { id: 'gridline-color', title: 'Gridline color', perSheet: false },
  { id: 'glyph-positions', title: 'Glyph positions', perSheet: false },
  { id: 'substitute-outlines', title: 'Substitute font outlines', perSheet: false },
  { id: 'glyph-rasterization', title: 'Glyph rasterization', perSheet: false },
  {
    id: 'clipped-line-scores',
    title: 'Line scores of clipped text',
    perSheet: false,
    unmeasured: [
      'a text run that starts or ends with a space glyph',
      'a Type 3 font with another font matrix',
      'a line that a clip cuts on both sides',
      'a clip edge inside a space',
      'a painted line that keeps one character other than a digit',
      'a top or bottom clip edge',
      'a cell of one character other than a digit',
      'rotated, mirrored, or vertical text',
      'repeated text that pairs across cells',
      'a cut line that differs only in white space',
    ],
  },
  {
    id: 'join-gap-limit',
    title: 'Join gap limit',
    perSheet: false,
    unmeasured: [
      'a gap of the pieces near the limit',
      'a join piece that moved along the row past the gap limit',
      'a join piece that moved off the row',
    ],
  },
  {
    id: 'reference-items-across-cells',
    title: 'Reference items across cells',
    perSheet: false,
    unmeasured: ['a reference item that holds the text of two cells'],
  },
  {
    id: 'conditional-rules',
    title: 'Conditional formatting rules',
    perSheet: false,
    unmeasured: [
      'a formula cell with a cached value, and a formula cell without one, which does not take part',
      'rules with equal `@priority`, which walk in document order',
      'boolean operands, and the order of numbers before booleans',
      'the order of strings by UTF-16 code unit after lowercase',
      'number equality without rounding',
      'a value equal to a bound of `between` or `notBetween`',
      'a stored value that cannot be read, such as a bad number, which compares as an empty value',
      'date values, and an empty value under a boolean operand',
      'the font family and the font size of a differential format, and text that clips in its row',
      '`patternType="none"`, another pattern type, and a gradient fill in a differential format',
      'differential borders',
      'a `solid` fill with only `x:fgColor` over a cell with no fill',
      'an `x:patternFill` with no color, an `x:bgColor` that does not resolve, and a `solid` fill with both `x:fgColor` and `x:bgColor`',
      'two shown true rules that set the same property, and a differential font that turns a property off',
      'a rule that is not shown with `stopIfTrue`, or that is true above a shown rule that sets the same property',
      'a true rule without `@dxfId`, or with a `@dxfId` that no `x:dxf` has, with `stopIfTrue`',
      'a fill with only `x:fgColor` before a true rule with `x:bgColor` over one cell',
      'an empty merged range under a rule',
      'a rich text run under a true rule',
      'a differential fill over a cell with a pattern fill or a gradient fill',
      'a differential font over a cell with its own font properties',
      'a differential number format over a cell with its own number format',
      'a differential color that is `auto` or a system color (`indexed` 64 or 65)',
      'underline styles other than `single` in a differential font',
      'a differential `x:vertAlign`',
      'a rule that covers a cell of a merged range other than its top-left cell',
    ],
  },
  { id: 'cut-references', title: 'Cut references', perSheet: false },
  { id: 'cut-cell-attachments', title: 'Cut cell attachments', perSheet: false },
  { id: 'pasted-values', title: 'Pasted rich text and errors', perSheet: false },
  { id: 'pasted-quotes', title: 'Quoted pasted text', perSheet: false },
  { id: 'copied-line-breaks', title: 'Copied line breaks', perSheet: false },
  { id: 'shadow-root-clipboard', title: 'Clipboard keys in a shadow root', perSheet: false },
  { id: 'sheet-titles', title: 'Sheet titles in the document properties', perSheet: false },
  { id: 'legacy-drawing-anchors', title: 'Legacy drawing anchors', perSheet: false },
  {
    id: 'wrapped-formula-values',
    title: 'Wrapped formula values',
    perSheet: false,
    unmeasured: [
      'an `mc:Choice` whose namespace the consumer understands',
      'a cell with no `mc:Fallback`',
      'a value from a branch with no `x:f`',
      'an `mc:Fallback` with a value and no `x:f`, beside an `mc:Choice` with a formula',
    ],
  },
  // A run of `?` after a literal or a dash in a zero section pads it. Only a section with a fill
  // and a dash gets its `?` back. Beside another literal, or with no fill, a decrease to no
  // places removes the last `?`, and an increase does not add it back.
  { id: 'decimal-padding', title: 'Decimal padding beside a literal', perSheet: false },
  // A decimal step adds or removes the last decimal placeholder of a section. When a literal or
  // a second point splits the places, as in `0.00" "0`, a step can move a place across the
  // literal, so a decrease and an increase do not return the code.
  { id: 'decimal-split-places', title: 'Decimal places split by a literal', perSheet: false },
]);
