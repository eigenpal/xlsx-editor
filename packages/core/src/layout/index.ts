// The `layout` lane: DOM-free grid geometry and cell text layout, in points, with an injected
// text measurer. `layoutSheet` is the entry point that paint reads: target-neutral records of the
// visible window of a worksheet.

export * from './units.ts';
export * from './approximations.ts';
// What a host needs to measure text: the measurer contract, the HarfBuzz measurer and its
// runtime, the fixed measurer for tests, and the font substitutes and fallbacks that a host
// supplies faces for. Internals stay unexported.
export {
  createFixedMeasurer,
  DEFAULT_FALLBACK_FAMILY,
  EAST_ASIAN_FALLBACK_FAMILIES,
  eastAsianFallbackFamilies,
  type FaceMetrics,
  FONT_SUBSTITUTES,
  type FontRequest,
  type ResolvedFace,
  type TextMeasurer,
} from './text/measurer.ts';
export {
  createHarfBuzzMeasurer,
  type FontSource,
  HOST_SUBSTITUTE_ORIGIN,
  initializeHarfBuzz,
  isHarfBuzzReady,
  type RegisteredFace,
} from './text/harfbuzz-measurer.ts';
export { setHarfBuzzWasmUrl } from './harfbuzz-wasm-binary.ts';
export { breakOpportunities, isIdeographic, type KinsokuLists } from './text/line-break.ts';
export { AxisIndex, type AxisSegment } from './axis.ts';
export {
  type AutoRowHeight,
  baseColumnWidth,
  columnPixels,
  columnWidthForPixels,
  type GridMetrics,
  rowHeightForPixels,
  sheetGeometry,
  type SheetGeometry,
} from './grid-geometry.ts';
export {
  autoRowHeight,
  type FontKey,
  type FontPixels,
  fontPixels,
  gridMetrics,
  type GridProfile,
  pixelWidth,
  PRINT_PROFILE,
  SCREEN_PROFILE,
  zoomProfile,
} from './metrics.ts';
export {
  type CellRange,
  frozenPanes,
  type FrozenPanes,
  type MergeIndex,
  mergedRanges,
  parseRange,
} from './panes.ts';
export {
  type CellTextInput,
  type CellTextLayout,
  layoutCellText,
  type Rect,
  type TextContext,
  type TextLine,
  type TextRunInput,
  type TextSpan,
  type TextStyle,
  type ValueKind,
} from './cell-text.ts';
export { fitNumber } from './number-fit.ts';
export { fitColumnWidths, fitRowHeights, type FitColumnOptions, type FitOptions } from './fit.ts';
export {
  drawingBounds,
  type DrawingLayout,
  drawingPartOf,
  MAX_ANCHORS,
  type PictureLayout,
  type ShapeLayout,
  sheetDrawings,
  type UnsupportedDrawing,
} from './drawings.ts';
export {
  type AxisEntry,
  type CellBox,
  type CellTextBox,
  layoutSheet,
  type LayoutOptions,
  type LayoutProblem,
  MAX_WINDOW,
  OVERFLOW_REACH,
  type PaneLayout,
  type SheetLayout,
  type Viewport,
} from './sheet-layout.ts';
