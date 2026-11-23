// The `output` lane: paint from layout records (design D7). `paintSheet` draws a sheet window
// into a canvas, the target that the viewer and the fidelity harness use, and `accessibleGrid`
// gives the same window as text for assistive technology. Pictures decode from package bytes.

export { MAX_CANVAS_SIDE, paintSheet, type PaintOptions, type PaintStats } from './paint-sheet.ts';
export { accessibleGrid, cellText, type AccessibleGridLabels } from './accessible-grid.ts';
export {
  createPictureStore,
  MAX_PICTURE_PIXELS,
  type PictureState,
  type PictureStore,
  type PictureWindow,
} from './pictures.ts';
export { registerFontFaces, type FontFaceHandle } from './font-faces.ts';
export {
  borderLine,
  type BorderLine,
  cssColor,
  DEFAULT_GRIDLINE_COLOR,
  fillPaint,
  type FillPaint,
} from './paint-style.ts';
