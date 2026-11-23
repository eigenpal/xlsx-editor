// The `store` lane: package read and save, the canonical tree, the document store, and its ops.
// DOM-free.

export * from './cell-ref.ts';
export * from './limits.ts';
export * from './package/namespaces.ts';
export * from './package/tree.ts';
export * from './package/query.ts';
export * from './package-write/serialize.ts';
export * from './package-write/write.ts';
export * from './package/package.ts';
// The `limits` parameter type of `readXlsx`, as DOCX Editor's core `store` entry exports it.
export type { ZipLimits } from '@eigenpal/ooxml-core';
// The `limits` parameter type and the rejection type of `parseXmlPart`, for the same reason.
export type { XmlLimits, XmlRejection } from '@eigenpal/ooxml-core';
export * from './workbook/workbook-index.ts';
// Not `sheetIndexTestRecorder`: a work counter for tests stays off the entry.
export { cellKey, keyAddress, sheetIndex } from './workbook/sheet-index.ts';
export type { CellEntry, SheetIndex, SheetProblem } from './workbook/sheet-index.ts';
export * from './workbook/inherited-format.ts';
export * from './workbook/xstring.ts';
export * from './formula/tokenizer.ts';
export * from './formula/format.ts';
export * from './formula/grammar.ts';
export { sharedFormulaText } from './store/shared-formulas.ts';
export { arrayFormulaIn, arrayRangeCount, cellChildBlocker, isIsoDate } from './store/cells.ts';
// Internal to the engine lanes: these take and give tree types (`XmlPart`, `Element`).
export { sheetMayHoldFormula } from './store/calc.ts';
export { cellHasFormula, wrappedFormula } from './workbook/cell-formula.ts';
export { STRUCTURE_SUBJECT_KINDS } from './store/ops.ts';
export type {
  CellInput,
  HistoryOp,
  InverseOp,
  OpRejection,
  StructureSubject,
  StructureSubjectKind,
  StructuralOpName,
  WorkbookOp,
} from './store/ops.ts';
export * from './store/format-patch.ts';
export {
  MAX_FORMAT_NEW_CELLS,
  clearFormatsOverflow,
  formatRangeOverflow,
} from './store/format-range.ts';
export {
  clearBlocker,
  clearRangeBlocker,
  rangeBounds,
  wrappedValueCell,
  type RangeBox,
} from './store/clear-range.ts';
export { firstValueLost, mergedRefMeeting, type CellOverride } from './store/merge-cells.ts';
export { MAX_STYLE_RECORDS } from './store/style-records.ts';
export { MAX_COLUMN_WIDTH, MAX_RESIZE_NEW_ROWS, MAX_ROW_HEIGHT } from './store/axis-size.ts';
export {
  STRUCTURE_PERMIT,
  cellLocked,
  lockedCellIn,
  protectionRefusal,
  readSheetProtection,
  readWorkbookProtection,
  type SheetProtection,
  type WorkbookProtection,
} from './store/sheet-protection.ts';
export * from './store/document-store.ts';
export { structureBlocker } from './structure/structure.ts';
export { workbookPartBlocker } from './structure/blockers.ts';
export { sheetPrefixText } from './structure/sheet-refs.ts';
export * from './package-write/save.ts';
export * from './workbook/hyperlinks.ts';
