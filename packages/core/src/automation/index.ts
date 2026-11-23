// The `automation` lane: the DOM-free edits that lower to store ops, after DOCX Editor's
// `core/src/automation` commands. The format commands land with task 7.4 of the active change,
// and the structural commands with task 7.5.

export {
  clearOps,
  formatOps,
  togglePatch,
  type ClearKind,
  type CommandOps,
  type FormatToggle,
} from './format-commands.ts';
export { mergeOps, unmergeOp, type MergeOptions } from './merge-commands.ts';
export {
  isWholeSheet,
  selectionBand,
  structureOp,
  type SelectionBand,
} from './structure-commands.ts';
// The automation host, after PPTX Editor's and DOCX Editor's `core/src/automation` hosts: the
// protocol that `@xlsx-editor.dev/editor-api` speaks, and the headless host over XLSX bytes.
export type {
  AutomationBatchRequest,
  AutomationBatchResponse,
  AutomationCapabilities,
  AutomationCellValue,
  AutomationCellWrite,
  AutomationError,
  AutomationErrorCode,
  AutomationHost,
  AutomationOpenResult,
  AutomationOperation,
  AutomationOperationResult,
  AutomationRangeRead,
  AutomationSaveResult,
  AutomationValue,
  AutomationWorksheetRead,
  AutomationWorksheetSelector,
} from './protocol.ts';
export {
  MAX_ARRAY_RANGES,
  MAX_BATCH_CELLS,
  MAX_BATCH_LISTED_WORKSHEETS,
  MAX_BATCH_MERGE_CHECKS,
  MAX_BATCH_OPERATIONS,
  MAX_BATCH_WRITE_CELLS,
  MAX_CELL_CHILDREN,
  openAutomationHost,
} from './host.ts';
export { MAX_RANGE_CELLS } from './range-address.ts';
export type { WorkbookOpenDiagnosticCode, WorkbookOpenLimit } from './open-diagnostics.ts';
