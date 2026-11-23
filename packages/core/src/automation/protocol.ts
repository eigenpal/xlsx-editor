// The automation host protocol, after PPTX Editor's `packages/core/src/automation/protocol.ts`
// at d4502552, which adapts DOCX Editor's `packages/core/src/automation/protocol.ts` at adf0e178.
// Differences: the operations are the workbook subset of this repository (worksheets and cell
// ranges), a worksheet is named by an opaque id instead of a minted handle, and there is no
// selection, image, or event capability.
//
// EVERYTHING HERE IS TRANSPORT-SHAPED. A request is data and a response is data, so the same
// host can sit behind a worker port, an HTTP boundary, or nothing at all. Nothing in this file
// names a tree node or a store.

import type { WorkbookOpenDiagnosticCode, WorkbookOpenLimit } from './open-diagnostics.ts';

/**
 * What a host supports, fixed at construction and frozen.
 *
 * @public
 */
export interface AutomationCapabilities {
  /** Reading and editing the workbook. Every operation needs it. */
  readonly document: boolean;
  /** Serializing the current workbook back to `.xlsx` bytes. */
  readonly save: boolean;
}

/**
 * Why an operation was refused, as a value to branch on.
 *
 * @public
 */
export type AutomationErrorCode =
  /** Nothing was applied: `expectedRevision` did not match the host's revision. */
  | 'stale-revision'
  /** An argument outside its range: an address, a value, a size. */
  | 'invalid-argument'
  /** The operation, argument form, or content is outside the subset that this host implements. */
  | 'unsupported'
  /** The operation would read or write more cells than one operation or batch may. */
  | 'resource-limit'
  /** The host has been disposed. Every later call fails this way. */
  | 'disposed'
  /** The store refused the transaction; nothing was applied. */
  | 'transaction-refused'
  /** The operation is not one this protocol version defines. */
  | 'unknown-operation'
  /** A lookup by key found nothing, and the caller did not ask for null. */
  | 'item-not-found';

/**
 * One refused operation. Plain data, not an `Error`: it crosses transports. `code` is the stable
 * part. `detail` selects a diagnostic in the object model, which never copies it into a message.
 *
 * @public
 */
export interface AutomationError {
  readonly code: AutomationErrorCode;
  /** Human-readable, for a log. Never parsed. */
  readonly message: string;
  /** Machine-ish specifics: the store's rejection code, the name of the bad argument. */
  readonly detail?: string;
  /**
   * For a refusal of the content of a cell: the A1 reference of the first such cell, such as
   * `B3`. The host makes it from the cell position, never from file text.
   */
  readonly cell?: string;
}

/**
 * One worksheet, as plain values. `id` is opaque, and it follows the worksheet through a rename
 * or a move.
 *
 * @public
 */
export interface AutomationWorksheetRead {
  readonly id: string;
  readonly name: string;
}

/**
 * A cell value as the file stores it. An empty cell reads as `""` and an error value as its
 * text, such as `#DIV/0!`.
 *
 * @public
 */
export type AutomationCellValue = number | string | boolean;

/**
 * A value to store in one cell: a number, text as written, a boolean, `empty` to clear the
 * value, or `null` to leave the cell as it is.
 *
 * @public
 */
export type AutomationCellWrite =
  | { readonly type: 'number'; readonly value: number }
  | { readonly type: 'string'; readonly value: string }
  | { readonly type: 'boolean'; readonly value: boolean }
  | { readonly type: 'empty' }
  | null;

/**
 * One range, as plain values, rows first.
 *
 * @public
 */
export interface AutomationRangeRead {
  /** The address with its sheet, such as `Sheet1!A1:B2`. */
  readonly address: string;
  readonly values?: readonly (readonly AutomationCellValue[])[];
  /** The number format code of each cell, such as `General` or `0.00%`. */
  readonly numberFormat?: readonly (readonly string[])[];
}

/**
 * The worksheet that an operation addresses: by its id, by its name (matched without letter
 * case), or the active worksheet.
 *
 * @public
 */
export type AutomationWorksheetSelector =
  { readonly id: string } | { readonly name: string } | { readonly active: true };

/**
 * One operation of a batch. A range read of values or formats after a value write in the same
 * batch is refused: the batch reads the workbook as it was before the batch.
 *
 * @public
 */
export type AutomationOperation =
  /** The worksheets of the workbook in tab order. Chartsheets are not worksheets. */
  | { readonly op: 'listWorksheets' }
  /** One worksheet. With `orNull`, a selector that matches none answers `null`. */
  | {
      readonly op: 'getWorksheet';
      readonly sheet: AutomationWorksheetSelector;
      readonly orNull?: boolean;
    }
  /** Read a range of a worksheet (`A1` or `A1:B2`). */
  | {
      readonly op: 'readRange';
      readonly sheet: AutomationWorksheetSelector;
      readonly address: string;
      readonly values?: boolean;
      readonly numberFormat?: boolean;
    }
  /** Store values in a range. The array has exactly the rows and columns of the range. */
  | {
      readonly op: 'writeValues';
      readonly sheet: AutomationWorksheetSelector;
      readonly address: string;
      readonly values: readonly (readonly AutomationCellWrite[])[];
    };

/**
 * What one operation answered.
 *
 * @public
 */
export type AutomationValue =
  null | AutomationWorksheetRead | readonly AutomationWorksheetRead[] | AutomationRangeRead;

/**
 * The result of one operation: its value, its refusal, or `skipped` when another operation of
 * the batch was refused.
 *
 * @public
 */
export type AutomationOperationResult =
  | { readonly status: 'ok'; readonly value: AutomationValue }
  | { readonly status: 'error'; readonly error: AutomationError }
  | { readonly status: 'skipped' };

/**
 * A batch: operations in order, and the revision that the caller last read. When the host is at
 * another revision, the batch is refused with `stale-revision` and nothing applies.
 *
 * @public
 */
export interface AutomationBatchRequest {
  readonly operations: readonly AutomationOperation[];
  readonly expectedRevision?: number;
}

/**
 * The answer to a batch. All or nothing: when `ok` is false, nothing was applied.
 *
 * @public
 */
export interface AutomationBatchResponse {
  readonly ok: boolean;
  readonly results: readonly AutomationOperationResult[];
  /** The host revision after the batch. */
  readonly revision: number;
  /** The batch committed a change. */
  readonly changed: boolean;
}

/**
 * The bytes of a save, or why there are none.
 *
 * @public
 */
export type AutomationSaveResult =
  | { readonly ok: true; readonly bytes: Uint8Array }
  | { readonly ok: false; readonly error: AutomationError };

/**
 * A host over opened bytes, or why the bytes do not open: `code` says whether the bytes are
 * malformed or pass a reader limit, `reason` is the public open diagnostic, and `limit` names the
 * reader limit that the bytes passed.
 *
 * @public
 */
export type AutomationOpenResult =
  | { readonly ok: true; readonly host: AutomationHost }
  | {
      readonly ok: false;
      readonly code: 'invalid-argument' | 'resource-limit';
      readonly reason: WorkbookOpenDiagnosticCode;
      readonly limit?: WorkbookOpenLimit;
    };

/**
 * A workbook automation host.
 *
 * @public
 */
export interface AutomationHost {
  readonly capabilities: AutomationCapabilities;
  /** The current revision of the workbook. */
  revision(): number;
  /** Run a batch. Never throws for a refusal: the response names it. */
  execute(request: AutomationBatchRequest): AutomationBatchResponse;
  /** The current workbook as `.xlsx` bytes. */
  save(): AutomationSaveResult;
  /** Release the host. Idempotent. */
  dispose(): void;
}
