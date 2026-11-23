/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See LICENSE.md in this package.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// After PPTX Editor's `packages/editor-api/src/runtime/errors.ts` at d4502552, which adapts DOCX
// Editor's `packages/editor-api/src/runtime/errors.ts` at 5993bd223945. Differences: the codes of
// the workbook subset (no `ValueNotLoaded`, `ObjectInUse`, or `NotImplemented`, because the
// subset has no client results, tracked objects, or members that it declares and does not
// implement), the diagnostics of the workbook host, and no `nearest` name suggestion. The
// `limit` field and the open hints follow DOCX Editor's `errors.ts` and
// `packages/core/src/automation/diagnostic-hints.ts` at 6c11927b. A limit hint says that the
// limit is fixed, as DOCX Editor's `open.xml-depth-limit` hint does, because `createServer` has
// no `limits` option.
//
// One error class, a stable code per refusal, and one fixed sentence per code. Nothing from the
// engine or the file reaches a message: a host detail only selects a public diagnostic, whose
// fixed hint becomes `debugInfo.innerError`.

import type {
  WorkbookOpenDiagnosticCode,
  WorkbookOpenLimit,
} from '@xlsx-editor.dev/core/automation';

/**
 * What went wrong, as a value a consumer may branch on.
 *
 * @public
 */
export type XlsxEditorErrorCode =
  /** A property was read before a `load(...)` for it completed in a `sync()`. */
  | 'PropertyNotLoaded'
  /**
   * The object cannot be addressed: its run ended, it is a null object, or the workbook no
   * longer holds it.
   */
  | 'InvalidObjectPath'
  /** An argument or load option this API does not accept. */
  | 'InvalidArgument'
  /** The workbook or the request exceeds a bounded limit. */
  | 'ResourceLimitExceeded'
  /** The collection has no such item: `getItem` with a name that no worksheet has. */
  | 'ItemNotFound'
  /**
   * The call, argument, or content is outside the supported subset for this host, and nothing
   * changed. Examples: a range address with a sheet name, a cell that holds a formula, or a
   * formula write.
   */
  | 'NotSupported'
  /** The request context's `run` has finished, so it can no longer be used. */
  | 'InvalidRequestContext'
  /** The runtime was disposed. Every later operation fails this way. */
  | 'RuntimeDisposed'
  /**
   * The batch was built from an older read, so nothing was applied: another writer changed the
   * workbook after this context read it. Read again before you retry.
   */
  | 'StaleDocument'
  /** The workbook refused the change, or answered something this runtime cannot use. */
  | 'GeneralException';

const MESSAGES: Readonly<Record<XlsxEditorErrorCode, string>> = Object.freeze({
  PropertyNotLoaded:
    'the property has not been loaded. Call load(...) and await context.sync() before reading it.',
  InvalidObjectPath:
    'the object cannot be addressed. Use each proxy only in the run that made it, and check ' +
    'isNullObject after a sync.',
  InvalidArgument: 'the argument is not one this API accepts.',
  ResourceLimitExceeded: 'the workbook or the request exceeds a supported resource limit.',
  ItemNotFound: 'the collection has no such item.',
  NotSupported:
    'the operation or argument is outside the supported subset for this host or content.',
  InvalidRequestContext: 'the request context has finished. Start another run to continue.',
  RuntimeDisposed: 'the runtime has been disposed.',
  StaleDocument:
    'the workbook changed after this context read it, so nothing was applied. Read again ' +
    'before you decide.',
  GeneralException: 'the workbook could not complete the request.',
});

/**
 * A public reason for a refusal: a stable code and a fixed hint.
 *
 * @public
 */
export interface XlsxEditorDiagnostic {
  /** A stable public diagnostic code, such as `write.text-reads-as-value`. */
  readonly code: string;
  /**
   * The fixed hint for the code: the cause, then the action. Never workbook content. A refusal of
   * the content of a cell ends it with the reference of that cell, such as `First refused cell:
   * B3.`
   */
  readonly hint: string;
}

/** The diagnostic of an argument refusal that names what the API accepts there. */
const ARGUMENT_NOT_ACCEPTED = 'argument.not-accepted';

const TOO_MANY_CELLS = 'One range read or write reaches at most 100,000 cells. Use smaller ranges.';

const FIXED_LIMIT = 'This limit is fixed in this API.';

/** The hint of `range.formula-cell`, which a host check and a store check both give. */
const FORMULA_CELL =
  'A cell in the range holds a formula, or a result of an array formula or a data table. ' +
  'This API does not calculate formulas, so it does not read the values of such cells or ' +
  'write values to them. Read a range that holds no formula. In a write, give null for each ' +
  'cell that holds a formula, and write the cells around an array formula or data table range ' +
  'in ranges that do not meet it.';

/** The hint of each open diagnostic. An open diagnostic is its own host detail. */
const OPEN_HINTS: Readonly<Record<WorkbookOpenDiagnosticCode, string>> = {
  'open.zip-corrupt':
    'The file is not a valid ZIP archive, or a compressed entry is damaged. Ask for a fresh ' +
    'copy of the file.',
  'open.zip-invalid-entry-name':
    'The archive contains an entry with an unsafe path, such as ".." or a leading "/". The ' +
    'editor refuses it for safety.',
  'open.zip-entry-limit': `The archive has more entries than the open limit allows. ${FIXED_LIMIT}`,
  'open.zip-size-limit': `The archive expands to more than the total size limit. ${FIXED_LIMIT}`,
  'open.zip-ratio-limit':
    'The file expands to many times its compressed size, which is how a decompression bomb ' +
    `looks. ${FIXED_LIMIT}`,
  'open.xml-malformed':
    'An XML part of the workbook is not well formed. Opening and saving the file in a desktop ' +
    'editor usually repairs it.',
  'open.xml-dtd-forbidden':
    'An XML part declares a DTD or an entity. The editor refuses these to block XML entity ' +
    'attacks.',
  'open.xml-part-size-limit': `One XML part is larger than the per-part size limit. ${FIXED_LIMIT}`,
  'open.xml-depth-limit': `An XML part nests deeper than the depth limit. ${FIXED_LIMIT}`,
  'open.xml-element-limit': `An XML part has more elements than the element limit. ${FIXED_LIMIT}`,
  'open.invalid-limits': 'The limits of the reader are not valid numbers.',
  'open.content-types-invalid':
    'The file has no valid [Content_Types].xml part, so it is not a usable Office package. ' +
    'Save it again as .xlsx.',
  'open.missing-workbook':
    'The package has no workbook part. The file may be a different Office format, such as ' +
    '.docx or .pptx, renamed to .xlsx.',
  'open.failed':
    'The workbook could not be opened for a reason without a specific code. Report it with the ' +
    'engine version.',
};

/**
 * The public diagnostics, by the host detail that selects them. A detail that is not here gives
 * no `innerError`: an unknown engine reason never becomes public API.
 */
const DIAGNOSTICS: ReadonlyMap<string, XlsxEditorDiagnostic> = new Map(
  (
    [
      ['gone', 'edit.stale-target', 'The workbook no longer holds the worksheet. Read it again.'],
      [
        'name',
        'read.worksheet-not-found',
        'No worksheet has that name. Load items/name of the worksheet collection to list the ' +
          'names, or use getItemOrNullObject.',
      ],
      [
        'null-object',
        'edit.null-object',
        'The object is a null object: its lookup found nothing. Check isNullObject first.',
      ],
      [
        'no-active',
        'read.no-active-worksheet',
        'The workbook has no visible worksheet to be the active one.',
      ],
      ['formula-cell', 'range.formula-cell', FORMULA_CELL],
      [
        'cell-children',
        'range.cell-too-large',
        'The cell holds more than 1,000 elements, which this API does not read or write. The ' +
          'limit is fixed.',
      ],
      [
        'merged-cell',
        'write.merged-cell',
        'A merged range covers the cell and hides it, because the cell is not the top-left ' +
          'cell of the merged range. Write the value to the top-left cell, or give null or an ' +
          'empty string for this cell.',
      ],
      [
        'merged-ranges',
        'write.too-many-merged-ranges',
        'The worksheet has more than 100,000 merged ranges, so this API does not write its ' +
          'values. The limit is fixed.',
      ],
      [
        'formula-ranges',
        'range.formula-ranges-unknown',
        'The worksheet has an array formula or data table range that this API cannot read, or ' +
          'more than 10,000 of them, so any cell can hold a formula result. This API does not ' +
          'read or write the values of this worksheet.',
      ],
      [
        'read-after-write',
        'sync.read-after-write',
        'A load of values or number formats follows a write in the same sync. Call ' +
          'context.sync() after the writes, then load.',
      ],
      [
        'unreadable-value',
        'read.unreadable-value',
        'A cell in the range stores a value that this API cannot read, such as a value in ' +
          'an alternate form.',
      ],
      [
        'unreadable-format',
        'read.unreadable-number-format',
        'A cell in the range has a number format with no code that this API can give: a ' +
          'built-in format whose code depends on the locale (ECMA-376 Part 1, 18.8.30, ids 27 ' +
          'to 36 and 50 to 58), or a format that the workbook does not define. Read the values ' +
          'only.',
      ],
      [
        'unreadable-sheet',
        'read.unreadable-sheet',
        'The worksheet part of the file cannot be read.',
      ],
      [
        'address',
        'range.address-not-supported',
        'Use the address of one cell, such as A1, or of a rectangle, such as A1:B2, with no ' +
          'sheet name. Whole rows, whole columns, and names are outside the subset.',
      ],
      [
        'address-outside',
        'range.address-outside-grid',
        'The address names a cell outside the grid of 16,384 columns and 1,048,576 rows.',
      ],
      ['range-cells', 'range.too-many-cells', TOO_MANY_CELLS],
      ['range-too-large', 'range.too-many-cells', TOO_MANY_CELLS],
      [
        'whole-sheet',
        'range.whole-sheet',
        'getRange() with no address is the whole worksheet, which is outside the subset. Give ' +
          'an address such as A1:B2.',
      ],
      [
        'operations',
        'sync.too-many-operations',
        'One sync sends at most 10,000 operations. Sync more often.',
      ],
      [
        'batch-cells',
        'sync.too-many-cells',
        'The range reads and writes of one sync reach at most 1,000,000 cells in all. Sync ' +
          'more often.',
      ],
      [
        'merge-checks',
        'sync.too-many-merge-checks',
        'The value writes of one sync check at most 50,000,000 merged ranges in all, because ' +
          'each write checks each merged range of its worksheet. Write larger ranges in fewer ' +
          'writes, or sync after each part. The limit is fixed.',
      ],
      [
        'write-cells',
        'sync.too-many-written-cells',
        'The value writes of one sync reach at most 10,000 cells in all. Split a larger range ' +
          'into parts, and sync after each part.',
      ],
      [
        'batch-worksheets',
        'sync.too-many-worksheets',
        'The loads of the worksheet list of one sync give at most 100,000 worksheets in all. ' +
          'Load items once, and keep the proxies for the rest of the run.',
      ],
      [
        'shape',
        'write.array-shape',
        'Give one inner array for each row of the range, and one item for each column.',
      ],
      [
        'broadcast',
        'write.single-value',
        'A single value for a whole range is outside the subset. Give a two-dimensional array ' +
          'with one item for each cell.',
      ],
      [
        'formula',
        'write.formula-not-supported',
        'Text that starts with = is a formula, and this API does not write formulas. Write a ' +
          'number, a boolean, or other text.',
      ],
      [
        'typed-text',
        'write.text-reads-as-value',
        'The text reads as a number, a date, a boolean, or an error value, or it starts with ' +
          "+, -, or ', which a typed entry can read as a formula, a number, or a text mark. " +
          'This API stores text only as written, so it refuses such text. To store a number or ' +
          'a boolean, write it as a number or a boolean. Text that reads as a value, such as a ' +
          'code with a leading zero, cannot be written in this subset. In a write-back of ' +
          'loaded values, give null for each cell that you do not change.',
      ],
      [
        'load-tree',
        'load.property-tree-not-supported',
        'A property tree, such as { values: true }, is outside the subset. Name the properties ' +
          "instead, such as load('values') or load({ select: 'values' }).",
      ],
      [
        'load-option',
        'load.option-not-supported',
        'The load option expand is outside the subset. Name the properties with select ' +
          "instead, such as load({ select: 'items/name' }) on the worksheet collection.",
      ],
      ['text-too-long', 'write.text-too-long', 'A cell holds at most 32,767 characters of text.'],
      [
        'invalid-character',
        'write.invalid-character',
        'The text has a character that the file format cannot store, such as a control character.',
      ],
      [
        'locked',
        'edit.locked',
        'The worksheet protects the cell, so it cannot be changed. Remove the protection of ' +
          'the worksheet, or leave the locked cell out of the write.',
      ],
      ['shared-formula', 'range.formula-cell', FORMULA_CELL],
      [
        'cell-alternate-content',
        'edit.alternate-content',
        'The cell holds content in an alternate form that this API cannot rewrite.',
      ],
      [
        'unsupported-package',
        'edit.unsupported-content',
        'The workbook holds content that this API cannot change, such as a part that cannot be read.',
      ],
      ...Object.entries(OPEN_HINTS).map(([code, hint]) => [code, code, hint] as const),
    ] as const
  ).map(([detail, code, hint]) => [detail, Object.freeze({ code, hint })])
);

/** The public diagnostic of a host detail, or undefined when the detail is not public. */
export function diagnosticOf(detail: string | undefined): XlsxEditorDiagnostic | undefined {
  return detail === undefined ? undefined : DIAGNOSTICS.get(detail);
}

/**
 * The fields a {@link XlsxEditorError} is constructed from.
 *
 * @public
 */
export interface XlsxEditorErrorInit {
  /** Which refusal this is. The stable thing to branch on. */
  readonly code: XlsxEditorErrorCode;
  /** The reader limit that a workbook passed, for `ResourceLimitExceeded` at open. */
  readonly limit?: WorkbookOpenLimit;
  /** The consumer-facing path of the object or property involved. */
  readonly target?: string;
  /** The revision the context had read at, for `StaleDocument`. */
  readonly expectedRevision?: number;
  /** The revision the workbook was at, for `StaleDocument`. */
  readonly actualRevision?: number;
  /** The public reason, when the engine gave one this API knows. */
  readonly diagnostic?: XlsxEditorDiagnostic;
}

/**
 * Why the engine refused, as a public diagnostic code with its fixed hint.
 *
 * @public
 */
export interface XlsxEditorInnerError {
  /** A stable public diagnostic code, such as `range.address-not-supported`. */
  readonly code: string;
  /**
   * The fixed hint for the code: the cause, then the action. Never workbook content. A refusal of
   * the content of a cell ends it with the reference of that cell, such as `First refused cell:
   * B3.`
   */
  readonly message: string;
}

/**
 * Debug details in the shape of Office.js `OfficeExtension.DebugInfo`.
 *
 * @public
 */
export interface XlsxEditorDebugInfo {
  /** The same code as {@link XlsxEditorError.code}. */
  readonly code: XlsxEditorErrorCode;
  /** The same message as the error's own. */
  readonly message: string;
  /** The consumer-facing path of the object or property involved, when there is one. */
  readonly errorLocation?: string;
  /** Why the engine refused, when it knows a specific reason. */
  readonly innerError?: XlsxEditorInnerError;
}

/**
 * Every refusal this runtime throws.
 *
 * Branch on {@link XlsxEditorError.code}, never on the message. Codes are stable public API.
 * Nothing from the engine or the file appears in the message: a consumer gets a stable code, a
 * fixed sentence, and `target`, the consumer-facing path of the call.
 *
 * @example
 * ```ts
 * try {
 *   await context.sync();
 * } catch (error) {
 *   if (isXlsxEditorError(error) && error.code === 'StaleDocument') {
 *     // Read again: another writer changed the workbook first.
 *   }
 * }
 * ```
 *
 * @public
 */
export class XlsxEditorError extends Error {
  /** Which refusal this is. Stable across versions; the thing to branch on. */
  readonly code: XlsxEditorErrorCode;
  /** For `ResourceLimitExceeded` at open: the reader limit that the workbook passed. */
  readonly limit?: WorkbookOpenLimit;
  /** The consumer-facing path of the object or property involved, when there is one. */
  readonly target?: string;
  /** For `StaleDocument`: the revision the context had last read at. */
  readonly expectedRevision?: number;
  /** For `StaleDocument`: the revision the workbook was at. */
  readonly actualRevision?: number;
  /** Debug details, as Office.js `OfficeExtension.Error.debugInfo` has them. */
  readonly debugInfo: XlsxEditorDebugInfo;

  constructor(init: XlsxEditorErrorInit) {
    const base =
      init.code === 'InvalidArgument' && init.diagnostic?.code === ARGUMENT_NOT_ACCEPTED
        ? `${MESSAGES[init.code]} ${init.diagnostic.hint}`
        : MESSAGES[init.code];
    const message = init.target ? `${base} (${init.target})` : base;
    super(message);
    this.name = 'XlsxEditorError';
    this.code = init.code;
    this.debugInfo = Object.freeze({
      code: init.code,
      message,
      ...(init.target === undefined ? {} : { errorLocation: init.target }),
      ...(init.diagnostic === undefined
        ? {}
        : {
            innerError: Object.freeze({
              code: init.diagnostic.code,
              message: init.diagnostic.hint,
            }),
          }),
    });
    if (init.limit !== undefined) this.limit = init.limit;
    if (init.target !== undefined) this.target = init.target;
    if (init.expectedRevision !== undefined) this.expectedRevision = init.expectedRevision;
    if (init.actualRevision !== undefined) this.actualRevision = init.actualRevision;
  }
}

/**
 * Whether a caught value is one of ours. By `name` as well as by `instanceof`: a consumer can
 * end up with two copies of this module, and `instanceof` fails across them.
 *
 * @public
 */
export function isXlsxEditorError(value: unknown): value is XlsxEditorError {
  if (value instanceof XlsxEditorError) return true;
  return (
    typeof value === 'object' &&
    value !== null &&
    (value as { name?: unknown }).name === 'XlsxEditorError' &&
    typeof (value as { code?: unknown }).code === 'string'
  );
}

/** The message this runtime uses for a code, for tests. */
export function messageFor(code: XlsxEditorErrorCode): string {
  return MESSAGES[code];
}

export function fail(init: XlsxEditorErrorInit): never {
  throw new XlsxEditorError(init);
}

/** A cell reference that a hint can name: column letters and a row number, nothing else. */
const CELL_REFERENCE = /^[A-Z]{1,3}[1-9][0-9]{0,6}$/;

/**
 * A refusal with the public diagnostic of a host detail, when it has one. When the host names
 * the cell that the refusal is about, the hint ends with it, such as `First refused cell: B3.`
 */
export function refusal(
  code: XlsxEditorErrorCode,
  target: string | undefined,
  detail?: string,
  cell?: string
): XlsxEditorError {
  const known = diagnosticOf(detail);
  const diagnostic =
    known && cell !== undefined && CELL_REFERENCE.test(cell)
      ? { code: known.code, hint: `${known.hint} First refused cell: ${cell}.` }
      : known;
  return new XlsxEditorError({
    code,
    ...(target === undefined ? {} : { target }),
    ...(diagnostic ? { diagnostic } : {}),
  });
}

/**
 * Refuse an argument with `InvalidArgument`, naming what the API accepts there, such as
 * `a string`. The message and the diagnostic `argument.not-accepted` both carry it.
 */
/**
 * An `InvalidArgument` refusal that names what the call accepts. When the refused argument is the
 * value of one cell, the hint ends with that cell, such as `First refused cell: B3.`
 */
export function argumentRefusal(
  target: string | undefined,
  accepted: string,
  cell?: string
): XlsxEditorError {
  const where =
    cell !== undefined && CELL_REFERENCE.test(cell) ? ` First refused cell: ${cell}.` : '';
  return new XlsxEditorError({
    code: 'InvalidArgument',
    ...(target === undefined ? {} : { target }),
    diagnostic: Object.freeze({ code: ARGUMENT_NOT_ACCEPTED, hint: `Use ${accepted}.${where}` }),
  });
}

export function failArgument(target: string | undefined, accepted: string): never {
  throw argumentRefusal(target, accepted);
}

/** Names as a list for a hint: `'a'`, `'a' or 'b'`, or `'a', 'b', or 'c'`. */
export function listOf(names: readonly string[]): string {
  const quoted = names.map((n) => `'${n}'`);
  if (quoted.length <= 2) return quoted.join(' or ');
  return `${quoted.slice(0, -1).join(', ')}, or ${quoted.at(-1)}`;
}
