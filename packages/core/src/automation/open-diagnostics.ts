// The public reason of a refused open, after DOCX Editor's `OPEN_CODES` and `openDiagnosticCode`
// (`packages/core/src/automation/diagnostics.ts`, codes in `diagnostic-vocabulary.ts`, at
// 6c11927b), and the reader limit that it passed, after the `limit` names of DOCX Editor's
// `createServer` (`packages/editor-api/src/runtime/server.ts` at 6c11927b). Differences: the XLSX
// reader gives a zip or an XML reason beside its layer, so there is one table for each layer (the
// two layers share `bad-name`); a package with no workbook part is `open.missing-workbook`, not
// `open.missing-main-document`; and there are no relationship, part count, or unsupported source
// codes, because this reader has no such refusal. The editor's failure `limit` uses the same
// table (`editor/load-steps.ts`).
//
// Only the RESULT of a lookup leaves the engine. A reader `detail` can name a part of the file,
// so it never becomes part of the answer.

import type { ZipRejection } from '@eigenpal/ooxml-core';
import type { PackageRejection } from '../store/package/package.ts';
import type { XmlPartRejection } from '../store/package/tree.ts';

/**
 * Why a workbook could not be opened, as a stable public code.
 *
 * @public
 */
export type WorkbookOpenDiagnosticCode =
  | 'open.zip-corrupt'
  | 'open.zip-invalid-entry-name'
  | 'open.zip-entry-limit'
  | 'open.zip-size-limit'
  | 'open.zip-ratio-limit'
  | 'open.xml-malformed'
  | 'open.xml-dtd-forbidden'
  | 'open.xml-part-size-limit'
  | 'open.xml-depth-limit'
  | 'open.xml-element-limit'
  | 'open.sheet-name-limit'
  | 'open.invalid-limits'
  | 'open.content-types-invalid'
  | 'open.missing-workbook'
  | 'open.failed';

/**
 * The reader limit that a refused open passed, with the names of DOCX Editor's `limit`.
 *
 * @public
 */
export type WorkbookOpenLimit =
  | 'zip.maxEntries'
  | 'zip.maxTotalBytes'
  | 'zip.maxRatio'
  | 'xml.maxBytes'
  | 'xml.maxElements'
  | 'xml.maxDepth'
  | 'workbook.maxSheetNameLength';

/**
 * Every zip and XML reason of the reader, by type: a new reader reason fails the type check here
 * until it has a public code. `too-large` is resolved by its limit, as in DOCX Editor.
 */
const ZIP_CODES: Readonly<Record<Exclude<ZipRejection, 'too-large'>, WorkbookOpenDiagnosticCode>> =
  {
    'inflate-error': 'open.zip-corrupt',
    'overlapping-entries': 'open.zip-corrupt',
    'bad-name': 'open.zip-invalid-entry-name',
    'too-many-entries': 'open.zip-entry-limit',
    'invalid-limits': 'open.invalid-limits',
  };

const XML_CODES: Readonly<
  Record<Exclude<XmlPartRejection, 'too-large'>, WorkbookOpenDiagnosticCode>
> = {
  'parse-error': 'open.xml-malformed',
  'no-root': 'open.xml-malformed',
  'unbound-prefix': 'open.xml-malformed',
  'bad-name': 'open.xml-malformed',
  'dtd-forbidden': 'open.xml-dtd-forbidden',
  'entity-forbidden': 'open.xml-dtd-forbidden',
  'too-deep': 'open.xml-depth-limit',
  'too-many-elements': 'open.xml-element-limit',
  'invalid-limits': 'open.invalid-limits',
};

type XmlPartLimit = 'xml.maxBytes' | 'xml.maxElements' | 'xml.maxDepth';

/** The XML part limits, by the reason that the bounded reader gives. */
const XML_LIMITS: Readonly<Record<string, XmlPartLimit>> = {
  'too-large': 'xml.maxBytes',
  'too-many-elements': 'xml.maxElements',
  'too-deep': 'xml.maxDepth',
};

function own<T>(table: object, key: string): T | undefined {
  return Object.hasOwn(table, key) ? (table as Record<string, T>)[key] : undefined;
}

/** The XML part limit of a reason of the bounded reader, or undefined when it is no limit. */
export function xmlLimitOf(reason: string): XmlPartLimit | undefined {
  return own<XmlPartLimit>(XML_LIMITS, reason);
}

/** The reader limit that a refused open passed, or undefined when it passed none. */
export function openLimitOf(error: PackageRejection): WorkbookOpenLimit | undefined {
  if (error.reason === 'xml') return xmlLimitOf(error.detail);
  if (error.reason === 'workbook') return error.limit;
  if (error.reason !== 'zip') return undefined;
  return error.detail === 'too-many-entries' ? 'zip.maxEntries' : error.limit;
}

/**
 * The public code of a refused open, and the reader limit that it passed, if any. `too-large` is
 * the one reason two layers share: the archive reports it WITH the limit it passed, a single XML
 * part without one.
 */
export function openDiagnosticOf(error: PackageRejection): {
  readonly code: WorkbookOpenDiagnosticCode;
  readonly limit?: WorkbookOpenLimit;
} {
  const limit = openLimitOf(error);
  return { code: codeOf(error, limit), ...(limit === undefined ? {} : { limit }) };
}

function codeOf(
  error: PackageRejection,
  limit: WorkbookOpenLimit | undefined
): WorkbookOpenDiagnosticCode {
  switch (error.reason) {
    case 'no-content-types':
      return 'open.content-types-invalid';
    case 'not-a-workbook':
      return 'open.missing-workbook';
    case 'workbook':
      return 'open.sheet-name-limit';
    case 'zip':
    case 'xml':
      if (error.detail === 'too-large') {
        if (limit === 'zip.maxRatio') return 'open.zip-ratio-limit';
        return limit === 'zip.maxTotalBytes' ? 'open.zip-size-limit' : 'open.xml-part-size-limit';
      }
      return own(error.reason === 'zip' ? ZIP_CODES : XML_CODES, error.detail) ?? 'open.failed';
  }
}
