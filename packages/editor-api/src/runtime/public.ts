/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See LICENSE.md in this package.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The public surface of the package entry: the error class and its types.

export {
  isXlsxEditorError,
  XlsxEditorError,
  type XlsxEditorDebugInfo,
  type XlsxEditorDiagnostic,
  type XlsxEditorErrorCode,
  type XlsxEditorErrorInit,
  type XlsxEditorInnerError,
} from './errors.ts';
export type { LoadOption } from './load-options.ts';
export { ClientObject, RequestContext, type ClientRequestContext } from './request-context.ts';
export type {
  RunCallback,
  WorkbookCapabilities,
  XlsxEditorRuntime,
  XlsxEditorServerRuntime,
} from './runtime.ts';
export { Range } from '../model/range.ts';
export { Workbook, Worksheet, WorksheetCollection } from '../model/workbook.ts';
