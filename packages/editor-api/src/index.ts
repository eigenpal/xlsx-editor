/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See LICENSE.md in this package.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// After PPTX Editor's `packages/editor-api/src/index.ts` at d4502552. Differences: the workbook
// subset of the object model, and the XLSX engine under it.
/**
 * Workbook automation for XLSX: a batching object model over the XLSX Editor engine.
 *
 * Work is described against proxy objects, and nothing reaches the workbook until
 * `context.sync()`, which sends the queued calls as one atomic batch: all of it applies, or none
 * of it. Reading a property that no `load` asked for is an error, not a silent `undefined`.
 *
 * This entry needs no browser. It opens XLSX bytes, drives them, and saves them back, so a
 * server, a worker, or a build script can use the whole API.
 *
 * @example Edit a workbook on a server
 * ```ts
 * import { XlsxEditor } from '@xlsx-editor.dev/editor-api';
 *
 * const runtime = await XlsxEditor.createServer(bytes);
 * try {
 *   await runtime.run(async (context) => {
 *     const sheet = context.workbook.worksheets.getItem('Sheet1');
 *     sheet.getRange('A1:B1').values = [['Region', 'Total']];
 *     await context.sync();
 *   });
 *   const edited = await runtime.save();
 * } finally {
 *   runtime.dispose();
 * }
 * ```
 *
 * The API shape follows a documented subset of the Office.js workbook object model. It is not
 * Office.js, it does not run in an add-in host, and every type in it is authored here.
 *
 * @packageDocumentation
 */

import { createServer } from './runtime/server.ts';
import type { XlsxEditorServerRuntime } from './runtime/runtime.ts';

export * from './runtime/public.ts';

/**
 * The entry point. A runtime over an editor that is open in a page is outside the subset.
 *
 * @public
 */
export interface XlsxEditorNamespace {
  /**
   * A runtime over XLSX bytes. It also offers `save()`. Malformed bytes refuse with
   * `InvalidArgument`, and bytes over a reader limit with `ResourceLimitExceeded`.
   */
  createServer(bytes: Uint8Array | ArrayBuffer): Promise<XlsxEditorServerRuntime>;
}

/**
 * The entry point: open a workbook and get a runtime to work against.
 *
 * @example
 * ```ts
 * const runtime = await XlsxEditor.createServer(bytes);
 * const names = await runtime.run(async (context) => {
 *   const sheets = context.workbook.worksheets.load('items/name');
 *   await context.sync();
 *   return sheets.items.map((sheet) => sheet.name);
 * });
 * ```
 *
 * @public
 */
export const XlsxEditor: XlsxEditorNamespace = Object.freeze({ createServer });
