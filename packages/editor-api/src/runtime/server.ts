/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See LICENSE.md in this package.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// After PPTX Editor's `packages/editor-api/src/runtime/server.ts` at d4502552, which follows DOCX
// Editor's `packages/editor-api/src/runtime/server.ts` at 5993bd223945. Differences: no fonts and
// no pictures, because the workbook subset measures nothing; core opens the bytes. A refused open
// gives its diagnostic and `limit`, as DOCX Editor's `createServer` at 6c11927b does; there are
// no `limits`, `modules`, or `onDiagnostic` options, so each limit is fixed.

import { openAutomationHost } from '@xlsx-editor.dev/core/automation';
import { diagnosticOf, fail, failArgument } from './errors.ts';
import { createRuntime, type XlsxEditorServerRuntime } from './runtime.ts';

/**
 * A runtime over XLSX bytes. Opening is a bounded parse: zip entry, size, and ratio caps, part
 * name checks, and XML with no DTD. Malformed input refuses with `InvalidArgument`; input over
 * a cap refuses with `ResourceLimitExceeded`, and `limit` names the cap, such as
 * `zip.maxEntries`. Each refusal gives an open diagnostic, such as `open.xml-malformed`, in
 * `debugInfo.innerError`. No message carries file content.
 */
export async function createServer(
  bytes: Uint8Array | ArrayBuffer
): Promise<XlsxEditorServerRuntime> {
  // The tag check also accepts the bytes of another window, which `instanceof` misses.
  const tag = Object.prototype.toString.call(bytes);
  if ((tag !== '[object Uint8Array]' && tag !== '[object ArrayBuffer]') || bytes.byteLength === 0)
    failArgument(
      'createServer',
      'the bytes of an XLSX file as a non-empty Uint8Array or ArrayBuffer'
    );
  // The runtime keeps no reference to the caller's array.
  const opened = openAutomationHost(bytes.slice());
  if (opened.ok) return createRuntime(opened.host);
  const code = opened.code === 'resource-limit' ? 'ResourceLimitExceeded' : 'InvalidArgument';
  const diagnostic = diagnosticOf(opened.reason);
  return fail({
    code,
    target: 'createServer',
    ...(opened.limit === undefined ? {} : { limit: opened.limit }),
    ...(diagnostic ? { diagnostic } : {}),
  });
}
