/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See LICENSE.md in this package.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The workbooks and helpers of the runtime tests. Each workbook is built here, at test time.

import { expect } from 'bun:test';
import { buildXlsx, c, row } from '../../../../scripts/fixtures/builder.ts';
import { isXlsxEditorError, type XlsxEditorError, type XlsxEditorErrorCode } from '../index.ts';

/**
 * Three worksheets: `Data` with each kind of stored value, `My Sheet` (a name that needs
 * quotes), and the hidden `Hidden`. `protectData` protects every cell of `Data`.
 */
export function sampleWorkbook(
  options: { calcPr?: string; protectData?: boolean } = {}
): Uint8Array {
  return buildXlsx({
    strings: ['shared text'],
    sheets: [
      {
        name: 'Data',
        rows:
          row(1, [
            c('A1', 42),
            c('B1', 0, { t: 's' }),
            c('C1', 'inline text'),
            c('D1', true),
            c('E1', '#DIV/0!', { t: 'e' }),
          ]) +
          row(2, ['<c r="A2" s="1"><v>0.25</v></c>', c('B2', 12, { f: 'A1/3.5' })]) +
          row(3, ['<c r="A3" s="2"><v>45000</v></c>']),
        ...(options.protectData ? { beforeMerges: '<sheetProtection sheet="1"/>' } : {}),
      },
      { name: 'My Sheet', rows: row(1, [c('A1', 'x')]) },
      { name: 'Hidden', state: 'hidden', rows: '' },
    ],
    styles: {
      cellXfs: [
        '<xf numFmtId="10" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>',
        '<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>',
      ],
    },
    ...(options.calcPr ? { calcPr: options.calcPr } : {}),
  });
}

/**
 * Await `promise` and expect it to refuse with `code` (and the diagnostic, when given). Gives the
 * error.
 */
export async function expectRefusal(
  promise: Promise<unknown>,
  code: XlsxEditorErrorCode,
  diagnostic?: string
): Promise<XlsxEditorError> {
  let caught: unknown;
  try {
    await promise;
  } catch (error) {
    caught = error;
  }
  expect(isXlsxEditorError(caught)).toBe(true);
  const error = caught as { code: string; debugInfo: { innerError?: { code: string } } };
  expect(error.code).toBe(code);
  if (diagnostic !== undefined) expect(error.debugInfo.innerError?.code).toBe(diagnostic);
  return caught as XlsxEditorError;
}

/** Expect that no text of an error (message, target, and debug info) holds any of `texts`. */
export function expectNoEcho(error: XlsxEditorError, texts: readonly string[]): void {
  const all = JSON.stringify([error.message, error.target, error.debugInfo]);
  for (const text of texts) expect(all).not.toContain(text);
}

/** Call `fn` and expect it to throw `code` at once. */
export function expectThrows(fn: () => unknown, code: XlsxEditorErrorCode): void {
  let caught: unknown;
  try {
    fn();
  } catch (error) {
    caught = error;
  }
  expect(isXlsxEditorError(caught)).toBe(true);
  expect((caught as { code: string }).code).toBe(code);
}
