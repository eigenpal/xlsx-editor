/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See LICENSE.md in this package.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Every refusal reason that the core host can give reaches the caller with its public code and
// its diagnostic. After PPTX Editor's
// `packages/editor-api/scripts/__tests__/refusal-reasons.test.ts` at d4502552, which scans the
// core for its refusal reasons and checks that each one maps to a diagnostic. Differences: the
// reasons of this host are the literal details of `refuseWith` and `fail` in
// `core/src/automation/host.ts`, so the scan reads that one file; a store refusal reason is
// listed by its type (`OpRejection['code']`), so a new store reason fails the type check here;
// and each reason is checked against its exact public code and diagnostic code, so a swap of two
// diagnostics fails too.

import { describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import {
  MAX_BATCH_MERGE_CHECKS,
  MAX_CELL_CHILDREN,
  type AutomationErrorCode,
} from '@xlsx-editor.dev/core/automation';
import type { OpRejection } from '@xlsx-editor.dev/core/store';
import { isXlsxEditorError, type XlsxEditorErrorCode } from '../index.ts';
import { hostRefusal } from '../runtime/request-context.ts';

type Expected = readonly [XlsxEditorErrorCode, string | undefined];

/**
 * Each host reason, by `code detail`, as its public code and diagnostic code. A reason with no
 * diagnostic is a protocol error that the runtime never sends: a malformed batch or selector.
 */
const HOST: Readonly<Record<string, Expected>> = {
  'item-not-found gone': ['InvalidObjectPath', 'edit.stale-target'],
  'item-not-found name': ['ItemNotFound', 'read.worksheet-not-found'],
  'item-not-found no-active': ['ItemNotFound', 'read.no-active-worksheet'],
  'invalid-argument sheet': ['InvalidArgument', undefined],
  'invalid-argument address': ['InvalidArgument', 'range.address-outside-grid'],
  'unsupported address': ['NotSupported', 'range.address-not-supported'],
  'resource-limit range-cells': ['ResourceLimitExceeded', 'range.too-many-cells'],
  'unsupported unreadable-sheet': ['NotSupported', 'read.unreadable-sheet'],
  'unsupported formula-ranges': ['NotSupported', 'range.formula-ranges-unknown'],
  'unsupported formula-cell': ['NotSupported', 'range.formula-cell'],
  'unsupported merged-cell': ['NotSupported', 'write.merged-cell'],
  'resource-limit cell-children': ['ResourceLimitExceeded', 'range.cell-too-large'],
  'resource-limit merged-ranges': ['ResourceLimitExceeded', 'write.too-many-merged-ranges'],
  'resource-limit merge-checks': ['ResourceLimitExceeded', 'sync.too-many-merge-checks'],
  'unsupported unreadable-value': ['NotSupported', 'read.unreadable-value'],
  'unsupported unreadable-format': ['NotSupported', 'read.unreadable-number-format'],
  'invalid-argument values': ['InvalidArgument', 'write.array-shape'],
  'transaction-refused text-too-long': ['InvalidArgument', 'write.text-too-long'],
  'transaction-refused locked': ['NotSupported', 'edit.locked'],
  'resource-limit batch-cells': ['ResourceLimitExceeded', 'sync.too-many-cells'],
  'resource-limit write-cells': ['ResourceLimitExceeded', 'sync.too-many-written-cells'],
  'resource-limit batch-worksheets': ['ResourceLimitExceeded', 'sync.too-many-worksheets'],
  'unsupported read-after-write': ['NotSupported', 'sync.read-after-write'],
  'unknown-operation ': ['GeneralException', undefined],
  'disposed ': ['RuntimeDisposed', undefined],
  'invalid-argument operations-list': ['InvalidArgument', undefined],
  'resource-limit operations': ['ResourceLimitExceeded', 'sync.too-many-operations'],
  'stale-revision ': ['StaleDocument', undefined],
};

/**
 * Each store refusal reason (`transaction-refused`), as its public code and diagnostic code. The
 * reasons of edits that this API never sends, such as a merge or a format patch, stay generic.
 */
const STORE: Readonly<Record<OpRejection['code'], Expected>> = {
  'text-too-long': ['InvalidArgument', 'write.text-too-long'],
  'invalid-character': ['InvalidArgument', 'write.invalid-character'],
  locked: ['NotSupported', 'edit.locked'],
  'shared-formula': ['NotSupported', 'range.formula-cell'],
  'cell-alternate-content': ['NotSupported', 'edit.alternate-content'],
  'unsupported-package': ['NotSupported', 'edit.unsupported-content'],
  'range-too-large': ['ResourceLimitExceeded', 'range.too-many-cells'],
  'stale-revision': ['StaleDocument', undefined],
  'unknown-sheet': ['GeneralException', undefined],
  'bad-reference': ['GeneralException', undefined],
  'bad-value': ['GeneralException', undefined],
  'bad-size': ['GeneralException', undefined],
  'merge-one-cell': ['GeneralException', undefined],
  'merge-overlap': ['GeneralException', undefined],
  'merge-drops-values': ['GeneralException', undefined],
  'bad-position': ['GeneralException', undefined],
  'invalid-sheet': ['GeneralException', undefined],
  'unreadable-formula': ['GeneralException', undefined],
  'formula-syntax': ['GeneralException', undefined],
  'r1c1-formula': ['GeneralException', undefined],
  'unsupported-function': ['GeneralException', undefined],
  'unknown-formula-reference': ['GeneralException', undefined],
  'bad-format': ['GeneralException', undefined],
  'unreadable-number-format': ['GeneralException', undefined],
  'too-many-rows': ['GeneralException', undefined],
  'resize-anchor': ['GeneralException', undefined],
  'clear-too-large': ['GeneralException', undefined],
  'style-limit': ['GeneralException', undefined],
  'invalid-part': ['GeneralException', undefined],
  'grid-full': ['GeneralException', undefined],
  'last-sheet': ['GeneralException', undefined],
  'nothing-to-redo': ['GeneralException', undefined],
  'nothing-to-undo': ['GeneralException', undefined],
  'sheet-name-ambiguous': ['GeneralException', undefined],
  'sheet-name-invalid': ['GeneralException', undefined],
  'sheet-name-taken': ['GeneralException', undefined],
  'sheet-name-too-long': ['GeneralException', undefined],
  'sheet-range-blocked': ['GeneralException', undefined],
  'structure-blocked': ['GeneralException', undefined],
};

/** The `code detail` of each refusal in the host source. A dynamic code is read as both codes. */
function hostReasons(): string[] {
  const source = readFileSync(
    join(import.meta.dir, '../../../core/src/automation/host.ts'),
    'utf8'
  );
  const found = new Set<string>();
  const calls = [
    /\brefuseWith\(\s*([\w.'-]+),\s*'[^']*',?\s*(?:([\w.'-]+),?\s*)?(?:\w+,?\s*)?\)/g,
    /\bfail\(\s*[\w.[\]!]+,\s*('[\w-]+'),\s*'[^']*',?\s*(?:([\w.'-]+),?\s*)?\)/g,
  ];
  for (const call of calls)
    for (const [, code, detail = "''"] of source.matchAll(call)) {
      // A detail that is not a literal is a store reason: `STORE` lists each one.
      if (!detail.startsWith("'")) continue;
      const codes = code!.startsWith("'")
        ? [code!.slice(1, -1)]
        : code === 'parsed.code'
          ? ['invalid-argument', 'unsupported']
          : [code!];
      for (const each of codes) found.add(`${each} ${detail.slice(1, -1)}`);
    }
  return [...found].sort();
}

function check(code: AutomationErrorCode, detail: string | undefined): Expected {
  const error = hostRefusal({ code, message: 'm', ...(detail ? { detail } : {}) }, 'target');
  return [error.code, error.debugInfo.innerError?.code];
}

describe('refusal reasons', () => {
  test('the scan finds each refusal of the host, and the table lists exactly those', () => {
    const reasons = hostReasons();
    expect(reasons).toContain('item-not-found name');
    expect(reasons).toContain('unsupported address');
    expect(reasons).toContain('resource-limit operations');
    expect(reasons).toEqual(Object.keys(HOST).sort());
  });

  test('each host reason has its public code and diagnostic', () => {
    for (const [key, expected] of Object.entries(HOST)) {
      const [code, detail] = key.split(' ') as [AutomationErrorCode, string];
      expect([key, ...check(code, detail || undefined)]).toEqual([key, ...expected]);
    }
  });

  test('each store reason has its public code and diagnostic', () => {
    for (const [detail, expected] of Object.entries(STORE))
      expect([detail, ...check('transaction-refused', detail)]).toEqual([detail, ...expected]);
  });

  test('a limit hint states the number of its limit', () => {
    const hintOf = (detail: string) =>
      hostRefusal({ code: 'resource-limit', message: 'm', detail }, 't').debugInfo.innerError
        ?.message;
    const number = (n: number) => n.toLocaleString('en-US');
    expect(hintOf('cell-children')).toContain(`more than ${number(MAX_CELL_CHILDREN)} elements`);
    // The core limit test pins `MAX_MERGES`, the layout limit, to 100,000.
    expect(hintOf('merged-ranges')).toContain('more than 100,000 merged ranges');
    expect(hintOf('merge-checks')).toContain(`${number(MAX_BATCH_MERGE_CHECKS)} merged ranges`);
  });

  test('the hint ends with the cell that the host names, when it is a cell reference', () => {
    const hintOf = (detail: string, cell?: string) =>
      hostRefusal(
        { code: 'unsupported', message: 'm', detail, ...(cell === undefined ? {} : { cell }) },
        't'
      ).debugInfo.innerError?.message;
    const plain = hintOf('formula-cell')!;
    expect(hintOf('formula-cell', 'B3')).toBe(`${plain} First refused cell: B3.`);
    expect(hintOf('formula-cell', 'XFD1048576')).toBe(`${plain} First refused cell: XFD1048576.`);
    // Anything that is not a plain cell reference stays out of the hint.
    for (const cell of ['b3', 'B0', 'B03', 'AAAA1', 'B12345678', 'A1:B2', '$B$3', ' B3', 'B3 ', ''])
      expect([cell, hintOf('formula-cell', cell)]).toEqual([cell, plain]);
    // A refusal with no diagnostic gets no hint, with a cell or not.
    const bare = hostRefusal({ code: 'disposed', message: 'm', cell: 'B3' }, 't');
    expect(bare.debugInfo.innerError).toBeUndefined();
    expect(bare.message).not.toContain('B3');
  });

  test('the error guard takes an error of this API, and no other error', () => {
    expect(isXlsxEditorError(hostRefusal({ code: 'disposed', message: 'm' }, 't'))).toBe(true);
    // An error from another copy of the package keeps its name and its code.
    expect(isXlsxEditorError({ name: 'XlsxEditorError', code: 'ItemNotFound' })).toBe(true);
    const foreign = Object.assign(new Error('m'), { code: 'ItemNotFound' });
    for (const value of [foreign, { name: 'XlsxEditorError' }, { code: 'x' }, null, 'm'])
      expect([value, isXlsxEditorError(value)]).toEqual([value, false]);
  });

  test('each diagnostic has a hint that says what to do, and one hint for each code', () => {
    const hints = new Map<string, string>();
    for (const [key] of Object.entries(HOST)) {
      const [code, detail] = key.split(' ') as [AutomationErrorCode, string];
      const inner = hostRefusal({ code, message: 'm', detail }, 't').debugInfo.innerError;
      if (inner) hints.set(`${inner.code}\n${inner.message}`, key);
    }
    for (const detail of Object.keys(STORE)) {
      const inner = hostRefusal({ code: 'transaction-refused', message: 'm', detail }, 't')
        .debugInfo.innerError;
      if (inner) hints.set(`${inner.code}\n${inner.message}`, detail);
    }
    const byCode = new Map<string, string[]>();
    for (const key of hints.keys()) {
      const [code, hint] = key.split('\n') as [string, string];
      expect([code, hint.length >= 20]).toEqual([code, true]);
      byCode.set(code, [...(byCode.get(code) ?? []), hint]);
    }
    const twice = [...byCode].filter(([, list]) => list.length > 1).map(([code]) => code);
    expect(twice).toEqual([]);
  });
});
