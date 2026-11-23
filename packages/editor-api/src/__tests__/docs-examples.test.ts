/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See LICENSE.md in this package.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The examples of `packages/editor-api/README.md` and `docs/site/content/editor-api/index.mdx`,
// run as written against a synthetic workbook, after PPTX Editor's
// `packages/editor-api/src/__tests__/docs-examples.test.ts` at d4502552. File reads and writes
// become bytes in memory. Each test also checks that its example is still in its page.

import { expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { buildXlsx, c, row } from '../../../../scripts/fixtures/builder.ts';
import { XlsxEditor, type XlsxEditorServerRuntime } from '../index.ts';

const README = readFileSync(new URL('../../README.md', import.meta.url), 'utf8');
const PAGE = readFileSync(
  new URL('../../../../docs/site/content/editor-api/index.mdx', import.meta.url),
  'utf8'
);

const book = () =>
  buildXlsx({
    sheets: [
      {
        name: 'Data',
        rows: row(2, [c('A2', 'North'), c('B2', 10)]) + row(3, [c('A3', 'South')]),
      },
    ],
  });

async function valuesOf(bytes: Uint8Array, address: string): Promise<unknown> {
  const runtime: XlsxEditorServerRuntime = await XlsxEditor.createServer(bytes);
  try {
    return await runtime.run(async (context) => {
      const range = context.workbook.worksheets.getItem('Data').getRange(address).load('values');
      await context.sync();
      return range.values;
    });
  } finally {
    runtime.dispose();
  }
}

test('the README example reads two cells, writes a header row, and saves', async () => {
  for (const line of [
    "const units = sheet.getRange('B2:B3').load('values');",
    "sheet.getRange('A1:C1').values = [['Region', 'Units', 'Checked']];",
    "const checked = units.values.map(([value]) => [value !== '']);",
    "sheet.getRange('C2:C3').values = checked;",
  ])
    expect(README).toContain(line);
  let saved: Uint8Array | undefined;
  const runtime = await XlsxEditor.createServer(book());
  try {
    await runtime.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem('Data');
      const units = sheet.getRange('B2:B3').load('values');
      await context.sync();
      sheet.getRange('A1:C1').values = [['Region', 'Units', 'Checked']];
      const checked = units.values.map(([value]) => [value !== '']);
      sheet.getRange('C2:C3').values = checked;
      await context.sync();
    });
    saved = await runtime.save();
  } finally {
    runtime.dispose();
  }
  expect(await valuesOf(saved, 'A1:C3')).toEqual([
    ['Region', 'Units', 'Checked'],
    ['North', 10, true],
    ['South', '', false],
  ]);
});

test('the guide example writes a header row and saves', async () => {
  expect(PAGE).toContain("sheet.getRange('A1:C1').values = [['Region', 'Units', 'Share']];");
  let saved: Uint8Array | undefined;
  const runtime = await XlsxEditor.createServer(book());
  try {
    await runtime.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem('Data');
      sheet.getRange('A1:C1').values = [['Region', 'Units', 'Share']];
      await context.sync();
    });
    saved = await runtime.save();
  } finally {
    runtime.dispose();
  }
  expect(await valuesOf(saved, 'A1:C1')).toEqual([['Region', 'Units', 'Share']]);
});
