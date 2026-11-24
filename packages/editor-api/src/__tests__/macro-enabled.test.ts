/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See LICENSE.md in this package.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// A macro-enabled workbook through the server runtime: it opens, a value write saves, and the
// saved package keeps the macro-enabled main-part content type and the macro part, which is
// never run (the Media Types stream, `[Content_Types].xml`, Part 2, 7.2.3).

import { expect, test } from 'bun:test';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { buildXlsx, c, row } from '../../../../scripts/fixtures/builder.ts';
import { XlsxEditor } from '../index.ts';

const SML = 'application/vnd.openxmlformats-officedocument.spreadsheetml.';
const MACRO_WORKBOOK = 'application/vnd.ms-excel.sheet.macroEnabled.main+xml';
const VBA_REL = 'http://schemas.microsoft.com/office/2006/relationships/vbaProject';
/** Bytes that stand for a macro part. Nothing reads or runs them. */
const MACRO_BYTES = new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1, 1, 2, 3]);

function macroWorkbook(): Uint8Array {
  const files = unzipSync(buildXlsx({ sheets: [{ name: 'Data', rows: row(1, [c('A1', 1)]) }] }));
  files['[Content_Types].xml'] = strToU8(
    strFromU8(files['[Content_Types].xml']!)
      .replace(`${SML}sheet.main+xml`, MACRO_WORKBOOK)
      .replace(
        '</Types>',
        '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/></Types>'
      )
  );
  files['xl/_rels/workbook.xml.rels'] = strToU8(
    strFromU8(files['xl/_rels/workbook.xml.rels']!).replace(
      '</Relationships>',
      `<Relationship Id="rIdMacros" Type="${VBA_REL}" Target="vbaProject.bin"/></Relationships>`
    )
  );
  files['xl/vbaProject.bin'] = MACRO_BYTES;
  return zipSync(files);
}

test('a save keeps the macro-enabled content type and the macro part', async () => {
  const runtime = await XlsxEditor.createServer(macroWorkbook());
  try {
    await runtime.run(async (context) => {
      context.workbook.worksheets.getItem('Data').getRange('B2').values = [[7]];
      await context.sync();
    });
    const files = unzipSync(await runtime.save());
    expect(files['xl/vbaProject.bin']).toEqual(MACRO_BYTES);
    const types = strFromU8(files['[Content_Types].xml']!);
    expect(types).toContain(`ContentType="${MACRO_WORKBOOK}"`);
    expect(types).not.toContain(`${SML}sheet.main+xml`);
    expect(strFromU8(files['xl/_rels/workbook.xml.rels']!)).toContain(VBA_REL);
    expect(strFromU8(files['xl/worksheets/sheet1.xml']!)).toContain('<v>7</v>');
  } finally {
    runtime.dispose();
  }
});
