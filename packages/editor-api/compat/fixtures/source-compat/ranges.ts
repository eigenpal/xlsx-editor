/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See LICENSE.md in this package.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
/**
 * Office.js workbook code in the call patterns of the Office.js documentation, written against
 * the `XlsxEditor` namespace. See README.md in this directory.
 *
 * This file is never executed. `compat/tsconfig.json` checks it against the public types.
 */
import type * as XlsxEditor from '../../../src/index.ts';

declare const runtime: XlsxEditor.XlsxEditorServerRuntime;

export async function listWorksheetNames(): Promise<string[]> {
  return runtime.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load('items/name');
    await context.sync();
    return sheets.items.map((sheet) => sheet.name);
  });
}

export async function writeHeaderAndTotals(): Promise<void> {
  await runtime.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem('Sheet1');
    const header = sheet.getRange('A1:C1');
    header.values = [['Region', 'Units', 'Total']];
    const data = sheet.getRange('A2:C3');
    data.values = [
      ['North', 10, '=B2*4'],
      ['South', 12, '=B3*4'],
    ];
    await context.sync();
  });
}

export async function readActiveRange(): Promise<unknown[][]> {
  return runtime.run(async (context) => {
    const range = context.workbook.worksheets.getActiveWorksheet().getRange('B2:C4');
    range.load(['address', 'values', 'numberFormat']);
    await context.sync();
    console.log(range.address);
    console.log(JSON.stringify(range.numberFormat, null, 4));
    return range.values;
  });
}

export async function readWithContextLoad(): Promise<string> {
  return runtime.run(async (context) => {
    const range = context.workbook.worksheets.getActiveWorksheet().getRange('A1');
    context.load(range, 'address');
    await context.sync();
    return range.address;
  });
}

export async function findOptionalSheet(name: string): Promise<boolean> {
  return runtime.run(async (context) => {
    const sheet = context.workbook.worksheets.getItemOrNullObject(name);
    sheet.load({ select: 'name' });
    await context.sync();
    if (sheet.isNullObject) return false;
    console.log(sheet.name);
    return true;
  });
}

export async function passThrough(): Promise<number> {
  return runtime.run(async (context) => context.sync(42));
}
