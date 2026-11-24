// The long and unusual sheet names of the hang tests (`cases.ts`). `cases.ts` lists them in
// `FILE_CASES` after its own first sheet, so each case names `HOSTILE_SHEET` first.

import { buildXlsx, type FixtureSheet } from '../fixtures/builder.ts';
import type { HangCase } from './cases.ts';

/** The cases of long and unusual sheet names. `first` is the first sheet, `rows` its rows. */
export const sheetNameCases = (first: FixtureSheet, rows: string): readonly HangCase[] => {
  const withSheets =
    (...names: string[]) =>
    () =>
      buildXlsx({ sheets: [first, ...names.map((name) => ({ name, rows }))] });
  return [
    {
      id: 'sheet-names-object-keys',
      covers: 'worksheets are named __proto__, constructor, and prototype',
      expect: 'open',
      firstCell: 'first',
      bytes: withSheets('__proto__', 'constructor', 'prototype'),
    },
    {
      // `x:sheet/@name` (`CT_Sheet`, ECMA-376 Part 1, 18.2.19) has no length limit in the
      // standard; the read refuses a stored name of more than 32,767 UTF-16 code units.
      id: 'sheet-name-past-limit',
      covers: 'a sheet name of 1000000 code units, past the limit of a stored sheet name',
      expect: 'refuse',
      refusedWith: {
        reason: 'workbook',
        detail: 'sheet-name-too-long',
        limit: 'workbook.maxSheetNameLength',
      },
      bytes: withSheets('N'.repeat(1_000_000)),
    },
    {
      id: 'sheet-name-at-limit',
      covers: 'a sheet name of 32767 code units, the longest stored sheet name that opens',
      expect: 'open',
      firstCell: 'first',
      bytes: withSheets('N'.repeat(32_767)),
    },
  ];
};
