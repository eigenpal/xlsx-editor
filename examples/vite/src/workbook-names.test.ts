// The demo shows each workbook name without bidirectional format controls, in its messages and
// in its workbook list.

import { describe, expect, test } from 'bun:test';
import { SAMPLE } from './sample.ts';
import { listedName, titleOf } from './workbook-names.ts';

describe('the shown name of a workbook', () => {
  test('a message names the sample for people and drops controls from a picked file name', () => {
    expect(titleOf(SAMPLE)).toBe('the sample workbook');
    expect(titleOf('a.xlsx')).toBe('a.xlsx');
    expect(titleOf('a‮xslx.exe')).toBe('axslx.exe');
    expect(titleOf('a\u0085b.xlsx')).toBe('ab.xlsx');
  });

  test('the workbook list drops controls from a picked file name', () => {
    expect(listedName('a.xlsx')).toBe('a.xlsx');
    expect(listedName('a‮xslx.exe')).toBe('axslx.exe');
  });
});
