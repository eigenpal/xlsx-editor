// The demo shows each workbook name without bidirectional format controls, in its messages and
// in its top bar.

import { describe, expect, test } from 'bun:test';
import { SAMPLE } from './sample.ts';
import { shownName, titleOf } from './workbook-names.ts';

describe('the shown name of a workbook', () => {
  test('a message names the sample for people and drops controls from a picked file name', () => {
    expect(titleOf(SAMPLE)).toBe('the sample workbook');
    expect(titleOf('a.xlsx')).toBe('a.xlsx');
    expect(titleOf('a‮xslx.exe')).toBe('axslx.exe');
    expect(titleOf('a\u0085b.xlsx')).toBe('ab.xlsx');
  });

  test('the top bar shows the name apart from its extension, and drops controls', () => {
    expect(shownName(SAMPLE)).toEqual({ name: 'Sample workbook', ext: '.xlsx' });
    expect(shownName('a.xlsx')).toEqual({ name: 'a', ext: '.xlsx' });
    expect(shownName('Plan.XLTX')).toEqual({ name: 'Plan', ext: '.xltx' });
    expect(shownName('a‮xslx.exe')).toEqual({ name: 'axslx.exe', ext: '' });
    expect(shownName('.xlsx')).toEqual({ name: '.xlsx', ext: '' });
    expect(shownName('notes')).toEqual({ name: 'notes', ext: '' });
  });
});
