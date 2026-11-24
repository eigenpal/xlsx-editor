// The demo shows each workbook name without bidirectional format controls, in its messages and
// in its top bar. A name that the user types in the top bar drops the same characters.

import { describe, expect, test } from 'bun:test';
import { SAMPLE } from './sample.ts';
import { renamedTo, shownExtension, shownName, titleOf, UNTITLED } from './workbook-names.ts';

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

  test('the extension is the one that Save writes for the package kind', () => {
    expect(shownExtension(shownName('a.xlsx'), 'template')).toBe('.xltx');
    expect(shownExtension(shownName('a.xltx'), null)).toBe('.xltx');
    expect(shownExtension(shownName('notes'), 'workbook')).toBe('.xlsx');
    expect(shownExtension(shownName('notes'), 'template')).toBe('.xltx');
    expect(shownExtension(shownName('notes'), null)).toBe('');
    expect(shownExtension(shownName('notes'), 'workbook')).toBe('.xlsx');
  });
});

describe('the name that the user types in the top bar', () => {
  test('drops controls, end spaces, and a typed extension', () => {
    expect(renamedTo('  Budget  ', 'Sample workbook')).toBe('Budget');
    expect(renamedTo('Budget.XLSX', 'Sample workbook')).toBe('Budget');
    expect(renamedTo('Plan.xltx ', 'Plan')).toBe('Plan');
    expect(renamedTo('a\u202eb\u0007c', 'x')).toBe('abc');
    // Only one extension at the end goes: the rest of the name is the user's.
    expect(renamedTo('notes.csv', 'x')).toBe('notes.csv');
  });

  test('an empty name keeps the name before it, or takes the untitled name', () => {
    expect(renamedTo('', 'Sample workbook')).toBe('Sample workbook');
    expect(renamedTo('   ', 'Budget')).toBe('Budget');
    expect(renamedTo('.xlsx', 'Budget')).toBe('Budget');
    expect(renamedTo('\u200f', '')).toBe(UNTITLED);
  });
});
