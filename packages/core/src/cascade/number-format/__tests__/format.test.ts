// The number format engine with the ECMA-376 built-in codes: the spec scenarios, the parts of
// format codes, both date systems, `t="d"` ISO dates, and hostile codes that must fall back to
// `General` with a problem, in linear time.

import { describe, expect, test } from 'bun:test';
import {
  builtinFormatCode,
  dateToSerial,
  EN_US,
  formatCodeProblem,
  formatValue,
  isDateFormat,
  serialToDate,
} from '../index.ts';

const fmt = (value: number | string | boolean, code: string, date1904 = false) =>
  formatValue(value, code, { locale: EN_US, date1904 });
const text = (value: number | string | boolean, code: string, date1904 = false) =>
  fmt(value, code, date1904).text;

describe('spec scenarios', () => {
  test('a built-in format: 1234.5 with id 4', () => {
    expect(text(1234.5, builtinFormatCode(4, EN_US)!)).toBe('1,234.50');
  });

  test('the same serial in two date systems', () => {
    expect(text(1, 'yyyy-mm-dd', false)).toBe('1900-01-01');
    expect(text(1, 'yyyy-mm-dd', true)).toBe('1904-01-02');
  });
});

describe('built-in formats', () => {
  test('the ECMA-376 table and unknown ids', () => {
    expect(builtinFormatCode(0, EN_US)).toBe('General');
    expect(builtinFormatCode(14, EN_US)).toBe('mm-dd-yy');
    expect(builtinFormatCode(5, EN_US)).toBeUndefined();
    expect(builtinFormatCode(164, EN_US)).toBeUndefined();
    expect(builtinFormatCode(1.5, EN_US)).toBeUndefined();
    expect(builtinFormatCode(14, { ...EN_US, builtins: { 14: 'dd/mm/yyyy' } })).toBe('dd/mm/yyyy');
  });

  test('negative numbers in the parenthesized built-ins', () => {
    expect(text(-1234, builtinFormatCode(37, EN_US)!)).toBe('(1,234)');
    expect(text(1234, builtinFormatCode(37, EN_US)!)).toBe('1,234 ');
    const red = fmt(-1234.5, builtinFormatCode(40, EN_US)!);
    expect(red).toEqual({ text: '(1,234.50)', color: 'FF0000' });
  });
});

describe('General', () => {
  test('the 11-character rule', () => {
    expect(text(0, 'General')).toBe('0');
    expect(text(1234567.891, 'General')).toBe('1234567.891');
    expect(text(12345678901, 'General')).toBe('12345678901');
    expect(text(123456789012, 'General')).toBe('1.23457E+11');
    expect(text(1e-10, 'General')).toBe('1E-10');
    expect(text(1.5e100, 'General')).toBe('1.5E+100');
    expect(text(-1 / 3, 'General')).toBe('-0.333333333');
  });

  test('booleans and text', () => {
    expect(text(true, '0.00')).toBe('TRUE');
    expect(text('abc', '0.00')).toBe('abc');
    expect(text('abc', '@')).toBe('abc');
  });
});

describe('number sections', () => {
  test('rounding is half away from zero on the shown digits', () => {
    expect(text(2.5, '0')).toBe('3');
    expect(text(-2.5, '0')).toBe('-3');
    expect(text(1.005, '0.00')).toBe('1.01');
    expect(text(0.96, '0.0')).toBe('1.0');
    expect(text(-0.2, '0')).toBe('-0');
  });

  test('placeholders, literals, grouping, and scaling', () => {
    expect(text(5, '0,000')).toBe('0,005');
    expect(text(1234567, '#,##0')).toBe('1,234,567');
    expect(text(0.5, '.00')).toBe('.50');
    expect(text(12.5, '.00')).toBe('12.50');
    expect(text(5551234, '000-0000')).toBe('555-1234');
    expect(text(1234567, '#,##0,"K"')).toBe('1,235K');
    expect(text(0.125, '0.0%')).toBe('12.5%');
    expect(text(5, '[$€-407] 0')).toBe('€ 5');
  });

  test('scientific notation', () => {
    expect(text(12345, '0.00E+00')).toBe('1.23E+04');
    expect(text(0.00012, '0.0E+00')).toBe('1.2E-04');
    expect(text(99999, '0.00E+00')).toBe('1.00E+05');
    expect(text(0, '0.00E+00')).toBe('0.00E+00');
    expect(text(123456, '##0.0E+0')).toBe('123.5E+3');
    expect(text(12345, '0.0E-0')).toBe('1.2E4');
  });

  test('fractions', () => {
    expect(text(1.75, '?/?')).toBe('7/4');
    expect(text(2, '# ?/?')).toBe('2    ');
    expect(text(0.5, '# ?/8')).toBe(' 4/8');
    expect(text(0.999, '# ?/?')).toBe('1    ');
  });

  test('sections, conditions, and colors', () => {
    expect(text(0, '0;-0;"zero"')).toBe('zero');
    expect(text(-5, '0;(0)')).toBe('(5)');
    // The reference PDFs paint `[Red]` text red (1 0 0) and `[Blue]` text blue (0 0 1), measured
    // on number-formats.pdf.
    expect(fmt(7, '[Blue]0').color).toBe('0000FF');
    expect(fmt(7, '[Color10]0').color).toBe('008000');
    expect(text(50, '[>=100]"big";"small"')).toBe('small');
    expect(text(-5, '[>=100]0;0')).toBe('-5');
    // A number through a text section aligns as text (number-formats.pdf).
    expect(fmt(42, '@')).toEqual({ text: '42', alignAsText: true });
  });

  test('skip and fill', () => {
    expect(fmt(5, '_(0_)')).toEqual({
      text: ' 5 ',
      pads: [
        { index: 0, char: '(' },
        { index: 2, char: ')' },
      ],
    });
    expect(fmt(-5, '* 0')).toEqual({ text: '-5', fill: { char: ' ', index: 1 } });
    // A `?` placeholder is a space as wide as a digit, after the sign of a negative value.
    expect(fmt(1.5, '?.??')).toEqual({ text: '1.5 ', pads: [{ index: 3, char: '0' }] });
    expect(fmt(-2.4, '# ??/??').pads).toEqual([
      { index: 3, char: '0' },
      { index: 7, char: '0' },
    ]);
    expect(fmt(0.5, 'h_)').pads).toEqual([{ index: 2, char: ')' }]);
    expect(fmt('ab', '_(@_)').pads).toEqual([
      { index: 0, char: '(' },
      { index: 3, char: ')' },
    ]);
    expect(fmt('ab', '@*-')).toEqual({ text: 'ab', fill: { char: '-', index: 2 } });
  });
});

describe('dates and times', () => {
  test('date parts and names', () => {
    expect(text(45352, 'dddd, mmmm d, yyyy')).toBe('Friday, March 1, 2024');
    expect(text(45352, 'ddd d mmm yy')).toBe('Fri 1 Mar 24');
    expect(text(45352, 'mmmmm')).toBe('M');
    expect(text(45352, builtinFormatCode(14, EN_US)!)).toBe('03-01-24');
  });

  test('minutes and months', () => {
    expect(text(45352.5105, 'm/d h:mm')).toBe('3/1 12:15');
    expect(text(0.0104166, 'mm:ss')).toBe('15:00');
  });

  test('times round and carry', () => {
    expect(text(0.999999, 'yyyy-mm-dd h:mm:ss', false)).toBe('1900-01-01 0:00:00');
    expect(text(0.75, 'h:mm A/P')).toBe('6:00 P');
    expect(text(0.0001, 'h:mm:ss.000')).toBe('0:00:08.640');
    expect(text(2.25, '[mm]:ss')).toBe('3240:00');
  });

  test('the 1900 leap-year serial and serial 0', () => {
    expect(text(0, 'yyyy-mm-dd')).toBe('1900-01-00');
    expect(text(60, 'yyyy-mm-dd ddd')).toBe('1900-02-29 Wed');
    expect(serialToDate(61, false)).toEqual({ year: 1900, month: 3, day: 1, weekday: 4 });
    expect(serialToDate(0, true)).toEqual({ year: 1904, month: 1, day: 1, weekday: 5 });
  });

  test('a value that is not a date fills the cell with #', () => {
    expect(fmt(-1, 'yyyy-mm-dd')).toEqual({ text: '', fill: { char: '#', index: 0 } });
    expect(fmt(3e6, 'yyyy-mm-dd')).toEqual({ text: '', fill: { char: '#', index: 0 } });
  });

  test('isDateFormat reads the first section', () => {
    expect(isDateFormat('yyyy-mm-dd')).toBe(true);
    expect(isDateFormat('[$-409]h:mm AM/PM;@')).toBe(true);
    expect(isDateFormat('[h]:mm')).toBe(true);
    expect(isDateFormat('0.00')).toBe(false);
    expect(isDateFormat('"date"0')).toBe(false);
    expect(isDateFormat('General')).toBe(false);
    expect(isDateFormat('"x'.padEnd(300, 'y'))).toBe(false);
  });
});

describe('t="d" ISO dates', () => {
  test('dates and date-times convert to serials in both systems', () => {
    expect(dateToSerial('2024-03-01', false)).toBe(45352);
    expect(dateToSerial('2024-03-01', true)).toBe(45352 - 1462);
    expect(dateToSerial('1900-01-01', false)).toBe(1);
    expect(dateToSerial('1900-02-28', false)).toBe(59);
    expect(dateToSerial('1900-03-01', false)).toBe(61);
    expect(dateToSerial('1904-01-01', true)).toBe(0);
    expect(dateToSerial('2024-03-01T18:00:00Z', false)).toBe(45352.75);
    expect(dateToSerial('2024-03-01T06:00:00.5', false)).toBeCloseTo(45352.25 + 0.5 / 86400, 12);
    expect(text(dateToSerial('2024-03-01', true)!, 'yyyy-mm-dd', true)).toBe('2024-03-01');
  });

  test('invalid dates give undefined', () => {
    for (const bad of [
      '',
      '2024-3-1',
      '2024-02-30',
      '2023-02-29',
      '2024-13-01',
      '1899-12-31',
      'abcd-ef-gh',
      '2024-03-01T25:00',
      '2024-03-01X10:00',
      '2024-03-01T10:00:00.',
      '2024-03-01'.padEnd(100, '0'),
    ])
      expect(dateToSerial(bad, false)).toBeUndefined();
    expect(dateToSerial('1903-12-31', true)).toBeUndefined();
  });
});

describe('hostile format codes', () => {
  const hostile: Record<string, string> = {
    'a 10 000 character code': '0'.repeat(10_000),
    '50 sections': Array.from({ length: 50 }, () => '0').join(';'),
    '1 000 brackets': '[Red]'.repeat(1000) + '0',
    'brackets over the section cap in a short code': '[Red][Blue][>0][$-409][Green]0',
    'nested quotes': '"a""b"""""'.repeat(30),
    'an unterminated quote': '0 "abc',
    'an unterminated bracket': '[Red 0',
    'an invalid condition': '[>=abc]0',
    'an unknown color index': '[Color99]0',
  };
  for (const [name, code] of Object.entries(hostile)) {
    test(`${name} falls back to General with a problem`, () => {
      const start = performance.now();
      const out = fmt(1234.5, code);
      expect(out.text).toBe('1234.5');
      expect(out.problem?.code).toBeString();
      expect(formatCodeProblem(code)).toEqual(out.problem!);
      expect(performance.now() - start).toBeLessThan(50);
    });
  }

  test('each refused code reports its problem code', () => {
    const codeOf = (code: string) => formatCodeProblem(code)?.code;
    expect(codeOf(hostile['a 10 000 character code']!)).toBe('format-code-too-long');
    expect(codeOf(hostile['50 sections']!)).toBe('format-too-many-sections');
    expect(codeOf(hostile['brackets over the section cap in a short code']!)).toBe(
      'format-too-many-brackets'
    );
    for (const name of [
      'an unterminated quote',
      'an unterminated bracket',
      'an invalid condition',
      'an unknown color index',
    ])
      expect(codeOf(hostile[name]!)).toBe('format-code-unparsable');
  });

  test('a value that is not finite shows #NUM! also with a refused code', () => {
    expect(text(Infinity, '0 "abc')).toBe('#NUM!');
    expect(text(NaN, '0'.repeat(300))).toBe('#NUM!');
  });

  test('a fraction of 1e21 or more shows decimal digits, not exponent notation', () => {
    expect(text(1e21, '?/?')).toBe('1000000000000000000000/1');
  });

  test('an elapsed hour token pads to its length', () => {
    expect(text(1.5, '[hhh]:mm')).toBe('036:00');
    expect(text(0.5, '[h]:mm')).toBe('12:00');
  });

  test('a long valid code at the cap formats in linear time', () => {
    const code = '0'.padEnd(255, '#');
    expect(formatCodeProblem(code)).toBeUndefined();
    expect(text(5, '#'.repeat(254) + '0')).toBe('5');
  });

  test('a huge value does not size a loop or a string past the double range', () => {
    expect(text(1e308, '#,##0').length).toBeLessThan(420);
    expect(text(1e-308, '0.' + '0'.repeat(250)).length).toBe(252);
    expect(text(Number.MAX_VALUE, '0.00E+00')).toBe('1.80E+308');
    expect(text(Infinity, '0')).toBe('#NUM!');
  });

  test('prototype keys in brackets are not colors', () => {
    expect(fmt(1, '[constructor]0')).toEqual({ text: '1' });
    expect(fmt(1, '[__proto__]0')).toEqual({ text: '1' });
  });
});
