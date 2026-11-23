import { describe, expect, test } from 'bun:test';
import { EN_US, type NumberLocale } from '../../cascade/index.ts';
import { englishText } from '../english.ts';
import { numberFormatListView } from '../number-format-list.ts';
import { CUSTOM_NUMBER_FORMAT_PREFIX } from '../number-format-commands.ts';

const custom = (code: string) => `${CUSTOM_NUMBER_FORMAT_PREFIX}${code}`;
const view = (value: string | undefined, locale: NumberLocale = EN_US) => {
  const { entries: _, ...rest } = numberFormatListView(value, englishText, locale);
  return rest;
};

describe('numberFormatListView', () => {
  test('lists the entries with their names', () => {
    const { entries } = numberFormatListView(undefined, englishText);
    expect(entries.map((e) => e.value)).toEqual([
      'general',
      'number',
      'currency',
      'accounting',
      'percent',
      'scientific',
      'fraction',
      'shortDate',
      'longDate',
      'time',
      'text',
    ]);
    expect(entries[0]).toEqual({ value: 'general', label: 'General' });
  });

  test('an entry with no places gives its name as the text and the name', () => {
    expect(view('general')).toEqual({ text: 'General', name: 'General' });
    expect(view('shortDate')).toEqual({ text: 'Short date', name: 'Short date' });
  });

  test('an entry with places names them', () => {
    expect(view('number')).toEqual({ text: 'Number', name: 'Number, 2 decimal places' });
  });

  test('a code that differs from an entry only in its places names the entry, current', () => {
    expect(view(custom('0.000'))).toEqual({
      text: 'Number',
      name: 'Number, 3 decimal places',
      current: 'number',
    });
    expect(view(custom('0.0%'))).toEqual({
      text: 'Percentage',
      name: 'Percentage, 1 decimal place',
      current: 'percent',
    });
  });

  test('any other code shows Custom and its code, and names places that a step changes', () => {
    expect(view(custom('#,##0.00'))).toEqual({
      text: 'Custom: #,##0.00',
      name: 'Custom: #,##0.00, 2 decimal places',
    });
    expect(view(custom('@" kg"'))).toEqual({ text: 'Custom: @" kg"', name: 'Custom: @" kg"' });
  });

  test('no value gives empty text', () => {
    expect(view(undefined)).toEqual({ text: '', name: '' });
    expect(view('not a value')).toEqual({ text: '', name: '' });
  });

  test('a locale with no valid currency symbol leaves out currency and accounting', () => {
    const locale: NumberLocale = { ...EN_US, currencySymbol: '' };
    const values = numberFormatListView('general', englishText, locale).entries.map((e) => e.value);
    expect(values).not.toContain('currency');
    expect(values).not.toContain('accounting');
  });
});
