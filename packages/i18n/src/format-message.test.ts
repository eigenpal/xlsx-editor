// `formatMessage`: the placeholder and plural rules of `t()`, for one message, with the own
// values of `vars` only.
import { describe, expect, test } from 'bun:test';
import { formatMessage } from './index.ts';

describe('formatMessage', () => {
  test('fills placeholders and picks plural branches', () => {
    expect(formatMessage('Sheet {name}', { name: 'Data' })).toBe('Sheet Data');
    expect(formatMessage('{count, plural, one {# sheet} other {# sheets}}', { count: 1 })).toBe(
      '1 sheet'
    );
    expect(formatMessage('{count, plural, =0 {none} other {#}}', { count: 0 })).toBe('none');
    expect(formatMessage('{count, plural, few {# few} other {#}}', { count: 3 }, 'pl')).toBe(
      '3 few'
    );
  });

  test('a placeholder with no own value stays as written', () => {
    expect(formatMessage('{missing}', { other: 1 })).toBe('{missing}');
    expect(formatMessage('{constructor} {toString}', {})).toBe('{constructor} {toString}');
    expect(formatMessage('{count, plural, other {#}} {__proto__}', {})).toBe(
      '{count, plural, other {#}} {__proto__}'
    );
  });

  test('a message with no values is returned as written', () => {
    expect(formatMessage('Sheet {name}')).toBe('Sheet {name}');
  });
});
