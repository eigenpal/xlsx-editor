// The break opportunities of a very long text: each layout reads the same ones.

import { describe, expect, test } from 'bun:test';
import { breaksOf } from '../cell-text.ts';

const long = (letter: string) => `${letter} `.repeat(40_000);

describe('the break opportunities of a long text', () => {
  test('a long text keeps its break opportunities between layouts', () => {
    const text = long('a');
    expect(breaksOf(text)).toBe(breaksOf(text));
  });

  test('a short text keeps none', () => {
    expect(breaksOf('a b')).not.toBe(breaksOf('a b'));
  });

  test('the cache holds only the last two long texts', () => {
    const first = long('b');
    const kept = breaksOf(first);
    breaksOf(long('c'));
    expect(breaksOf(first)).toBe(kept);
    breaksOf(long('d'));
    breaksOf(long('e'));
    expect(breaksOf(first)).not.toBe(kept);
  });
});
