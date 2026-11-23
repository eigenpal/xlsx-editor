import { describe, expect, test } from 'bun:test';
import { fontsMissingNotice } from './font-messages.ts';

describe('the font messages of the demo', () => {
  test('counts the fonts that did not load', () => {
    expect(fontsMissingNotice(1)).toStartWith('1 font face did not load');
    expect(fontsMissingNotice(12)).toStartWith('12 font faces did not load');
  });
});
