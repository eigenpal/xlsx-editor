// The start of a long cell text that a content field holds.

import { describe, expect, test } from 'bun:test';
import { contentFieldText, MAX_CONTENT_FIELD_CHARS } from '../cell-editor.ts';

describe('contentFieldText', () => {
  test('keeps a text at the limit', () => {
    const text = 'a'.repeat(MAX_CONTENT_FIELD_CHARS);
    expect(contentFieldText(text)).toBe(text);
  });

  test('cuts a longer text at the limit', () => {
    const text = 'a'.repeat(MAX_CONTENT_FIELD_CHARS + 10);
    expect(contentFieldText(text)).toHaveLength(MAX_CONTENT_FIELD_CHARS);
  });

  test('keeps no half of a surrogate pair at the cut', () => {
    // The pair starts at the last unit that the field holds.
    const text = `${'a'.repeat(MAX_CONTENT_FIELD_CHARS - 1)}😀b`;
    const shown = contentFieldText(text);
    expect(shown).toBe('a'.repeat(MAX_CONTENT_FIELD_CHARS - 1));
    expect(/[\uD800-\uDBFF]$/.test(shown)).toBe(false);
  });
});
