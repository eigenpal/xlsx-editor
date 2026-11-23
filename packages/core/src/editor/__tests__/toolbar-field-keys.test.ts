// The keys that a field in the toolbar keeps (`fieldOwnsKey`), with plain objects in place of
// elements. The adapter tests walk the same rule in a bar.

import { describe, expect, test } from 'bun:test';
import { fieldOwnsKey, type FieldLike } from '../toolbar-layout.ts';

const KEYS = ['ArrowLeft', 'ArrowRight', 'Home', 'End'] as const;

/** The keys of `KEYS` that `field` keeps. */
const kept = (field: FieldLike) =>
  KEYS.filter((key) => fieldOwnsKey(field as EventTarget, key, key === 'ArrowRight'));

const text = (start: number | null, end = start, value = 'ab'): FieldLike => ({
  localName: 'input',
  type: 'text',
  value,
  selectionStart: start,
  selectionEnd: end,
});

describe('a text field', () => {
  test('keeps Home and End, and an arrow until the caret is at that end', () => {
    expect(kept(text(0))).toEqual(['ArrowRight', 'Home', 'End']);
    expect(kept(text(1))).toEqual(['ArrowLeft', 'ArrowRight', 'Home', 'End']);
    expect(kept(text(2))).toEqual(['ArrowLeft', 'Home', 'End']);
  });

  test('with selected text, keeps both arrows', () => {
    expect(kept(text(0, 2))).toEqual(['ArrowLeft', 'ArrowRight', 'Home', 'End']);
  });

  test('with no type, is a text field', () => {
    expect(kept({ localName: 'input', value: '', selectionStart: 0, selectionEnd: 0 })).toEqual([
      'Home',
      'End',
    ]);
  });

  test('whose caret the page cannot read, such as an email field, keeps every key', () => {
    for (const type of ['email', 'number', 'date'])
      expect(kept({ localName: 'input', type, selectionStart: null, selectionEnd: null })).toEqual([
        ...KEYS,
      ]);
  });
});

describe('other fields', () => {
  test('a text area and an editable element keep every key', () => {
    expect(kept({ ...text(2), localName: 'textarea' })).toEqual([...KEYS]);
    expect(kept({ localName: 'div', isContentEditable: true })).toEqual([...KEYS]);
  });

  test('a list keeps Home and End, and gives the arrows to the bar', () => {
    expect(kept({ localName: 'select' })).toEqual(['Home', 'End']);
  });

  test('a radio keeps the arrows, and a range keeps every key, which change their value', () => {
    expect(kept({ localName: 'input', type: 'radio' })).toEqual(['ArrowLeft', 'ArrowRight']);
    expect(kept({ localName: 'input', type: 'range' })).toEqual([...KEYS]);
  });

  test('an input that acts as a button, a button, and nothing keep no key', () => {
    for (const type of ['checkbox', 'button', 'submit', 'reset', 'image', 'color', 'file'])
      expect(kept({ localName: 'input', type })).toEqual([]);
    expect(kept({ localName: 'button' })).toEqual([]);
    expect(fieldOwnsKey(null, 'Home', false)).toBe(false);
  });
});
