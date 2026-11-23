// The key rules of a chrome field that edits the open cell edit (`field-keys.ts`).

import { describe, expect, test } from 'bun:test';
import {
  composesKey,
  editMoveOf,
  fieldHistoryKey,
  refusesEditText,
  refusesReference,
} from '../index.ts';

const keys = (
  key: string,
  mods: Partial<Record<'shift' | 'alt' | 'ctrl' | 'meta', boolean>> = {}
) => ({
  key,
  shiftKey: !!mods.shift,
  altKey: !!mods.alt,
  ctrlKey: !!mods.ctrl,
  metaKey: !!mods.meta,
});

describe('field keys', () => {
  test('Enter and Tab move as in the cell editor; Alt+Enter is a line break', () => {
    expect(editMoveOf(keys('Enter'))).toBe('down');
    expect(editMoveOf(keys('Enter', { shift: true }))).toBe('up');
    expect(editMoveOf(keys('Tab'))).toBe('right');
    expect(editMoveOf(keys('Tab', { shift: true }))).toBe('left');
    expect(editMoveOf(keys('Enter', { alt: true }))).toBeNull();
    expect(editMoveOf(keys('Enter', { ctrl: true }))).toBe('down');
    expect(editMoveOf(keys('Enter', { meta: true, shift: true }))).toBe('up');
    expect(editMoveOf(keys('a'))).toBeNull();
  });

  test('a composing key belongs to the input method', () => {
    expect(composesKey({ isComposing: true }, false)).toBe(true);
    expect(composesKey({ keyCode: 229 }, false)).toBe(true);
    expect(composesKey({}, true)).toBe(true);
    expect(composesKey({ keyCode: 13 }, false)).toBe(false);
  });

  test('history keys step the workbook only when the field holds no typing', () => {
    expect(fieldHistoryKey(keys('z', { ctrl: true }), false)).toBe('undo');
    expect(fieldHistoryKey(keys('z', { meta: true, shift: true }), false)).toBe('redo');
    expect(fieldHistoryKey(keys('y', { ctrl: true }), false)).toBe('redo');
    expect(fieldHistoryKey(keys('z', { ctrl: true }), true)).toBeNull();
    expect(fieldHistoryKey(keys('z'), false)).toBeNull();
  });

  test('a refusal marks the reference field or the edit text, never both', () => {
    for (const code of ['go-to-reference', 'go-to-name', 'go-to-sheet'] as const) {
      expect([refusesReference(code), refusesEditText(code)]).toEqual([true, false]);
    }
    for (const code of ['composing', 'edit-pending', 'edit-dropped'] as const) {
      expect([refusesReference(code), refusesEditText(code)]).toEqual([false, false]);
    }
    for (const code of ['formula-syntax', 'locked', 'unsupported-function'] as const) {
      expect([refusesReference(code), refusesEditText(code)]).toEqual([false, true]);
    }
  });
});
