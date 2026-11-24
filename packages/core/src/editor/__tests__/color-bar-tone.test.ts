// The tone of the color bar under a color icon (`colorBarTone` in `toolbar-layout.ts`).
import { describe, expect, test } from 'bun:test';
import { colorBarTone } from '../toolbar-layout.ts';

describe('colorBarTone', () => {
  test('a color near white is light, and a color near black is dark', () => {
    expect(colorBarTone('FFFFFF')).toBe('light');
    expect(colorBarTone('F2F2F2')).toBe('light');
    expect(colorBarTone('000000')).toBe('dark');
    expect(colorBarTone('002060')).toBe('dark');
  });

  test('a color between them takes no edge', () => {
    expect(colorBarTone('FF0000')).toBeUndefined();
    expect(colorBarTone('FFFF00')).toBeUndefined();
    expect(colorBarTone('808080')).toBeUndefined();
  });

  test('a value that is not a six-digit hex color takes no edge', () => {
    expect(colorBarTone(null)).toBeUndefined();
    expect(colorBarTone(undefined)).toBeUndefined();
    expect(colorBarTone('FFF')).toBeUndefined();
    expect(colorBarTone('red')).toBeUndefined();
  });
});
