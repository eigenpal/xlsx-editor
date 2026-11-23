// The start theme of the demo (`ThemeToggle.tsx`): the address's `?theme=`, else the theme that
// the user picked before, else light. Storage that is off never stops the demo.

import { afterEach, describe, expect, test } from 'bun:test';
import { initialTheme, rememberTheme } from './ThemeToggle.tsx';

const real = Object.getOwnPropertyDescriptor(globalThis, 'localStorage');

/** A stand-in for `localStorage`: a map, or storage that throws, as in a private window. */
function useStorage(kind: 'map' | 'off') {
  const items = new Map<string, string>();
  const off = () => {
    throw new Error('storage is off');
  };
  Object.defineProperty(globalThis, 'localStorage', {
    configurable: true,
    value:
      kind === 'map'
        ? {
            getItem: (k: string) => items.get(k) ?? null,
            setItem: (k: string, v: string) => void items.set(k, v),
          }
        : { getItem: off, setItem: off },
  });
}

afterEach(() => {
  if (real) Object.defineProperty(globalThis, 'localStorage', real);
  else delete (globalThis as { localStorage?: unknown }).localStorage;
});

describe('the start theme of the demo', () => {
  test('starts light, and keeps a pick for the next visit', () => {
    useStorage('map');
    expect(initialTheme('')).toBe('light');
    rememberTheme('dark');
    expect(initialTheme('')).toBe('dark');
  });

  test('the address wins over the kept pick, and an unknown value is ignored', () => {
    useStorage('map');
    rememberTheme('dark');
    expect(initialTheme('?theme=light')).toBe('light');
    expect(initialTheme('?theme=blue')).toBe('dark');
  });

  test('storage that is off starts light and keeps nothing', () => {
    useStorage('off');
    expect(() => rememberTheme('dark')).not.toThrow();
    expect(initialTheme('')).toBe('light');
    expect(initialTheme('?theme=dark')).toBe('dark');
  });
});
