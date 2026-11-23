// The address of the other demo in the React | Vue switcher (`adapters.ts`): it keeps the page's
// query, sets `theme` to the theme on screen, and works with the development and hosted bases.

import { describe, expect, test } from 'bun:test';
import { ADAPTER_BASES, adapterHref, linkHref, sanitizeHref } from './adapters.ts';

describe('the address of the other demo', () => {
  test('sets the theme on screen over the theme of the query', () => {
    expect(adapterHref('/vue/', '', 'dark')).toBe('/vue/?theme=dark');
    expect(adapterHref('/vue/', '?theme=dark', 'light')).toBe('/vue/?theme=light');
  });

  test('keeps the rest of the query', () => {
    expect(adapterHref('/react/', '?a=1&theme=light&b=two', 'dark')).toBe(
      '/react/?a=1&theme=dark&b=two'
    );
  });

  test('works with the development and the hosted bases', () => {
    expect(adapterHref(ADAPTER_BASES.vue.dev, '?a=1', 'light')).toBe(
      'http://localhost:5181/?a=1&theme=light'
    );
    expect(adapterHref(ADAPTER_BASES.react.dev, '', 'dark')).toBe(
      'http://localhost:5180/?theme=dark'
    );
    expect(adapterHref(ADAPTER_BASES.vue.build, '', 'light')).toBe('/vue/?theme=light');
    expect(adapterHref(ADAPTER_BASES.react.build, '', 'light')).toBe('/react/?theme=light');
  });

  test('a query cannot turn the address into a script URL, and an inert one falls back', () => {
    const href = adapterHref('/vue/', '?x=javascript:alert(1)', 'light');
    expect(linkHref(sanitizeHref(href), '/vue/')).toBe(href);
    expect(linkHref(sanitizeHref('javascript:alert(1)'), '/vue/')).toBe('/vue/');
  });
});
