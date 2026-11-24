// `Icon`: the twin of the React test (`packages/react/src/__tests__/icon.test.tsx`).

import { afterEach, expect, test } from 'bun:test';
import { h } from 'vue';
import { CHROME_ICON_PATHS } from '@xlsx-editor.dev/core/editor';
import { Icon } from '../src/index.ts';
import { mount, unmountAll } from './mount.ts';

afterEach(unmountAll);

test('draws a chrome icon by name, decorative, with the default class', () => {
  const view = mount(() => h(Icon, { name: 'file_download' }));
  const svg = view.container.querySelector('svg')!;
  expect(svg.getAttribute('class')).toBe('xlsx-icon');
  expect(svg.getAttribute('aria-hidden')).toBe('true');
  expect(svg.getAttribute('viewBox')).toBe('0 -960 960 960');
  expect([...svg.querySelectorAll('path')].map((p) => p.getAttribute('d'))).toEqual([
    ...CHROME_ICON_PATHS.file_download,
  ]);
});

test('draws host path data in place of a name, with a host class', () => {
  const view = mount(() => h(Icon, { name: 'undo', paths: ['M0 0h10v10z'], className: 'mine' }));
  const svg = view.container.querySelector('svg')!;
  expect(svg.getAttribute('class')).toBe('xlsx-icon mine');
  expect(svg.querySelector('path')!.getAttribute('d')).toBe('M0 0h10v10z');
});

test('draws nothing for a name that is not a chrome icon, also a key of Object', () => {
  for (const name of ['constructor', '__proto__', 'nope']) {
    const view = mount(() => h(Icon, { name: name as 'undo' }));
    expect(view.container.querySelectorAll('path').length).toBe(0);
    view.unmount();
  }
});
