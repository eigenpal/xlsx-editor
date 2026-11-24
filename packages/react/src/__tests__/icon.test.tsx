/** @jsxImportSource react */
// `Icon`: a chrome icon by name, host path data, and the class of the svg.

import { afterEach, expect, test } from 'bun:test';
import { cleanup, render } from '@testing-library/react';
import { CHROME_ICON_PATHS } from '@xlsx-editor.dev/core/editor';
import { Icon } from '../index.ts';

afterEach(cleanup);

test('draws a chrome icon by name, decorative, with the default class', () => {
  const { container } = render(<Icon name="file_download" />);
  const svg = container.querySelector('svg')!;
  expect(svg.getAttribute('class')).toBe('xlsx-icon');
  expect(svg.getAttribute('aria-hidden')).toBe('true');
  expect(svg.getAttribute('viewBox')).toBe('0 -960 960 960');
  expect([...svg.querySelectorAll('path')].map((p) => p.getAttribute('d'))).toEqual([
    ...CHROME_ICON_PATHS.file_download,
  ]);
});

test('draws host path data in place of a name, with a host class', () => {
  const { container } = render(<Icon name="undo" paths={['M0 0h10v10z']} className="mine" />);
  const svg = container.querySelector('svg')!;
  expect(svg.getAttribute('class')).toBe('xlsx-icon mine');
  expect(svg.querySelector('path')!.getAttribute('d')).toBe('M0 0h10v10z');
});

test('draws nothing for a name that is not a chrome icon, also a key of Object', () => {
  for (const name of ['constructor', '__proto__', 'nope']) {
    const { container, unmount } = render(<Icon name={name as 'undo'} />);
    expect(container.querySelectorAll('path').length).toBe(0);
    unmount();
  }
});
