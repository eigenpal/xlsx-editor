/** @jsxImportSource react */
// The React | Vue switcher (`AdapterSwitcher.tsx`) in happy-dom: the current demo is marked and
// is not a link, the other demo is a link with the query and the theme on screen, and the
// switcher is one Tab stop.

import { afterEach, describe, expect, test } from 'bun:test';
import { cleanup, render } from '@testing-library/react';
import { AdapterSwitcher } from '../AdapterSwitcher.tsx';

afterEach(cleanup);

/** The elements in `root` that the Tab key stops on. */
const tabStops = (root: HTMLElement) =>
  [...root.querySelectorAll<HTMLElement>('a[href], button, input, select, [tabindex]')].filter(
    (el) => el.tabIndex >= 0
  );

describe('the React | Vue switcher', () => {
  test('marks the React demo and links to the Vue demo with the theme on screen', () => {
    const view = render(<AdapterSwitcher current="react" theme="dark" />);
    const nav = view.getByRole('navigation', { name: 'Framework' });
    const current = nav.querySelector('[aria-current="page"]');
    expect(current?.textContent).toBe('React');
    expect(current?.tagName).toBe('SPAN');
    const link = view.getByRole('link', { name: 'Vue' });
    // Bun sets no `import.meta.env.DEV`, so the switcher takes the hosted bases.
    expect(link.getAttribute('href')).toBe('/vue/?theme=dark');
    expect(tabStops(nav)).toEqual([link]);
  });

  test('marks the Vue demo and links back to the React demo', () => {
    const view = render(<AdapterSwitcher current="vue" theme="light" />);
    const nav = view.getByRole('navigation', { name: 'Framework' });
    expect(nav.querySelector('[aria-current="page"]')?.textContent).toBe('Vue');
    const link = view.getByRole('link', { name: 'React' });
    expect(link.getAttribute('href')).toBe('/react/?theme=light');
    expect(tabStops(nav)).toEqual([link]);
  });
});
