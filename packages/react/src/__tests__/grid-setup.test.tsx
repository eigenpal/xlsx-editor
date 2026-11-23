// The development warning for a grid that has no styles.

import { afterEach, describe, expect, test } from 'bun:test';
import { cleanup, render, waitFor } from '@testing-library/react';
import { buildXlsx, c, row } from '../../../../scripts/fixtures/builder.ts';
import { gridSetupProblem } from '@xlsx-editor.dev/core/editor';
import { XlsxEditor } from '../index.ts';

afterEach(cleanup);

describe('gridSetupProblem', () => {
  test('names the stylesheet when the scroller does not scroll', () => {
    expect(gridSetupProblem({ overflowY: 'visible' })).toContain('styles.css');
  });

  test('reports nothing for a styled grid', () => {
    expect(gridSetupProblem({ overflowY: 'auto' })).toBeNull();
  });

  test('reports nothing when the host computes no styles', () => {
    expect(gridSetupProblem({ overflowY: '' })).toBeNull();
  });
});

test('in development, a grid with no stylesheet logs the warning', async () => {
  const warnings: string[] = [];
  const warn = console.warn;
  const computed = window.getComputedStyle;
  console.warn = (message: string) => warnings.push(String(message));
  window.getComputedStyle = ((element: Element) =>
    element.classList.contains('xlsx-scroller')
      ? ({ overflowY: 'visible' } as CSSStyleDeclaration)
      : computed.call(window, element)) as typeof window.getComputedStyle;
  try {
    const document = buildXlsx({ sheets: [{ name: 'Data', rows: row(1, [c('A1', 1)]) }] });
    render(<XlsxEditor document={document} style={{ height: 600 }} />);
    await waitFor(() =>
      expect(warnings.some((text) => text.includes('@xlsx-editor.dev/react/styles.css'))).toBe(true)
    );
  } finally {
    console.warn = warn;
    window.getComputedStyle = computed;
  }
});

test('in development, a stylesheet that loads after the grid gives no warning', async () => {
  const warnings: string[] = [];
  const warn = console.warn;
  const computed = window.getComputedStyle;
  const state = Object.getOwnPropertyDescriptor(window.document, 'readyState');
  let styled = false;
  let frames = 0;
  console.warn = (message: string) => warnings.push(String(message));
  Object.defineProperty(window.document, 'readyState', {
    configurable: true,
    get: () => 'loading',
  });
  window.getComputedStyle = ((element: Element) => {
    if (!element.classList.contains('xlsx-scroller')) return computed.call(window, element);
    frames += 1;
    return { overflowY: styled ? 'auto' : 'visible' } as CSSStyleDeclaration;
  }) as typeof window.getComputedStyle;
  try {
    const document = buildXlsx({ sheets: [{ name: 'Data', rows: row(1, [c('A1', 1)]) }] });
    const { container } = render(<XlsxEditor document={document} style={{ height: 600 }} />);
    await waitFor(() => expect(container.querySelector('.xlsx-scroller')).not.toBeNull());
    await new Promise((resolve) => window.requestAnimationFrame(resolve));
    expect(frames).toBe(0);
    styled = true;
    window.dispatchEvent(new window.Event('load'));
    await waitFor(() => expect(frames).toBeGreaterThan(0));
    expect(warnings.some((text) => text.includes('styles.css'))).toBe(false);
  } finally {
    console.warn = warn;
    window.getComputedStyle = computed;
    if (state) Object.defineProperty(window.document, 'readyState', state);
    else delete (window.document as { readyState?: string }).readyState;
  }
});
