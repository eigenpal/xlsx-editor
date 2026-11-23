// XLSX Editor copy of pptx-editor/scripts/check-css-scope.test.ts at 3bf89a88 (EigenPal).
// Intended difference: the class and name prefix (`xlsx-` for `pptx-`, with the article "an"
// for "a").

import { describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { cssScopeErrors, nestedSelectorScoped, selectorScoped } from './check-css-scope.mjs';

const root = join(import.meta.dir, '..');

describe('check-css-scope', () => {
  test('the shipped chrome stylesheet styles nothing outside the editor', () => {
    const css = readFileSync(join(root, 'packages/core/src/styles/editor.css'), 'utf8');
    expect(cssScopeErrors(css)).toEqual([]);
  });

  test('a selector is scoped when it starts with an xlsx- class and stays inside', () => {
    for (const selector of [
      '.xlsx-editor',
      '.xlsx-editor.dark .xlsx-editor',
      'fieldset.xlsx-dialog__section',
      '.xlsx-menu:is(.a, .b) > button',
      '.xlsx-button:not(:disabled):active',
      '.xlsx-canvas:has(~ .xlsx-find:not([hidden]))',
      '.xlsx-swatches__section + .xlsx-swatches__section',
      '.xlsx-menu li + li',
      '.xlsx-a > * + *',
      '*.xlsx-x',
    ])
      expect([selector, selectorScoped(selector)]).toEqual([selector, true]);
    for (const selector of [
      'button',
      '*',
      ':root',
      '.dark .xlsx-editor',
      'div > .xlsx-menu',
      'body:has(.xlsx-dialog[open])',
      ':not(.xlsx-x)',
      ':is(body, .xlsx-x)',
      ':where(.xlsx-x)',
      'div[data-x=".xlsx-y"]',
      '.xlsx-editor ~ div',
      '.xlsx-editor + *',
      '.xlsx-editor +',
      '> .xlsx-x',
    ])
      expect([selector, selectorScoped(selector)]).toEqual([selector, false]);
  });

  test('a nested selector may not step to a host sibling', () => {
    expect(nestedSelectorScoped('&:hover')).toBe(true);
    expect(nestedSelectorScoped('.xlsx-a + .xlsx-b')).toBe(true);
    expect(nestedSelectorScoped('button')).toBe(true);
    expect(nestedSelectorScoped('& + p')).toBe(false);
    expect(nestedSelectorScoped('+ p')).toBe(false);
    expect(nestedSelectorScoped('~ .xlsx-x')).toBe(true);
    expect(nestedSelectorScoped('&.xlsx-on > span')).toBe(true);
    expect(nestedSelectorScoped('body:has(&)')).toBe(false);
    expect(nestedSelectorScoped(':is(body, &)')).toBe(false);
    expect(nestedSelectorScoped('div:has(> &)')).toBe(false);
    expect(nestedSelectorScoped('.host &')).toBe(false);
  });

  test('host-wide rules, names, layers, and imports fail with their line', () => {
    const css = [
      '.xlsx-editor { color: red; }',
      'body, .xlsx-x { margin: 0; }',
      '@media (min-width: 1px) { html { color: red; } .xlsx-y { color: red; } }',
      '@keyframes spin { to { rotate: 1turn; } }',
      '@keyframes xlsx-spin { to { rotate: 1turn; } }',
      "@import 'host.css';",
      "@font-face { font-family: X; src: url('x.woff2'); }",
      '.xlsx-z { color: red; &:hover { color: red; } @starting-style { opacity: 0; } }',
      '.xlsx-z { & + p { color: red; } }',
      '.xlsx-a { @media (x) { body:has(&) { margin: 0; } } }',
      '@layer base, components;',
      '@layer xlsx-chrome { .xlsx-a { color: red; } }',
      '@scope (.xlsx-editor) { button { color: red; } }',
      '@scope (body) { .xlsx-a { color: red; } }',
      ".xlsx-q[data-a='}'] { content: '{'; }",
      '/* body { color: red; } */',
      '.xlsx-s::after { content: "/*"; } body { color: red; } /* */',
    ].join('\n');
    expect(cssScopeErrors(css)).toEqual([
      '2: "body" is not scoped to an xlsx- class',
      '3: "html" is not scoped to an xlsx- class',
      '4: @keyframes spin needs an xlsx- name',
      '6: @import is not allowed here',
      '7: @font-face is not allowed',
      '9: "& + p" is not scoped to an xlsx- class',
      '10: "body:has(&)" is not scoped to an xlsx- class',
      '11: @layer base needs an xlsx- name',
      '11: @layer components needs an xlsx- name',
      '14: @scope (body) needs an xlsx- root',
      '17: "body" is not scoped to an xlsx- class',
    ]);
  });
});
