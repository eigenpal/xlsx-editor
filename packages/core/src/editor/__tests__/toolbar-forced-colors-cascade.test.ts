// Which rule of `packages/core/src/styles/editor.css` sets the text color of a toolbar element,
// with and without forced colors. A text test of each rule cannot see that a stronger rule of the
// normal styles beats a rule of the forced-colors block, so this test runs the cascade: it reads
// every rule that matches the element (happy-dom `matches`), and the rule with the highest
// selector weight wins, then the later rule. A state such as `:hover` is an attribute here.

import { describe, expect, test } from 'bun:test';
import { Window } from 'happy-dom';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';

const css = readFileSync(join(import.meta.dir, '../../styles/editor.css'), 'utf8');
const FORCED = '(forced-colors: active)';

interface Rule {
  readonly selector: string;
  readonly media: string | null;
  readonly order: number;
  readonly declarations: ReadonlyMap<string, string>;
}

/** Splits `text` at each `separator` outside parentheses, brackets, and quotes. */
function splitTop(text: string, separator: string): string[] {
  const parts: string[] = [];
  let depth = 0;
  let quote = '';
  let start = 0;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i]!;
    if (quote) {
      if (ch === quote) quote = '';
    } else if (ch === "'" || ch === '"') quote = ch;
    else if (ch === '(' || ch === '[') depth++;
    else if (ch === ')' || ch === ']') depth--;
    else if (ch === separator && depth === 0) {
      parts.push(text.slice(start, i));
      start = i + 1;
    }
  }
  parts.push(text.slice(start));
  return parts.map((p) => p.trim()).filter(Boolean);
}

/** The rules of a stylesheet in source order, each with the `@media` condition around it. */
function parseRules(text: string, media: string | null = null, rules: Rule[] = []): Rule[] {
  let i = 0;
  while (i < text.length) {
    const open = text.indexOf('{', i);
    if (open < 0) break;
    let depth = 1;
    let close = open + 1;
    for (; close < text.length && depth > 0; close++) {
      if (text[close] === '{') depth++;
      else if (text[close] === '}') depth--;
    }
    const prelude = text.slice(i, open).trim();
    const body = text.slice(open + 1, close - 1);
    if (prelude.startsWith('@media')) parseRules(body, prelude.slice(6).trim(), rules);
    else if (!prelude.startsWith('@')) {
      const declarations = new Map<string, string>();
      for (const d of splitTop(body, ';')) {
        const colon = d.indexOf(':');
        declarations.set(d.slice(0, colon).trim(), d.slice(colon + 1).trim());
      }
      for (const selector of splitTop(prelude, ','))
        rules.push({
          selector: selector.replace(/\s+/g, ' '),
          media,
          order: rules.length,
          declarations,
        });
    }
    i = close;
  }
  return rules;
}

type Weight = [number, number, number];
/** The part of a happy-dom element that the cascade reads. */
interface Matcher {
  matches(selector: string): boolean;
}

/** The selector weight (Selectors Level 4): `:is`, `:not`, and `:has` take their strongest
 * argument, and `:where` takes none. */
function weight(selector: string): Weight {
  const w: Weight = [0, 0, 0];
  const add = (v: Weight) => v.forEach((n, k) => (w[k] += n));
  let i = 0;
  const ident = () => {
    const m = /^[\w-]+/.exec(selector.slice(i));
    i += m ? m[0].length : 0;
    return m ? m[0] : '';
  };
  while (i < selector.length) {
    const ch = selector[i]!;
    if (ch === '#' || ch === '.') {
      i++;
      ident();
      w[ch === '#' ? 0 : 1] += 1;
    } else if (ch === '[') {
      i = selector.indexOf(']', i) + 1;
      w[1] += 1;
    } else if (ch === ':') {
      const element = selector[i + 1] === ':';
      i += element ? 2 : 1;
      const name = ident();
      let args: string | null = null;
      if (selector[i] === '(') {
        let depth = 0;
        const start = i + 1;
        for (; i < selector.length; i++) {
          if (selector[i] === '(') depth++;
          else if (selector[i] === ')' && --depth === 0) break;
        }
        args = selector.slice(start, i++);
      }
      if (element || name === 'before' || name === 'after') w[2] += 1;
      else if (name === 'where') continue;
      else if (args !== null && ['is', 'not', 'has'].includes(name))
        add(splitTop(args, ',').map(weight).sort(compare).at(-1)!);
      else w[1] += 1;
    } else if (/[\w-]/.test(ch)) {
      ident();
      w[2] += 1;
    } else i++;
  }
  return w;
}

function compare(a: Weight, b: Weight): number {
  return a[0] - b[0] || a[1] - b[1] || a[2] - b[2];
}

const RULES = parseRules(css.replace(/\/\*[\s\S]*?\*\//g, ''));
/** A user state as an attribute, so that `matches` can test it on a detached element. */
const STATE = /:(hover|focus-visible|focus|active)(?![\w-])/g;

/** The value of `property` that the cascade gives the element, and the rule that sets it. */
function cascade(element: Matcher, property: string, forced: boolean) {
  let best: { rule: Rule; weight: Weight } | null = null;
  for (const rule of RULES) {
    if (!rule.declarations.has(property) || rule.selector.includes('::')) continue;
    if (rule.media !== null && !(forced && rule.media === FORCED)) continue;
    if (!element.matches(rule.selector.replace(STATE, '[data-state-$1]'))) continue;
    const w = weight(rule.selector);
    if (!best || compare(w, best.weight) >= 0) best = { rule, weight: w };
  }
  return { value: best?.rule.declarations.get(property), selector: best?.rule.selector };
}

/** One window for every case: each case builds a detached tree in its document. */
const DOCUMENT = new Window().document;

/** The element marked `data-test` in `markup`, inside an editor root. */
function editor(markup: string): Matcher {
  const root = DOCUMENT.createElement('div');
  root.className = 'xlsx-editor';
  root.innerHTML = markup;
  return root.querySelector('[data-test]')!;
}

function toolbar(markup: string): Matcher {
  return editor(`<div class="xlsx-toolbar"><div class="xlsx-toolbar-group">${markup}</div></div>`);
}

const DISABLED = {
  'a bar button': '<button class="xlsx-toolbar-button" aria-disabled="true" data-test></button>',
  'a pressed bar button': `<button class="xlsx-toolbar-button" aria-pressed="true" data-active
    aria-disabled="true" data-test></button>`,
  'a field': '<button class="xlsx-toolbar-select" aria-disabled="true" data-test></button>',
  'a color trigger': '<button class="xlsx-toolbar-color" disabled data-test></button>',
  'a focused bar button': `<button class="xlsx-toolbar-button" aria-disabled="true"
    data-state-focus-visible data-test></button>`,
  'a panel row': '<button class="xlsx-toolbar-row" aria-disabled="true" data-test></button>',
  'a group menu row': `<button class="xlsx-toolbar-row" role="menuitem" aria-disabled="true"
    data-test></button>`,
  'the more formats row': `<div class="xlsx-toolbar-list"><button class="xlsx-toolbar-list-more"
    aria-disabled="true" data-test></button></div>`,
  'an active panel row': `<button class="xlsx-toolbar-row" data-active aria-disabled="true"
    data-test></button>`,
  'a natively disabled row': '<button class="xlsx-toolbar-row" disabled data-test></button>',
  'a focused panel row': `<button class="xlsx-toolbar-row" aria-disabled="true"
    data-state-focus-visible data-test></button>`,
  'a hovered pressed bar button': `<button class="xlsx-toolbar-button" data-active
    aria-disabled="true" data-state-hover data-test></button>`,
  "a host's own button": '<button aria-disabled="true" data-test></button>',
};

/** Disabled controls outside the toolbar, which the same forced-colors rule covers. */
const OUTSIDE = {
  'a sheet menu row':
    '<button class="xlsx-sheet-menu-item" aria-disabled="true" data-test></button>',
  'a sheet tab': `<div class="xlsx-sheet-tab" aria-selected="true" aria-disabled="true"
    data-test></div>`,
  'the add sheet button': '<button class="xlsx-sheet-add" aria-disabled="true" data-test></button>',
  'a focused add sheet button': `<button class="xlsx-sheet-add" aria-disabled="true"
    data-state-focus-visible data-test></button>`,
  'a zoom button': `<div class="xlsx-zoom"><button aria-disabled="true" data-test></button></div>`,
};

describe('the text color that wins on a toolbar element', () => {
  test('the stylesheet parses into rules, and the weights follow the selector rules', () => {
    // Every rule outside `@container` (no toolbar rule is in one), and the forced-colors block.
    expect(RULES.length).toBeGreaterThan(250);
    expect(
      RULES.some((r) => r.media === FORCED && r.declarations.get('color') === 'GrayText')
    ).toBe(true);
    expect(weight(".xlsx-toolbar button:is(:disabled, [aria-disabled='true'])")).toEqual([0, 2, 1]);
    expect(weight('.a:where(.b, .c > .d)::before')).toEqual([0, 1, 1]);
    expect(weight('.a :not(.b, #c)')).toEqual([1, 1, 0]);
  });

  test('no rule that the cascade skips sets a toolbar color', () => {
    // The cascade reads plain rules and `@media` blocks only. Other media, other blocks such as
    // `@supports` and `@layer`, and `!important` would change the answer without a sign.
    const skipped = RULES.filter(
      (r) =>
        r.media !== null &&
        r.media !== FORCED &&
        r.declarations.has('color') &&
        /xlsx-(toolbar|sheet|zoom)/.test(r.selector)
    );
    expect(skipped.map((r) => r.selector)).toEqual([]);
    expect(css).not.toMatch(/@(supports|layer)\b/);
    expect(css).not.toMatch(/!important/);
  });

  for (const [name, markup] of Object.entries(DISABLED))
    test(`forced colors give ${name} GrayText when it is disabled`, () => {
      const found = cascade(toolbar(markup), 'color', true);
      expect(found).toEqual({ value: 'GrayText', selector: expect.any(String) });
    });

  for (const [name, markup] of Object.entries(OUTSIDE))
    test(`forced colors give ${name} GrayText when it is disabled`, () => {
      expect(cascade(editor(markup), 'color', true).value).toBe('GrayText');
    });

  test('forced colors keep the 0.35 fade of a disabled control, as PPTX Editor does', () => {
    for (const name of ['a bar button', 'a field', 'a color trigger'] as const)
      expect(cascade(toolbar(DISABLED[name]), 'opacity', true).value).toBe('0.35');
    expect(cascade(toolbar(DISABLED['a focused bar button']), 'opacity', true).value).toBe('1');
    expect(cascade(toolbar(DISABLED['a panel row']), 'opacity', true).value).toBeUndefined();
    expect(cascade(editor(OUTSIDE['the add sheet button']), 'opacity', true).value).toBe('0.35');
    expect(cascade(editor(OUTSIDE['a focused add sheet button']), 'opacity', true).value).toBe('1');
  });

  test('without forced colors, a disabled row takes the subtle color and a control fades', () => {
    for (const name of ['a panel row', 'an active panel row', 'the more formats row'] as const)
      expect(cascade(toolbar(DISABLED[name]), 'color', false).value).toBe('var(--doc-text-subtle)');
    expect(cascade(toolbar(DISABLED['a focused bar button']), 'color', false).value).toBe(
      'var(--doc-text-subtle)'
    );
    expect(cascade(toolbar(DISABLED['a bar button']), 'color', false).value).toBe(
      'var(--doc-text-muted)'
    );
  });

  test('forced colors keep HighlightText on a hovered pressed button and its check mark', () => {
    const pressed = toolbar(
      '<button class="xlsx-toolbar-button" aria-pressed="true" data-state-hover data-test></button>'
    );
    expect(cascade(pressed, 'color', true).value).toBe('HighlightText');
    const option = toolbar(`<div class="xlsx-toolbar-list"><div role="option"
      aria-selected="true"><svg class="xlsx-toolbar-check" data-test></svg></div></div>`);
    expect(cascade(option, 'color', true).value).toBe('inherit');
  });
});

describe('the colors that win on a checked menu row', () => {
  const row = (attrs: string) =>
    editor(`<div class="xlsx-menubar"><div class="xlsx-menu-panel" role="menu">
      <button class="xlsx-menu-row" role="menuitemcheckbox" aria-checked="true" ${attrs}
        data-test></button></div></div>`);

  test('forced colors give a focused or hovered checked row HighlightText on Highlight', () => {
    for (const state of ['data-state-focus', 'data-state-hover']) {
      expect(cascade(row(state), 'color', true).value).toBe('HighlightText');
      expect(cascade(row(state), 'background', true).value).toBe('Highlight');
    }
  });

  test('forced colors give a disabled checked row GrayText, and a checked row at rest its text color', () => {
    expect(cascade(row('aria-disabled="true"'), 'color', true).value).toBe('GrayText');
    expect(cascade(row(''), 'color', true).value).toBe('inherit');
  });

  test('without forced colors, a disabled checked row takes the subtle color', () => {
    expect(cascade(row('aria-disabled="true"'), 'color', false).value).toBe(
      'var(--doc-text-subtle)'
    );
    expect(cascade(row(''), 'color', false).value).toBe('var(--doc-primary)');
  });
});
