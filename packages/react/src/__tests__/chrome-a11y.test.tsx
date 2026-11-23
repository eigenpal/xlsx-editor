/** @jsxImportSource react */
// Accessibility of the editor chrome in a high contrast theme, after
// pptx-editor/packages/react/src/__tests__/chrome-a11y.test.tsx at 0057839b (EigenPal): a rule
// of the forced-colors block names each state and sets its system colors, the element in each
// state matches that rule and the element out of it does not, and each color swatch is a 24px
// target. The Vue twin is packages/vue/test/chrome-a11y.test.ts. The stylesheet checks are
// shared, so they run in this file only. Difference from PPTX Editor: the stylesheet check
// reads the declarations of each rule, not only its selector text.

import { readFileSync } from 'node:fs';
import { afterEach, describe, expect, test } from 'bun:test';
import { act, cleanup, fireEvent, render, waitFor, within } from '@testing-library/react';
import { buildXlsx, c, row } from '../../../../scripts/fixtures/builder.ts';
import { XlsxEditor } from '../index.ts';
import { FONTS, workbook } from './toolbar-stubs.tsx';
import { FORCED_STATES } from '../../../core/src/editor/__tests__/forced-colors-states.ts';

afterEach(cleanup);

/** The stylesheet with no comments. */
const css = readFileSync(
  new URL('../../../core/src/styles/editor.css', import.meta.url),
  'utf8'
).replace(/\/\*[\s\S]*?\*\//g, '');

const squash = (text: string) =>
  text.replace(/\s+/g, ' ').replace(/\( /g, '(').replace(/ \)/g, ')').trim();

/** The selectors of a selector list, split at the commas outside parentheses. */
function selectorsOf(list: string): string[] {
  const out: string[] = [];
  let depth = 0;
  let start = 0;
  for (let i = 0; i < list.length; i += 1) {
    if (list[i] === '(') depth += 1;
    else if (list[i] === ')') depth -= 1;
    else if (list[i] === ',' && depth === 0) {
      out.push(squash(list.slice(start, i)));
      start = i + 1;
    }
  }
  out.push(squash(list.slice(start)));
  return out;
}

interface Rule {
  /** The offset of the rule in the stylesheet, for source order. */
  readonly at: number;
  readonly selectors: readonly string[];
  readonly declarations: readonly string[];
}

/**
 * The style rules of `css` from `from` to the brace that closes the region, in source order.
 * An at-rule, or a rule with a nested block, is not read: `nested` counts them.
 */
function rulesIn(from: number): { rules: Rule[]; nested: number; end: number } {
  const rules: Rule[] = [];
  let nested = 0;
  let depth = 0;
  let start = from;
  let prelude = '';
  let body = from;
  for (let i = from; i < css.length; i += 1) {
    if (css[i] === ';' && depth === 0) {
      // A statement at-rule, such as `@layer name;`, ends at its semicolon.
      start = i + 1;
    } else if (css[i] === '{') {
      if (depth === 0) [prelude, body] = [css.slice(start, i), i + 1];
      depth += 1;
    } else if (css[i] === '}') {
      if (depth === 0) return { rules, nested, end: i };
      depth -= 1;
      if (depth > 0) continue;
      const text = css.slice(body, i);
      if (prelude.trim().startsWith('@') || text.includes('{')) nested += 1;
      else {
        const declarations = text.split(';').map(squash);
        rules.push({
          at: start,
          selectors: selectorsOf(prelude),
          declarations: declarations.filter((d) => d !== ''),
        });
      }
      start = i + 1;
    }
  }
  return { rules, nested, end: css.length };
}

/** The top-level style rules of the stylesheet, outside every at-rule. */
const topRules = rulesIn(0).rules;

/**
 * The rules of every forced-colors block of the stylesheet, in source order. Each block runs
 * from its `@media` line to its matching brace, so a rule outside the blocks is not read.
 * `nested` counts the at-rules and nested blocks inside them, which are not read. A hand parser,
 * because the CSSOM of happy-dom drops system color values.
 */
const { forcedRules, nested } = ((): { forcedRules: Rule[]; nested: number } => {
  const opening = '@media (forced-colors: active) {';
  const all: Rule[] = [];
  let count = 0;
  for (let at = css.indexOf(opening); at >= 0; at = css.indexOf(opening, at + 1)) {
    const block = rulesIn(at + opening.length);
    all.push(...block.rules);
    count += block.nested;
  }
  return { forcedRules: all, nested: count };
})();

/** The declarations of every rule of `rules` that names `selector`, in source order. */
const declarationsOf = (selector: string, rules: readonly Rule[] = forcedRules): string[] =>
  rules.filter((r) => r.selectors.includes(squash(selector))).flatMap((r) => r.declarations);

/**
 * The top-level rules and the forced-colors rules in source order: in a high contrast theme,
 * the later of two rules with one selector wins.
 */
const appliedRules = [...topRules, ...forcedRules].sort((a, b) => a.at - b.at);

/** Properties whose name starts with the name of another property, but which it does not set. */
const NOT_LONGHANDS = new Set(['outline-offset', 'color-scheme']);

/** The name parts of a box side or axis, as in `border-top` and `padding-inline-start`. */
const SIDES = new Set(['top', 'right', 'bottom', 'left', 'inline', 'block', 'start', 'end']);
const PHYSICAL = new Set(['top', 'right', 'bottom', 'left']);

/** A property as its family, the box sides it names, and its aspect (empty for all aspects). */
function partsOf(name: string) {
  const [family = '', ...rest] = name.split('-');
  return {
    family,
    sides: rest.filter((p) => SIDES.has(p)),
    aspect: rest.filter((p) => !SIDES.has(p)).join('-'),
  };
}

/**
 * Whether a declaration of `name` sets some of `property`: a shorthand or a longhand of it, as
 * `background` and `background-color`, or `border-top` and `border-color`. Two sides overlap
 * unless both are physical or both logical and they differ. A shorthand with no aspect does not
 * set a radius.
 */
function overlaps(name: string, property: string): boolean {
  if (NOT_LONGHANDS.has(name) || NOT_LONGHANDS.has(property)) return false;
  const [a, b] = [partsOf(name), partsOf(property)];
  if (a.family !== b.family) return false;
  const radius = [a.aspect, b.aspect].filter((x) => x.includes('radius')).length;
  if (radius === 1) return false;
  if (a.aspect !== '' && b.aspect !== '' && a.aspect !== b.aspect) return false;
  const [x, y] = [a.sides[0], b.sides[0]];
  if (x === undefined || y === undefined || x === y) return true;
  return PHYSICAL.has(x) !== PHYSICAL.has(y);
}

const forcedSet = new Set(forcedRules);

/**
 * The value of `property` for `selector`: the last declaration of it wins, as in the cascade.
 * A later declaration that sets some of it, such as `background-color` after `background` or
 * `border-top` after `border-color`, replaces it with a conflict that no expected value equals.
 * Read from the applied rules, the value must come from a forced-colors rule, because a
 * top-level rule also paints the normal theme.
 */
function valueOf(
  selector: string,
  property: string,
  rules: readonly Rule[] = appliedRules
): string | undefined {
  let value: string | undefined;
  for (const rule of rules) {
    if (!rule.selectors.includes(squash(selector))) continue;
    for (const d of rule.declarations) {
      const colon = d.indexOf(':');
      const name = d.slice(0, colon).trim();
      if (name === property)
        value =
          rules === appliedRules && !forcedSet.has(rule)
            ? `outside the forced-colors block: ${d}`
            : d.slice(colon + 1).trim();
      else if (overlaps(name, property)) value = `overridden by ${d}`;
    }
  }
  return value;
}

async function open(bytes = workbook()) {
  const view = render(
    <div style={{ height: 400 }}>
      <XlsxEditor document={bytes} fonts={FONTS} />
    </div>
  );
  // Query only the mounted container: the lists and menus render inside it.
  const scoped = within(view.container);
  await waitFor(() => expect(scoped.getAllByRole('tab').length).toBeGreaterThan(0), {
    timeout: 5000,
  });
  return scoped;
}

describe('the stylesheet', () => {
  test('every forced-colors rule is read, and none nests or is important', () => {
    expect(forcedRules.length).toBeGreaterThan(Object.keys(FORCED_STATES).length);
    expect(nested).toBe(0);
    // An `!important` declaration, also outside the block, wins over the order that the state
    // check reads.
    expect(css.match(/!\s*important/gi)).toBeNull();
  });

  test('the forced-colors rules give each state its system colors', () => {
    for (const [state, { selector, declarations }] of Object.entries(FORCED_STATES)) {
      for (const d of declarations) {
        const colon = d.indexOf(':');
        const property = d.slice(0, colon);
        const want = { state, property, value: d.slice(colon + 1).trim() };
        expect({ state, property, value: valueOf(selector, property) }).toEqual(want);
      }
    }
  });

  test('a hovered control in a highlighted state keeps the state color', () => {
    // The toolbar hover rules have a specificity of (0,5,1) at most; this rule has (0,6,1).
    const prefix =
      ".xlsx-toolbar button:is(:hover, :focus-visible):not(:disabled):not([aria-disabled='true']):is(";
    const all = forcedRules.filter((r) => r.selectors.some((s) => s.startsWith(prefix)));
    expect(all).toHaveLength(1);
    const hovered = all[0];
    expect(hovered?.declarations).toEqual([
      'forced-color-adjust: none',
      'background: Highlight',
      'color: HighlightText',
    ]);
    expect(hovered?.selectors).toHaveLength(1);
    const inner = hovered!.selectors[0]!;
    expect(inner.endsWith(')')).toBe(true);
    expect(selectorsOf(inner.slice(prefix.length, -1))).toEqual([
      ".xlsx-toolbar-row[role='menuitem']",
      ".xlsx-toolbar-no-fill[aria-selected='true']",
      ".xlsx-toolbar-button[aria-pressed='true']",
      '.xlsx-toolbar-button[data-active]',
      ".xlsx-toolbar-row[aria-pressed='true']",
      '.xlsx-toolbar-row[data-active]',
    ]);
  });

  test('the item under a list never takes the fill of the selected option', () => {
    const fills = forcedRules
      .filter((r) => r.selectors.some((s) => s.includes('.xlsx-toolbar-list-more')))
      .flatMap((r) => r.declarations)
      .filter((d) => d.startsWith('background'));
    expect(fills).toEqual([]);
  });

  test('a highlighted state gives every part inside it the highlight text color', () => {
    const { pressed, selectedTab, selectedOption } = FORCED_STATES;
    for (const state of [pressed, selectedTab, selectedOption]) {
      const parts = `${state.selector} *`;
      expect([valueOf(parts, 'color'), valueOf(parts, 'fill')]).toEqual([
        'inherit',
        'currentColor',
      ]);
    }
  });

  test('a disabled toggle never takes the pressed state, and an active one does', () => {
    // The adapters render no disabled toggle in the pressed state, so plain elements hold the
    // selector shape here.
    const { pressed, disabledButton } = FORCED_STATES;
    const bar = document.createElement('div');
    bar.className = 'xlsx-toolbar';
    const button = (attributes: Record<string, string>) => {
      const el = document.createElement('button');
      el.className = 'xlsx-toolbar-button';
      for (const [name, value] of Object.entries(attributes)) el.setAttribute(name, value);
      return bar.appendChild(el);
    };
    const disabled = button({ 'aria-pressed': 'true', 'aria-disabled': 'true' });
    expect(disabled.matches(pressed.selector)).toBe(false);
    expect(disabled.matches(disabledButton.selector)).toBe(true);
    expect(button({ 'data-active': '' }).matches(pressed.selector)).toBe(true);
    expect(button({ 'aria-pressed': 'false' }).matches(pressed.selector)).toBe(false);
  });

  test('a color swatch is a 24px target', () => {
    // Every rule, also in the forced-colors blocks, that sizes the palette or a swatch.
    const sized = (selector: string, name: RegExp) =>
      declarationsOf(selector, appliedRules).filter((d) => name.test(d.slice(0, d.indexOf(':'))));
    expect(sized('.xlsx-toolbar-palette', /^grid-template/)).toEqual([
      'grid-template-columns: repeat(7, 24px)',
    ]);
    expect(
      sized('.xlsx-toolbar-swatch', /^(min-|max-)?(width|height|inline-size|block-size)$/)
    ).toEqual(['width: 24px', 'height: 24px']);
  });
});

describe('the chrome in each state matches its forced-colors rule', () => {
  const { pressed, selectedTab, selectedOption, focusedOption, disabledButton } = FORCED_STATES;

  test('a pressed toggle, the selected tab, and a disabled button', async () => {
    const view = await open(
      buildXlsx({
        sheets: [
          { name: 'S', rows: row(1, [c('A1', 1)]) },
          { name: 'T', rows: row(1, [c('A1', 2)]) },
        ],
      })
    );
    const bold = view.getByRole('button', { name: 'Bold' });
    expect(bold.matches(pressed.selector)).toBe(false);
    fireEvent.click(bold);
    expect(bold.getAttribute('aria-pressed')).toBe('true');
    expect(bold.matches(pressed.selector)).toBe(true);
    expect(view.getByRole('tab', { name: 'S' }).matches(selectedTab.selector)).toBe(true);
    expect(view.getByRole('tab', { name: 'T' }).matches(selectedTab.selector)).toBe(false);
    const { focusedTab, focusedSelectedTab } = FORCED_STATES;
    const [first, second] = [
      view.getByRole('tab', { name: 'S' }),
      view.getByRole('tab', { name: 'T' }),
    ];
    act(() => first.focus());
    expect(first.matches(focusedTab.selector)).toBe(true);
    expect(first.matches(focusedSelectedTab.selector)).toBe(true);
    act(() => second.focus());
    expect(second.matches(focusedTab.selector)).toBe(true);
    expect(second.matches(focusedSelectedTab.selector)).toBe(false);
    expect(view.getByRole('button', { name: 'Redo' }).matches(disabledButton.selector)).toBe(true);
    expect(view.getByRole('button', { name: 'Undo' }).matches(disabledButton.selector)).toBe(false);
  });

  test('the selected and the focused list option', async () => {
    const view = await open();
    fireEvent.click(view.getByRole('button', { name: /^Number format/ }));
    const general = view.getByRole('option', { name: 'General' });
    const number = view.getByRole('option', { name: 'Number' });
    expect(general.getAttribute('aria-selected')).toBe('true');
    expect(general.matches(selectedOption.selector)).toBe(true);
    expect(number.matches(selectedOption.selector)).toBe(false);
    act(() => number.focus());
    expect(number.matches(focusedOption.selector)).toBe(true);
    expect(general.matches(focusedOption.selector)).toBe(false);
    // The item under the list takes the ring of a focused option, never the selected fill.
    const more = view.getByRole('button', { name: 'More number formats…' });
    expect(more.matches(FORCED_STATES.focusedListMore.selector)).toBe(false);
    act(() => more.focus());
    expect(more.matches(FORCED_STATES.focusedListMore.selector)).toBe(true);
    expect(more.matches(selectedOption.selector)).toBe(false);
    // The selected option with focus: its ring takes the text color of its fill.
    act(() => general.focus());
    expect(general.matches(FORCED_STATES.focusedSelectedOption.selector)).toBe(true);
    expect(number.matches(FORCED_STATES.focusedSelectedOption.selector)).toBe(false);
  });

  test('the selected color swatch and the selected no fill', async () => {
    const { swatch, selectedSwatch, selectedNoFill, focusedSwatch } = FORCED_STATES;
    const view = await open();
    fireEvent.click(view.getByRole('button', { name: 'Font color' }));
    const red = view.getByRole('option', { name: 'Red' });
    expect(red.matches(swatch.selector)).toBe(true);
    expect(red.matches(selectedSwatch.selector)).toBe(false);
    fireEvent.click(red);
    fireEvent.click(view.getByRole('button', { name: 'Font color' }));
    expect(view.getByRole('option', { name: 'Red' }).matches(selectedSwatch.selector)).toBe(true);
    expect(view.getByRole('option', { name: 'Blue' }).matches(selectedSwatch.selector)).toBe(false);
    fireEvent.click(view.getByRole('button', { name: 'Fill color' }));
    // A cell with no fill: the no fill option is the selected one.
    const none = view.getByRole('option', { name: 'No fill' });
    expect(none.getAttribute('aria-selected')).toBe('true');
    expect(none.matches(selectedNoFill.selector)).toBe(true);
    expect(none.matches(selectedSwatch.selector)).toBe(false);
    expect(none.matches(focusedSwatch.selector)).toBe(false);
    act(() => none.focus());
    expect(none.matches(focusedSwatch.selector)).toBe(true);
  });

  test('a focused color swatch differs from the selected one', async () => {
    const { selectedSwatch, focusedSwatch } = FORCED_STATES;
    const view = await open();
    fireEvent.click(view.getByRole('button', { name: 'Font color' }));
    fireEvent.click(view.getByRole('option', { name: 'Red' }));
    fireEvent.click(view.getByRole('button', { name: 'Font color' }));
    const red = view.getByRole('option', { name: 'Red' });
    const blue = view.getByRole('option', { name: 'Blue' });
    act(() => blue.focus());
    expect(blue.matches(focusedSwatch.selector)).toBe(true);
    expect(blue.matches(selectedSwatch.selector)).toBe(false);
    expect(red.matches(selectedSwatch.selector)).toBe(true);
    expect(red.matches(focusedSwatch.selector)).toBe(false);
    // The selected swatch keeps the line inside its edge while it has focus.
    act(() => red.focus());
    expect(red.matches(focusedSwatch.selector)).toBe(true);
    expect(red.matches(selectedSwatch.selector)).toBe(true);
  });

  test('the focused and the disabled sheet menu item', async () => {
    const { focusedMenuItem, disabledMenuItem } = FORCED_STATES;
    const view = await open();
    fireEvent.contextMenu(view.getByRole('tab', { name: 'S' }));
    // The only sheet cannot be deleted.
    const remove = view.getByRole('menuitem', { name: 'Delete' });
    const rename = view.getByRole('menuitem', { name: 'Rename' });
    expect(remove.matches(disabledMenuItem.selector)).toBe(true);
    expect(rename.matches(disabledMenuItem.selector)).toBe(false);
    act(() => rename.focus());
    expect(rename.matches(focusedMenuItem.selector)).toBe(true);
    expect(rename.matches(FORCED_STATES.focusedMenuRow.selector)).toBe(true);
    act(() => remove.focus());
    expect(remove.matches(focusedMenuItem.selector)).toBe(false);
    expect(remove.matches(FORCED_STATES.focusedMenuRow.selector)).toBe(false);
  });
});
