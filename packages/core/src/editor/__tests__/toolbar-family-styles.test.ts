// The toolbar rules of `packages/core/src/styles/editor.css` that take the values of
// pptx-editor/packages/core/src/styles/editor.css at d4502552 (EigenPal): the separator, the
// control states, the picker field, the caret, and the list rows. A change of one of these
// values is a change of the family look, so it changes this test too.

import { describe, expect, test } from 'bun:test';
import { Window } from 'happy-dom';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';

const css = readFileSync(join(import.meta.dir, '../../styles/editor.css'), 'utf8');
/** The stylesheet with each run of white space as one space, so a line wrap changes nothing. */
const flat = css.replace(/\s+/g, ' ');

/** The two places of a group separator: after a group, and before a host's loose children. */
const SEPARATED = [
  '.xlsx-toolbar-group + .xlsx-toolbar-group:not([data-toolbar-fixed])',
  '.xlsx-toolbar > [data-toolbar-fixed]:not(:first-child)',
];
/** The separator rule. `:where` keeps its weight at one class, so a host group class can replace it. */
const SEPARATOR = `.xlsx-toolbar-group:where( ${SEPARATED.join(', ')} )::before`;

/** The declarations of the first rule whose selector is exactly `selector`. */
function rule(selector: string): string {
  const escaped = selector.replace(/[.*+?^${}()|[\]\\]/g, '\\$&').replace(/\s+/g, '\\s+');
  const match = new RegExp(`(?:^|\\}|\\*/)\\s*${escaped}\\s*\\{([^}]*)\\}`, 'm').exec(css);
  if (!match) throw new Error(`no rule for ${selector}`);
  return match[1]!;
}

describe('the toolbar family styles', () => {
  test('the chrome root and the exit cue fall back to the family font stack', () => {
    const stack =
      "font-family: var( --doc-font-ui, -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif );";
    expect(flat.split(stack).length - 1).toBe(2);
  });

  test('a short line separates two groups, and a group has no border', () => {
    const separator = rule(SEPARATOR);
    for (const d of [
      'width: 1px',
      'height: 18px',
      'margin: 0 5px',
      'background: var(--doc-border)',
    ])
      expect(separator).toContain(d);
    expect(css).not.toMatch(/\.xlsx-toolbar-group \+ \.xlsx-toolbar-group\s*\{[^}]*border-left/);
    // Forced colors draw a background in `Canvas`, so the line takes `CanvasText` there.
    expect(flat).toMatch(
      /@media \(forced-colors: active\) \{[^@]*\.xlsx-toolbar > \[data-toolbar-fixed\]:not\(:first-child\) \)::before \{ background: CanvasText; \}/
    );
  });

  test('a host separator draws the same line, across a row of the More panel', () => {
    const line = rule('.xlsx-toolbar-separator');
    for (const d of [
      'width: 1px',
      'height: 18px',
      'margin: 0 5px',
      'background: var(--doc-border)',
    ])
      expect(line).toContain(d);
    const across = rule(".xlsx-toolbar-separator[aria-orientation='horizontal']");
    for (const d of ['width: auto', 'height: 1px']) expect(across).toContain(d);
    expect(flat).toMatch(
      /@media \(forced-colors: active\) \{[^@]*\.xlsx-toolbar-separator \{ background: CanvasText; \}/
    );
  });

  test('the block of loose host children has its line whether or not the More button is shown', () => {
    // The bar measures that block, so a width that changes with the More button makes the two
    // layouts flip without end.
    rule(SEPARATOR);
    const document = new Window().document;
    const bar = document.createElement('div');
    bar.className = 'xlsx-toolbar';
    const element = (className: string, fixed = false) => {
      const e = document.createElement('div');
      e.className = className;
      if (fixed) e.setAttribute('data-toolbar-fixed', '');
      return e;
    };
    const lined = (e: { matches(s: string): boolean }) => SEPARATED.some((s) => e.matches(s));
    const first = element('xlsx-toolbar-group');
    const second = element('xlsx-toolbar-group');
    const more = element('xlsx-toolbar-more');
    const fixed = element('xlsx-toolbar-group', true);
    // Every group fits: the block comes after a group.
    bar.append(first, second, fixed);
    expect([lined(first), lined(second), lined(fixed)]).toEqual([false, true, true]);
    // Some groups move into the More panel: the block comes after the More button.
    second.remove();
    bar.insertBefore(more, fixed);
    expect([lined(first), lined(more), lined(fixed)]).toEqual([false, false, true]);
    // No group fits.
    first.remove();
    expect(lined(fixed)).toBe(true);
  });

  test('a control has a muted glyph at rest, the text color on hover, and the press scale', () => {
    const control = rule('.xlsx-toolbar-button,\n.xlsx-toolbar-select,\n.xlsx-toolbar-color');
    for (const d of [
      'height: 30px',
      'min-width: 30px',
      'border-radius: var(--doc-radius-control, 7px)',
      'color: var(--doc-text-muted)',
    ])
      expect(control).toContain(d);
    expect(flat).toContain(
      ".xlsx-editor :where(.xlsx-toolbar-button, .xlsx-toolbar-select, .xlsx-toolbar-color):hover:not(:disabled):not( [aria-disabled='true'] ) { color: var(--doc-text); }"
    );
    // A panel row keeps its own color on hover, so the active row stays marked. The hover also
    // reaches a `Toolbar.Action` outside the toolbar, by its own class.
    expect(flat).toContain(
      ".xlsx-toolbar button:hover:not(:disabled):not([aria-disabled='true']), .xlsx-toolbar-button:hover:not(:disabled):not([aria-disabled='true']) { background: var(--doc-bg-hover); }"
    );
    // A button and a color trigger press in; a field does not, as the family field does not.
    expect(flat).toContain(
      ".xlsx-editor :is(.xlsx-toolbar-button, .xlsx-toolbar-color):active:not(:disabled):not([aria-disabled='true']) { transform: scale(0.94); }"
    );
    expect(flat).not.toMatch(/\.xlsx-toolbar-select[^{]*:active/);
    // Reduced motion stops the press scale.
    expect(css).toMatch(
      /@media \(prefers-reduced-motion: reduce\) \{\s*\.xlsx-editor :is\(\.xlsx-toolbar-button[^}]*transition: none;[^}]*\}[^}]*transform: none;/
    );
  });

  test('a disabled bar control fades, and forced colors keep the system disabled color', () => {
    // One flat selector for each control, as in PPTX Editor.
    expect(flat).toContain(
      ".xlsx-toolbar-button:is(:disabled, [aria-disabled='true']), .xlsx-toolbar-select:is(:disabled, [aria-disabled='true']), .xlsx-toolbar-color:is(:disabled, [aria-disabled='true']) { opacity: 0.35; }"
    );
    // The fade does not fade the focus ring of a disabled control, which stays focusable.
    expect(flat).toContain(
      ".xlsx-toolbar [aria-disabled='true']:focus-visible, .xlsx-toolbar-button[aria-disabled='true']:focus-visible { opacity: 1; color: var(--doc-text-subtle); }"
    );
    // A disabled panel row takes the subtle text color and no fade, so its reason stays readable.
    expect(flat).toContain(
      ".xlsx-toolbar-row:is(:disabled, [aria-disabled='true']) { color: var(--doc-text-subtle); }"
    );
    // Forced colors repeat each disabled selector, also the focused one, with `GrayText`, and
    // no rule of the forced-colors block resets the fade.
    expect(flat).toMatch(
      /\.xlsx-toolbar \[aria-disabled='true'\]:focus-visible, \.xlsx-toolbar button:is\(:disabled, \[aria-disabled='true'\]\), \.xlsx-toolbar-button:is\([^{]*\.xlsx-toolbar-row:is\(:disabled, \[aria-disabled='true'\]\), \.xlsx-toolbar-list-more\[aria-disabled='true'\],[^{]*\{ color: GrayText; \}/
    );
  });

  test('a picker is a light field with a small muted caret', () => {
    const field = rule('.xlsx-toolbar-select');
    for (const d of [
      'padding: 0 8px',
      'border: 1px solid var(--doc-border)',
      'background: var(--doc-surface)',
      'color: var(--doc-text)',
      'font-size: 13px',
    ])
      expect(field).toContain(d);
    // 152px keeps 116px for the text, as before the padding and caret grew: 2 borders, 2
    // paddings of 8px, the 14px caret, and a 4px gap.
    expect(rule('.xlsx-toolbar-select-wide')).toContain('width: 152px');
    // In a panel row, the wide field takes the room beside its label, so the row does not wrap.
    const inPanel = rule('.xlsx-toolbar-labelled > .xlsx-toolbar-select-wide');
    for (const d of ['flex: 1 1 0', 'min-width: 30px', 'max-width: 152px'])
      expect(inPanel).toContain(d);
    // 72px leaves 36px for the size: 2 borders, 2 paddings of 8px, the 14px caret, a 4px gap.
    expect(rule('.xlsx-toolbar-select-narrow')).toContain('width: 72px');
    const caret = rule('.xlsx-toolbar svg.xlsx-toolbar-caret');
    for (const d of ['width: 14px', 'height: 14px', 'opacity: 0.6']) expect(caret).toContain(d);
  });

  test('a list row is 32px tall, and a check mark, not a fill, marks the selected option', () => {
    const option = rule(".xlsx-toolbar-list [role='option']");
    for (const d of [
      'min-height: 32px',
      'padding: 6px 10px',
      'font-size: 13px',
      'line-height: 20px',
      'border-radius: var(--doc-radius-control, 6px)',
    ])
      expect(option).toContain(d);
    const check = rule('.xlsx-toolbar svg.xlsx-toolbar-check');
    for (const d of ['flex: none', 'margin-inline-start: auto', 'color: var(--doc-text-muted)'])
      expect(check).toContain(d);
    expect(css).not.toMatch(/^\.xlsx-toolbar-list \[role='option'\]\[aria-selected='true'\]\s*\{/m);
  });

  test('the number format list shows its 11 formats, and every row starts its text at 10px', () => {
    expect(rule('.xlsx-toolbar-list')).toContain('max-height: 360px');
    const more = rule('.xlsx-toolbar-list-more');
    for (const d of ['min-height: 32px', 'padding: 6px 10px', 'font-size: 13px'])
      expect(more).toContain(d);
    // Forced colors: the edge of the current option and its padding add up to the 10px padding.
    expect(flat).toContain(
      ".xlsx-toolbar-list [role='option'][aria-current='true'] { border-inline-start: 3px solid Highlight; padding-inline-start: 7px; }"
    );
    expect(flat).toContain(
      ".xlsx-toolbar-list [role='option'][aria-current='true']:is(:hover, :focus) { border-inline-start-width: 5px; padding-inline-start: 5px; }"
    );
  });

  test('the radius fallbacks are the family values', () => {
    expect(css).not.toMatch(
      /--doc-radius-control, 4px|--doc-radius-panel, 8px|--doc-radius-pill, 14px/
    );
    expect(css).toContain('var(--doc-radius-pill, 999px)');
    expect(css).toContain('var(--doc-radius-panel, 12px)');
    // A menu row takes the family row radius.
    for (const selector of ['.xlsx-sheet-menu-item', '.xlsx-toolbar-row'])
      expect(rule(selector)).toContain('border-radius: var(--doc-radius-control, 6px)');
  });

  test('a disabled add-sheet button fades as a disabled toolbar button, with no hover', () => {
    const faded = rule(".xlsx-sheet-add[aria-disabled='true']");
    for (const d of ['opacity: 0.35', 'cursor: default']) expect(faded).toContain(d);
    const focused = rule(".xlsx-sheet-add[aria-disabled='true']:focus-visible");
    for (const d of ['opacity: 1', 'color: var(--doc-text-subtle)']) expect(focused).toContain(d);
    expect(css).not.toMatch(/\.xlsx-sheet-add:hover\s*\{/);
    expect(() => rule(".xlsx-sheet-add:hover:not([aria-disabled='true'])")).not.toThrow();
    // In forced colors, the focused disabled button keeps `GrayText` over its focus rule.
    expect(flat).toMatch(
      /@media \(forced-colors: active\) \{[^@]*\.xlsx-sheet-add\[aria-disabled='true'\]:focus-visible,[^{]*\{ color: GrayText; \}/
    );
  });

  test('a disabled sheet tab takes the subtle text color and no hover', () => {
    const disabled = rule(".xlsx-sheet-tab[aria-disabled='true']");
    for (const d of ['color: var(--doc-text-subtle)', 'cursor: default'])
      expect(disabled).toContain(d);
    expect(css).not.toMatch(/\.xlsx-sheet-tab:hover\s*\{/);
    // `:where` keeps the hover rule at two classes, under the selected tab rule.
    expect(() => rule(".xlsx-sheet-tab:hover:where(:not([aria-disabled='true']))")).not.toThrow();
  });
});
