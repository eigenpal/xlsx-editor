// The state rules of the forced-colors block of `packages/core/src/styles/editor.css`, shared
// by the chrome accessibility tests of both adapters
// (`packages/react/src/__tests__/chrome-a11y.test.tsx` and `packages/vue/test/chrome-a11y.test.ts`).
// The React test checks that a rule of the block names each selector and sets its declarations;
// both tests check that the element in each state matches the selector, and that the element
// out of that state does not. It lives here, not beside the stylesheet, because the package
// ships `src/styles`.

/** One state: the selector of its rule, and the declarations that the rule must set. */
export interface ForcedState {
  readonly selector: string;
  readonly declarations: readonly string[];
}

const HIGHLIGHT = ['forced-color-adjust: none', 'background: Highlight', 'color: HighlightText'];

/** Each state rule, by state. */
export const FORCED_STATES = {
  focusRing: {
    selector: '.xlsx-editor :focus-visible',
    declarations: ['outline-color: Highlight'],
  },
  formulaBarFocus: {
    selector: '.xlsx-cell-content-box:has(.xlsx-cell-content:focus-visible)',
    declarations: ['outline-color: Highlight'],
  },
  gridFocus: {
    selector: '.xlsx-surface:not([data-no-focus-ring]):has(.xlsx-scroller:focus-visible)::after',
    declarations: ['box-shadow: none', 'outline: 2px solid Highlight', 'outline-offset: -2px'],
  },
  selection: {
    selector: '.xlsx-selection-range',
    declarations: [
      'forced-color-adjust: none',
      'outline: 2px solid Highlight',
      'outline-offset: -1px',
      'box-shadow: inset 0 0 0 2px Canvas',
    ],
  },
  activeCell: {
    selector: '.xlsx-active-cell',
    declarations: [
      'forced-color-adjust: none',
      'border-color: Highlight',
      'box-shadow: inset 0 0 0 1px Canvas',
    ],
  },
  unfocusedActiveCell: {
    selector: '.xlsx-surface:not(:focus-within) .xlsx-active-cell',
    declarations: ['border-color: CanvasText'],
  },
  focusedTab: {
    selector: '.xlsx-sheet-tab:focus-visible',
    declarations: ['outline-offset: -4px'],
  },
  focusedSelectedTab: {
    selector: ".xlsx-sheet-tab[aria-selected='true']:focus-visible",
    declarations: ['outline-color: HighlightText'],
  },
  pressed: {
    selector:
      ".xlsx-toolbar-button:is([aria-pressed='true'], [data-active]):not([aria-disabled='true'])",
    declarations: HIGHLIGHT,
  },
  selectedTab: { selector: ".xlsx-sheet-tab[aria-selected='true']", declarations: HIGHLIGHT },
  selectedOption: {
    selector: ".xlsx-toolbar-list [role='option'][aria-selected='true']",
    declarations: HIGHLIGHT,
  },
  focusedOption: {
    selector: ".xlsx-toolbar-list [role='option']:is(:hover, :focus)",
    declarations: ['outline: 2px solid CanvasText', 'outline-offset: -2px'],
  },
  focusedSelectedOption: {
    selector: ".xlsx-toolbar-list [role='option'][aria-selected='true']:is(:hover, :focus)",
    declarations: ['outline-color: HighlightText'],
  },
  focusedListMore: {
    selector: ".xlsx-toolbar-list-more:is(:hover, :focus):not([aria-disabled='true'])",
    declarations: ['outline: 2px solid CanvasText', 'outline-offset: -2px'],
  },
  focusedMenuItem: {
    selector: '.xlsx-sheet-menu-item:is(:hover, :focus-visible):not([aria-disabled])',
    declarations: HIGHLIGHT,
  },
  focusedMenuRow: {
    selector: '.xlsx-sheet-menu-item:focus-visible:not([aria-disabled])',
    declarations: ['outline: 2px solid HighlightText', 'outline-offset: -4px'],
  },
  focusedToolbarMenuRow: {
    selector: ".xlsx-toolbar-row[role='menuitem']:focus-visible:not([aria-disabled='true'])",
    declarations: ['outline: 2px solid HighlightText', 'outline-offset: -4px'],
  },
  disabledButton: {
    selector: ".xlsx-toolbar button:is(:disabled, [aria-disabled='true'])",
    declarations: ['color: GrayText'],
  },
  disabledMenuItem: {
    selector: '.xlsx-sheet-menu-item[aria-disabled]',
    declarations: ['color: GrayText'],
  },
  swatch: {
    selector: '.xlsx-toolbar-swatch',
    declarations: ['forced-color-adjust: none', 'box-shadow: inset 0 0 0 1px CanvasText'],
  },
  colorBar: {
    selector: '.xlsx-toolbar-color-bar',
    declarations: ['forced-color-adjust: none', 'box-shadow: inset 0 0 0 1px CanvasText'],
  },
  colorChip: {
    selector: '.xlsx-toolbar-color-chip',
    declarations: ['forced-color-adjust: none', 'box-shadow: inset 0 0 0 1px CanvasText'],
  },
  selectedSwatch: {
    selector: ".xlsx-toolbar .xlsx-toolbar-swatch[aria-pressed='true']",
    declarations: [
      'outline: 3px solid Highlight',
      'box-shadow: inset 0 0 0 2px Canvas, inset 0 0 0 4px Highlight',
    ],
  },
  selectedNoFill: {
    selector: ".xlsx-toolbar .xlsx-toolbar-color-clear[aria-pressed='true']",
    declarations: [...HIGHLIGHT, 'outline: none'],
  },
  focusedSwatch: {
    selector: '.xlsx-toolbar :is(.xlsx-toolbar-swatch, .xlsx-toolbar-color-clear):focus-visible',
    declarations: ['outline: 2px solid CanvasText', 'outline-offset: 1px'],
  },
  selectedHeading: {
    selector: '.xlsx-head-selected',
    declarations: [...HIGHLIGHT, 'border-color: CanvasText'],
  },
  resizeGuide: {
    selector: '.xlsx-resize-guide',
    declarations: ['forced-color-adjust: none', 'background: Highlight'],
  },
} as const satisfies Record<string, ForcedState>;
