/**
 * XLSX support matrix: the single source of truth for feature-support claims (AGENTS.md, "Docs
 * site"). The docs site renders it through `<FeatureMatrix>`, `<FeatureSummary>`, and
 * `<FeatureBadge>` (`docs/site/README.md`).
 *
 * Same role as docx-editor/docs/site/data/word-features.ts (EigenPal), with the same `id`,
 * `name`, `category`, `notes`, and `docsLink` fields and the same honesty and notes rules.
 * Differences:
 * - One `status` for each row, not the `editing`, `rendering`, and `roundTrip` axes, and other
 *   status values. The row's category says what the status claims: rendering, viewer behavior,
 *   an edit, a refusal, or a save.
 * - No `tier` field. No package of this repository has a tier yet.
 * - Each row carries typed evidence, and `bun run check:support-matrix` ties it to the fixture
 *   catalog, the save-verified records of design D5a, the known fidelity gaps of design D8, the
 *   test files, and the known approximations. A claim without evidence fails the gate.
 *
 * Honesty rule: when in doubt, downgrade. A `partial` that turns out to be `supported` is good
 * news; a `supported` that turns out to be partial loses trust.
 *
 * Notes rule: notes render inside a table cell, so keep them short. Write Simplified Technical
 * English: active voice, one idea for each sentence, 26 words or fewer for each sentence. Put
 * file names, XML names, and attribute names in code font. Name the observable behavior, not the
 * internal lane or code path. A pull request that changes user-visible behavior changes its rows.
 */

import { FORMULA_BAR_FEATURES } from './support-matrix-formula-bar.ts';
import { NUMBER_FORMAT_FEATURES } from './support-matrix-number-format.ts';
import { STRUCTURE_FEATURES } from './support-matrix-structure.ts';
import { MERGE_FEATURES } from './support-matrix-merge.ts';
import { FORMULA_SAVE_FEATURES } from './support-matrix-formula-save.ts';
import { MACRO_ENABLED_FEATURES } from './support-matrix-macro-enabled.ts';
import { LIMIT_FEATURES } from './support-matrix-limits.ts';
import { TOOLBAR_FEATURES } from './support-matrix-toolbar.ts';
import { PANE_FEATURES } from './support-matrix-panes.ts';
import { REFUSAL_FEATURES } from './support-matrix-refusals.ts';
import { CUSTOMIZATION_FEATURES } from './support-matrix-customization.ts';

export const SUPPORT_STATUSES = ['supported', 'partial', 'unsupported'] as const;
export type SupportStatus = (typeof SUPPORT_STATUSES)[number];

/**
 * What a row claims. Rendering categories (`values`, `formatting`, `grid`, `drawings`) claim the
 * printed result, so a supported or partial row cites a fidelity result and lists its fixture in
 * `fixtures`. `viewer` covers screen behavior that a printed reference cannot show. `editing` and
 * `save` claim saved bytes, so a supported row cites save-verified fixtures. `refusals` claim
 * that an edit writes nothing.
 */
export const SUPPORT_CATEGORIES = [
  'values',
  'formatting',
  'grid',
  'drawings',
  'viewer',
  'editing',
  'refusals',
  'save',
  'collaboration',
] as const;
export type SupportCategory = (typeof SUPPORT_CATEGORIES)[number];

/** The fidelity checks of `scripts/fidelity/thresholds.ts` (`FailureKind`). */
export type FidelityCheck = 'pixels' | 'line-breaks' | 'baseline' | 'line-position' | 'extra-text';

/**
 * A fidelity result: the fixture's page passes every check except `gaps`, which must equal the
 * known gaps that `scripts/fidelity/thresholds.ts` records for it.
 */
export interface FidelityEvidence {
  readonly fixture: string;
  readonly gaps?: readonly FidelityCheck[];
}

export interface SupportEvidence {
  /** Fixtures that show the feature, by catalog name (`scripts/fixtures/catalog.ts`). */
  readonly fixtures?: readonly string[];
  /** Fixtures whose saved and edited copies must both be `verified` by a current record. */
  readonly saveVerified?: readonly string[];
  /** The fidelity result of each fixture page that shows the feature. */
  readonly fidelity?: readonly FidelityEvidence[];
  /** Test files, relative to the repository root, with `/` separators. */
  readonly tests?: readonly string[];
  /** Ids from `packages/core/src/layout/approximations.ts`. */
  readonly approximations?: readonly string[];
}

export interface SupportFeature {
  /** Stable key, for example `values.strings`. Never rename. */
  readonly id: string;
  readonly name: string;
  readonly category: SupportCategory;
  readonly status: SupportStatus;
  readonly notes?: string;
  /** The docs page that covers the feature, such as `/docs/workbook-limits`. */
  readonly docsLink?: string;
  readonly evidence: SupportEvidence;
}

export const SUPPORT_CATEGORY_LABELS: Readonly<Record<SupportCategory, string>> = {
  values: 'Cell values',
  formatting: 'Cell formats',
  grid: 'Grid structure',
  drawings: 'Drawings',
  viewer: 'Viewer',
  editing: 'Editing',
  refusals: 'Refused edits',
  save: 'Save',
  collaboration: 'Collaboration',
};

const CORE = 'packages/core/src';
const EVERY_FIXTURE = [
  'basic-values',
  'formulas',
  'preserved-parts',
  'number-formats',
  'number-formats-more',
  'dates-1904',
  'colors-fonts',
  'text-layout',
  'grid-structure',
  'structure',
  'sheet-ops',
  'sheet-add',
  'clipboard-edit',
  'resize-edit',
  'formula-overflow',
  'wrapped-formula',
  'wrapped-only-formula',
  'fills-patterns',
  'cell-styles',
  'conditional-rules',
  'macro-enabled',
] as const;
const FORMAT_EDIT_TESTS = [
  `${CORE}/store/__tests__/format-ops.test.ts`,
  `${CORE}/automation/__tests__/format-commands.test.ts`,
  `${CORE}/editor/__tests__/format-editing.test.ts`,
];
/** The save evidence of the format edit rows: one edit of every format kind, and one of each. */
const FORMAT_EDIT_NOTE =
  'The save check covers one edit of `colors-fonts` that applies every format kind together, and a feature edit of this format kind alone. A format edit that would create more than 16,384 new cells is refused with `range-too-large`, so a whole column or the whole sheet cannot be formatted.';
/** Rendering evidence: the fidelity results, with each fixture also listed in `fixtures`. */
const shows = (...fidelity: FidelityEvidence[]) => ({
  fixtures: fidelity.map((f) => f.fixture),
  fidelity,
});
const TEXT_LAYOUT: FidelityEvidence = { fixture: 'text-layout' };
const GRID_STRUCTURE_GAPS: FidelityEvidence = { fixture: 'grid-structure', gaps: ['baseline'] };

export const supportFeatures: readonly SupportFeature[] = [
  // --- Cell values ------------------------------------------------------------
  {
    id: 'values.scalars',
    name: 'Numbers, booleans, and errors',
    category: 'values',
    status: 'supported',
    evidence: {
      ...shows({ fixture: 'basic-values' }),
      tests: [`${CORE}/cascade/__tests__/values.test.ts`],
    },
  },
  {
    id: 'values.strings',
    name: 'Shared and inline strings',
    category: 'values',
    status: 'supported',
    evidence: {
      ...shows({ fixture: 'basic-values' }),
      tests: [`${CORE}/cascade/__tests__/values.test.ts`],
    },
  },
  {
    id: 'values.rich-text',
    name: 'Rich text in cells',
    category: 'values',
    status: 'partial',
    notes:
      'Each run (`x:r`) of a shared string shows with its own font, size, color, and style (`x:rPr`). A run takes the cell font for each property that it does not set.',
    evidence: {
      ...shows(TEXT_LAYOUT),
      tests: [
        `${CORE}/cascade/__tests__/values.test.ts`,
        `${CORE}/cascade/__tests__/cell-format.test.ts`,
        `${CORE}/layout/__tests__/cell-text.test.ts`,
        `${CORE}/output/__tests__/paint-sheet.test.ts`,
      ],
      approximations: [
        'row-auto-height',
        'clipped-line-scores',
        'join-gap-limit',
        'reference-items-across-cells',
      ],
    },
  },
  {
    id: 'values.cached-formulas',
    name: 'Formula results',
    category: 'values',
    status: 'supported',
    notes:
      'A cell with a direct `x:f` shows the cached value that the file stores. The engine does not calculate.',
    evidence: {
      ...shows({ fixture: 'basic-values' }, { fixture: 'formulas' }),
      tests: [
        `${CORE}/cascade/__tests__/values.test.ts`,
        `${CORE}/store/__tests__/security.test.ts`,
      ],
    },
  },
  {
    id: 'values.wrapped-formulas',
    name: 'Wrapped formula values',
    category: 'values',
    status: 'partial',
    notes:
      'A cell can keep its formula and its cached value inside `mc:AlternateContent`. The formula is the first `x:f` in the wrappers, read with `mc:Fallback` first, then each `mc:Choice`. The value is the `x:v` of the branch of that formula, else of the first branch with an `x:v` and no `x:f`. A direct `x:v` comes first. The fixture `wrapped-formula` measures the value of `mc:Fallback` when the `mc:Choice` requires a namespace that no consumer understands. The other branch cases have no reference PDF (approximation "Wrapped formula values"). `editing.formula-bar` gives what the formula bar shows.',
    evidence: {
      ...shows({ fixture: 'wrapped-formula' }),
      approximations: ['wrapped-formula-values'],
      tests: [
        `${CORE}/cascade/__tests__/values.test.ts`,
        `${CORE}/editor/__tests__/stale-notice.test.ts`,
      ],
    },
  },
  {
    id: 'values.number-formats',
    name: 'Built-in and custom number formats',
    category: 'values',
    status: 'partial',
    notes:
      'Sections, conditions, colors, fractions, scientific notation, and dates and times work. A `[$-locale]` identifier, a calendar selector, and a numeral system selector change nothing. All 152 checked lines that the paint draws in `number-formats` match the reference PDF, 1 of them only up to its cell clip.',
    evidence: {
      ...shows({ fixture: 'number-formats' }, { fixture: 'number-formats-more' }),
      tests: [
        `${CORE}/cascade/number-format/__tests__/format.test.ts`,
        `${CORE}/cascade/number-format/__tests__/measured.test.ts`,
        'scripts/fidelity/harness.test.ts',
      ],
      approximations: ['clipped-line-scores', 'glyph-rasterization'],
    },
  },
  {
    id: 'values.dates-1904',
    name: '1904 date system',
    category: 'values',
    status: 'supported',
    evidence: {
      ...shows({ fixture: 'dates-1904' }),
      tests: [`${CORE}/cascade/number-format/__tests__/format.test.ts`],
    },
  },
  {
    id: 'values.formula-evaluation',
    name: 'Formula calculation',
    category: 'values',
    status: 'unsupported',
    notes: 'The engine never evaluates a formula. It shows the cached value.',
    evidence: {
      tests: [
        `${CORE}/store/__tests__/security.test.ts`,
        `${CORE}/cascade/__tests__/values.test.ts`,
      ],
    },
  },
  // --- Cell formats -----------------------------------------------------------
  {
    id: 'formatting.colors',
    name: 'Theme, indexed, and RGB colors with tint',
    category: 'formatting',
    status: 'partial',
    notes:
      'The `colors-fonts` reference PDF measures theme, indexed, RGB, and automatic colors (`x:color/@theme`, `@indexed`, `@rgb`, and `@auto`) on fills and fonts. It measures tints (`@tint`) on theme and RGB fill colors.',
    evidence: {
      ...shows({ fixture: 'colors-fonts' }),
      tests: [`${CORE}/cascade/__tests__/cell-format.test.ts`],
      approximations: ['tint-rounding'],
    },
  },
  {
    id: 'formatting.fonts',
    name: 'Fonts and theme fonts',
    category: 'formatting',
    status: 'partial',
    notes:
      'The fixtures name the open faces that the fidelity harness draws (Carlito and Liberation Sans). When the host passes no fonts, a stored family such as Calibri draws with its metric-compatible substitute, whatever fonts the machine has. Unit tests check that mapping, and no fidelity fixture measures it.',
    evidence: {
      ...shows({ fixture: 'colors-fonts' }),
      tests: [
        `${CORE}/cascade/__tests__/cell-format.test.ts`,
        `${CORE}/layout/text/__tests__/harfbuzz-measurer.test.ts`,
        `${CORE}/output/__tests__/paint-sheet.test.ts`,
      ],
      approximations: [
        'glyph-positions',
        'substitute-outlines',
        'glyph-rasterization',
        'clipped-line-scores',
      ],
    },
  },
  {
    id: 'formatting.borders',
    name: 'Borders',
    category: 'formatting',
    status: 'partial',
    notes:
      'The `colors-fonts` reference PDF measures one thin bottom border (`x:bottom/@style` set to `thin`). The other `ST_BorderStyle` values have no measurement.',
    evidence: {
      ...shows({ fixture: 'colors-fonts' }),
      tests: [`${CORE}/output/__tests__/paint-sheet.test.ts`],
      approximations: ['border-line-styles'],
    },
  },
  {
    id: 'formatting.fills',
    name: 'Fills',
    category: 'formatting',
    status: 'partial',
    notes:
      'The `colors-fonts` reference PDF measures solid fills (`x:patternFill/@patternType` set to `solid`). The `fills-patterns` reference PDF measures each pattern tile, linear gradients at five angles, and path gradients (`x:gradientFill`). Pattern and gradient paint differs from it at the pixel level, as the approximation names.',
    evidence: {
      ...shows({ fixture: 'colors-fonts' }, { fixture: 'fills-patterns' }),
      tests: [
        `${CORE}/cascade/__tests__/cell-format.test.ts`,
        `${CORE}/output/__tests__/paint-sheet.test.ts`,
        `${CORE}/output/__tests__/paint-fills.test.ts`,
        'scripts/fidelity/pdf-fills.test.ts',
      ],
      approximations: ['fill-pixels'],
    },
  },
  {
    id: 'formatting.cell-style-formats',
    name: 'Cell formats under cell style formats',
    category: 'formatting',
    status: 'partial',
    notes:
      'The `cell-styles` reference PDF and measured record show cell formats that omit a font, fill, border, number format, or alignment. Each shows record 0 or the default alignment, for each apply attribute state under a style format that applies all records. Under a style format that applies none, the cases have no apply attributes. None shows the record of its cell style format. The measured record checks resolution, not the face that paint draws. The `cell-styles` page differs from a pdf.js raster of its reference PDF in 388 pixels. Their cause is not known.',
    evidence: {
      ...shows({ fixture: 'cell-styles' }),
      saveVerified: ['cell-styles'],
      tests: [`${CORE}/cascade/__tests__/cell-format.test.ts`],
      approximations: ['glyph-rasterization'],
    },
  },
  {
    id: 'formatting.alignment',
    name: 'Horizontal and vertical alignment and indent',
    category: 'formatting',
    status: 'partial',
    evidence: {
      ...shows(TEXT_LAYOUT),
      tests: [
        `${CORE}/layout/__tests__/cell-text.test.ts`,
        `${CORE}/layout/__tests__/sheet-layout.test.ts`,
      ],
      approximations: ['unmeasured-text-layout', 'clipped-line-scores', 'join-gap-limit'],
    },
  },
  {
    id: 'formatting.wrap',
    name: 'Wrapped text and shrink to fit',
    category: 'formatting',
    status: 'partial',
    notes:
      'A row without a set height (`x:row/@ht`) grows to hold wrapped lines (`x:alignment/@wrapText`). No reference PDF measures that height. Wrapped text in a merged range does not change the row height. Only text wraps: a number shows on one line, also with `wrapText`, `justify`, or `distributed`. A number that does not fit shows its General rounding or the `#` fill, as in a cell that does not wrap. A line break in text that does not wrap shows as one space. No reference PDF shows a number with `wrapText` or a line break in text that does not wrap.',
    evidence: {
      ...shows(TEXT_LAYOUT),
      tests: [
        `${CORE}/layout/__tests__/cell-text.test.ts`,
        `${CORE}/layout/__tests__/sheet-layout.test.ts`,
        `${CORE}/layout/__tests__/format-layout.test.ts`,
        `${CORE}/layout/__tests__/fit.test.ts`,
      ],
      approximations: ['unmeasured-text-layout', 'row-auto-height'],
    },
  },
  {
    id: 'formatting.overflow',
    name: 'Text overflow into empty cells',
    category: 'formatting',
    status: 'partial',
    evidence: {
      ...shows(TEXT_LAYOUT, { fixture: 'formula-overflow' }),
      tests: [
        `${CORE}/layout/__tests__/cell-text.test.ts`,
        `${CORE}/layout/__tests__/sheet-layout.test.ts`,
      ],
      approximations: ['overflow-reach', 'unmeasured-text-layout', 'clipped-line-scores'],
    },
  },
  {
    id: 'formatting.rotation',
    name: 'Rotated and stacked text',
    category: 'formatting',
    status: 'partial',
    notes:
      'The reference PDF measures rotated text (`x:alignment/@textRotation`) by 45, 90, and -45 degrees with general alignment at the bottom of the cell. Other rotated cases have no measurement. Stacked text sits 1 point right.',
    evidence: {
      ...shows(TEXT_LAYOUT),
      tests: [
        `${CORE}/layout/__tests__/cell-text.test.ts`,
        `${CORE}/layout/__tests__/sheet-layout.test.ts`,
      ],
      approximations: ['measured-text-offsets', 'unmeasured-text-layout', 'unmeasured-rotation'],
    },
  },
  {
    id: 'formatting.conditional-formatting',
    name: 'Conditional formatting',
    category: 'formatting',
    status: 'partial',
    notes:
      'The grid paints `cellIs` rules (`x:cfRule`) with literal operands. Other rule types, and `cellIs` operands that need a formula, are not shown, and a notice shows for the sheet; the approximation lists every case. A save keeps every rule, and structural edits are refused.',
    evidence: {
      ...shows(
        { fixture: 'preserved-parts' },
        { fixture: 'conditional-rules', gaps: ['line-position'] }
      ),
      saveVerified: ['conditional-rules'],
      approximations: ['conditional-rules', 'measured-text-offsets', 'tint-rounding'],
      tests: [
        `${CORE}/cascade/__tests__/conditional.test.ts`,
        `${CORE}/layout/__tests__/sheet-layout.test.ts`,
        `${CORE}/store/__tests__/preserved-content.test.ts`,
        `${CORE}/editor/__tests__/toolbar-commands.test.ts`,
        `${CORE}/editor/__tests__/toolbar-conditional.test.ts`,
        `packages/react/src/__tests__/toolbar.test.tsx`,
        `packages/react/src/__tests__/editor-command.test.tsx`,
      ],
    },
  },
  // --- Grid structure ---------------------------------------------------------
  {
    id: 'grid.column-widths',
    name: 'Column widths',
    category: 'grid',
    status: 'partial',
    evidence: {
      ...shows(GRID_STRUCTURE_GAPS),
      tests: [`${CORE}/layout/__tests__/grid-geometry.test.ts`],
      approximations: ['pixel-density'],
    },
  },
  {
    id: 'grid.row-heights',
    name: 'Row heights',
    category: 'grid',
    status: 'partial',
    notes:
      'A font size, font name, or wrap edit removes a fitted row height, so the row fits again. A value edit, a paste, or a clear does the same for each row that it changes, whatever the row holds. The row then fits all of its cells again, so a row with extra space gets shorter. A height set by hand stays.',
    evidence: {
      ...shows(GRID_STRUCTURE_GAPS, TEXT_LAYOUT),
      tests: [
        `${CORE}/layout/__tests__/grid-geometry.test.ts`,
        `${CORE}/layout/__tests__/axis.test.ts`,
        `${CORE}/layout/__tests__/format-layout.test.ts`,
        `${CORE}/store/__tests__/format-ops-cases.test.ts`,
        `${CORE}/store/__tests__/row-refit.test.ts`,
        `${CORE}/editor/__tests__/row-refit.test.ts`,
      ],
      approximations: ['row-auto-height', 'screen-geometry'],
    },
  },
  {
    id: 'grid.hidden',
    name: 'Hidden rows and columns',
    category: 'grid',
    status: 'partial',
    evidence: {
      ...shows(GRID_STRUCTURE_GAPS),
      tests: [
        `${CORE}/layout/__tests__/grid-geometry.test.ts`,
        `${CORE}/layout/__tests__/axis.test.ts`,
      ],
    },
  },
  {
    id: 'grid.merged-cells',
    name: 'Merged cells',
    category: 'grid',
    status: 'partial',
    evidence: {
      ...shows(GRID_STRUCTURE_GAPS),
      tests: [
        `${CORE}/layout/__tests__/panes.test.ts`,
        `${CORE}/layout/__tests__/sheet-layout.test.ts`,
        `${CORE}/layout/__tests__/sheet-layout-limits.test.ts`,
      ],
      approximations: ['measured-text-offsets'],
    },
  },
  // --- Drawings ---------------------------------------------------------------
  {
    id: 'drawings.pictures',
    name: 'Pictures',
    category: 'drawings',
    status: 'partial',
    notes: 'PNG, JPEG, GIF, and BMP. An external image is never loaded.',
    evidence: {
      ...shows(GRID_STRUCTURE_GAPS),
      tests: [
        `${CORE}/layout/__tests__/drawings.test.ts`,
        `${CORE}/output/__tests__/paint-sheet.test.ts`,
      ],
      approximations: ['drawing-objects'],
    },
  },
  {
    id: 'drawings.shapes',
    name: 'Preset and custom shapes',
    category: 'drawings',
    status: 'partial',
    notes: 'Solid fills and lines only. Shape text is not drawn.',
    evidence: {
      ...shows(GRID_STRUCTURE_GAPS),
      tests: [
        `${CORE}/layout/__tests__/drawings.test.ts`,
        `${CORE}/output/__tests__/paint-sheet.test.ts`,
      ],
      approximations: ['drawing-objects'],
    },
  },
  {
    id: 'drawings.charts',
    name: 'Charts',
    category: 'drawings',
    status: 'unsupported',
    notes:
      'A save keeps the chart part. The grid does not draw it, and structural edits are refused. A sheet that the editor added in the session can be renamed and deleted, unless a chart can refer to it.',
    evidence: {
      tests: [
        `${CORE}/layout/__tests__/drawings.test.ts`,
        `${CORE}/store/__tests__/structure-refusals.test.ts`,
        `${CORE}/store/__tests__/sheet-ops-added.test.ts`,
      ],
    },
  },
  // --- Viewer -----------------------------------------------------------------
  ...PANE_FEATURES,
  {
    id: 'values.stale-formulas',
    name: 'Stale formula results',
    category: 'viewer',
    status: 'supported',
    notes:
      'When the workbook sets `x:calcPr/@fullCalcOnLoad`, or after an edit, the editor shows the stored formula results. If the shown sheet holds a formula cell, the editor marks them as stale in four places. On a sheet with no formula cell, the editor marks nothing. A format edit, a clear of formats, and a pasted cell format mark the results when a formula calls `CELL` or a macro function `GET.*`. The snapshot sets `stale`. A notice in the status strip above the grid says that formula results can be out of date. The notice also counts formula cells that the grid does not show. These are an `x:c` inside `mc:AlternateContent` (also in a wrapped row), outside a row, or inside another cell. A shown cell with its formula in an `mc:AlternateContent` branch gets the stale mark. The formula bar shows the formula of such a cell (`editing.formula-bar`). When a stale formula cell is the active cell, the live region reads its shown result, its formula, and "not recalculated". The accessible text of each stale formula cell in the grid ends with the same words. The canvas paints no mark on a stale cell. The printed page shows no notice, so no reference PDF covers it.',
    evidence: {
      tests: [
        `${CORE}/cascade/__tests__/values.test.ts`,
        `${CORE}/store/__tests__/cell-format-index.test.ts`,
        `${CORE}/layout/__tests__/sheet-layout.test.ts`,
        `${CORE}/editor/__tests__/cell-editing.test.ts`,
        `${CORE}/editor/__tests__/stale-notice.test.ts`,
        `${CORE}/output/__tests__/paint-sheet.test.ts`,
        'packages/react/src/__tests__/adapter.test.tsx',
      ],
    },
  },
  {
    id: 'viewer.selection',
    name: 'Keyboard navigation and selection',
    category: 'viewer',
    status: 'supported',
    notes:
      'Arrows, Tab, Enter, Home, Page Up, and Page Down move the active cell. Shift extends the selection, and Mod with an arrow jumps to the edge of the data. Mod+Home and Mod+End go to the ends. Page Up, Page Down, and Mod+End skip hidden rows and columns, as the arrows do. F6 and Shift+F6 move the keyboard among the toolbar, the grid, and the footer. Outside a selection of many cells, Shift+Tab in column A and Tab in the last column move the focus out of the grid. A selection that is one merged cell is one cell for Tab and Enter, also when they commit a cell edit. From any cell, Escape with nothing to cancel says how to leave, and the next Tab or Shift+Tab moves the focus out of the grid. Any other key except a modifier key alone, a press anywhere on the page, or a focus change cancels that exit. While the exit waits for Tab, a cue over the foot of the grid shows the same words. The accessible description of the grid names this exit and F6. Alt+Page Up and Alt+Page Down change the sheet. A press on a row heading, a column heading, or the corner selects whole rows, whole columns, or the whole sheet. A press or a drag with Shift extends the selection from its anchor to the pressed cell, row, or column. The active cell stays when the selection holds it. Otherwise the anchor becomes the active cell. Mod+A also selects the whole sheet. A press with Control (Command on Apple platforms) on a cell or a heading adds a new range to the selection. A drag from that press extends the new range. A workbook opens with the ranges of `x:selection/@sqref` and the active cell of `@activeCell` selected, in the range that `@activeCellId` names, at most 64 ranges. A sheet with many merged ranges can open with fewer ranges, as each range grows over the merged ranges that it meets. A drag on a heading edge resizes only the band of whole columns or rows that holds that edge. Each range shows the selection tint and outline, and the headings of every range show as selected. Enter and Tab move through every range in order. A clear and a format command apply to every range in one undo step. Copy, cut, paste, merge, the decimal steps, and row or column inserts and deletes refuse a selection of several ranges. Their controls are disabled with the reason. The editor has no add mode on Shift+F8. The live region of the grid says the active cell or the selection when the selection moves, or when the focus enters the grid. A step that moves the focus and says its own words, such as a press on a sheet tab, says only those words. If a toolbar menu applies a format and focuses the grid, the live region says the change, or the active cell when its text stays. An edit can change the shown text of the active cell, such as a number format or a typed entry. The live region then says the active cell and its new text once, also on a selection of many cells. An edit that keeps the shown text says nothing, except a decimal step or a number format change. That edit says that the decimal places increased or decreased, or that the number format changed. An undo or a redo names its step, and the active cell when the step changed its text. When Enter or Tab moves the active cell inside a selection of many cells, the live region says the selection with the new active cell. When its words have more than 65,536 UTF-16 code units, the live region says their start and that the text continues.',
    evidence: {
      tests: [
        `${CORE}/editor/__tests__/xlsx-editor.test.ts`,
        `${CORE}/editor/__tests__/several-ranges.test.ts`,
        `${CORE}/editor/__tests__/live-announce.test.ts`,
        `${CORE}/editor/__tests__/structure-editing.test.ts`,
        `${CORE}/editor/__tests__/cell-editing.test.ts`,
        `${CORE}/editor/__tests__/grid-exit.test.ts`,
        `${CORE}/editor/__tests__/grid-entry.test.ts`,
        'packages/react/src/__tests__/adapter.test.tsx',
        'packages/react/src/__tests__/grid-entry.test.tsx',
        'packages/react/src/__tests__/editor-command.test.tsx',
        'packages/vue/test/grid-entry.test.ts',
        'packages/vue/test/editor-command.test.ts',
      ],
    },
  },
  {
    id: 'viewer.gridlines',
    name: 'Gridlines',
    category: 'viewer',
    status: 'partial',
    evidence: {
      tests: [`${CORE}/output/__tests__/paint-sheet.test.ts`],
      approximations: ['gridline-color'],
    },
  },
  {
    id: 'viewer.sheets',
    name: 'Sheet tabs and hidden sheets',
    category: 'viewer',
    status: 'supported',
    evidence: { tests: [`${CORE}/editor/__tests__/xlsx-editor.test.ts`] },
  },
  {
    id: 'viewer.toolbar',
    name: 'Editing toolbar',
    category: 'viewer',
    status: 'supported',
    notes:
      'Each slot runs an editor command. A disabled slot stays focusable and gives its reason. The hidden prop leaves out slots, and a group with every slot hidden disappears. Groups that do not fit move to an overflow panel. Slot parts replace or hide a slot where it stands, and host groups are measured and move to the panel with the built-in groups. With preset off, the toolbar renders only its children. With an `onSave` handler, Control or Command and S in the editor runs the save, also when the save control is hidden. Without one, the browser keeps the shortcut.',
    docsLink: '/docs/customization/toolbar',
    evidence: {
      tests: [
        `${CORE}/editor/__tests__/toolbar-commands.test.ts`,
        `${CORE}/editor/__tests__/toolbar-arrange.test.ts`,
        `packages/react/src/__tests__/toolbar.test.tsx`,
        `packages/react/src/__tests__/toolbar-parts.test.tsx`,
        `packages/vue/test/toolbar-parts.test.ts`,
      ],
    },
  },
  ...TOOLBAR_FEATURES,
  {
    id: 'viewer.view-mode',
    name: 'View mode',
    category: 'viewer',
    status: 'supported',
    notes:
      'In view mode, every edit command and edit key is refused and writes nothing. The toolbar disables its edit controls with the reason. Selection, sheet tabs, zoom, and save stay available.',
    evidence: {
      tests: [
        `${CORE}/editor/__tests__/view-mode.test.ts`,
        `packages/react/src/__tests__/editor-command.test.tsx`,
      ],
    },
  },
  ...CUSTOMIZATION_FEATURES,
  ...LIMIT_FEATURES,
  // --- Editing ----------------------------------------------------------------
  {
    id: 'editing.cell-values',
    name: 'Edit numbers, dates, text, and booleans',
    category: 'editing',
    status: 'supported',
    notes:
      'Typed text reads as a number in these forms: a plain decimal, a percent (`12%`), and a currency amount with `$`, `€`, `£`, or `¥` (`$5.50`). A number with group separators (`1,234.5`), a negative in parentheses (`(100)`), and a fraction with a whole part (`1 1/2`) also read as numbers. So do a date (`2026-11-12`), a time (`13:45`, `1:45 PM`), an elapsed time over 23 hours (`36:00`), and a date and a time. The decimal and group separators come from the number locale. A date in a form other than ISO 8601 reads in the day and month order of the locale short date (built-in format 14). The short date has an order only when it is a day, a month, and a year with one separator. With no order, such a date stays text, and an ISO 8601 date takes the code `yyyy-mm-dd`. A `.` separates a date only when that short date uses `.`, so `1.2.10` stays text in US English. A day and a month with no year (`1/2`) is a date in the current year, and `0 1/2` is a fraction. A date is a serial number in the date system of the workbook (ECMA-376 Part 1, 18.17.4.1). In a cell that stores a date value (`t="d"`), ISO 8601 text stays a date value. On a cell with the General format, the entry also sets the number format of its form, in the same undo step. That format is a built-in format (Part 1, 18.8.30) or a custom code. A currency amount and an elapsed time with no seconds take a custom code. A date and a time takes built-in format 22 when format 22 is a plain date in the order of the short date and one clock. Typed seconds add `:ss`, and a typed AM or PM replaces its clock. With another date order, it takes the short date code and the typed clock. With no date order, it takes `yyyy-mm-dd` and the clock. A workbook can define a built-in id with a code other than the one that the number locale shows. Then the cell takes a custom id with the locale code. A sheet whose protection does not allow formatting cells keeps the cell format and stores the number. So does a workbook whose list of cell formats or number formats is full, or whose styles part the editor cannot read. So does a number locale code that the editor cannot read. A cell with another number format keeps it. A cell with the Text format (`@`) keeps the entry as text, but a formula stays a formula. On a re-edit, the cell editor shows a date as ISO 8601 (`2026-11-12`) and a time on a 24-hour clock (`13:45`). The cell editor shows an elapsed time in hours (`36:00`), and a percent with `%` (`12%`). The formula bar shows the same text, also for a cell that the opened workbook holds. A currency amount shows as a plain number (`5.5`). So does a time in a format that shows fractions of a second, when the shown fraction is not zero. A two-digit year from 00 to 29 is in the 2000s, and from 30 to 99 in the 1900s. A plain number typed in a percent cell is not scaled: `60` stores 60. `TRUE` and `FALSE` read as booleans. A leading apostrophe keeps the rest as text, and the apostrophe is not stored. An error literal such as `#N/A`, a month name, and text of more than 64 characters stay text. A plain-text paste reads each cell the same way.',
    evidence: {
      saveVerified: ['basic-values', 'clipboard-edit', 'text-layout', 'grid-structure'],
      tests: [
        `${CORE}/editor/__tests__/cell-editing.test.ts`,
        `${CORE}/editor/__tests__/typed-entry.test.ts`,
        `${CORE}/store/__tests__/store-ops.test.ts`,
      ],
    },
  },
  {
    id: 'editing.cell-editor',
    name: 'Cell editor input and input methods',
    category: 'editing',
    status: 'supported',
    notes:
      'Typing starts an edit, F2 edits in place, Enter and Tab commit and move, and Escape cancels. Alt+Enter adds a line break. A commit of text with a line break turns on wrap text (`x:alignment/@wrapText`) of the cell in the same undo step. It does not when the cell format wraps already, or when the sheet protection does not allow formatting cells. An input method composes text in the cell editor, and a composition commits nothing until it ends.',
    evidence: {
      saveVerified: ['basic-values'],
      tests: [
        `${CORE}/editor/__tests__/cell-editing.test.ts`,
        `${CORE}/editor/__tests__/cell-editor-cases.test.ts`,
        `${CORE}/editor/__tests__/line-break-wrap.test.ts`,
      ],
    },
  },
  ...FORMULA_BAR_FEATURES,
  {
    id: 'editing.sheet-protection',
    name: 'Sheet protection',
    category: 'editing',
    status: 'partial',
    notes:
      'On a protected sheet (`x:sheetProtection/@sheet`), the editor refuses a value edit on a locked cell with `locked`. It refuses a format edit unless `formatCells` allows it. A paste that gives a cell another cell format is a format edit. A plain-text paste of a typed form is not refused for its format: it stores the number and keeps the cell format. It refuses a cut whose paste could not clear the source, so no later paste fails for it. It refuses a column width unless `formatColumns` allows it, and a row height unless `formatRows` allows it. A heading edge whose size is locked shows no resize edge. It refuses a row or column insertion or deletion unless its attribute allows it. When `x:workbookProtection/@lockStructure` is true, it refuses a sheet add, rename, delete, or move. The password hash is never read or checked, and it is kept. It refuses a merge or an unmerge of cells with the rule `merge`, whatever the protection attributes state. Protected ranges, `lockWindows`, and `lockRevision` are not enforced.',
    evidence: {
      tests: [
        `${CORE}/store/__tests__/sheet-protection.test.ts`,
        `${CORE}/editor/__tests__/typed-entry.test.ts`,
        `${CORE}/editor/__tests__/toolbar-commands.test.ts`,
        `${CORE}/editor/__tests__/clipboard.test.ts`,
        `${CORE}/editor/__tests__/resize.test.ts`,
        `${CORE}/editor/__tests__/merge-editing.test.ts`,
      ],
    },
  },
  {
    id: 'editing.formulas',
    name: 'Enter formulas',
    category: 'editing',
    status: 'supported',
    notes:
      'The formula (`x:f`) is stored with no cached value (`x:v`). An edit to a shared formula master unshares its children. When the `x:f` of a child is inside `mc:AlternateContent` in its cell, an edit or a clear of its master is refused. A clear that also clears that child is refused with `cell-alternate-content`, as is any edit of a cell with a formula there. A typed formula can call only the predefined functions of ECMA-376 Part 1, 18.17.7; a call to any other function is refused with `unsupported-function`. A formula that the workbook holds keeps its stored text.',
    evidence: {
      saveVerified: ['basic-values', 'formulas'],
      tests: [
        `${CORE}/store/__tests__/formula-edits.test.ts`,
        `${CORE}/store/__tests__/formula-scan.test.ts`,
        `${CORE}/editor/__tests__/cell-editing.test.ts`,
        `${CORE}/store/formula/__tests__/tokenizer.test.ts`,
      ],
    },
  },
  {
    id: 'editing.undo-redo',
    name: 'Undo and redo',
    category: 'editing',
    status: 'supported',
    notes:
      'Undo restores the package exactly, one transaction for each step. Undo and redo also restore the selection of the step and show its sheet. The save check covers a saved edit, not a save after an undo.',
    evidence: {
      saveVerified: ['basic-values'],
      tests: [
        `${CORE}/store/__tests__/store-ops.test.ts`,
        `${CORE}/store/__tests__/format-ops.test.ts`,
        `${CORE}/store/__tests__/structure-ops.test.ts`,
        `${CORE}/editor/__tests__/toolbar-commands.test.ts`,
        `${CORE}/editor/__tests__/xlsx-editor.test.ts`,
      ],
    },
  },
  {
    id: 'editing.clear',
    name: 'Clear values, formats, or both',
    category: 'editing',
    status: 'supported',
    notes:
      'A clear of formats writes a plain cell where a row or column format shows, with the limit of a format edit. A clear of formats marks the cached results for recalculation when a formula or a defined name calls `CELL` or a macro function `GET.*`. On a protected sheet, a clear of formats keeps whether each cell is locked or hidden: the cell gets format 0 with its own protection. A clear of values that would break a formula range is refused with `shared-formula`. That is a cell of an array or data table range, or a shared master that cannot unshare. The formula can be inside `mc:AlternateContent`. Clear contents and Clear all are then disabled. Their reason names the range, or the shared formula master and the cell that blocks it. When the formula of a shared child would move out of the grid, those controls stay enabled. The clear is then refused when it runs. The paste of a cut clears its source cells as Clear all does, and the save check of the fixture `clipboard-edit` covers those cells. The save check opened a copy of `colors-fonts` after a clear of formats and a clear of contents.',
    evidence: {
      saveVerified: ['colors-fonts'],
      tests: [
        `${CORE}/store/__tests__/formula-edits.test.ts`,
        `${CORE}/store/__tests__/format-ops.test.ts`,
        `${CORE}/store/__tests__/sheet-protection.test.ts`,
        `${CORE}/editor/__tests__/format-editing.test.ts`,
        `${CORE}/editor/__tests__/cell-editing.test.ts`,
        `${CORE}/editor/__tests__/clipboard.test.ts`,
      ],
    },
  },
  {
    id: 'editing.font-style',
    name: 'Bold, italic, and underline',
    category: 'editing',
    status: 'supported',
    notes: FORMAT_EDIT_NOTE,
    evidence: { saveVerified: ['colors-fonts'], tests: FORMAT_EDIT_TESTS },
  },
  {
    id: 'editing.font-name-size',
    name: 'Font and font size',
    category: 'editing',
    status: 'supported',
    notes: FORMAT_EDIT_NOTE,
    evidence: { saveVerified: ['colors-fonts'], tests: FORMAT_EDIT_TESTS },
  },
  {
    id: 'editing.font-color',
    name: 'Font color',
    category: 'editing',
    status: 'supported',
    notes: `${FORMAT_EDIT_NOTE} The color picker writes six hex digits (\`x:color/@rgb\`), also for a theme color or one of its tints and shades. **Automatic** writes \`x:color/@auto\` (ECMA-376 Part 1, 18.8.19).`,
    evidence: { saveVerified: ['colors-fonts'], tests: FORMAT_EDIT_TESTS },
  },
  {
    id: 'editing.fill',
    name: 'Fill color',
    category: 'editing',
    status: 'supported',
    notes: `${FORMAT_EDIT_NOTE} The color picker writes a solid fill of six hex digits, also for a theme color or one of its tints and shades.`,
    evidence: { saveVerified: ['colors-fonts'], tests: FORMAT_EDIT_TESTS },
  },
  {
    id: 'editing.borders',
    name: 'Borders',
    category: 'editing',
    status: 'partial',
    notes:
      'The save check covers one thin outline border, in the edit of `colors-fonts` that applies every format kind together. The other line styles paint without a measurement.',
    evidence: {
      saveVerified: ['colors-fonts'],
      tests: FORMAT_EDIT_TESTS,
      approximations: ['border-line-styles'],
    },
  },
  {
    id: 'editing.alignment',
    name: 'Alignment and wrap',
    category: 'editing',
    status: 'supported',
    notes: FORMAT_EDIT_NOTE,
    evidence: {
      saveVerified: ['colors-fonts'],
      tests: [...FORMAT_EDIT_TESTS, `${CORE}/editor/__tests__/toolbar-commands.test.ts`],
    },
  },
  {
    id: 'editing.number-format',
    name: 'Number format',
    category: 'editing',
    status: 'supported',
    notes: FORMAT_EDIT_NOTE,
    evidence: { saveVerified: ['colors-fonts'], tests: FORMAT_EDIT_TESTS },
  },
  ...NUMBER_FORMAT_FEATURES,
  {
    id: 'editing.row-column-formats',
    name: 'Empty cells in a formatted row or column',
    category: 'editing',
    status: 'partial',
    notes:
      'An empty cell shows the format of its row or column. A format edit on it starts from that format, and the format moves with a structural edit. A clear of formats writes a plain cell in each such position. One clear writes at most 16,384 new cells and refuses a larger range. No reference PDF and no save check covers a row or column format.',
    evidence: {
      tests: [
        `${CORE}/layout/__tests__/format-layout.test.ts`,
        `${CORE}/store/__tests__/format-ops-cases.test.ts`,
        `${CORE}/store/__tests__/structure-calc.test.ts`,
      ],
    },
  },
  ...STRUCTURE_FEATURES,
  ...MERGE_FEATURES,
  {
    id: 'editing.rows',
    name: 'Insert and delete rows',
    category: 'editing',
    status: 'supported',
    notes:
      'References, merged cells (`x:mergeCell`), hyperlinks, defined names (`x:definedName`), and drawings move with the rows. A drawing part that the editor cannot read refuses the edit with `structure-blocked`. A drawing whose relationship part stays as bytes refuses it too. So does a drawing that relates to a part other than an image or a hyperlink, such as a diagram. The toolbar then disables the command with the reason. A formula or a defined name inside `mc:AlternateContent` keeps its text in each branch. An edit that would change that text is refused with `structure-blocked`, which names the cell, the name, or the wrapper of a workbook name list. So is an edit that moves or deletes a cell of a shared or array formula with a member there. So is an edit that changes the formula of any member of such a shared formula. A shared formula with master copies in two cells, or with more than four master texts, refuses the edit too. A data table formula refuses every edit of its sheet, also when it is inside `mc:AlternateContent`. The save check opened a copy of `structure` after a row delete, and a copy after a row insert. A stored sheet name of more than 31 characters refuses the edit with `sheet-name-too-long`. When a file gives the sheet name to another sheet too, in any case, the edit is refused with `sheet-name-ambiguous`, which names the other sheet. A hyperlink target such as `#Data!A5` moves with its cell.',
    evidence: {
      saveVerified: ['structure'],
      tests: [
        `${CORE}/store/__tests__/structure-ops.test.ts`,
        `${CORE}/store/__tests__/structure-wrapped.test.ts`,
        `${CORE}/editor/__tests__/structure-editing.test.ts`,
        `${CORE}/automation/__tests__/structure-commands.test.ts`,
        `${CORE}/store/__tests__/sheet-name-fold.test.ts`,
        `${CORE}/store/__tests__/sheet-ops-links.test.ts`,
        `${CORE}/store/__tests__/drawing-kept-as-bytes.test.ts`,
      ],
    },
  },
  {
    id: 'editing.columns',
    name: 'Insert and delete columns',
    category: 'editing',
    status: 'supported',
    notes:
      'References, merged cells (`x:mergeCell`), hyperlinks, defined names (`x:definedName`), and drawings move with the columns. A drawing part that the editor cannot read refuses the edit with `structure-blocked`. A drawing whose relationship part stays as bytes refuses it too. So does a drawing that relates to a part other than an image or a hyperlink, such as a diagram. The toolbar then disables the command with the reason. A formula or a defined name inside `mc:AlternateContent` keeps its text in each branch. An edit that would change that text is refused with `structure-blocked`, which names the cell, the name, or the wrapper of a workbook name list. So is an edit that moves or deletes a cell of a shared or array formula with a member there. So is an edit that changes the formula of any member of such a shared formula. A shared formula with master copies in two cells, or with more than four master texts, refuses the edit too. A data table formula refuses every edit of its sheet, also when it is inside `mc:AlternateContent`. The save check opened a copy of `structure` after a column insert, and a copy after a column delete. A stored sheet name of more than 31 characters refuses the edit with `sheet-name-too-long`. When a file gives the sheet name to another sheet too, in any case, the edit is refused with `sheet-name-ambiguous`, which names the other sheet. A hyperlink target such as `#Data!A5` moves with its cell.',
    evidence: {
      saveVerified: ['structure'],
      tests: [
        `${CORE}/store/__tests__/structure-ops.test.ts`,
        `${CORE}/store/__tests__/structure-wrapped.test.ts`,
        `${CORE}/editor/__tests__/structure-editing.test.ts`,
        `${CORE}/automation/__tests__/structure-commands.test.ts`,
        `${CORE}/store/__tests__/sheet-name-fold.test.ts`,
        `${CORE}/store/__tests__/sheet-ops-links.test.ts`,
        `${CORE}/store/__tests__/drawing-kept-as-bytes.test.ts`,
      ],
    },
  },
  {
    id: 'editing.automation-api',
    name: 'Workbook automation object model',
    category: 'editing',
    status: 'partial',
    notes:
      'Read cell values and number formats, and write plain values, on a server with `@xlsx-editor.dev/editor-api`, under the EigenPal Pro License. It implements a small Office.js subset: worksheets by name, ranges by `A1` address, `values`, `numberFormat` (read), and `address`. A read gives the stored value. A formula cell, a formula write, a number format write, and text that reads as another value are refused. The package is private and not published to npm. Every other member is absent from the types. No fixture has a save check of an edit through the API.',
    docsLink: '/docs/editor-api',
    evidence: {
      tests: [
        'packages/editor-api/src/__tests__/runtime.test.ts',
        'packages/editor-api/src/__tests__/rules.test.ts',
        'packages/editor-api/src/__tests__/docs-examples.test.ts',
        'packages/editor-api/scripts/__tests__/conformance.test.ts',
        `${CORE}/automation/__tests__/host.test.ts`,
        `${CORE}/automation/__tests__/host-rules.test.ts`,
      ],
    },
  },
  // --- Refused edits ----------------------------------------------------------
  ...REFUSAL_FEATURES,
  // --- Save -------------------------------------------------------------------
  {
    id: 'save.round-trip',
    name: 'Open and save with no repair prompt',
    category: 'save',
    status: 'supported',
    notes: 'Each fixture opens with no repair prompt after a save and after its edit.',
    evidence: { saveVerified: EVERY_FIXTURE, tests: [`scripts/save-check/records.test.ts`] },
  },
  {
    id: 'save.preserved-content',
    name: 'Content that the engine does not model',
    category: 'save',
    status: 'supported',
    notes: 'A save keeps unknown parts and elements. Macros and embedded objects stay inert.',
    evidence: {
      saveVerified: ['preserved-parts'],
      tests: [
        `${CORE}/store/__tests__/preserved-content.test.ts`,
        `${CORE}/store/__tests__/cell-children.test.ts`,
        `${CORE}/store/__tests__/package-read.test.ts`,
      ],
    },
  },
  ...MACRO_ENABLED_FEATURES,
  {
    id: 'save.namespace-declarations',
    name: 'Prefixes that Markup Compatibility attributes name',
    category: 'save',
    status: 'partial',
    notes:
      'An edit keeps the namespace declarations of the prefixes that `mc:Ignorable` and the other Markup Compatibility attributes name, also on an element that it builds again. A workbook with an undeclared prefix in a required part is refused on open; an optional part with one stays as bytes. No save check covers an edit of an element with these attributes.',
    evidence: { tests: [`${CORE}/store/__tests__/namespace-declarations.test.ts`] },
  },
  {
    id: 'save.hyperlinks',
    name: 'Hyperlinks kept and moved',
    category: 'save',
    status: 'supported',
    notes:
      'A save keeps each hyperlink (`x:hyperlink`) and its relationship. An inserted or deleted row or column moves the hyperlink with its cell, or removes it with the cell. A location, or a target such as `#Data!A5`, moves with the cell that it names. An external target passes the URL allowlist (`http`, `https`, `mailto`, `tel`, `ftp`) before any link can use it. The grid does not open a hyperlink.',
    evidence: {
      saveVerified: ['structure'],
      tests: [
        `${CORE}/store/__tests__/structure-ops.test.ts`,
        `${CORE}/store/__tests__/structure-cases.test.ts`,
        `${CORE}/store/__tests__/sheet-ops-links.test.ts`,
        `${CORE}/store/__tests__/security.test.ts`,
        `scripts/fixtures/structure.test.ts`,
      ],
    },
  },
  ...FORMULA_SAVE_FEATURES,
  {
    id: 'save.tables-pivots',
    name: 'Tables, pivot tables, and data validation',
    category: 'save',
    status: 'unsupported',
    notes:
      'A save keeps them (`x:table`, `x:pivotTableDefinition`, and `x:dataValidations`). The grid does not show them, and structural edits are refused. A sheet that the editor added in the session can be renamed and deleted, unless a pivot cache can refer to it.',
    evidence: {
      tests: [
        `${CORE}/store/__tests__/structure-refusals.test.ts`,
        `${CORE}/store/__tests__/sheet-ops-added.test.ts`,
        `${CORE}/store/__tests__/preserved-content.test.ts`,
      ],
    },
  },
  {
    id: 'save.external-content',
    name: 'External links, data connections, and macros',
    category: 'save',
    status: 'unsupported',
    notes:
      'The engine never loads or runs them (`x:externalLink`, `x:connections`, and `vbaProject.bin`). A save keeps their parts byte for byte, and writes their relationship parts in a normalized form. A macro-enabled workbook opens, and its macro part stays inert (`save.macro-enabled`). A data connection refuses a rename and a delete of each sheet of the file. A sheet that the editor added in the session takes both, unless the connection can refer to it.',
    evidence: {
      tests: [
        `${CORE}/store/__tests__/package-read.test.ts`,
        `${CORE}/store/__tests__/security.test.ts`,
        `${CORE}/store/__tests__/sheet-ops-added.test.ts`,
      ],
    },
  },
  // --- Collaboration ----------------------------------------------------------
  {
    id: 'collaboration.sessions',
    name: 'Real-time collaboration',
    category: 'collaboration',
    status: 'unsupported',
    notes:
      'The editor has no real-time collaboration. Each editor instance edits its own copy of the workbook.',
    evidence: {},
  },
];

/** Each row by its `id`, for `<FeatureBadge id="...">`, as `wordFeatureById` in DOCX Editor. */
export const supportFeatureById: Readonly<Record<string, SupportFeature>> = Object.fromEntries(
  supportFeatures.map((row) => [row.id, row])
);
