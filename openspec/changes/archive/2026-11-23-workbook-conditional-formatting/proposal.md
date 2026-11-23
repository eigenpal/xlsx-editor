## Why

The change `typed-ooxml-workbook-editor` keeps conditional formatting through a save, but the grid does not paint it. The reference PDF of `preserved-parts` draws the value 5 of cell B1 in bold, because a `cellIs` rule with `greaterThan 3` names a differential format that sets bold. The editor draws it in the regular face.

## What Changes

- Read the `x:cfRule` elements of a worksheet within fixed limits, and sort them by `@priority`, then document order.
- Evaluate `cellIs` rules whose operands are literals against the stored value of each visible cell. Every other rule is not shown, and each such rule that the reader builds has one reason.
- Lay the differential format (`x:dxf`) of each true rule over the cell format: the font, the fill, and the number format.
- Show a notice for each sheet whose conditional formatting the grid can leave out, and a tooltip hint on each format control whose state a true rule changes at the active cell.
- Add the fixture `conditional-rules` with its reference PDF and save check.
- Add the approximation "Conditional formatting rules", and move the support matrix row `formatting.conditional-formatting` from unsupported to partial.

## Capabilities

### New Capabilities

- `workbook-conditional-formatting`: The rule reader, the evaluation of `cellIs` rules with literal operands, and the differential format overlay in layout.

### Modified Capabilities

None. Structural edits on a sheet with conditional formatting stay refused, as the first change states, so its scenarios do not change.

## Non-goals

Rules that need a formula (`expression`, references, functions), the text, blank, error, duplicate, top, and average rule types, color scales, data bars, icon sets, extension rules, differential borders, differential gradient fills and pattern fills other than `solid`, a list of each rule that is not shown with its reason, conditional formatting edits, and moving rules on structural edits. These belong to later changes.

## Impact

The change adds `packages/core/src/cascade/conditional.ts`, `overlayDxf` in `cascade/cell-format.ts`, and `StyleLists.dxfs` in the store, and calls the overlay from `resolveCell` in `layout/sheet-layout.ts`. The editor snapshot gets the fields `conditionalHidden` and `conditionalFormat`, and the toolbar state gets `hint`. `packages/react` shows the notice and the hint, and `useEditorCommand` gets `hint`. `packages/i18n` adds the strings `status.conditionalHidden`, `status.conditionalHiddenShort`, `hint.conditionalFormat`, `toolbar.hintTitle`, and the refusal `copy-conditional-limit`. The copy text and the announcement of the active cell apply the rules. The fixture adds a known line position gap of 2 points for right-aligned italic and bold-italic numbers. Its cause has no measurement: no reference PDF has italic text from a cell format. The package `@xlsx-editor.dev/core` is private and not published, so no published API changes. Row heights do not change. Collaboration stays out of scope, as design D10 of the first change states.
