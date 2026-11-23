## 1. Rules

- [x] 1.1 Read the rules of a worksheet with the limits of design D2, and sort them by priority.
- [x] 1.2 Evaluate `cellIs` rules with literal operands by design D3, with the reasons of the rules that are not shown.

## 2. Overlay

- [x] 2.1 Add `StyleLists.dxfs` and `overlayDxf`, and lay the formats of the true rules over each resolved cell in layout, by design D4.
- [x] 2.2 Test in layout that B1 of `preserved-parts` shows bold, as its reference PDF draws it, and test that the toolbar shows the cell format.

## 3. Documentation

- [x] 3.1 Add the approximation "Conditional formatting rules" to both lists, and move the support matrix row to partial.

## 4. Measurements

- [x] 4.1 Add the fixture `conditional-rules` with its reference PDF and save check, test the face of each value in layout, and remove the cases that it measures from the approximation.
- [x] 4.2 Measure the other operators, empty cells, underline, theme colors, and the fill cases, and evaluate the cells without `x:c` in layout. The cases that no reference PDF measures stay in the approximation "Conditional formatting rules" for a later change.

## 5. Signals

- [x] 5.1 Show a notice for each sheet with a rule that is not shown, and give the format controls a hint when a true rule changes the active cell, with i18n strings and tests.
