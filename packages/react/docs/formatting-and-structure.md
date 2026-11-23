# Format cells and change the structure

This guide describes cell formats, the format state of the snapshot, row and column inserts and deletes, and merged cells. For setup, see the [package README](../README.md).

## Format cells

To format the selected cells, call `applyFormat()` on the editor with the properties to set. Each cell keeps the properties that you do not set. The editor stores the result as a cell format (`x:xf` in `x:cellXfs`, ECMA-376 Part 1, 18.8.45 and 18.8.10), and the cell names it by index in `x:c/@s`. One call is one undo step:

<!-- sample: guide-apply-format -->

```tsx
ref.current?.getEditor()?.applyFormat({
  bold: true,
  fill: { type: 'solid', color: 'FFF2CC' },
  borders: { edges: 'outline', style: 'thin' },
  numberFormat: { code: '0.00' },
});
```

Give one of `id` and `code` in `numberFormat`. A patch that gives both is refused with `bad-format`. The `id` is a number format id (`x:numFmt/@numFmtId`): a built-in id of ECMA-376 Part 1, 18.8.30, or a custom id of the workbook. The `code` is a format code (`x:numFmt/@formatCode`). A border `style` is an `ST_BorderStyle` value (ECMA-376 Part 1, 18.18.3), such as `thin`.

The editor also has `toggleFormat()` for bold, italic, and underline, `clearValues()`, `clearFormats()`, and `clearAll()`. A format command that would create more than 16,384 new cells (one whole row), such as a format on a whole column or on the whole sheet, is refused with the reason `range-too-large`. A clear of formats writes a plain cell in each empty position that shows a row or column format, with the same limit, and is refused past it with the reason `clear-too-large`. The snapshot holds these refusals in `formatRefusal` and `clearRefusal`.

## Read the format state

The editor snapshot holds the format of the active cell in `activeFormat`. Its `numberFormat` holds the format id, the code of the id in the number locale, and `hasCode`. `hasCode` is `false` when the id has no code in the workbook or the number locale, so the cell shows General.

`decimalSteps` holds the state of the two decimal steps on the selection, in `increase` and `decrease`. Each is `null` when at least one selected cell can step. Otherwise it holds the refusal that the step gives, with the reason in `reason`, as [Decimal places](toolbar-and-keyboard.md#decimal-places) describes. `selectionSingle` is `true` when the selection is one cell or exactly one merged range.

The state reads the stored cells, and it is not the enabled state of the controls. View mode, input method composition, a protected sheet, and a cell edit with changed text also disable both controls. So read the enabled state only from `toolbarCommandState()` or `useEditorCommand()`, as for every other control.

If a true conditional formatting rule changes the state of a format control at the active cell, `conditionalFormat` holds the format state that the grid shows; otherwise it is `null`. A rule that changes only strike, superscript or subscript (`x:vertAlign`), or the underline style gives `null`, because no control shows them. A differential format does not apply alignment, so the alignment controls never get a hint. The format commands read and set the cell format in `activeFormat`.

If the sheet has conditional formatting that the grid can leave out, `conditionalHidden` is `true`, and the editor shows a notice. The causes are a rule that the engine does not evaluate, rules past the read limits, more shown rules or ranges than the caps of one cell or one row can take, rules in `x:extLst`, and a layout that ran out of its check budget. After a budget cut, the notice stays until the shown sheet changes or another workbook opens, and `conditionalFormat` is `null` until then.

## Insert and delete rows and columns

To insert or delete rows and columns, select whole rows or columns by their headings, or select a cell, and call `insertRows()`, `deleteRows()`, `insertColumns()`, or `deleteColumns()` on the editor. Rows go in above the selection, and columns to the left of it. Every reference in the workbook moves with the cells: formulas on every sheet, defined names, merged ranges, hyperlinks, column widths, and drawings. A reference to a deleted cell becomes `#REF!`.

If the workbook holds content that refers to cells and that the editor cannot move, such as conditional formatting, data validation, a table, or a chart, the command is refused with the reason `structure-blocked`, and the message names that content in words, such as "conditional formatting on sheet Data" or "the defined name Rates". The part path stays in the refusal's `subject` and `detail`, for logs.

To show a refusal in your own chrome, call `refusalText(refusal, t, snapshot().numberLocale)` of `@xlsx-editor.dev/core/editor`. The number locale groups the digits of a number in the message, such as a cell limit. Inside `XlsxEditor.Root`, get `t` from `useTranslation()` of `@xlsx-editor.dev/react`. Outside it, call `translatorFor(locale)` of `@xlsx-editor.dev/react` with the same `locale` that you pass to the root. A refusal is a typed code with fields, not a catalog key, so a host can branch on the code and on `lock`, and `refusalText` picks the one whole message for it.

A refusal with the code `locked` has `lock` set to the protection rule that refused it. The sheet protection is `x:sheetProtection` (ECMA-376 Part 1, 18.3.1.85), and the workbook protection is `x:workbookProtection` (ECMA-376 Part 1, 18.2.29). For the rules, see [View mode and protection](view-mode-and-protection.md). The message names that rule:

- `cell`: a locked cell, or a new cell format that would change the protection of a cell.
- `format`: the sheet does not allow formatting cells.
- `structure`: the sheet does not allow the insert or the delete.
- `size`: the sheet does not allow a column width or a row height.
- `workbook`: the workbook protection does not allow a sheet add, rename, delete, or move.
- `merge`: the sheet is protected, so no merge or unmerge is allowed.

A `locked` refusal of a cut or a paste also has `command` set to `cut` or `paste`, so its message names that command. If the whole sheet is selected, the command is refused with the reason `whole-sheet`. Other edits stay available.

## Merge cells

To merge the selection into one cell, call `mergeCells()` on the editor. A merged range is an `x:mergeCell` in the `x:mergeCells` of the worksheet (ECMA-376 Part 1, 18.3.1.54 and 18.3.1.55). To also center its text, call `mergeCells({ center: true })`. The merged cell keeps the value of its top-left cell only. If other cells in the selection hold values, the merge is refused with the reason `merge-drops-values`, and the refusal's `subject` names the first of those cells, such as `C1`.

To merge anyway, call `mergeCells({ dropValues: true })`: the merge then clears those cells in the same undo step. That clear can refuse with the reasons of a clear of values, such as `shared-formula` for a part of an array formula range. A `shared-formula` or `cell-alternate-content` refusal of that clear has `command` set to `merge`.

A merge of one cell is refused with `merge-one-cell`, and a selection that overlaps a merged range with `merge-overlap`, which names that range. With `center`, a range too large to format is refused with `range-too-large`. To unmerge, call `unmergeCells()`: it unmerges each merged range that the selection meets.

To enable your own merge controls, call `mergeCommandState(snapshot, command)` with `merge`, `mergeCenter`, or `unmerge`. It gives `enabled` and a `disabledReason` catalog key, as `toolbarCommandState` does. For an enabled merge, `dropsValuesAt` names the first cell that would lose its value, or is null. The state reads the snapshot fields `selectionMerged` and `selectionDropsValuesAt`. An open cell edit counts as it would be after its commit. If `dropsValuesAt` is set, ask the user before the merge:

<!-- sample: guide-merge-button -->

<!-- prettier-ignore -->
```tsx
import { mergeCommandState, useEditorSnapshot, useXlsxEditor } from '@xlsx-editor.dev/react';

function MergeButton({ confirm }: { confirm: (cell: string) => boolean }) {
  const editor = useXlsxEditor();
  const state = mergeCommandState(useEditorSnapshot(), 'merge');
  const merge = () => {
    const cell = state.dropsValuesAt;
    if (cell === null) editor?.mergeCells();
    else if (confirm(cell)) editor?.mergeCells({ dropValues: true });
  };
  return (
    <button type="button" disabled={!state.enabled} onClick={merge}>
      Merge
    </button>
  );
}
```
