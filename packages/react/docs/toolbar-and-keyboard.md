# Toolbar and keyboard

This guide describes the toolbar of `@xlsx-editor.dev/react` and the keyboard of the editor. For setup and the component props, see the [package README](../README.md).

## Toolbar

`XlsxEditor` shows a toolbar over the grid. To get the workbook when the user selects **Save** in the toolbar, pass `onSave`. It receives the workbook as bytes of the kind that was opened: `snapshot().workbookKind` is `workbook` or `template`, so you can choose the file extension. With `onSave`, Control or Command and S in the editor also saves, also when you hide the save control. Without `onSave`, the save control is disabled, and the browser keeps the shortcut:

<!-- sample: guide-on-save -->

```tsx
<XlsxEditor
  document={bytes}
  onSave={(edited, saved) => {
    if (saved.stillOpen) upload(edited);
  }}
  style={{ height: 600 }}
/>
```

The bytes arrive after the press, because the save runs in short tasks. A press during a running save of the same revision joins that save, so `onSave` runs once. `saved.stillOpen` is `false` when another open, `close()`, or `destroy()` ended the workbook during the save. `saved.revision` is the revision that the bytes hold, so an edit made during the save stays unsaved. When saves of different revisions overlap, `onSave` never receives older bytes after newer bytes.

A browser opens a file dialog only soon after a user action. The bytes of a large workbook can arrive after that time, and `showSaveFilePicker` then rejects with a `SecurityError`. If this happens, keep the bytes and ask the user to select a control of your own, and open the dialog from that press.

In your own layout, place `XlsxEditor.Toolbar` inside `XlsxEditor.Viewport`. It takes these props:

| Prop | Type | Default | What it does |
| --- | --- | --- | --- |
| `onSave` | `(bytes: Uint8Array<ArrayBuffer>, saved: SavedWorkbook) => void` | none | Receives the workbook bytes of the save control, and of Control or Command and S in the editor, also with the save control hidden. Without it, the save control is disabled, and the browser keeps the shortcut. A refused save calls nothing. Check `saved.stillOpen` before you store the bytes: it is `false` when the workbook on screen is no longer the one that was saved. |
| `onSaveStart` | `(saving: Promise<SavedWorkbook \| null>) => void` | none | Called once when a press of the save control, or of Control or Command and S in the editor, starts a delivery to `onSave`. It gets the promise of the result, so you can show that the save runs. A press that joins the running save of an earlier press and a refused save do not call it. A press that joins a save from `saveInTasks` calls it, because that save has no delivery yet. The promise resolves to `null` when the save is refused later. |
| `opening` | `boolean` | `false` | Tells the toolbar that your page opens another workbook while the editor still shows the old one, for example during a download or a file read before `load`. While it is `true`, the save control is disabled with the reason that a workbook is loading, and Control or Command and S in the editor starts no save work. After such a press, and only then, the snapshot holds the refusal `loading`, with `command: 'save'`. A press reads the value of the latest render. When the prop changes to `false`, the toolbar forgets that refusal and keeps any other refusal. |
| `overflow` | `boolean` | `true` | Keeps the toolbar on one row and moves the groups that do not fit into the **More** panel. `false` wraps the controls onto more rows. It has no effect with `preset={false}`. |
| `hidden` | `ChromeSlotId[]` | `[]` | The slots to leave out. A group whose slots are all hidden is left out. It has no effect with `preset={false}`. |
| `preset` | `boolean` | `true` | `false` renders only the children, in their order, with no measure and no **More** panel. |
| `children` | `ReactNode` | none | Parts that change the toolbar. See [Toolbar parts](#toolbar-parts). |
| `className` | `string` | none | Added after the `xlsx-toolbar` class. |

To leave out the history group and the save control:

<!-- sample: guide-hidden-slots -->

```tsx
<XlsxEditor.Toolbar hidden={['history.undo', 'history.redo', 'file.save']} />
```

The toolbar shows the groups of `CHROME_GROUPS` in this order. Each slot id is the group id and the control id:

| Group | Slot ids |
| --- | --- |
| History | `history.undo`, `history.redo` |
| Font | `font.family`, `font.size`, `font.bold`, `font.italic`, `font.underline`, `font.color` |
| Fill | `fill.color` |
| Borders | `borders.outline`, `borders.all`, `borders.none` |
| Alignment | `alignment.left`, `alignment.center`, `alignment.right`, `alignment.top`, `alignment.middle`, `alignment.bottom`, `alignment.wrap` |
| Number format | `number.format`, `number.custom`, `number.increaseDecimals`, `number.decreaseDecimals` |
| Insert and delete (one menu) | `cells.insertRows`, `cells.insertColumns`, `cells.deleteRows`, `cells.deleteColumns` |
| Clear (one menu) | `clear.contents`, `clear.formats`, `clear.all` |
| File | `file.save` |

A control with the `dialog` shape in `CHROME_GROUPS`, such as `number.custom`, opens a dialog. Its `execute` needs the value that the dialog produces, so render it as a dialog opener, not as a plain button.

The font size list also takes a typed size from 1 to 409 points. A number format that is not in the list shows as `Custom:` and its format code.

The number format list has these entries: General, Number, Currency, Accounting, Percentage, Scientific, Fraction, Short date, Long date, Time, and Text. An entry with a built-in id of ECMA-376 Part 1, 18.8.30, writes that id. If the workbook gives the id another code, the entry writes the code of the id in the number locale instead. Currency (`"$"#,##0.00`) and Accounting take the currency symbol from `numberLocale.currencySymbol`. The symbol goes before the number in every number locale. Long date writes `dddd, mmmm d, yyyy` in every number locale, with English day and month names.

A code that differs from an entry only in its decimal places shows the entry's name, such as Percentage for `0.0%`. The list marks that entry as current, not as picked. A pick of the current entry writes the code of that entry: Percentage on a `0.000%` cell writes `0.00%`. On one cell or one merged range that already shows the picked entry, a pick writes nothing and adds no undo step.

The Number entry has no thousands separator, so the list shows built-in ids 3 (`#,##0`) and 4 (`#,##0.00`) as Custom. An entry shows only its name, with no title, so a screen reader does not read its code. The accessible name and the title of the closed list give the name and the decimal places, such as "Percentage, 1 decimal place", but never the code of an entry.

To build your own list, call `numberFormatListView(value, t, snapshot.numberLocale)` of `@xlsx-editor.dev/core/editor`, which `@xlsx-editor.dev/react` also exports. Pass the `value` of `useEditorCommand('number.format')`. The result holds the translated `entries`, the `text` of the closed list, the `name` for its accessible name and title, and the `current` entry. The toolbar list shows itself from the same function.

### Decimal places

`number.increaseDecimals` and `number.decreaseDecimals` add or remove one decimal place in each number section of the code. The other parts of the code stay the same: sections, colors, conditions, literals, the text section, and the exponent. A decrease removes the decimal point with the last place, so `0.` steps up to `0.0` and back to `0`. These rules apply to the code:

- In the zero section of a code with three or more sections, a run of `?` can follow a quoted literal or a dash. Such a run, as `"-"??` in the Accounting code, pads the literal. A step adds or removes one `?` there, and no decimal point.
- With a literal other than a dash, or with no fill, a decrease to no places removes the last `?` of that padding. An increase does not add it back. This is the known approximation `decimal-padding`.
- A run of `?` before a quoted literal, as `??" kg"`, steps as a number.
- A literal or a second point can split the places, as in `0.00" "0`. A step can then move a place to the other side of the literal. This is the known approximation `decimal-split-places`.
- A step does not change a date or time, text, a fraction, or a General section with a color, a condition, a literal, or other sections.
- A cell keeps its format when the step would make a comma scale the number. It also keeps it when the step would add or remove a group separator. So `0.00,0` does not step down.
- A cell with a control character in its code keeps its format, and so does a cell whose new code would be longer than 255 characters. The other selected cells step.
- An increase stops at 30 places in a section.

On General, a step writes one place more or one place less than the grid shows for the number in the cell. A narrow column shows fewer places, such as `3.141593` for 3.14159265358. If the grid shows the number in exponent form, such as `1E+20`, the step writes a scientific code, such as `0.0E+00`. With no text measurer, or when the cell shows only `#`, the step reads the General text at full width. Such a cell can still show only `#` after the step, until the column is wide enough. On one empty General cell, an increase writes `0.0`, and creates the cell when it has no `x:c` element.

On a selection, each selected cell steps from its own number format, in one undo step. These rules apply:

- A cell whose code cannot change keeps its format, so a selection of a date and a number changes only the number.
- A cell that holds text, a boolean, or an error does not step. If no selected cell holds a number, both controls are disabled.
- A date cell (`t="d"`) steps from the serial number that the grid shows for it. A `t="d"` value that the editor cannot read shows as text, so it does not step.
- Among more cells, an empty cell keeps its format, and so does an empty position. One empty cell steps from its own code.
- A merged range steps as its top-left cell.
- Every General number of a column takes the step of the first General number of that column. So the column lines up. The first is the active cell in its own column, and else the first in row order. A General number in exponent form steps alone.

The step reads only the cells that exist, at most 16,384 inside the selection. On a whole column, a cell that you create later takes the column format, which the step does not change. The number format list has no such limit. Like the other format commands, it is refused only when it would create more than 16,384 new cells.

After a step or a number format change, the live region says the new text of the active cell. On a selection of more than one cell, it first says that the decimal places or the number format changed. A number format change that also commits typed text is the exception: the live region then says only the new text of the active cell. If the text stays the same, or changes only in its spaces, the live region says only that the decimal places or the number format changed.

If no selected cell can step, the control is disabled, and the command is refused with `decimals-unchanged`. Past 16,384 cells that exist, it is refused with `decimals-too-large`. The refusal holds the disabled reason of the control in `reason`. For `decimals-unchanged`, `refusalText` shows that reason. For `decimals-too-large`, it shows the cell limit.

The reason names the problem of the active cell when it holds a number, and else the problem of the first selected cell with a number. It speaks of the cell when `selectionSingle` is `true`, and else of a selected cell.

While a cell edit holds changed text, both controls are disabled with `editPending`, and the command is refused with `edit-pending`. This differs from Bold and the number format list, which commit the typed text first. The difference is deliberate. A step reads the stored format of each cell. The commit of typed text can change that format, such as `12%` on a General cell. So you commit or cancel the edit, and then see what the step changes.

To step from your own control, call `useEditorCommand('number.increaseDecimals').execute()`. A host with no React calls `runToolbarCommand(editor, 'number.increaseDecimals')` of `@xlsx-editor.dev/core/editor`. The step has no method on the editor instance. It is a toolbar command, as indent is in DOCX Editor's command map, and the command keeps its state and its refusals in one place.

#### Decimal step state in your own chrome

To show why a decimal step is disabled, render the control from `useEditorCommand('number.increaseDecimals')`. Use `isEnabled` for the disabled state, and show `disabledReason`, which is already translated. A host with no React reads `toolbarCommandState(snapshot, 'number.increaseDecimals').disabledReason` of `@xlsx-editor.dev/core/editor`, and translates it with `disabledReasonText(reason, t)`.

To get the code that a list entry writes, call `numberFormatChoiceCode(value, locale)`. To get the list value of `activeFormat.numberFormat`, call `numberFormatValue(format, snapshot.numberLocale)`. A value that is no entry starts with `CUSTOM_NUMBER_FORMAT_PREFIX`.

### Number format dialog

The last item of the number format list, **More number formats…**, opens the number format dialog of the `number.custom` slot. If you pass `hidden={['number.format']}`, the slot shows as its own button. In the dialog, type a format code (ECMA-376 Part 1, 18.8.31), or pick one of the custom codes of the workbook. The preview shows the active cell's value through the code, as the grid would show it. If the active cell is empty, the preview shows a sample number. If the editor cannot apply the code, the dialog shows the reason, marks the field invalid, and **Apply** changes nothing.

**Apply** sets the number format of the selection in one undo step. On one cell, an **Apply** of the code that the dialog opened with changes nothing, so the cell keeps its format. A new code gets a custom number format from id 164, and a code that the workbook already holds uses its id. A code equal to a built-in code is written as the built-in id, so the cell shows the number locale's code for that id, as the preview does.

To apply a code from your own control, pass it to `useEditorCommand('number.custom').execute(code)`. To show why a code cannot be applied before you call `execute`, call `numberFormatCodeProblem(code)`. To build your own number format dialog, use the functions that the toolbar dialog uses. `@xlsx-editor.dev/react` exports each of them:

1. In the event handler that opens the dialog, call `numberFormatDialogStart(editor)`, and keep the result as `start`. It commits an open cell edit, so do not call it during render. Keep that object itself: for a copy, `numberFormatDialogApplyStep` returns `closed`. It gives the code to show first in `start.initial`. If it returns `null`, the slot is disabled or the commit is refused: do not open the dialog.

   Render your opener from `useEditorCommand('number.custom')`, disabled when `isEnabled` is `false` and with `disabledReason` as its reason. Then a `null` start means a refused commit: show `refusalText(snapshot.refusal, t, snapshot.numberLocale)`, or keep `XlsxEditor.StatusMessage` in the layout.

1. To list the custom codes of the workbook, call `workbookNumberFormatCodes(editor)` once for each change of the editor state, such as in a `useMemo` keyed on `useEditorSnapshot()`, not for each key. It lists at most `MAX_LISTED_NUMBER_FORMATS` codes. The codes come from the file: show them as text only.
1. As the user types, call `previewNumberFormat(editor, code)`. If `ok` is `false`, show its `problem`; it has no `text`. Otherwise show `text`. If `stale` is `true`, the value is a formula result that an edit can have changed: show it with `t('grid.staleResult', { text })`.

   If the active cell is empty, the preview formats `NUMBER_FORMAT_SAMPLE` and sets `sample`. If the code has a color, `color` holds it as `RRGGBB` from the engine's palette. A rich text cell shows its stored text for every code.

1. Open the dialog as a native `dialog` element with `showModal()`, as the sample does. The browser moves focus into it, keeps Tab in it, sends `cancel` on Escape, and returns focus to the opener when it closes. Give the element `aria-modal="true"`: then Control or Command and S in it does not save the workbook. If you build the dialog from another element, focus its first control, keep Tab in it with `trapTabWithin(dialog, event)` on `keydown`, and close it on Escape.
1. While the dialog is open, close it when `numberFormatDialogCurrent(editor, start)` is `false`: another workbook loaded, the selection moved, the number locale changed, or the slot is disabled. Read it with `useEditorState`: the selector reads the editor, and `useEditorState` calls it again on each change of the editor state.
1. On **Apply**, call `numberFormatDialogApplyStep(editor, start, code)`, and act on its `kind`. For `closed` and `unchanged`, close the dialog. For `apply`, run `execute(code)`, and close the dialog when it returns `true`. If it returns `false`, the editor refused the edit: show `refusalText(snapshot.refusal, t, snapshot.numberLocale)`, or keep `XlsxEditor.StatusMessage` in the layout. For `problem`, keep the dialog open.

To show a problem, translate its key with its parameters. Each key is in `NUMBER_FORMAT_CODE_PROBLEM_KEYS` and in the catalog of `@xlsx-editor.dev/i18n`.

The following sample is a minimal dialog component with the close check and the **Apply** step. It leaves out the preview and the list of steps 2 and 3. `start` comes from step 1, and the opener passes `close`, a function from `useCallback`:

<!-- sample: guide-number-format-dialog -->

```tsx
import { useCallback, useEffect, useLayoutEffect, useRef, useState } from 'react';
import {
  numberFormatDialogApplyStep,
  numberFormatDialogCurrent,
  refusalText,
  useEditorCommand,
  useEditorState,
  useTranslation,
  useXlsxEditor,
  type NumberFormatDialogStart,
} from '@xlsx-editor.dev/react';

function NumberFormatHostDialog({
  start,
  close,
}: {
  start: NumberFormatDialogStart;
  close: () => void;
}) {
  const editor = useXlsxEditor()!;
  const { t } = useTranslation();
  const command = useEditorCommand('number.custom');
  const [code, setCode] = useState(start.initial);
  const [error, setError] = useState<string | null>(null);
  const ref = useRef<HTMLDialogElement>(null);
  // The browser focuses the field, keeps Tab in the dialog, and returns focus on close.
  useLayoutEffect(() => {
    const dialog = ref.current!;
    dialog.showModal();
    return () => dialog.close();
  }, []);
  const current = useEditorState(
    useCallback(() => numberFormatDialogCurrent(editor, start), [editor, start])
  );
  useEffect(() => {
    if (!current) close();
  }, [current, close]);
  function apply() {
    const step = numberFormatDialogApplyStep(editor, start, code);
    if (step.kind === 'problem') {
      setError(t(step.problem.key, step.problem.params));
    } else if (step.kind !== 'apply' || command.execute(code)) {
      close();
    } else {
      // The editor refused the edit: show its reason.
      const snapshot = editor.snapshot();
      setError(
        snapshot.refusal
          ? refusalText(snapshot.refusal, t, snapshot.numberLocale)
          : t('dialogs.numberFormat.notApplied')
      );
    }
  }
  return (
    <dialog
      ref={ref}
      aria-modal="true"
      aria-label={t('dialogs.numberFormat.title')}
      onCancel={(e) => {
        e.preventDefault();
        close();
      }}
    >
      <input
        aria-label={t('dialogs.numberFormat.code')}
        value={code}
        onChange={(e) => setCode(e.currentTarget.value)}
      />
      {error ? <p role="alert">{error}</p> : null}
      <button type="button" onClick={apply}>
        {t('common.apply')}
      </button>
    </dialog>
  );
}
```

The `value` of `number.custom` is `General` for a built-in id that has no code in the editor's number locale. If you send that value back through `execute`, the cells change to the `General` format: on a range, the dialog's **Apply** of that unchanged value does the same. A built-in id whose locale code is the value keeps its id.

### Control state

The toolbar has no state of its own. Each control reads its state from `toolbarCommandState()` of `@xlsx-editor.dev/core/editor`, so the toolbar and your own controls show the same answer. The sheet tabs read the state of the add button and the sheet menu rows from `sheetCommandState()` of the same entry, which `@xlsx-editor.dev/react` also exports. A disabled control shows its label and the engine's reason as its tooltip (`toolbar.disabledTitle`), and the reason as its accessible description. The reasons are keys of the `disabledReason` section of the catalog in `@xlsx-editor.dev/i18n`:

- `editorNotReady`: no workbook is open yet.
- `loading`: a workbook is loading. The save control also takes this reason while the `opening` prop is `true`.
- `nothingToUndo` and `nothingToRedo`: the history has no step in that direction.
- `editPending`: the open cell edit has changed text. Its commit would clear the redo steps, and a decimal step would read a format that the commit can change.
- `rangeTooLarge`: a format command, or a merge and center, would create more than one whole row of new cells.
- `clearTooLarge`: a clear of formats would create more than one whole row of new cells.
- `structureBlocked`, `structureBlockedUnnamed`, `structureUnavailable`, and `wholeSheet`: the sheet or the selection refuses an insert or a delete.
- `noSave`: the toolbar has no `onSave`.
- `viewMode`: the editor is in view mode.
- `composing`: an input method composes in the open cell editor, and every control waits for it to finish.
- `formatLocked`: the sheet is protected, and its protection does not allow formatting cells, so the format and clear-format commands are disabled.
- `stylesUnread`: the styles part stays as bytes past a read limit, so the format commands and merge and center are disabled. On a sheet that is not protected, Clear formats and Clear all stay enabled, because they give the default format.
- `stylesUnreadProtected`: the styles part stays as bytes, and the sheet is protected. The protection of each cell is in that part, so Clear formats and Clear all are disabled.
- `structureLocked`: the sheet is protected, and its protection does not allow the insert or the delete of rows or columns that the command makes.
- `workbookLocked`: the workbook protection does not allow a sheet add, rename, delete, or move, so `sheetCommandState()` disables them.
- `mergeLocked`: the sheet is protected, so `mergeCommandState()` disables merge and unmerge.
- `mergeOneCell`: one cell is selected, so there is nothing to merge.
- `mergeOverlap`: the selection overlaps a merged range, so a merge waits for an unmerge.
- `noMergedRange`: the selection has no merged range to unmerge.
- `lastSheet`: the sheet is the last visible sheet, so `sheetCommandState()` disables Delete.
- `notWorksheet`: the sheet is not a worksheet, such as a chartsheet or a macro sheet, or its part is missing. `sheetCommandState()` disables Rename, Delete, and both moves of it, and its tab shows the reason.
- `sheetEditBlocked`: content in the workbook can refer to any sheet. That is a chart, a pivot table, or a sheet that is not a worksheet. It is also a workbook or package relationship outside the allowlist, such as data connections. `sheetCommandState()` then disables Rename and Delete of every sheet, and each worksheet tab shows the reason. The snapshot holds the refusal in `sheetEditRefusal`.
- `locked`: the sheet is protected, and a locked cell in the selection holds content, so Clear contents and Clear all would change it.
- The keys that start with `decimals`: no selected cell can take a decimal step. The key names the cause, such as `decimalsDate`. The engine picks the key for the selection, so you show `disabledReason` and never map a cause to a key.
- `cellAlternateContent`: a cell of the selection holds a value and also holds data in a form that the editor cannot update (`selectionContentBlocked`), so Clear contents and Clear all are disabled.
- `cellAlternateContentClearAll`: the reason of `cellAlternateContent` for Clear all while Clear formats is enabled. It also points to Clear formats.
- `formulaRange`: a clear of the selection would change a formula range that the editor cannot rewrite or remove (`selectionFormulaRangeBlocked`), so Clear contents and Clear all are disabled. The clear command gives `shared-formula` for the same selection. A clear that would move the formula of a shared child out of the grid is refused only when it runs, so the field is null for it.
- `formulaRangeClearAll`: the reason of `formulaRange` for Clear all while Clear formats is enabled. It also points to Clear formats.
- `formulaCell` and `formulaCellClearAll`: the reasons of `formulaRange` and `formulaRangeClearAll` when one cell's formula blocks the clear of a shared formula master. They name the master to leave out (`selectionFormulaRangeMasterCell`) and the cell whose formula blocks it (`selectionFormulaRangeCauseCell`).
- `formulaUnreadable` and `formulaUnreadableClearAll`: the reasons of `formulaCell` and `formulaCellClearAll` when the formula of the master cannot be read (`selectionFormulaRangeCauseUnreadable` is `true`).
- `notWired`: the slot id is not in the registry.

An enabled control can also have a `hint` object, whose `key` is a key of the `hint` section of the catalog. The toolbar shows the hint after the label in the tooltip, and as the accessible description. The key `hint.conditionalFormat` means that a true conditional formatting rule shows a different state of this control at the active cell. The control still reads and sets the cell format.

The key `hint.decimalsLargeSelection` is on a decimal step when the selection is too large to format. Then the other format controls are disabled with `rangeTooLarge`, and the step still changes the selected cells that hold numbers. When a layout of the grid runs out of its conditional formatting check budget, the notice shows, and no control has this hint until another sheet or workbook shows.

If a cell edit holds changed text, undo is available and cancels the edit, as `undo()` on the editor instance does. Redo is not available while the edit holds changed text, with the reason `editPending`, and `redo()` refuses it with `edit-pending`.

### Toolbar parts

If you build your own layout with `XlsxEditor.Toolbar`, the toolbar takes parts as its children. `XlsxEditor` renders its own toolbar and does not take parts. The [Toolbar](https://github.com/eigenpal/xlsx-editor/blob/main/docs/site/content/customization/toolbar.mdx) page of the docs site shows each part with React and Vue code.

- A named part, such as `XlsxEditor.Toolbar.Bold`, stands for one slot. `<XlsxEditor.Toolbar.Bold hidden />` is `<XlsxEditor.Toolbar.Slot slotId="font.bold" hidden />`. A named part changes the built-in control with `icon`, `asChild`, and `className`. The children of a button part are its icon.
- `XlsxEditor.Toolbar.Slot` with a `slotId` replaces that slot with its children, where the slot stands. With `hidden`, it removes the slot and its children, as the `hidden` list does. With `icon`, `asChild`, or `className` and no other children, it changes the built-in control.
- `XlsxEditor.Toolbar.Group` with a new `id` adds a host group. The bar measures it with its own groups, and moves it into the **More** panel when it does not fit. In the panel, the group shows as a section with its heading: `label`, else the translated `labelKey`, else the `id`. A host group follows every built-in group, or the group that `after` names. A group with a built-in `id`, such as `file`, shows its children after its own controls. With `hidden`, it removes the group and its children.

Pass the parts as direct children of the toolbar or of a Fragment. The toolbar does not read a part that a component of your own renders. In the preset bar, such a part shows in the block of other children, so a `Toolbar.Group` there never moves into the **More** panel.

The last part for a slot or a group wins. If a slot is in the `hidden` list, the list wins over a `Toolbar.Slot` with children. In a menu group, such as Clear, the replaced slot's content shows after the menu, and the menu keeps the other slots.

The bar moves whole groups into the **More** panel in collapse order. The built-in groups leave from the end of the bar: `file` has the priority 10, and `history` 90. A host group with no `priority` leaves before every built-in group. A lower `priority` leaves first, and a `pinned` group stays at every width. In the panel, the content of a `Toolbar.Slot` and of a `Toolbar.Group` shows in the section of its group, or their `overflowContent` when they have it.

- `XlsxEditor.Toolbar.Action` is a button for an action of your own, with `label`, `icon`, `onSelect`, `active`, `disabled`, and `disabledReason`. A disabled action stays focusable and gives its reason. In a group in the **More** panel, it is a row with its icon and label.
- `XlsxEditor.Toolbar.Separator` draws a rule between two controls.

Other children, such as a status text, show in one block after the groups and the **More** button. The block always stays in the bar, and the bar measures the groups in the room that is left.

The parts take these props:

| Part | Prop | Type | What it does |
| --- | --- | --- | --- |
| `Toolbar.Slot` | `slotId` | `ChromeSlotId` | The slot that the part names. Required. |
| `Toolbar.Slot` | `hidden` | `boolean` | Removes the slot where it stands. |
| `Toolbar.Slot` | `icon` | `ReactNode` | An icon in place of the registry icon. Button slots only. |
| `Toolbar.Slot` | `asChild` | `boolean` | Renders the built-in control as the one element in the children. Button slots only. |
| `Toolbar.Slot` | `className` | `string` | A class after the class of the control, or of the trigger of a list or a palette. |
| `Toolbar.Slot` | `overflowContent` | `() => ReactNode` | Content for the slot's row in the **More** panel. |
| `Toolbar.Slot` | `children` | `ReactNode` | Content that takes the place of the slot. Without children, the part renders the built-in control of the slot. |
| `Toolbar.Group` | `id` | `string` | A built-in group id, such as `file`, or a new id for a host group. Required. |
| `Toolbar.Group` | `label` | `string` | The heading of a host group in the **More** panel, and its accessible name. |
| `Toolbar.Group` | `labelKey` | `string` | A catalog key for the heading, translated with the toolbar's locale. `label` wins over it. |
| `Toolbar.Group` | `after` | `string` | The id of the group that a host group follows. |
| `Toolbar.Group` | `priority` | `number` | The collapse priority. A lower number leaves the bar first. |
| `Toolbar.Group` | `pinned` | `boolean` | Keeps the group in the bar at every width. |
| `Toolbar.Group` | `hidden` | `boolean` | Removes the group from the bar and from the **More** panel. |
| `Toolbar.Group` | `overflowContent` | `() => ReactNode` | Content for the group's section of the **More** panel. |
| `Toolbar.Group` | `className` | `string` | A class after `xlsx-toolbar-group`. The group draws its separator line as `::before`, so the `gap`, `padding`, `border`, and `background` of the class also apply around the line. The built-in line rule has the weight of one class. A `::before` rule of your class overrides the line when it comes after the editor stylesheet. To remove the line, set `content: none`. |
| `Toolbar.Group` | `children` | `ReactNode` | The controls of the group. A host group with no children is left out. |

The following sample replaces the save control, removes the underline control, and adds a host group after the history group:

<!-- sample: guide-toolbar-parts -->

```tsx
<XlsxEditor.Toolbar>
  <XlsxEditor.Toolbar.Slot slotId="file.save">
    <SaveToCloudButton />
  </XlsxEditor.Toolbar.Slot>
  <XlsxEditor.Toolbar.Slot slotId="font.underline" hidden />
  <XlsxEditor.Toolbar.Group id="share" label="Share" after="history">
    <ShareButton />
  </XlsxEditor.Toolbar.Group>
</XlsxEditor.Toolbar>
```

With `preset={false}`, the toolbar renders only its children. There, a `Toolbar.Slot` with no children renders the built-in control of its slot, and a `Toolbar.Group` renders its children as one group. The parts render in the order that you give, so `after` has no effect:

<!-- sample: guide-toolbar-own -->

```tsx
<XlsxEditor.Toolbar preset={false}>
  <XlsxEditor.Toolbar.Group id="text" label="Text">
    <XlsxEditor.Toolbar.Slot slotId="font.bold" />
    <XlsxEditor.Toolbar.Slot slotId="font.italic" />
  </XlsxEditor.Toolbar.Group>
  <XlsxEditor.Toolbar.Slot slotId="history.undo" />
</XlsxEditor.Toolbar>
```

The arrow keys move between the controls of this bar too, as in the preset bar.

A `Toolbar.Slot` with no children in a host group places a second copy of a built-in control. To move the control into the host group, also hide its slot with a `Toolbar.Slot` with `hidden`, or with the `hidden` list.

In a development build, the toolbar prints a warning once for each part that cannot do what it says. A build is a development build when `import.meta.env.DEV` is `true`, or, where that flag is absent, when `process.env.NODE_ENV` is not `production`. A bundle that sets neither prints no warning. Each warning starts with `[xlsx-editor]`. The toolbar warns about these parts:

- A `Toolbar.Slot` whose `slotId` is not a slot of `CHROME_GROUPS`. In the preset bar, its children show in the block of other children.
- A `Toolbar.Slot` with children whose slot the `hidden` list names.
- A `Toolbar.Slot` with children whose group is hidden.
- A `label`, `labelKey`, or `after` on a built-in group.
- An `icon` or `asChild` on a list, a palette, or the dialog opener.
- An `asChild` part whose children are not one element.
- An `icon` or `className` on a `Toolbar.Slot` with content.
- An `after` that names no group, that names its own group, or that leads into a cycle of anchors.
- An `after` that names a group that stands last because its own anchors name no group.

### Controls of your own

To place controls of your own anywhere in the layout, call `useEditorCommand(slot)` inside `XlsxEditor.Root`. It returns the same state that the toolbar shows:

| Field | Type | What it holds |
| --- | --- | --- |
| `execute` | `(value?: string) => boolean` | Runs the command. Pass the picked value of a value slot, such as `'14'` for `font.size`. Returns `true` when the command ran and `false` when the editor refused it. |
| `isActive` | `boolean` | The command is applied at the active cell, such as bold on a bold cell. |
| `isEnabled` | `boolean` | The editor runs the command now. |
| `disabledReason` | `string \| null` | The engine's reason when the command is disabled, translated. |
| `hint` | `string \| null` | The engine's hint for an enabled command, translated, such as the text of `hint.conditionalFormat`. Null on a disabled command. |
| `value` | `string \| null` | What a value slot shows, such as the font name, the size in points, or the format code of the active cell for `number.custom`. |

The component re-renders only when the state of its slot changes, and the result keeps its identity while its fields are equal:

<!-- sample: guide-bold-button -->

```tsx
import { useEditorCommand } from '@xlsx-editor.dev/react';

function BoldButton() {
  const bold = useEditorCommand('font.bold');
  return (
    <button
      type="button"
      aria-pressed={bold.isActive}
      aria-disabled={!bold.isEnabled}
      title={bold.disabledReason ?? 'Bold'}
      onMouseDown={(e) => e.preventDefault()}
      onClick={() => bold.execute()}
    >
      Bold
    </button>
  );
}
```

Inside a toolbar, `file.save` hands the bytes to the toolbar's `onSave`, also in a control of your own that is a toolbar part's content, such as the content of a `Toolbar.Slot`. Outside a toolbar, `file.save` is disabled: call `save()` or `saveInTasks()` on the ref instead.

### Toolbar keyboard

| Keys | Action |
| --- | --- |
| Tab | Move to the next control of the toolbar. Each control is a stop, disabled controls too. |
| Left arrow, Right arrow, Home, End | Move between the controls of the toolbar, your own controls too. The keys stop at the first and the last control. A key that your control handles and cancels does not move the focus. |
| Enter, Space, or Down arrow on a list, palette, or menu button | Open it, with the keyboard on the selected item or the first item |
| Down arrow on an open list | Move the keyboard into it |
| Up arrow, Down arrow, Home, End | Move between the items of an open list or menu. In a palette, Left arrow and Right arrow move too. |
| Enter or Space on an item | Apply the item, and return the keyboard to the grid |
| Escape | Close the list, and return the keyboard to its button |

A disabled control keeps the keyboard: pressing it does nothing. A control of your own that cannot take the focus, such as a hidden input or a disabled control, is not a stop. If a control of your own moves the focus on when it gets the focus, for example to a field of its own or to a popup, the focus stays where the control moved it.

A field of your own in the toolbar keeps the keys that it needs:

- A text field keeps Home and End. It keeps Left arrow and Right arrow until the caret is at that end of the text, or while text is selected.
- A field whose caret the page cannot read, such as an email or a number field, keeps all four keys.
- A text area and an editable element keep all four keys.
- A `<select>` element keeps Home and End.
- A radio button keeps Left arrow and Right arrow, and a range input (`<input type="range">`) keeps all four keys, because they change the value.
- A control in an element of your own with the role `dialog` or `menu` keeps all four keys. Move the focus in it yourself. A toolbar inside such an element of your own keeps its keys.

To leave a field that keeps a key, press Tab.

## Keyboard

| Keys | Action |
| --- | --- |
| Arrow keys | Move the active cell |
| Shift and an arrow key | Extend the selection. A selection of whole rows or columns stays whole rows or columns. |
| Shift and Space | Select the whole rows of the selection |
| Control or Command and Space | Select the whole columns of the selection |
| Control or Command and an arrow key | Go to the edge of the data |
| Tab, Enter, and Shift with them | Move right, down, left, or up |
| Tab, Enter, and Shift with them, in a selection of more than one cell (a merged cell is one cell) | Move the active cell inside the selection, wrap at its edges, and keep the selection |
| Shift+Tab in column A, or Tab in the last column | Move the focus out of the grid, to the previous or next control |
| Escape, then Tab or Shift+Tab | Move the focus out of the grid from any cell, to the next or previous control. If Escape ends a pending copy or cut or cancels a cell edit, press Escape again. Any other key except a modifier key alone, a press anywhere on the page, or a focus change cancels the exit |
| Page Up, Page Down | Move by one window of rows |
| Home | Go to column A of the row |
| Control or Command and Home | Go to cell A1 |
| Control or Command and End | Go to the last row and column that hold content |
| Alt and Page Up or Page Down | Show the previous or next sheet |
| Control or Command and A | Select the whole sheet |
| Control or Command and `+`, `-`, or `0`, when no whole row or column is selected | Zoom in, zoom out, or reset the zoom |
| Control or Command and `+` (or Shift and `=`), on whole rows or columns | Insert rows or columns |
| Control or Command and `-`, on whole rows or columns | Delete the selected rows or columns. Zoom out stays available when no whole row or column is selected. |
| A character | Start an edit with that character |
| F2, or a double-click on a cell | Edit the active cell |
| Backspace | Start an edit with empty text |
| Enter, Tab, and Shift with them | Commit the edit and move |
| Escape | Cancel the edit |
| Control or Command and S | Save, as the save control of the toolbar does, when the toolbar has `onSave`, also with the save control hidden. Only a key in the editor counts, and the editor takes the key from the browser. A refused save shows "The workbook was not saved" with the reason |
| Control or Command and Z, Control or Command and Shift and Z, or Control and Y, in the cell editor | Undo or redo the typing in the cell editor, as the browser's own text input binds the keys. The keys never cancel the edit and never change the workbook history |
| Alt and Enter, in the cell editor | Add a line break. Where the browser has no `insertText` command, the line break has no undo step in the cell editor |
| Delete | Clear the values of the selection |
| Control or Command and B, I, or U | Toggle bold, italic, or underline |
| Control or Command and C, X, or V | Copy, cut, or paste the selection |
| Left arrow, Right arrow, Home, End, on a sheet tab | Show the previous, next, first, or last sheet, and move the focus to its tab |
| F2 on a sheet tab | Rename the sheet |
| Shift+F10 or the Menu key on a sheet tab | Open the sheet actions |
| Control or Command and Z, Control or Command and Shift and Z, or Control or Command and Y, on a sheet tab or on **Add a sheet** | Undo or redo |
| Control or Command and Z | Undo, with no cell edit open. With an empty history, the editor shows "There is nothing to undo." once |
| Control or Command and Shift and Z, or Control or Command and Y | Redo, with no cell edit open. With nothing to redo, the editor shows "There is nothing to redo." once |
| Enter, Tab, and Shift with them, in the formula bar | Commit the edit and move, as in the cell editor. With no open edit, move the active cell |
| Alt and Enter, in the formula bar | Add a line break |
| Escape, in the formula bar | Cancel the edit, and return the keyboard to the grid |
| Enter, in the name box | Select the typed reference, and return the keyboard to the grid |
| Escape, in the name box | Drop the typed reference, and return the keyboard to the grid |
| Control or Command and Z, Control or Command and Shift and Z, or Control or Command and Y, in the formula bar or the name box | Undo or redo, with no cell edit open and no typed reference. Otherwise the keys undo or redo the typing in the field |
| F6, Shift+F6 | Move the keyboard to the next or previous region: the toolbar, the formula bar, the grid, and the footer with the sheet tabs and the zoom control. In the formula bar, an open cell edit commits first. The formula bar region starts at the name box |

A press on a cell selects it, and a drag selects the cells that it covers. With Shift, a press or a drag extends the selection from its anchor to the pressed cell, and a press on a row or column heading extends it to whole rows or columns. The active cell stays when the new selection holds it; otherwise the anchor becomes the active cell.

On a touch screen, a drag pans the sheet and a tap selects the cell under it. A two-finger pinch does not zoom the grid; use the zoom control. On a trackpad, a pinch zooms as Control and the mouse wheel do. The grid shows a focus ring while it has the keyboard; without the keyboard, the active cell keeps a muted border. The ring color is the `--doc-focus-ring` token, shared with DOCX Editor; its contrast against the chrome is below 3:1, and a change to it is made for the editor family as a whole.
