// The conditional formatting budget flag of the editor (change `workbook-conditional-formatting`,
// design D4): a layout that ran out of its check budget turns on the notice of the shown sheet.
// The format controls then get no hint, and the copy text and the announcement of the active
// cell use the cell format, since the grid can then show the cell format.

import { NO_SHEET_STATE, type SheetState } from './editor-model.ts';

/**
 * The workbook and the sheet where a layout ran out of its conditional formatting check
 * budget. The sheet is its part name, not its index, so the flag stays with its sheet when a
 * sheet op moves it, and never passes to the sheet that a delete shows at the same index. The
 * flag stays set until the shown sheet changes or another workbook opens (the next snapshot then
 * clears it), so a smaller window does not hide the notice and grow the grid into the budget
 * again.
 */
export class ConditionalCut {
  private at: { readonly store: object; readonly sheet: string } | null = null;

  /** `onMark` runs when a cut is new, so the editor publishes the notice. */
  constructor(private readonly onMark: () => void) {}

  /**
   * Whether the flag holds for this sheet. The copy text and the announcement of the active cell
   * then use the cell format, as the hint does.
   */
  holds(store: object, sheet: string | undefined): boolean {
    return this.at !== null && this.at.store === store && this.at.sheet === sheet;
  }

  /**
   * The sheet state with the cut: the notice on, and no `conditionalFormat`. A flag of another
   * sheet or workbook is cleared first, so a return to a sheet starts without it.
   */
  apply(state: SheetState, store: object | undefined, sheet: string | undefined): SheetState {
    if (this.at?.store !== store || this.at?.sheet !== sheet) this.at = null;
    if (this.at === null || state === NO_SHEET_STATE) return state;
    return { ...state, conditionalHidden: true, conditionalFormat: null };
  }

  /**
   * Record a cut of the shown sheet: a layout that ran out of its check budget. A new cut runs
   * `onMark`. A sheet without a part name has no cells to check, so it gets no flag.
   */
  mark(store: object, sheet: string | undefined): void {
    if (sheet === undefined || this.holds(store, sheet)) return;
    this.at = { store, sheet };
    this.onMark();
  }
}
