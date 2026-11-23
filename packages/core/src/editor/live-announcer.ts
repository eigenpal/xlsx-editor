// The live region of the grid: what the editor says, the words of the active cell after an edit,
// and the selection when the focus enters the grid. A store change can change the shown text of the
// active cell: a format command, a typed entry, or an automation edit, also on a selection of many
// cells. The editor then says the new words of the active cell once, at the next paint, unless the
// same command already said something: an undo or a redo names its step, a paste or a move names
// the new selection, and a sheet command names its sheet. Words equal to the words before the
// change say nothing, unless the focus entered the grid in the same step: the paint then says the
// selection. DOCX Editor has no live region for the selected text, which stays in the editable
// page. Its `editor/editor-announcer.ts` is a polite region that the editor owns and that says each
// message once, and its page indicator is a status region whose text changes only when its value
// changes, as here.

import { REANNOUNCE } from './cell-editing.ts';
import { hostTimers } from './host-realm.ts';
import type { SelectionWords } from './selection.ts';
import { TaskMark } from './task-mark.ts';

/** Words as a screen reader reads them: runs of spaces as one, with none at the ends. */
const heard = (words: string | undefined): string | undefined => words?.replace(/\s+/g, ' ').trim();

/** The live region of the grid, and the words of the active cell before an unsaid change. */
export class LiveAnnouncer {
  /**
   * The words of the active cell in the workbook before the first store change that nothing
   * said yet. It reads the active cell at the time of the call, so a step that moves the
   * selection compares one cell before and after the change.
   */
  #before: (() => string | undefined) | undefined;
  /** The text of the region at that change. */
  #heard: string | null = null;
  #pending = false;
  /** What the paint says when the unsaid change leaves the words of the active cell. */
  #confirm: string | undefined;
  /** The change can touch more cells than the active cell: say `#confirm` before new words. */
  #many = false;
  /** The focus entered the grid while the change waited for its paint. */
  #entry = false;
  /** The selection at the last message of this editor: `say` or `selected`. */
  #shown: SelectionWords | undefined;
  /** Words written in the task that runs now: the step that moves the focus said them. */
  readonly #written = new TaskMark(hostTimers(() => this.region()));

  /**
   * @param region - The live region, when the editor is attached.
   * @param selection - The words of the selection that shows now.
   */
  constructor(
    private readonly region: () => HTMLElement | undefined,
    private readonly selection: (sheetChanged?: boolean) => SelectionWords | undefined
  ) {}

  /**
   * Put `text` in the live region. The same text again is announced again. The region then
   * holds other words than the selection, such as an undo, a copy, or the new text of the
   * active cell, so a later move of the active cell inside the same selection says the cell.
   */
  say(text: string): void {
    this.#drop();
    const live = this.region();
    if (!live) return;
    live.textContent = live.textContent === text ? `${text}${REANNOUNCE}` : text;
    this.#written.note();
    this.#shown = this.selection();
  }

  /**
   * Say the selection after it changed, after its sheet when `sheetChanged`. Nothing is said
   * while the editor is detached.
   *
   * - The active cell moves inside the same selection, as on Enter inside a range: the region
   *   says the selection with the active cell (`again`).
   * - A store change, such as a paste: the region says `again` on the same selection, and on a
   *   new selection when the change changed the active cell.
   * - A select of the same selection that the region still holds, such as each move of a
   *   heading drag: the region says nothing. For one cell, `again` is `said`, and the same words
   *   keep an unsaid change, so the next paint says the cell when the change changed it.
   *
   * A sheet change always says its sheet.
   */
  selected(sheetChanged = false): void {
    const [live, words] = [this.region(), this.selection(sheetChanged)];
    if (!live || !words) return;
    const text = live.textContent;
    const held = text?.endsWith(REANNOUNCE) ? text.slice(0, -REANNOUNCE.length) : text;
    const last = this.#shown;
    const same = !sheetChanged && last?.range === words.range;
    const moved = same && last.at !== words.at;
    const named = words.again !== words.said;
    if (same && (held === last.said || held === last.again)) {
      if (named && (moved || this.#pending)) {
        this.say(words.again);
      }
      return;
    }
    const changed = this.#pending && (same || this.#before?.() !== words.cell);
    this.#drop();
    this.#shown = words;
    live.textContent = moved || changed ? words.again : words.said;
    this.#written.note();
  }

  /**
   * The focus entered the grid from outside it: say the selection with its active cell, as a grid
   * of the WAI-ARIA Authoring Practices reads the cell that has the focus. The same words again
   * are said again. Nothing is said when the step that moved the focus said its words, such as a
   * sheet tab that names its sheet, a move of Enter in the formula bar that names the next cell,
   * or a refused rename of a sheet that loses the focus. A step is one task (`TaskMark`), so a
   * key that a browser runs before the timer of the mark says nothing more. While a store change
   * waits for its paint, such as a format that a toolbar menu applies, the paint says the change,
   * or the selection when the change keeps the words of the active cell and has no confirmation.
   */
  entered(): void {
    if (this.#written.held) return;
    if (this.#pending) {
      this.#entry = true;
      return;
    }
    const words = this.selection();
    if (words) this.say(words.again);
  }

  /**
   * Empty the live region, so its words are not read again: a refusal that ends, or a switch
   * of the catalog. An unsaid change stays, and the paint then does not say it.
   */
  silence(): void {
    const live = this.region();
    if (live) live.textContent = '';
  }

  /**
   * Empty the live region and forget an unsaid change: another workbook opens, or a detach. A
   * detach keeps the selection (`forget` false), so after the next attach a move inside the
   * same selection still says the active cell.
   */
  clear(forget = true): void {
    this.#drop();
    if (forget) this.#shown = undefined;
    const live = this.region();
    if (live) live.textContent = '';
  }

  /**
   * Call before the editor follows a store change, with the words of the active cell in the
   * workbook before it: keep them for the first change that nothing said yet, so two changes
   * before one paint say the words once.
   */
  storeChanging(before: () => string | undefined): void {
    if (this.#pending) return;
    this.#before = before;
    this.#heard = this.region()?.textContent ?? null;
    this.#pending = true;
  }

  /**
   * Keep `text` for the store change that nothing said yet: the paint says it when the change
   * leaves the words of the active cell as they were, such as a decimal step that changes only
   * other cells. Changed words are said in its place, so the change is said once. With `many`,
   * a selection of more than one cell, changed words follow the text in one message, as DOCX
   * Editor joins the parts of one change, so a step on a range never names only one cell.
   */
  confirm(text: string, many = false): void {
    if (!this.#pending) return;
    this.#confirm = text;
    this.#many = many;
  }

  /** The words of the active cell, when a store change that nothing said yet changed them. */
  changed(words: () => string | undefined): string | undefined {
    if (!this.#pending) return undefined;
    const now = words();
    return now !== this.#before?.() ? now : undefined;
  }

  /**
   * At a paint: say the words of the active cell when a store change changed them. Two guards
   * keep one message for each command. A message of this editor drops the change at once:
   * `say`, `clear`, and `selected` when it writes new words. The empty region of `silence`
   * keeps the change but changes the text of the region, so the change is then not said. The
   * paint drops the change in each case, so the next store change starts a new one.
   */
  settle(words: () => string | undefined): void {
    if (!this.#pending) return;
    const [before, confirm, many, entry] = [
      this.#before?.(),
      this.#confirm,
      this.#many,
      this.#entry,
    ];
    this.#drop();
    if ((this.region()?.textContent ?? null) !== this.#heard) return;
    const now = words();
    // After a decimal step or a number format change, words that differ only in their spaces,
    // such as the padding of `?`, read the same, so the confirmation is said. Any other change
    // that changes the words says them, also when only their spaces change.
    const same = now === before || (confirm !== undefined && heard(now) === heard(before));
    // A format change says its words with the padding spaces of the code collapsed.
    const said = confirm !== undefined ? heard(now) : now;
    if (said !== undefined && !same) this.say(many && confirm ? `${confirm} ${said}` : said);
    else if (confirm !== undefined) this.say(confirm);
    else if (entry) {
      const words = this.selection();
      if (words) this.say(words.again);
    }
  }

  /** Forget the change, and the workbook that it keeps. */
  #drop(): void {
    this.#pending = false;
    this.#before = undefined;
    this.#confirm = undefined;
    this.#many = false;
    this.#entry = false;
  }
}
