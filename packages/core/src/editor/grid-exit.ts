// The key that leaves the grid from any cell (WCAG 2.1.2, No Keyboard Trap). In the grid, Tab
// and Shift+Tab move the active cell, so they leave the grid only at its first or last column
// (`grid-keys.ts`). Escape with nothing to cancel arms the exit: the editor says how to leave,
// and the next Tab or Shift+Tab goes to the browser, which moves the focus to the next or the
// previous control of the page. Any other key, a press anywhere on the page (the grid, its
// scrollbar, the toolbar, or a sheet tab, which keep the focus on the grid), or a focus that
// leaves the grid disarms it. A modifier key alone, such as the Shift of Shift+Tab, keeps it
// armed. While the exit is armed, a cue over the grid shows the hint (`showExitCue` in
// `grid-surface.ts`), and the live region says it once. The grid's accessible description
// (`grid.exitDescription`) names the keys before the user needs them, with F6 for a screen
// reader that takes Escape: each adapter's region keys move the focus to the next region of the
// editor (React `landmarks.ts`; a Vue twin keeps F6 for parity).
//
// Escape ends a pending copy or cut first; only an Escape that the key table leaves to the page arms
// the exit. An Escape in the cell editor cancels the edit, so the keys from an open edit are
// Escape, Escape, and Tab.
//
// Sister products: DOCX Editor has no exit key, because Tab types or indents in a document.
// PPTX Editor leaves text editing with Escape so that Tab moves again, and Tab past the last
// shape leaves the slide. A worksheet keeps Tab for the next cell, so Escape arms the exit here.
// The cue takes the form of DOCX Editor's header and footer edit hint (`.docx-hf-edit-hint`): a
// small pill over the region that never takes the pointer, with its text from the catalog.

/** The keys that are a modifier alone: they keep the exit armed. */
const MODIFIERS = new Set(['Shift', 'Control', 'Alt', 'AltGraph', 'Meta', 'CapsLock']);

/** The exit of one grid: armed by Escape, used by the next Tab. */
export class GridExit {
  #armed = false;

  /**
   * @param onArmedChange - Called with true when Escape arms the exit, to show and say how to
   * leave the grid, and with false when the exit is used or canceled, to hide the cue.
   */
  constructor(private readonly onArmedChange: (armed: boolean) => void) {}

  #set(armed: boolean): void {
    if (this.#armed === armed) return;
    this.#armed = armed;
    this.onArmedChange(armed);
  }

  /**
   * Read a keydown before `handleGridKey`. Returns true when the key leaves the grid: the key
   * table must then not act, so the browser moves the focus.
   */
  beforeKeyTable(e: KeyboardEvent, grid: EventTarget): boolean {
    if (e.target !== grid || MODIFIERS.has(e.key)) return false;
    // A held Escape keeps the exit as it is, so the hint comes once, when the exit arms.
    if (e.key === 'Escape' && e.repeat) return false;
    const leave = this.#armed && e.key === 'Tab' && !e.ctrlKey && !e.metaKey && !e.altKey;
    this.#set(false);
    return leave;
  }

  /**
   * Read a keydown after `handleGridKey`: an Escape that it left to the page arms the exit and
   * shows the hint. A held Escape shows it once.
   */
  afterKeyTable(e: KeyboardEvent, grid: EventTarget): void {
    if (e.target !== grid || e.key !== 'Escape' || e.defaultPrevented) return;
    if (e.ctrlKey || e.metaKey || e.altKey || e.shiftKey) return;
    if (e.isComposing || e.keyCode === 229) return;
    this.#set(true);
  }

  /** Disarm: a press anywhere on the page, or a focus that leaves the grid. */
  reset(): void {
    this.#set(false);
  }
}
