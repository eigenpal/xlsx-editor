// The cell editor: a text input over the active cell, the editing projection of task 7.2.
//
// It is only a projection of one cell of one revision. It holds the text that the user types,
// and on commit it hands the text to the editor, which turns it into one store transaction or a
// typed refusal. A refused commit keeps the editor open with the text, so an edit is never
// dropped without a word (AGENTS.md, "Editing projections").
//
// Keys: Enter commits and moves down (Shift+Enter up), Tab commits and moves right (Shift+Tab
// left), Escape cancels, Alt+Enter adds a line break. An edit that started by typing commits and
// moves on an arrow key; an edit that started with F2 moves the caret, and F2 switches between
// the two. While an input method composes, no key commits or cancels. The history keys
// (`historyShortcutOf`) belong to the input's own undo: they never reach the store history, and
// never cancel or close the edit. A first character that the editor sets has no undo step.
//
// The editor covers the whole edited cell, also when the cell is partly out of the window. When
// the cell scrolls out of the window, the editor stays focused with its text but is not seen
// (`data-offscreen`). The next key in it scrolls the cell back into view first; a key that
// would commit does only that, so an edit never commits to a cell that the user cannot see.
// A window too small to show any part of the cell lets the key act.
//
// The input element comes from the document of the stage it is mounted in (`host-realm.ts`).

import { historyShortcutOf } from './history-shortcut.ts';

/** How an edit started: by typing (arrows commit) or with F2 (arrows move the caret). */
export type CellEditMode = 'enter' | 'edit';

/**
 * Where `commitEdit` moves the active cell after the commit: as Enter (`down`), Shift+Enter
 * (`up`), Tab (`right`), and Shift+Tab (`left`).
 *
 * @public
 */
export type EditMove = 'down' | 'up' | 'right' | 'left';

/**
 * The options of `setEditText`.
 *
 * @public
 */
export interface EditTextOptions {
  /**
   * Whether an input method composes in the field that calls `setEditText`. While it is true,
   * each command that would commit the open edit is refused with `composing`.
   */
  readonly composing?: boolean;
}

/** The row and column step of each `EditMove`, with no prototype to read through. */
const EDIT_MOVES: Readonly<Record<EditMove, readonly [number, number]>> = Object.assign(
  Object.create(null) as Record<EditMove, readonly [number, number]>,
  { down: [1, 0], up: [-1, 0], right: [0, 1], left: [0, -1] } as const
);

/** The row and column step of `move`. Throws a `TypeError` for a value that is not an `EditMove`. */
export function editMoveStep(move: EditMove): readonly [number, number] {
  const step = typeof move === 'string' ? EDIT_MOVES[move] : undefined;
  if (!step) throw new TypeError(`Not an edit move: ${String(move)}`);
  return step;
}

/**
 * The most UTF-16 code units of a cell's text that a content field holds. Each layout of a page
 * reads the whole text of its fields, so a longer text shows shortened and read-only there.
 *
 * @public
 */
export const MAX_CONTENT_FIELD_CHARS = 65_536;

/**
 * The start of `text` that a content field holds, with no half of a surrogate pair at its end.
 * A result that is shorter than `text` is only for reading: the edit keeps the complete text.
 *
 * @public
 */
export function contentFieldText(text: string): string {
  if (text.length <= MAX_CONTENT_FIELD_CHARS) return text;
  const last = text.charCodeAt(MAX_CONTENT_FIELD_CHARS - 1);
  const halfPair = last >= 0xd800 && last <= 0xdbff;
  return text.slice(0, halfPair ? MAX_CONTENT_FIELD_CHARS - 1 : MAX_CONTENT_FIELD_CHARS);
}

/** What the cell editor asks of the editor. */
export interface CellEditorHooks {
  /**
   * Commit the text, then move the active cell by `move`. `inRange` (Enter and Tab) moves the
   * active cell inside a multi-cell selection and keeps it, as without an edit. Returns false
   * when refused.
   */
  commit(text: string, move: readonly [number, number] | null, inRange?: boolean): boolean;
  cancel(): void;
  /** The text of the editor changed (`typed`: by an input event), or its composition did. */
  change(typed: boolean): void;
  /** Scroll the edited cell into view: the editor is out of the window. */
  reveal(): void;
}

/** A box in the stage, in CSS pixels. */
export interface EditorBox {
  readonly left: number;
  readonly top: number;
  readonly width: number;
  readonly height: number;
  /** The font size of the text, in CSS pixels. */
  readonly fontSize: number;
  /** The CSS font family list of the text. */
  readonly fontFamily: string;
}

const ARROWS: Readonly<Record<string, readonly [number, number]>> = {
  ArrowUp: [-1, 0],
  ArrowDown: [1, 0],
  ArrowLeft: [0, -1],
  ArrowRight: [0, 1],
};

export class CellEditor {
  readonly input: HTMLTextAreaElement;
  #mode: CellEditMode = 'enter';
  #composing = false;
  #open = false;
  #offscreen = false;
  readonly #hooks: CellEditorHooks;
  readonly #off: (() => void)[] = [];

  constructor(stage: HTMLElement, label: string, hooks: CellEditorHooks) {
    this.#hooks = hooks;
    const input = stage.ownerDocument.createElement('textarea');
    input.className = 'xlsx-cell-editor';
    input.name = 'xlsx-cell-editor';
    input.rows = 1;
    input.hidden = true;
    input.spellcheck = false;
    // The text sets its own direction from its first strong character (`dir="auto"`): a text or
    // a formula whose first letter is of a right-to-left script edits right to left, and one
    // whose first letter is of another script edits left to right. The grid stays left to right.
    input.dir = 'auto';
    input.setAttribute('autocomplete', 'off');
    input.setAttribute('autocapitalize', 'off');
    input.setAttribute('aria-label', label);
    stage.append(input);
    this.input = input;
    const on = <K extends keyof HTMLElementEventMap>(
      type: K,
      fn: (e: HTMLElementEventMap[K]) => void
    ) => {
      input.addEventListener(type, fn);
      this.#off.push(() => input.removeEventListener(type, fn));
    };
    on('keydown', (e) => this.#onKey(e));
    on('input', () => this.#hooks.change(true));
    on('compositionstart', () => {
      this.#composing = true;
      this.#hooks.change(false);
    });
    on('compositionend', () => {
      this.#composing = false;
      this.#hooks.change(false);
    });
    // A press inside the editor places the caret; the grid must not take it as a cell press.
    on('pointerdown', (e) => e.stopPropagation());
  }

  get isOpen(): boolean {
    return this.#open;
  }

  get mode(): CellEditMode {
    return this.#mode;
  }

  get value(): string {
    return this.input.value;
  }

  /** Whether the edited cell is out of the window, so the editor is not seen. */
  get offscreen(): boolean {
    return this.#offscreen;
  }

  /** An input method composes in the editor: the composed text belongs to it. */
  get composing(): boolean {
    return this.#composing;
  }

  /**
   * Show the editor with `text` and the caret at the end. With `focus`, it takes the keyboard;
   * without, the keyboard stays where it is, as when the formula bar opens the edit.
   */
  open(text: string, mode: CellEditMode, focus = true): void {
    this.#open = true;
    this.#mode = mode;
    this.#composing = false;
    this.input.value = text;
    this.input.hidden = false;
    if (focus) this.input.focus({ preventScroll: true });
    this.input.setSelectionRange(text.length, text.length);
  }

  /** Replace the text of the open editor, as the formula bar types it. */
  setText(text: string): void {
    if (this.#open && this.input.value !== text) this.input.value = text;
  }

  /** Put the editor over the whole cell, or keep it unseen when the cell is out of the window. */
  place(box: EditorBox, visible: boolean): void {
    const s = this.input.style;
    s.left = `${box.left}px`;
    s.top = `${box.top}px`;
    s.minWidth = `${Math.max(0, box.width)}px`;
    s.minHeight = `${Math.max(0, box.height)}px`;
    s.fontSize = `${box.fontSize}px`;
    s.fontFamily = box.fontFamily;
    this.#offscreen = !visible;
    this.input.toggleAttribute('data-offscreen', !visible);
  }

  close(): void {
    this.#open = false;
    this.#composing = false;
    this.#offscreen = false;
    this.input.removeAttribute('data-offscreen');
    this.input.hidden = true;
    this.input.value = '';
  }

  dispose(): void {
    for (const off of this.#off.splice(0)) off();
    this.input.remove();
  }

  #onKey(e: KeyboardEvent): void {
    // F6 and Shift+F6 commit the edit and pass to the adapter, which moves the keyboard to the
    // next region. A refused commit keeps the edit open and the key.
    const f6 = e.key === 'F6' && !e.altKey && !e.ctrlKey && !e.metaKey;
    if (f6 && this.#open && !e.isComposing && !this.#composing)
      if (this.#hooks.commit(this.input.value, null)) return;
    // The grid's own keys never see a key typed in the editor.
    e.stopPropagation();
    if (!this.#open) return;
    // An unseen editor first brings its cell back into view; a key that commits does no more.
    // A window with no room for the cell cannot show it, and does not hold the key back.
    if (this.#offscreen) {
      this.#hooks.reveal();
      if (!this.#offscreen && !e.isComposing && !this.#composing && this.#commits(e)) {
        e.preventDefault();
        return;
      }
    }
    if (e.isComposing || this.#composing || e.keyCode === 229) return;
    // The history keys belong to the input's own undo while the editor is open: they never
    // step the store history, and never cancel or close the edit.
    if (historyShortcutOf(e)) return;
    const back = e.shiftKey;
    switch (e.key) {
      case 'Enter':
        e.preventDefault();
        if (e.altKey) {
          this.#insertBreak();
          return;
        }
        this.#hooks.commit(this.input.value, [back ? -1 : 1, 0], true);
        return;
      case 'Tab':
        e.preventDefault();
        this.#hooks.commit(this.input.value, [0, back ? -1 : 1], true);
        return;
      case 'Escape':
        e.preventDefault();
        this.#hooks.cancel();
        return;
      case 'F2':
        e.preventDefault();
        this.#mode = this.#mode === 'enter' ? 'edit' : 'enter';
        return;
    }
    const arrow = this.#arrowCommit(e);
    if (arrow) {
      e.preventDefault();
      this.#hooks.commit(this.input.value, arrow);
    }
  }

  /** The move of an arrow key that commits (an edit that started by typing), or null. */
  #arrowCommit(e: KeyboardEvent): readonly [number, number] | null {
    const arrow = Object.hasOwn(ARROWS, e.key) ? ARROWS[e.key]! : null;
    return arrow && this.#mode === 'enter' && !e.shiftKey && !e.ctrlKey && !e.metaKey
      ? arrow
      : null;
  }

  /**
   * Insert a line break at the caret. `insertText` is the path that the input's own undo
   * records; where the document has no such command or refuses it, the text is set directly,
   * and that break has no undo step in the input.
   */
  #insertBreak(): void {
    const doc = this.input.ownerDocument as Document & {
      execCommand?: (command: string, ui: boolean, value: string) => boolean;
    };
    let done = false;
    try {
      done = typeof doc.execCommand === 'function' && doc.execCommand('insertText', false, '\n');
    } catch {
      done = false;
    }
    if (done) return;
    this.input.setRangeText('\n', this.input.selectionStart, this.input.selectionEnd, 'end');
    this.#hooks.change(true);
  }

  /** Whether a key commits the edit: Enter (not Alt+Enter), Tab, or an arrow that commits. */
  #commits(e: KeyboardEvent): boolean {
    if (e.key === 'Enter') return !e.altKey;
    return e.key === 'Tab' || this.#arrowCommit(e) !== null;
  }
}
