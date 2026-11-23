// The DOM rules of the formula bar's content field that do not depend on a framework: where a
// press beside the one-row text area puts the caret, how the field scrolls to the line of its
// caret, the insert of a typed line break, and the history keys of a field with no typing of its
// own. Moved from the React adapter's `FormulaBar.tsx`, so the React and Vue formula bars act
// the same. Adapter support, not a stable API.

import { windowScale } from './host-frame.ts';
import { fieldHistoryKey, type FieldKeys } from './field-keys.ts';
import type { XlsxEditorInstance } from './xlsx-editor.ts';

/** The end of the line that the one-row field `el` shows, as an offset in its text. */
function shownLineEnd(el: HTMLTextAreaElement): number {
  const line = parseFloat(el.ownerDocument.defaultView?.getComputedStyle(el).lineHeight ?? '');
  const shown = line > 0 ? Math.round(el.scrollTop / line) : 0;
  let end = -1;
  for (let i = 0; i <= shown; i += 1) {
    end = el.value.indexOf('\n', end + 1);
    if (end < 0) return el.value.length;
  }
  return end;
}

/** The part of a mouse event that `pressFieldBox` reads. @internal */
export interface FieldBoxPress {
  readonly target: EventTarget | null;
  readonly button: number;
  readonly shiftKey: boolean;
  preventDefault(): void;
}

/**
 * A press on the box of the field `el` outside its text area, such as on the line count or the
 * gap beside the text, puts the caret at the end of the line that the field shows. With Shift,
 * on a field that has the keyboard, it extends the selection from its anchor to there, as a
 * press with Shift in the text does. @internal
 */
export function pressFieldBox(el: HTMLTextAreaElement | null, e: FieldBoxPress): void {
  if (!el || e.target === el || e.button !== 0) return;
  e.preventDefault();
  // The shown line is read first: a focus scrolls the field to its old caret.
  const at = shownLineEnd(el);
  const forward = el.selectionDirection !== 'backward';
  const anchor = forward ? el.selectionStart : el.selectionEnd;
  const extend = e.shiftKey && el.ownerDocument.activeElement === el;
  el.focus();
  if (extend && anchor > at) el.setSelectionRange(at, anchor, 'backward');
  else el.setSelectionRange(extend ? anchor : at, at, 'forward');
  showCaretLine(el, true);
}

/**
 * Insert `text` at the caret of `el` as typed text, which the field's own undo can remove.
 * @internal
 */
export function insertFieldText(el: HTMLTextAreaElement, text: string): boolean {
  const doc = el.ownerDocument as Document & {
    execCommand?: (command: string, ui: boolean, value: string) => boolean;
  };
  try {
    return typeof doc.execCommand === 'function' && doc.execCommand('insertText', false, text);
  } catch {
    return false;
  }
}

/**
 * Scroll the one-row field `el`, while it has the keyboard, to the whole line that holds the
 * caret. The browser scrolls only the caret into view, which can leave the field between two
 * lines, or on a line without the caret after a line break. With `reveal`, the field also
 * scrolls a caret at a line end, set from code as by a press on the line count, to the right
 * edge: the browser does not scroll such a caret into view. @internal
 */
export function showCaretLine(el: HTMLTextAreaElement | null, reveal = false): void {
  if (!el || el.ownerDocument.activeElement !== el) return;
  const style = el.ownerDocument.defaultView?.getComputedStyle(el);
  const line = parseFloat(style?.lineHeight ?? '');
  if (!style || !(line > 0)) return;
  // The caret is the end that moves: the start of a backward selection, else its end.
  const caret = el.selectionDirection === 'backward' ? el.selectionStart : el.selectionEnd;
  const index = el.value.slice(0, caret).split('\n').length - 1;
  el.scrollTop = index * line;
  // The field keeps the horizontal scroll of a longer line: it scrolls back to the start
  // when the text before the caret on this line fits, so no first character is hidden.
  const before = el.value.slice(el.value.lastIndexOf('\n', caret - 1) + 1, caret);
  const left = parseFloat(style.paddingLeft) || 0;
  const right = parseFloat(style.paddingRight) || 0;
  const width = before === '' ? 0 : lineWidth(el, style, before);
  if (!Number.isFinite(width)) return;
  if (width <= el.clientWidth - left - right) {
    el.scrollLeft = 0;
    return;
  }
  // A reveal puts the caret, at a line end, at the right edge, so the text before it shows
  // whatever the scroll was before.
  if (reveal) el.scrollLeft = width + left + right - el.clientWidth;
}

/**
 * The width of one line of `el`. A line with a tab is laid out in a hidden copy of the field's
 * text style, because only layout gives the tab stops of `tab-size`. That costs one layout for
 * each caret move or render on such a line, which is small next to the paint of the field.
 */
function lineWidth(el: HTMLElement, style: CSSStyleDeclaration, text: string): number {
  if (!text.includes('\t')) return textWidth(el, style, text);
  const host = el.parentElement;
  if (!host) return Number.POSITIVE_INFINITY;
  const copy = el.ownerDocument.createElement('div');
  copy.setAttribute('aria-hidden', 'true');
  const look = copy.style;
  look.position = 'absolute';
  look.visibility = 'hidden';
  look.left = '0';
  look.top = '0';
  look.whiteSpace = 'pre';
  // Each font property, not the shorthand, which resets kerning, variants, and features.
  for (const name of [
    'font-family',
    'font-size',
    'font-style',
    'font-weight',
    'font-stretch',
    'font-kerning',
    'font-variant',
    'font-feature-settings',
    'font-variation-settings',
  ])
    look.setProperty(name, style.getPropertyValue(name));
  look.letterSpacing = style.letterSpacing;
  look.wordSpacing = style.wordSpacing;
  look.textTransform = style.textTransform;
  look.textIndent = '0';
  look.setProperty('tab-size', style.getPropertyValue('tab-size') || '8');
  copy.textContent = text;
  host.appendChild(copy);
  const width = copy.getBoundingClientRect().width;
  copy.remove();
  // The box is in screen pixels, after any transform of an ancestor; the scroll of the field
  // is in layout pixels. The field's own scale converts one to the other.
  return width / windowScale(el).x;
}

/** One measuring context for each document, made when a field first needs it. */
const measures = new WeakMap<Document, CanvasRenderingContext2D | null>();

/** The width of `text` in the font of `el`, or infinity when the document cannot measure. */
function textWidth(el: HTMLElement, style: CSSStyleDeclaration, text: string): number {
  const doc = el.ownerDocument;
  if (!measures.has(doc)) measures.set(doc, doc.createElement('canvas').getContext('2d'));
  const context = measures.get(doc);
  if (!context) return Number.POSITIVE_INFINITY;
  context.font = style.font || `${style.fontSize} ${style.fontFamily}`;
  return context.measureText(text).width;
}

/** The part of a key event that `stepFieldHistory` reads. @internal */
export interface FieldHistoryKey extends FieldKeys {
  preventDefault(): void;
}

/**
 * Step the workbook history for a history key of a field with no typing of its own. True when
 * the key stepped the history. @internal
 */
export function stepFieldHistory(
  e: FieldHistoryKey,
  editor: Pick<XlsxEditorInstance, 'undo' | 'redo'>,
  typing: boolean
): boolean {
  const history = fieldHistoryKey(e, typing);
  if (!history) return false;
  e.preventDefault();
  if (history === 'undo') editor.undo();
  else editor.redo();
  return true;
}
