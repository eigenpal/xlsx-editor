// Move the references of formula text for a structural edit (design D4).
//
// The bounded tokenizer reads the text. A reference that points into the edited sheet moves by
// the rules of `axis.ts`: a single cell that the edit deletes becomes `#REF!`, a range that
// loses some of its rows or columns shrinks, and a range that loses all of them becomes `#REF!`.
// Two cells that the range operator joins on one sheet (`Data!A1:Data!A4`) move as one range.
// Absolute (`$`) references move too, because they name a position, not an offset. A string is
// text and never moves, and a name, a structured reference, an error, and a reference to another
// workbook keep their text.
//
// A reference that the move cannot place with certainty refuses the edit instead of guessing:
// text the tokenizer cannot read, a sheet range (`Sheet1:Sheet3!A1`) that spans the edited
// sheet, a reference with no sheet in text that has no host sheet (a defined name), an end of a
// range that `:` builds from other operands (`INDEX(A:A,1):A10`) when the edit deletes it, a move
// that makes the text longer than `MAX_FORMULA_LENGTH`.
//
// Only the moved references change: every other character keeps its text, so a formula that
// the edit does not touch is written back as it was.

import { columnLetters } from '../cell-ref.ts';
import {
  MAX_FORMULA_LENGTH,
  tokenizeFormula,
  type FormulaToken,
  type RefEnd,
  type ReferenceToken,
} from '../formula/tokenizer.ts';
import { sheetNameKey } from '../workbook/workbook-index.ts';
import { moveSpan, movePosition, type AxisEdit } from './axis.ts';

/** Where formula text lives, for a structural edit. Sheet names are folded (`sheetNameKey`). */
export interface ReferenceContext {
  readonly edit: AxisEdit;
  /** The edited sheet. */
  readonly sheet: string;
  /** The sheet that holds the text, or undefined for workbook-level text (a defined name). */
  readonly host: string | undefined;
  /** Every sheet by its tab position, for sheet ranges. */
  readonly order: ReadonlyMap<string, number>;
}

/** The moved text, or why the text cannot move with certainty. */
export type MovedText =
  | { readonly ok: true; readonly text: string; readonly changed: boolean }
  | { readonly ok: false; readonly reason: string };

/** Whether a reference points into the edited sheet: yes, no, or not with certainty. */
function pointsIntoSheet(token: ReferenceToken, ctx: ReferenceContext): boolean | 'unsure' {
  const prefix = token.sheet;
  if (!prefix) return ctx.host === undefined ? 'unsure' : ctx.host === ctx.sheet;
  if (prefix.book !== undefined) return prefix.book === '0' ? 'unsure' : false;
  if (prefix.sheets.length === 1) return sheetNameKey(prefix.sheets[0]!) === ctx.sheet;
  const a = ctx.order.get(sheetNameKey(prefix.sheets[0]!));
  const b = ctx.order.get(sheetNameKey(prefix.sheets[1]!));
  const t = ctx.order.get(ctx.sheet);
  if (a === undefined || b === undefined || t === undefined) return 'unsure';
  return t >= Math.min(a, b) && t <= Math.max(a, b) ? 'unsure' : false;
}

function endText(end: RefEnd): string {
  let out = '';
  if (end.col) out += `${end.col.absolute ? '$' : ''}${columnLetters(end.col.index + 1)}`;
  if (end.row) out += `${end.row.absolute ? '$' : ''}${end.row.index + 1}`;
  return out;
}

function withIndex(end: RefEnd, key: 'row' | 'col', index: number): RefEnd {
  const axis = end[key];
  return axis ? { ...end, [key]: { index, absolute: axis.absolute } } : end;
}

/**
 * The text of a moved reference: the same text when the edit does not move it, or `#REF!` with
 * its sheet prefix when the edit deletes all of it.
 */
export function movedReferenceText(token: ReferenceToken, e: AxisEdit): string {
  const key = e.axis === 'rows' ? 'row' : 'col';
  const prefix = token.sheet?.raw ?? '';
  const gone = `${prefix}#REF!`;
  if (token.shape === 'cell') {
    const p = token.from[key]!.index;
    const q = movePosition(e, p);
    if (q === null) return gone;
    return q === p ? token.text : `${prefix}${endText(withIndex(token.from, key, q))}`;
  }
  // Whole columns do not move with rows, and whole rows do not move with columns.
  if ((token.shape === 'columns' && key === 'row') || (token.shape === 'rows' && key === 'col'))
    return token.text;
  const to = token.to!;
  const a = token.from[key]!.index;
  const b = to[key]!.index;
  const lo = Math.min(a, b);
  const hi = Math.max(a, b);
  const span = moveSpan(e, lo, hi);
  if (span === null) return gone;
  if (span[0] === lo && span[1] === hi) return token.text;
  const from = withIndex(token.from, key, a <= b ? span[0] : span[1]);
  const end = withIndex(to, key, a <= b ? span[1] : span[0]);
  return `${prefix}${endText(from)}:${endText(end)}`;
}

/** The index of the next token after `k` in `dir` that is not a space, or -1. */
function neighbor(tokens: readonly FormulaToken[], k: number, dir: 1 | -1): number {
  for (let j = k + dir; j >= 0 && j < tokens.length; j += dir)
    if (tokens[j]!.kind !== 'space') return j;
  return -1;
}

const isRangeOperator = (t: FormulaToken | undefined) => t?.kind === 'operator' && t.text === ':';

/** The sheet that a reference names, folded, or the host sheet; undefined when not certain. */
function sheetOf(token: ReferenceToken, ctx: ReferenceContext): string | undefined {
  const p = token.sheet;
  if (!p) return ctx.host;
  return p.book === undefined && p.sheets.length === 1 ? sheetNameKey(p.sheets[0]!) : undefined;
}

/**
 * The texts of two cells joined by the range operator (`Data!A1:Data!A4`) on one sheet, moved
 * as one span: an edit that deletes one end shrinks the range, and one that deletes all of it
 * makes it `#REF!` (`from` holds that text, and `to` is empty).
 */
function movedPair(a: ReferenceToken, b: ReferenceToken, e: AxisEdit): [string, string] {
  const key = e.axis === 'rows' ? 'row' : 'col';
  const p = a.from[key]!.index;
  const q = b.from[key]!.index;
  const span = moveSpan(e, Math.min(p, q), Math.max(p, q));
  if (span === null) return [`${a.sheet?.raw ?? ''}#REF!`, ''];
  const ends = p <= q ? span : ([span[1], span[0]] as const);
  const text = (t: ReferenceToken, i: number) =>
    i === t.from[key]!.index
      ? t.text
      : `${t.sheet?.raw ?? ''}${endText(withIndex(t.from, key, i))}`;
  return [text(a, ends[0]), text(b, ends[1])];
}

/** Move the references of tokens. The text is the text the tokens were read from. */
export function moveTokens(
  text: string,
  tokens: readonly FormulaToken[],
  ctx: ReferenceContext
): MovedText {
  let out = '';
  let last = 0;
  let changed = false;
  const unsure = (t: FormulaToken): MovedText => ({
    ok: false,
    reason: `the reference ${t.text} cannot be moved with certainty`,
  });
  const put = (start: number, end: number, moved: string) => {
    out += text.slice(last, start) + moved;
    last = end;
    changed = true;
  };
  for (let k = 0; k < tokens.length; k += 1) {
    const token = tokens[k]!;
    if (token.kind !== 'reference') continue;
    const into = pointsIntoSheet(token, ctx);
    if (into === 'unsure') return unsure(token);
    const op = neighbor(tokens, k, 1);
    const m = op < 0 ? -1 : neighbor(tokens, op, 1);
    const other = m < 0 ? undefined : tokens[m];
    // Two cells joined by `:` on one sheet are one range.
    if (isRangeOperator(tokens[op]) && token.shape === 'cell' && other?.kind === 'reference') {
      const into2 = pointsIntoSheet(other, ctx);
      if (into2 === 'unsure') return unsure(other);
      const same = sheetOf(token, ctx) !== undefined && sheetOf(token, ctx) === sheetOf(other, ctx);
      if (other.shape === 'cell' && same) {
        if (into) {
          const [x, y] = movedPair(token, other, ctx.edit);
          const end = other.start + other.text.length;
          if (y === '') put(token.start, end, x);
          else {
            if (x !== token.text) put(token.start, token.start + token.text.length, x);
            if (y !== other.text) put(other.start, end, y);
          }
        }
        k = m;
        continue;
      }
    }
    if (!into) continue;
    const moved = movedReferenceText(token, ctx.edit);
    if (moved === token.text) continue;
    // One end of a range that the operator builds from other operands (`INDEX(A:A,1):A10`):
    // when the edit deletes it, the range cannot shrink with certainty.
    const joined = isRangeOperator(tokens[op]) || isRangeOperator(tokens[neighbor(tokens, k, -1)]);
    if (joined && moved.endsWith('#REF!')) return unsure(token);
    put(token.start, token.start + token.text.length, moved);
  }
  if (!changed) return { ok: true, text, changed };
  const moved = out + text.slice(last);
  // A longer reference (`#REF!`, more digits) can push the text past the length the tokenizer
  // reads: the engine could not read its own output, so the edit is refused.
  if (moved.length > MAX_FORMULA_LENGTH)
    return {
      ok: false,
      reason: `the moved formula would be longer than ${MAX_FORMULA_LENGTH} characters`,
    };
  return { ok: true, text: moved, changed };
}

/** Read formula text and move its references. */
export function moveReferences(text: string, ctx: ReferenceContext): MovedText {
  const read = tokenizeFormula(text);
  if (!read.ok)
    return { ok: false, reason: `the formula text cannot be read (${read.problem} at ${read.at})` };
  return moveTokens(text, read.tokens, ctx);
}
