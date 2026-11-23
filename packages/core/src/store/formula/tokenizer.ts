// The bounded A1 reference tokenizer of design D4.
//
// It reads the text of a formula without a leading `=`: the text of `x:f` (`CT_CellFormula`,
// sml.xsd), an `ST_Formula`. It splits the text into tokens: cell and range references
// (relative and `$` absolute, sheet-qualified with quoted and unquoted sheet names, whole
// columns and whole rows), string literals, numbers, booleans, error
// literals, function names, defined names, structured references, operators, and separators.
// A string literal is text, never a reference. Text that it cannot read ends the read with a
// typed reason and the offset where it stopped: a caller refuses the edit, and never guesses.
// It also refuses a code unit that XML 1.0 cannot hold (`xmlInvalidAt`), because `x:f` could
// not be saved, and R1C1 references (`R1C1`, `RC[1]`), because `x:f` holds A1 text only.
//
// Security (AGENTS.md, "Security"): the input is attacker-controlled. The read is one forward
// pass with no regular expression and no backtracking. The input length, the token count, and
// the nesting depth are capped, so the cost is linear in a capped length.

import { MAX_COLS, MAX_ROWS } from '../cell-ref.ts';
import { xmlInvalidAt } from '../workbook/xstring.ts';

/**
 * The longest formula text the tokenizer reads, in UTF-16 code units. This is the engine's own
 * cap: `ST_Formula` in sml.xsd has no `maxLength`.
 */
export const MAX_FORMULA_LENGTH = 8192;
/** The most tokens one formula can hold. */
export const MAX_FORMULA_TOKENS = 4096;
/** The deepest nesting of parentheses, braces, and structured-reference brackets. */
export const MAX_FORMULA_DEPTH = 64;

/** One axis of a reference: a 0-based index, and whether it has a `$` marker. */
export interface RefAxis {
  readonly index: number;
  readonly absolute: boolean;
}

/** One end of a reference. A whole-column end has no row; a whole-row end has no column. */
export interface RefEnd {
  readonly row?: RefAxis;
  readonly col?: RefAxis;
}

/** The sheet part of a reference or a name, as written: `Sheet1!`, `'My sheet'!`, `[1]S!`. */
export interface SheetPrefix {
  /** The text as written, with the `!`. Moving a reference keeps it verbatim. */
  readonly raw: string;
  /** The sheet names, unquoted: one, or two for a sheet range (`Sheet1:Sheet3!`). */
  readonly sheets: readonly string[];
  /** The external workbook of `[n]` or `[Book.xlsx]`, as written, when the prefix has one. */
  readonly book?: string;
}

interface TokenBase {
  /** The text of the token as written. */
  readonly text: string;
  /** The offset of the token in the formula text. */
  readonly start: number;
}

/** The kinds of token that carry no more than their text. */
export type PlainTokenKind =
  | 'number'
  | 'string'
  | 'boolean'
  | 'function'
  | 'structured'
  | 'operator'
  | 'open'
  | 'close'
  | 'arrayOpen'
  | 'arrayClose'
  | 'separator'
  | 'space';

/** A reference token: one cell, a range, whole columns, or whole rows. */
export interface ReferenceToken extends TokenBase {
  readonly kind: 'reference';
  readonly shape: 'cell' | 'range' | 'columns' | 'rows';
  readonly sheet?: SheetPrefix;
  readonly from: RefEnd;
  /** The second end of a range, whole columns, or whole rows. */
  readonly to?: RefEnd;
}

/** One token of a formula. */
export type FormulaToken =
  | ReferenceToken
  | (TokenBase & { readonly kind: 'name' | 'error'; readonly sheet?: SheetPrefix })
  | (TokenBase & { readonly kind: PlainTokenKind });

/** Why the tokenizer cannot read a formula. */
export type FormulaProblem =
  | 'empty'
  | 'too-long'
  | 'too-many-tokens'
  | 'too-deep'
  | 'unterminated-string'
  | 'unterminated-sheet-name'
  | 'unterminated-bracket'
  | 'unbalanced'
  | 'unexpected-character'
  | 'bad-number'
  | 'bad-name'
  | 'bad-error'
  /** A code unit that XML 1.0 cannot hold: a control character or a lone surrogate. */
  | 'invalid-character'
  /** An R1C1 reference, such as `R1C1` or `RC[-1]`. */
  | 'r1c1-reference'
  /** An operator or a group with no operand (`A1+`, `1+*2`, `()`): `checkFormulaGrammar`. */
  | 'missing-operand'
  /** Two operands with no operator between them (`1 2`, `A1"x"`): `checkFormulaGrammar`. */
  | 'missing-operator'
  /** A `,` or `;` outside a function call, a group, or an array: `checkFormulaGrammar`. */
  | 'misplaced-separator'
  /**
   * The range operator `:` with an operand that is not a reference (`1.5:2`, `"a":"b"`), or a
   * row range outside the grid (`0:3`, `1:1048577`).
   */
  | 'bad-range';

export type TokenizeResult =
  | { readonly ok: true; readonly tokens: readonly FormulaToken[] }
  | { readonly ok: false; readonly problem: FormulaProblem; readonly at: number };

/**
 * The error literals of ECMA-376 Part 1, 18.17.3, `#GETTING_DATA`, and the error values that
 * workbooks store for dynamic arrays, data types, and external data; longest first.
 */
const ERRORS = [
  '#GETTING_DATA',
  '#BLOCKED!',
  '#CONNECT!',
  '#UNKNOWN!',
  '#DIV/0!',
  '#VALUE!',
  '#SPILL!',
  '#FIELD!',
  '#NULL!',
  '#NAME?',
  '#CALC!',
  '#BUSY!',
  '#REF!',
  '#NUM!',
  '#N/A',
];

const isLetter = (c: number): boolean => (c >= 65 && c <= 90) || (c >= 97 && c <= 122);
const isDigit = (c: number): boolean => c >= 48 && c <= 57;
/** A character of a name, an unquoted sheet name, or a reference: `$` included. */
const isWordChar = (c: number): boolean =>
  isLetter(c) ||
  isDigit(c) ||
  c === 95 /* _ */ ||
  c === 46 /* . */ ||
  c === 92 /* \ */ ||
  c === 63 /* ? */ ||
  c === 36 /* $ */ ||
  c > 127;
const isNameStart = (c: number): boolean => isLetter(c) || c === 95 || c === 92 || c > 127;
/** Space, tab, line feed, and carriage return. */
const isSpace = (c: number): boolean => c === 32 || c === 9 || c === 10 || c === 13;

/** The end of a word that starts at `i`. */
function wordEnd(s: string, i: number): number {
  let j = i;
  while (j < s.length && isWordChar(s.charCodeAt(j))) j += 1;
  return j;
}

/** The digits from `i` to the end of `w` as a 1-based row number in the grid, or undefined. */
function rowNumber(w: string, i: number): number | undefined {
  const n = w.length - i;
  if (n < 1 || n > 7 || w.charCodeAt(i) === 48) return undefined;
  let row = 0;
  for (let k = i; k < w.length; k += 1) {
    const c = w.charCodeAt(k);
    if (!isDigit(c)) return undefined;
    row = row * 10 + (c - 48);
  }
  return row <= MAX_ROWS ? row : undefined;
}

/** Column letters from `i` as a 1-based column number in the grid, and where they end. */
function columnNumber(w: string, i: number): { col: number; end: number } | undefined {
  let col = 0;
  let k = i;
  while (k < w.length && isLetter(w.charCodeAt(k))) {
    if (k - i >= 3) return undefined;
    col = col * 26 + ((w.charCodeAt(k) | 32) - 96);
    k += 1;
  }
  return k > i && col <= MAX_COLS ? { col, end: k } : undefined;
}

/** A cell end (`$A$1`, `b7`), or undefined. */
function cellEnd(w: string): RefEnd | undefined {
  const colAbs = w.charCodeAt(0) === 36;
  const letters = columnNumber(w, colAbs ? 1 : 0);
  if (!letters) return undefined;
  const rowAbs = w.charCodeAt(letters.end) === 36;
  const row = rowNumber(w, letters.end + (rowAbs ? 1 : 0));
  if (row === undefined) return undefined;
  return {
    col: { index: letters.col - 1, absolute: colAbs },
    row: { index: row - 1, absolute: rowAbs },
  };
}

/** A whole-column end (`A`, `$XFD`), or undefined. */
function columnEnd(w: string): RefEnd | undefined {
  const abs = w.charCodeAt(0) === 36;
  const letters = columnNumber(w, abs ? 1 : 0);
  if (!letters || letters.end !== w.length) return undefined;
  return { col: { index: letters.col - 1, absolute: abs } };
}

/** A whole-row end (`5`, `$5`), or undefined. */
function rowEnd(w: string): RefEnd | undefined {
  const abs = w.charCodeAt(0) === 36;
  const row = rowNumber(w, abs ? 1 : 0);
  return row === undefined ? undefined : { row: { index: row - 1, absolute: abs } };
}

/** The end of an R1C1 row or column part at `j`: digits, or `[n]` with a sign; -1 if broken. */
function r1c1Part(s: string, j: number): number {
  if (s.charCodeAt(j) !== 91 /* [ */) {
    while (isDigit(s.charCodeAt(j))) j += 1;
    return j;
  }
  let k = j + 1;
  if (s.charCodeAt(k) === 43 || s.charCodeAt(k) === 45) k += 1;
  const digits = k;
  while (k - digits < 8 && isDigit(s.charCodeAt(k))) k += 1;
  return k > digits && s.charCodeAt(k) === 93 /* ] */ ? k + 1 : -1;
}

/**
 * The end of an R1C1 reference at `i` (`R`, `C`, `R2C3`, `RC[-1]`, `R[1]`) that covers at
 * least the word ending at `wordEndAt`, or -1.
 */
function r1c1End(s: string, i: number, wordEndAt: number): number {
  let j = i;
  let marks = 0;
  if ((s.charCodeAt(j) | 32) === 114 /* r */) {
    j = r1c1Part(s, j + 1);
    if (j < 0) return -1;
    marks += 1;
  }
  if ((s.charCodeAt(j) | 32) === 99 /* c */) {
    j = r1c1Part(s, j + 1);
    if (j < 0) return -1;
    marks += 1;
  }
  if (marks === 0 || j < wordEndAt || (j < s.length && isWordChar(s.charCodeAt(j)))) return -1;
  return j;
}

/**
 * Whether a sheet name must be written in quotes: it has the shape of an A1 cell in the grid
 * (`A1`), of an R1C1 reference (`R1C1`, `R`), or of a boolean (`TRUE`), so a reader could take
 * the unquoted name for that operand.
 */
export function sheetNameNeedsQuotes(name: string): boolean {
  const upper = name.toUpperCase();
  return (
    upper === 'TRUE' ||
    upper === 'FALSE' ||
    cellEnd(name) !== undefined ||
    r1c1End(name, 0, name.length) === name.length
  );
}

/**
 * Whether `w` is a defined name or a function name: a legal first character, no `$`, and not
 * a cell reference with a decimal part (`A1.5` is a broken number, not a name).
 */
function isName(w: string): boolean {
  if (w.length === 0 || !isNameStart(w.charCodeAt(0))) return false;
  const dot = w.indexOf('.');
  if (dot > 0 && dot < w.length - 1 && cellEnd(w.slice(0, dot))) {
    let digits = true;
    for (let k = dot + 1; k < w.length; k += 1) digits &&= isDigit(w.charCodeAt(k));
    if (digits) return false;
  }
  for (let k = 0; k < w.length; k += 1) if (w.charCodeAt(k) === 36) return false;
  return true;
}

class Reader {
  readonly tokens: FormulaToken[] = [];
  /** The open brackets: `(` and `{` as character codes. */
  readonly stack: number[] = [];
  i = 0;
  constructor(readonly s: string) {}

  /** Add a token. A problem when the token cap is passed, else null. */
  push(token: FormulaToken): Step {
    this.tokens.push(token);
    return this.tokens.length <= MAX_FORMULA_TOKENS ? null : 'too-many-tokens';
  }
}

type Step = FormulaProblem | null;

/**
 * Read an operand that starts with a word: a reference, a range, a name, a function name, a
 * boolean, or a structured reference. `sheet` is a prefix already read from `start`.
 */
function readWord(r: Reader, sheet: SheetPrefix | undefined, start: number): Step {
  const s = r.s;
  const begin = r.i;
  const end = wordEnd(s, begin);
  const w = s.slice(begin, end);
  const next = s.charCodeAt(end);
  if (!sheet && next === 33 /* ! */) {
    r.i = end + 1;
    return readQualified(r, { raw: s.slice(start, end + 1), sheets: [w] }, start);
  }
  if (!sheet && next === 58 /* : a sheet range `Sheet1:Sheet3!` */) {
    const end2 = wordEnd(s, end + 1);
    // A cell of the grid before `:` starts a range, not a sheet range (`A1:Data!B2`): the
    // grammar check refuses a range from a cell to a cell of another sheet at the colon. A word
    // past column XFD or row 1048576 (`ZZZ1`, `XFE1`) is no cell, so it names a sheet.
    if (end2 > end + 1 && s.charCodeAt(end2) === 33 && !cellEnd(w)) {
      r.i = end2 + 1;
      const sheets = [w, s.slice(end + 1, end2)];
      return readQualified(r, { raw: s.slice(start, end2 + 1), sheets }, start);
    }
  }
  // R1C1 text, unless it is an A1 cell (`R1`, `C2`) or a whole-column range (`R:R`).
  const r1 = r1c1End(s, begin, end);
  if (
    r1 > 0 &&
    s.charCodeAt(r1) !== 40 &&
    !cellEnd(w) &&
    !(next === 58 && (columnEnd(w) || rowEnd(w)))
  )
    return 'r1c1-reference';
  if (!sheet && next === 40 /* ( */) {
    if (!isName(w)) return 'bad-name';
    r.i = end;
    return r.push({ kind: 'function', text: w, start: begin });
  }
  if (!sheet && next === 91 /* [ */) {
    if (!isName(w)) return 'bad-name';
    r.i = end;
    return readStructured(r, begin);
  }
  const cell = cellEnd(w);
  const column = cell ? undefined : columnEnd(w);
  const row = cell || column ? undefined : rowEnd(w);
  const first = cell ?? column ?? row;
  if (first && next === 58 /* : a range `A1:B2`, `A:C`, or `1:3` */) {
    const end2 = wordEnd(s, end + 1);
    const w2 = s.slice(end + 1, end2);
    const after = s.charCodeAt(end2);
    const second = cell ? cellEnd(w2) : column ? columnEnd(w2) : rowEnd(w2);
    if (second && after !== 40 && after !== 33 && after !== 91) {
      r.i = end2;
      return r.push({
        kind: 'reference',
        shape: cell ? 'range' : column ? 'columns' : 'rows',
        text: s.slice(start, end2),
        start,
        from: first,
        to: second,
        ...(sheet ? { sheet } : {}),
      });
    }
  }
  r.i = end;
  if (cell)
    return r.push({
      kind: 'reference',
      shape: 'cell',
      text: s.slice(start, end),
      start,
      from: cell,
      ...(sheet ? { sheet } : {}),
    });
  const upper = w.toUpperCase();
  if (!sheet && (upper === 'TRUE' || upper === 'FALSE'))
    return r.push({ kind: 'boolean', text: w, start });
  if (!isName(w)) return 'bad-name';
  return r.push({ kind: 'name', text: s.slice(start, end), start, ...(sheet ? { sheet } : {}) });
}

/** After a sheet prefix: a reference, a range, a name, or `#REF!`. */
function readQualified(r: Reader, sheet: SheetPrefix, start: number): Step {
  const c = r.s.charCodeAt(r.i);
  if (c === 35 /* # */) return readError(r, start, sheet);
  if (r.i < r.s.length && isWordChar(c)) return readWord(r, sheet, start);
  return 'unexpected-character';
}

function readError(r: Reader, start: number, sheet?: SheetPrefix): Step {
  const s = r.s;
  for (const e of ERRORS) {
    if (s.length - r.i >= e.length && s.slice(r.i, r.i + e.length).toUpperCase() === e) {
      r.i += e.length;
      return r.push({
        kind: 'error',
        text: s.slice(start, r.i),
        start,
        ...(sheet ? { sheet } : {}),
      });
    }
  }
  return 'bad-error';
}

/** A structured reference from `[`: `Table1[Col]`, `[@Col]`, `T[[#Headers],[A]]`. */
function readStructured(r: Reader, start: number): Step {
  const s = r.s;
  let depth = 0;
  for (let j = r.i; j < s.length; j += 1) {
    const c = s.charCodeAt(j);
    if (c === 39 /* ' escapes the next character */) {
      j += 1;
      continue;
    }
    if (c === 91) {
      depth += 1;
      if (depth + r.stack.length > MAX_FORMULA_DEPTH) return 'too-deep';
    } else if (c === 93) {
      depth -= 1;
      if (depth === 0) {
        r.i = j + 1;
        return r.push({ kind: 'structured', text: s.slice(start, r.i), start });
      }
    }
  }
  return 'unterminated-bracket';
}

/** A quoted sheet prefix `'It''s'!`, with an optional `[n]` workbook index inside. */
function readQuotedSheet(r: Reader, start: number): Step {
  const s = r.s;
  let name = '';
  let j = r.i + 1;
  for (;;) {
    if (j >= s.length) return 'unterminated-sheet-name';
    if (s.charCodeAt(j) === 39) {
      if (s.charCodeAt(j + 1) !== 39) break;
      name += "'";
      j += 2;
      continue;
    }
    name += s[j];
    j += 1;
  }
  if (s.charCodeAt(j + 1) !== 33 || name.length === 0) return 'unterminated-sheet-name';
  r.i = j + 2;
  let book: string | undefined;
  if (name.charCodeAt(0) === 91) {
    const close = name.indexOf(']');
    if (close < 2) return 'unterminated-bracket';
    book = name.slice(1, close);
    name = name.slice(close + 1);
  }
  const colon = name.indexOf(':');
  const sheets = colon < 0 ? [name] : [name.slice(0, colon), name.slice(colon + 1)];
  return readQualified(
    r,
    { raw: s.slice(start, r.i), sheets, ...(book !== undefined ? { book } : {}) },
    start
  );
}

/** Characters that end an unquoted workbook name in `[Book.xlsx]Sheet!`. */
const BOOK_STOP = new Set([33, 34, 39, 40, 41, 44, 59, 91, 123, 125]);

/** The longest unquoted workbook name that the tokenizer reads, in characters. */
const MAX_BOOK_NAME = 255;

/**
 * `[n]Sheet!A1`, `[n]Jan:Mar!A1`, or `[Book.xlsx]Sheet!A1`: an external workbook index or name
 * before an unquoted sheet name or sheet range, as one prefix, or undefined. A structured reference never has a sheet name and `!` after its
 * closing bracket, so the two do not meet.
 */
function readBookPrefix(r: Reader, start: number): Step | undefined {
  const s = r.s;
  let j = r.i + 1;
  while (j < s.length && j - r.i <= MAX_BOOK_NAME) {
    const c = s.charCodeAt(j);
    if (c === 93 /* ] */ || c <= 32 || BOOK_STOP.has(c)) break;
    j += 1;
  }
  if (j === r.i + 1 || s.charCodeAt(j) !== 93) return undefined;
  const nameEnd = wordEnd(s, j + 1);
  if (nameEnd === j + 1) return undefined;
  const sheets = [s.slice(j + 1, nameEnd)];
  let bang = nameEnd;
  if (s.charCodeAt(nameEnd) === 58 /* : a sheet range `[1]Jan:Mar!` */) {
    const end2 = wordEnd(s, nameEnd + 1);
    if (end2 === nameEnd + 1) return undefined;
    sheets.push(s.slice(nameEnd + 1, end2));
    bang = end2;
  }
  if (s.charCodeAt(bang) !== 33) return undefined;
  const book = s.slice(r.i + 1, j);
  r.i = bang + 1;
  return readQualified(r, { raw: s.slice(start, r.i), sheets, book }, start);
}

/** A word of digits, with an optional `$`: the shape of a row number. */
const ROW_WORD = /^\$?[0-9]{1,16}$/;

/** A number (`1`, `.5`, `2.5E-3`), or a whole-row range that starts with digits (`1:3`). */
function readNumber(r: Reader): Step {
  const s = r.s;
  const start = r.i;
  const wEnd = wordEnd(s, start);
  if (s.charCodeAt(wEnd) === 58) {
    const end2 = wordEnd(s, wEnd + 1);
    const [a, b] = [s.slice(start, wEnd), s.slice(wEnd + 1, end2)];
    if (rowEnd(a) && rowEnd(b)) return readWord(r, undefined, start);
    // Two whole numbers around `:` are a row range; a row outside the grid is not one.
    if (ROW_WORD.test(a) && ROW_WORD.test(b)) return 'bad-range';
  }
  let j = start;
  let digits = 0;
  while (isDigit(s.charCodeAt(j))) {
    j += 1;
    digits += 1;
  }
  if (s.charCodeAt(j) === 46) {
    j += 1;
    while (isDigit(s.charCodeAt(j))) {
      j += 1;
      digits += 1;
    }
  }
  if (digits === 0) return 'bad-number';
  if ((s.charCodeAt(j) | 32) === 101 /* e */) {
    j += 1;
    if (s.charCodeAt(j) === 43 || s.charCodeAt(j) === 45) j += 1;
    let exp = 0;
    while (isDigit(s.charCodeAt(j))) {
      j += 1;
      exp += 1;
    }
    if (exp === 0) return 'bad-number';
  }
  if (j < s.length && isWordChar(s.charCodeAt(j))) return 'bad-number';
  r.i = j;
  return r.push({ kind: 'number', text: s.slice(start, j), start });
}

/** One token at `r.i`. */
function step(r: Reader): Step {
  const s = r.s;
  const start = r.i;
  const c = s.charCodeAt(start);
  const one = (kind: PlainTokenKind, length = 1): Step => {
    r.i += length;
    return r.push({ kind, text: s.slice(start, r.i), start });
  };
  if (isSpace(c)) {
    let j = start;
    while (j < s.length && isSpace(s.charCodeAt(j))) j += 1;
    return one('space', j - start);
  }
  switch (c) {
    case 34 /* " */: {
      let j = start + 1;
      for (;;) {
        if (j >= s.length) return 'unterminated-string';
        if (s.charCodeAt(j) !== 34) j += 1;
        else if (s.charCodeAt(j + 1) === 34) j += 2;
        else break;
      }
      return one('string', j + 1 - start);
    }
    case 35 /* # */:
      return readError(r, start);
    case 39 /* ' */:
      return readQuotedSheet(r, start);
    case 91 /* [ */: {
      const book = readBookPrefix(r, start);
      return book === undefined ? readStructured(r, start) : book;
    }
    case 40 /* ( */:
    case 123 /* { */:
      r.stack.push(c);
      if (r.stack.length > MAX_FORMULA_DEPTH) return 'too-deep';
      return one(c === 40 ? 'open' : 'arrayOpen');
    case 41 /* ) */:
    case 125 /* } */:
      if (r.stack.pop() !== (c === 41 ? 40 : 123)) return 'unbalanced';
      return one(c === 41 ? 'close' : 'arrayClose');
    case 44 /* , */:
      return one('separator');
    case 59 /* ; a row separator, only inside an array constant */:
      return r.stack[r.stack.length - 1] === 123 ? one('separator') : 'unexpected-character';
    case 60 /* < */: {
      const n = s.charCodeAt(start + 1);
      return one('operator', n === 61 || n === 62 ? 2 : 1);
    }
    case 62 /* > */:
      return one('operator', s.charCodeAt(start + 1) === 61 ? 2 : 1);
    case 43 /* + */:
    case 45 /* - */:
    case 42 /* * */:
    case 47 /* / */:
    case 94 /* ^ */:
    case 38 /* & */:
    case 37 /* % */:
    case 61 /* = */:
    case 58 /* : the range operator between operands that are not plain references */:
      return one('operator');
  }
  if (isDigit(c) || (c === 46 && isDigit(s.charCodeAt(start + 1)))) return readNumber(r);
  if (c === 36 || (isWordChar(c) && c !== 46 && c !== 63)) return readWord(r, undefined, start);
  return 'unexpected-character';
}

/** Read formula text (without the leading `=`). Never throws. */
export function tokenizeFormula(text: string): TokenizeResult {
  if (text.length > MAX_FORMULA_LENGTH) return { ok: false, problem: 'too-long', at: 0 };
  const invalid = xmlInvalidAt(text);
  if (invalid >= 0) return { ok: false, problem: 'invalid-character', at: invalid };
  const r = new Reader(text);
  while (r.i < text.length) {
    const at = r.i;
    const problem = step(r);
    if (problem) return { ok: false, problem, at };
    if (r.i <= at) return { ok: false, problem: 'unexpected-character', at };
  }
  if (r.stack.length > 0) return { ok: false, problem: 'unbalanced', at: text.length };
  if (r.tokens.every((t) => t.kind === 'space')) return { ok: false, problem: 'empty', at: 0 };
  return { ok: true, tokens: r.tokens };
}
