// Write tokens back to formula text: as entered (with function names, references, booleans,
// and error literals in uppercase), or with every relative reference moved by an offset
// (design D4, shared formula children). Strings, names, sheet prefixes, and structured
// references keep their text, except that an unquoted sheet name with the shape of a cell, an
// R1C1 reference, or a boolean is written in quotes.

import { columnLetters, MAX_COLS, MAX_ROWS } from '../cell-ref.ts';
import {
  sheetNameNeedsQuotes,
  tokenizeFormula,
  type FormulaProblem,
  type FormulaToken,
  type RefAxis,
  type RefEnd,
  type ReferenceToken,
  type SheetPrefix,
} from './tokenizer.ts';
import { checkFormulaGrammar, checkRangeEnds } from './grammar.ts';
import { PREDEFINED_FUNCTIONS } from './predefined-functions.generated.ts';
import {
  checkFormulaReferences,
  type FormulaReferenceProblem,
  type FormulaScope,
} from './references.ts';
import { sheetNameKey } from '../workbook/workbook-index.ts';

const NO_NAMES: ReadonlySet<string> = new Set();

/**
 * Why a formula cannot be written. `ambiguous-sheet-range`: an unquoted sheet range whose start
 * or end reads as a cell reference (`ambiguousSheetRange`). `name-rebinds`: a name with no sheet
 * prefix would bind to another defined name (`qualifyFormula`).
 */
export type FormulaWriteProblem =
  | FormulaProblem
  | FormulaReferenceProblem
  | 'out-of-grid'
  | 'unsupported-function'
  | 'ambiguous-sheet-range'
  | 'name-rebinds'
  | 'scoped-function';

/**
 * Whether a sheet range prefix is unquoted and one of its names reads as a cell, row, or R1C1
 * reference (`A1`, `$B$2`, `12`, `R1C1`). The tokenizer reads `A1:Sheet1!B2` as the sheet range
 * `A1:Sheet1`, but the text can also mean the range from `A1` to `Sheet1!B2`. An edit that
 * rewrites or qualifies such a prefix cannot keep its meaning, so it is refused.
 */
export function ambiguousSheetRange(prefix: SheetPrefix): boolean {
  if (prefix.sheets.length !== 2 || prefix.raw.startsWith("'")) return false;
  return prefix.sheets.some(
    (name) =>
      /^\$?[A-Za-z]{1,3}\$?[0-9]+$/.test(name) ||
      /^\$?[0-9]+$/.test(name) ||
      /^[Rr][0-9]*[Cc][0-9]*$/.test(name)
  );
}

export type FormulaTextResult =
  | { readonly ok: true; readonly text: string }
  | {
      readonly ok: false;
      readonly problem: FormulaWriteProblem;
      readonly at: number;
      /** What the problem names: the function name of `unsupported-function`. */
      readonly subject?: string;
    };

/** A row and column offset. */
export interface FormulaOffset {
  readonly rows: number;
  readonly cols: number;
}

function moveAxis(
  axis: RefAxis | undefined,
  by: number,
  limit: number
): RefAxis | null | undefined {
  if (!axis) return undefined;
  if (axis.absolute || by === 0) return axis;
  const index = axis.index + by;
  return index >= 0 && index < limit ? { index, absolute: false } : null;
}

function endText(end: RefEnd, offset: FormulaOffset): string | null {
  const col = moveAxis(end.col, offset.cols, MAX_COLS);
  const row = moveAxis(end.row, offset.rows, MAX_ROWS);
  if (col === null || row === null) return null;
  let out = '';
  if (col) out += `${col.absolute ? '$' : ''}${columnLetters(col.index + 1)}`;
  if (row) out += `${row.absolute ? '$' : ''}${row.index + 1}`;
  return out;
}

/**
 * A sheet prefix as written, or in quotes when it is unquoted and a sheet name in it has the
 * shape of a cell, an R1C1 reference, or a boolean (`TRUE!A1` is written `'TRUE'!A1`).
 */
function prefixText(sheet: SheetPrefix | undefined): string {
  if (!sheet) return '';
  if (sheet.book !== undefined || sheet.raw.startsWith("'")) return sheet.raw;
  return sheet.sheets.some(sheetNameNeedsQuotes) ? `'${sheet.sheets.join(':')}'!` : sheet.raw;
}

function referenceText(token: ReferenceToken, offset: FormulaOffset): string | null {
  const from = endText(token.from, offset);
  const to = token.to ? endText(token.to, offset) : '';
  if (from === null || to === null) return null;
  return `${prefixText(token.sheet)}${from}${token.to ? `:${to}` : ''}`;
}

/**
 * A function name in upper case, with its `_xlfn.` and `_xlws.` prefixes as authored: the
 * prefixes are part of the stored name (ECMA-376 Part 1, 18.17.7), and a rewrite keeps them.
 */
function functionText(text: string): string {
  let i = 0;
  while (/^_xl(fn|ws)\./i.test(text.slice(i, i + 6))) i += 6;
  return text.slice(0, i) + text.slice(i).toUpperCase();
}

/**
 * The text of tokens, with the relative references moved by `offset`. A reference that moves
 * out of the grid fails with `out-of-grid` and its offset.
 */
export function formatTokens(
  tokens: readonly FormulaToken[],
  offset: FormulaOffset = { rows: 0, cols: 0 }
): FormulaTextResult {
  let out = '';
  for (const token of tokens) {
    switch (token.kind) {
      case 'reference': {
        const text = referenceText(token, offset);
        if (text === null) return { ok: false, problem: 'out-of-grid', at: token.start };
        out += text;
        break;
      }
      case 'function':
        out += functionText(token.text);
        break;
      case 'boolean':
        out += token.text.toUpperCase();
        break;
      case 'error':
        out += `${prefixText(token.sheet)}${token.text.slice(token.sheet?.raw.length ?? 0).toUpperCase()}`;
        break;
      case 'name':
        out += `${prefixText(token.sheet)}${token.text.slice(token.sheet?.raw.length ?? 0)}`;
        break;
      default:
        out += token.text;
    }
  }
  return { ok: true, text: out };
}

/** Read formula text and write it back as entered, or the reason it cannot be read. */
export function normalizeFormula(text: string): FormulaTextResult {
  const read = tokenizeFormula(text);
  return read.ok ? formatTokens(read.tokens) : read;
}

/**
 * Read formula text that a user entered, check its grammar (`checkFormulaGrammar`) and, with
 * `scope`, that each sheet, external workbook, and table it names is one that the workbook
 * holds (`checkFormulaReferences`), and write it back as entered. Formula text from the file
 * goes through `normalizeFormula` or `moveFormula` instead, which never apply these checks.
 */
export function readFormulaInput(text: string, scope?: FormulaScope): FormulaTextResult {
  const read = tokenizeFormula(text);
  if (!read.ok) return read;
  const grammar =
    checkFormulaGrammar(read.tokens) ?? checkRangeEnds(read.tokens, scope?.names ?? NO_NAMES);
  if (grammar) return { ok: false, ...grammar };
  // A typed call to a function outside the predefined functions of ECMA-376 Part 1, 18.17.7,
  // is refused, in any case. A name with a stored-name prefix is not in the set either.
  // Formulas read from a file never pass here, so their stored text keeps its prefixes.
  for (const token of read.tokens)
    if (token.kind === 'function' && !PREDEFINED_FUNCTIONS.has(token.text.toUpperCase()))
      return { ok: false, problem: 'unsupported-function', at: token.start, subject: token.text };
  const refs = scope ? checkFormulaReferences(read.tokens, scope) : null;
  return refs ? { ok: false, ...refs } : formatTokens(read.tokens);
}

/** Formula text with every relative reference moved by `offset`. */
export function moveFormula(text: string, offset: FormulaOffset): FormulaTextResult {
  const read = tokenizeFormula(text);
  return read.ok ? formatTokens(read.tokens, offset) : read;
}

/**
 * The position of the first range whose start has no sheet prefix and whose end has an unquoted
 * prefix of one local sheet, joined by `:` with no space (`A1:S!B2`), or -1. The tokenizer reads
 * such text as a range from `A1` to `S!B2`, but the text can also mean the sheet range
 * `A1:S`. An edit that rewrites the prefix cannot keep its meaning, so it is refused, as for
 * `ambiguousSheetRange`.
 */
export function crossSheetRangeAt(tokens: readonly FormulaToken[]): number {
  for (let i = 0; i + 2 < tokens.length; i += 1) {
    const [a, op, b] = [tokens[i]!, tokens[i + 1]!, tokens[i + 2]!];
    if (a.kind !== 'reference' || a.sheet || op.kind !== 'operator' || op.text !== ':') continue;
    if (b.kind !== 'reference' && b.kind !== 'name') continue;
    const sheet = b.sheet;
    if (!sheet || sheet.book !== undefined || sheet.raw.startsWith("'")) continue;
    if (sheet.sheets.length !== 1) continue;
    if (op.start === a.start + a.text.length && b.start === op.start + 1) return a.start;
  }
  return -1;
}

/**
 * Copied formula text with every relative reference moved by `offset`, as `moveFormula`, but an
 * unquoted sheet range that `ambiguousSheetRange` reports fails with `ambiguous-sheet-range`:
 * in `A1:S!B2`, the tokenizer reads `A1:S` as a sheet range and would move only `B2`, but the
 * text can also mean the range from `A1` to `S!B2`. A paste refuses such text, as a cut to
 * another sheet does (`qualifyFormula`).
 */
export function moveCopiedFormula(text: string, offset: FormulaOffset): FormulaTextResult {
  const read = tokenizeFormula(text);
  if (!read.ok) return read;
  const cross = crossSheetRangeAt(read.tokens);
  if (cross >= 0) return { ok: false, problem: 'ambiguous-sheet-range', at: cross };
  for (const token of read.tokens) {
    const sheet = token.kind === 'reference' || token.kind === 'name' ? token.sheet : undefined;
    if (sheet && ambiguousSheetRange(sheet))
      return { ok: false, problem: 'ambiguous-sheet-range', at: token.start };
  }
  return formatTokens(read.tokens, offset);
}

/**
 * Formula text with each reference that has no sheet prefix qualified by `prefix`, the prefix
 * text with its `!` (`Data!`, `'My data'!`), unless `references` is false. A cut that moves a
 * formula to another sheet keeps the cells it referred to this way; a copy to another sheet
 * keeps its references on the target sheet and applies only the name rules that follow.
 *
 * A name with no sheet prefix that `localNames` holds (the names that the source sheet scopes,
 * as `sheetNameKey` folds them) gets the prefix too, so it does not bind to a name of the
 * target sheet. A function name counts as a name here, because a defined name can hold a
 * function (`myFn(1)`), but formula text cannot give a function call a sheet prefix: a
 * function name that `localNames` holds fails with `scoped-function`. Every other token keeps
 * its text. A sheet range
 * that `ambiguousSheetRange` reports fails with `ambiguous-sheet-range`. A name with no sheet
 * prefix that the target sheet scopes and the source sheet does not (`targetNames` holds it and
 * `localNames` does not) fails with `name-rebinds`: on the target sheet it would bind to the
 * target's name, and no prefix names the workbook scope.
 */
export function qualifyFormula(
  text: string,
  prefix: string,
  localNames: ReadonlySet<string> = new Set(),
  targetNames: ReadonlySet<string> = new Set(),
  references = true
): FormulaTextResult {
  const read = tokenizeFormula(text);
  if (!read.ok) return read;
  const cross = crossSheetRangeAt(read.tokens);
  if (cross >= 0) return { ok: false, problem: 'ambiguous-sheet-range', at: cross };
  let out = '';
  for (const token of read.tokens) {
    const sheet = token.kind === 'reference' || token.kind === 'name' ? token.sheet : undefined;
    if (sheet && ambiguousSheetRange(sheet))
      return { ok: false, problem: 'ambiguous-sheet-range', at: token.start };
    const key = sheetNameKey(token.text);
    const bare = (token.kind === 'name' && !token.sheet) || token.kind === 'function';
    const local = bare && localNames.has(key);
    if (local && token.kind === 'function')
      return { ok: false, problem: 'scoped-function', at: token.start };
    if (bare && !local && targetNames.has(key))
      return { ok: false, problem: 'name-rebinds', at: token.start };
    const qualify = token.kind === 'reference' ? references && !token.sheet : local;
    out += qualify ? prefix + token.text : token.text;
  }
  return { ok: true, text: out };
}
