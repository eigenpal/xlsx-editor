// A light grammar check of formula text that a user enters (design D4): after the tokenizer
// reads the text, the tokens must alternate between operands and operators. It refuses an
// operator with no operand (`A1+`, `+`, `1+*2`), an empty group (`()`), two operands with no
// operator (`1 2`), and a `,` or `;` outside a function call, a group, or an array constant.
//
// Rules:
// - `+` and `-` are unary where an operand is expected; `%` follows an operand.
// - A function call can have empty arguments (`F()`, `IF(A1,,1)`); a group and an array cannot.
// - In a group, `,` is the union operator (`SUM((A1,B1))`); at the top level it is refused.
// - A space between two references is the intersection operator (`A1:B2 B1:C3`).
// - `:` is the range operator of the formula syntax (ECMA-376 Part 1, 18.17.2), and both of its
//   operands are references: a cell or range reference, a name, `#REF!`, a structured
//   reference, or the result of a closed group or function call (`OFFSET(A1,1,1):B5`). A
//   number, a string, a boolean, an error other than `#REF!`, or an array constant next to `:`
//   is refused with `bad-range` (`="a":"b"`, `=1.5:2`, `={1,2}:A1`, `=A1:1`), and so is a
//   range end that is not a cell, column, or row of the grid nor a defined name (`=A1:A`,
//   `=A1:XFE1`; `checkRangeEnds`).
//
// It applies to user input only. Formula text from the file stays as it is (design D4): the
// store never refuses an edit because a formula in the file breaks these rules.
//
// Security (AGENTS.md, "Security"): one forward pass over the tokens, with a stack no deeper than
// the tokenizer's nesting cap.

import type { FormulaProblem, FormulaToken } from './tokenizer.ts';

/** Why the grammar check refused formula text, and the offset of the token. */
export interface GrammarProblem {
  readonly problem: FormulaProblem;
  readonly at: number;
}

const OPERAND_KINDS = new Set<FormulaToken['kind']>([
  'reference',
  'name',
  'error',
  'number',
  'string',
  'boolean',
  'structured',
]);

/** Tokens that can stand on either side of the intersection operator (a space). */
const REFERENCE_KINDS = new Set<FormulaToken['kind']>([
  'reference',
  'name',
  'error',
  'structured',
  'close',
  'function',
  'open',
]);

/**
 * A name with the shape of an A1 cell or column end: one to three letters and an optional row,
 * each with an optional `$`. The tokenizer reads such text as a name only when it is not a
 * reference: a column without a row next to a cell (`A1:A`, `A:A1`), or an end outside the
 * grid (`A1:XFE1`, `A1:A1048577`). Such text can also be a defined name (`tax`).
 */
const REF_SHAPE = /^\$?[A-Za-z]{1,3}(\$?[0-9]+)?$/;

/**
 * The first range end next to `:` that has the shape of a cell or column end
 * (`REF_SHAPE`) but is not one of the defined names `names` (in upper case): a broken range
 * end, refused with `bad-range`. Null when there is none.
 */
export function checkRangeEnds(
  tokens: readonly FormulaToken[],
  names: ReadonlySet<string>
): GrammarProblem | null {
  const colon = (t: FormulaToken | undefined) => t?.kind === 'operator' && t.text === ':';
  for (let k = 0; k < tokens.length; k += 1) {
    const t = tokens[k]!;
    if (t.kind !== 'name' || !REF_SHAPE.test(t.text) || names.has(t.text.toUpperCase())) continue;
    let before = k - 1;
    while (tokens[before]?.kind === 'space') before -= 1;
    let after = k + 1;
    while (tokens[after]?.kind === 'space') after += 1;
    if (colon(tokens[before]) || colon(tokens[after])) return { problem: 'bad-range', at: t.start };
  }
  return null;
}

/** Whether a token can be an operand of the range operator `:`. */
const rangeOperand = (t: FormulaToken | undefined, side: 'left' | 'right'): boolean =>
  t !== undefined &&
  (t.kind === 'reference' ||
    t.kind === 'name' ||
    t.kind === 'structured' ||
    (t.kind === 'error' && t.text.toUpperCase() === '#REF!') ||
    (side === 'left' ? t.kind === 'close' : t.kind === 'function' || t.kind === 'open'));

/** Whether a token is a reference with a sheet prefix. */
const qualified = (t: FormulaToken | undefined): boolean =>
  t?.kind === 'reference' && t.sheet !== undefined;

type Frame = 'call' | 'group' | 'array';

/** The first grammar problem of a token list from `tokenizeFormula`, or null. */
export function checkFormulaGrammar(tokens: readonly FormulaToken[]): GrammarProblem | null {
  const stack: Frame[] = [];
  let expectOperand = true;
  let prev: FormulaToken | undefined;
  let spaced = false;
  const fail = (problem: FormulaProblem, token: FormulaToken): GrammarProblem => ({
    problem,
    at: token.start,
  });
  for (let k = 0; k < tokens.length; k += 1) {
    const token = tokens[k]!;
    if (token.kind === 'space') {
      spaced = true;
      continue;
    }
    const top = stack[stack.length - 1];
    const startsOperand =
      OPERAND_KINDS.has(token.kind) ||
      token.kind === 'function' ||
      token.kind === 'open' ||
      token.kind === 'arrayOpen';
    if (!expectOperand && startsOperand) {
      // Two operands: only a space between two references is an operator (intersection).
      const intersect =
        spaced &&
        prev !== undefined &&
        REFERENCE_KINDS.has(prev.kind) &&
        REFERENCE_KINDS.has(token.kind);
      if (!intersect) return fail('missing-operator', token);
      expectOperand = true;
    }
    spaced = false;
    switch (token.kind) {
      case 'function':
        // The tokenizer reads a function name only before `(`; the `open` that follows is a call.
        stack.push('call');
        expectOperand = true;
        break;
      case 'open':
        if (prev?.kind !== 'function') stack.push('group');
        expectOperand = true;
        break;
      case 'arrayOpen':
        stack.push('array');
        expectOperand = true;
        break;
      case 'close':
      case 'arrayClose': {
        // A call can end with an empty argument: `F()` and `F(1,)`.
        const emptyArgument =
          top === 'call' && (prev?.kind === 'open' || prev?.kind === 'separator');
        if (expectOperand && !emptyArgument) return fail('missing-operand', token);
        stack.pop();
        expectOperand = false;
        break;
      }
      case 'separator':
        if (top === undefined) return fail('misplaced-separator', token);
        // A call can have an empty argument (`IF(A1,,1)`); a group and an array cannot.
        if (
          expectOperand &&
          !(top === 'call' && (prev?.kind === 'open' || prev?.kind === 'separator'))
        )
          return fail('missing-operand', token);
        if (token.text === ';' && top !== 'array') return fail('misplaced-separator', token);
        expectOperand = true;
        break;
      case 'operator':
        if (token.text === ':') {
          let n = k + 1;
          while (tokens[n]?.kind === 'space') n += 1;
          if (!rangeOperand(prev, 'left') || !rangeOperand(tokens[n], 'right'))
            return fail('bad-range', token);
          // A range from a cell to a cell of another sheet (`A1:Data!B2`) is no A1 range.
          if (prev?.kind === 'reference' && !prev.sheet && qualified(tokens[n]))
            return fail('bad-range', token);
          expectOperand = true;
        } else if (token.text === '%') {
          if (expectOperand) return fail('missing-operand', token);
        } else if (expectOperand) {
          if (token.text !== '+' && token.text !== '-') return fail('missing-operand', token);
        } else expectOperand = true;
        break;
      default:
        expectOperand = false;
    }
    prev = token;
  }
  if (expectOperand && prev) return fail('missing-operand', prev);
  return null;
}
