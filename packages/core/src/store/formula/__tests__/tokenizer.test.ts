import { describe, expect, test } from 'bun:test';
import {
  MAX_FORMULA_DEPTH,
  MAX_FORMULA_LENGTH,
  MAX_FORMULA_TOKENS,
  tokenizeFormula,
  type FormulaProblem,
  type FormulaToken,
  type ReferenceToken,
} from '../tokenizer.ts';
import { moveFormula, normalizeFormula, readFormulaInput } from '../format.ts';

function tokens(text: string): readonly FormulaToken[] {
  const r = tokenizeFormula(text);
  if (!r.ok) throw new Error(`${text}: ${r.problem} at ${r.at}`);
  return r.tokens;
}

const kinds = (text: string) => tokens(text).map((t) => t.kind);
const refs = (text: string) =>
  tokens(text).filter((t): t is ReferenceToken => t.kind === 'reference');
const problem = (text: string) => {
  const r = tokenizeFormula(text);
  return r.ok ? null : r.problem;
};

describe('tokenizeFormula: references', () => {
  test('a cell, relative and absolute', () => {
    const [a, b, c, d] = refs('A1+$B$2+C$3+$D4');
    expect(a!.from).toEqual({
      col: { index: 0, absolute: false },
      row: { index: 0, absolute: false },
    });
    expect(b!.from).toEqual({
      col: { index: 1, absolute: true },
      row: { index: 1, absolute: true },
    });
    expect([c!.from.col!.absolute, c!.from.row!.absolute]).toEqual([false, true]);
    expect([d!.from.col!.absolute, d!.from.row!.absolute]).toEqual([true, false]);
  });

  test('ranges, whole columns, and whole rows', () => {
    expect(refs('SUM(A1:B10)').map((r) => [r.shape, r.text])).toEqual([['range', 'A1:B10']]);
    expect(refs('SUM(A:C)')[0]!.shape).toBe('columns');
    expect(refs('SUM($A:$A)')[0]!.from.col).toEqual({ index: 0, absolute: true });
    expect(refs('SUM(2:5)')[0]!.shape).toBe('rows');
    expect(refs('SUM($2:$5)')[0]!.to!.row).toEqual({ index: 4, absolute: true });
    expect(refs('XFD1048576')[0]!.from).toEqual({
      col: { index: 16383, absolute: false },
      row: { index: 1048575, absolute: false },
    });
  });

  test('sheet-qualified references, quoted and unquoted', () => {
    const [a, b, c, d] = refs("Sheet1!A1+'My sheet'!B2:C3+'It''s'!D4+Data!E:E");
    expect(a!.sheet).toEqual({ raw: 'Sheet1!', sheets: ['Sheet1'] });
    expect(b!.sheet!.sheets).toEqual(['My sheet']);
    expect(b!.shape).toBe('range');
    expect(c!.sheet!.sheets).toEqual(["It's"]);
    expect(d!.shape).toBe('columns');
    expect(refs('Sheet1:Sheet3!A1')[0]!.sheet!.sheets).toEqual(['Sheet1', 'Sheet3']);
    expect(refs("'Jan 1:Mar 1'!A1")[0]!.sheet!.sheets).toEqual(['Jan 1', 'Mar 1']);
    expect(refs('[1]Prices!B2')[0]!.sheet).toEqual({
      raw: '[1]Prices!',
      sheets: ['Prices'],
      book: '1',
    });
    expect(refs("'[2]Other book'!A1")[0]!.sheet!.book).toBe('2');
    expect(refs('SUM([1]Jan:Mar!A1)')[0]!.sheet).toEqual({
      raw: '[1]Jan:Mar!',
      sheets: ['Jan', 'Mar'],
      book: '1',
    });
    expect(refs('Données!A1')[0]!.sheet!.sheets).toEqual(['Données']);
  });

  test('lowercase references are references', () => {
    expect(refs('a1+b$2')).toHaveLength(2);
  });

  test('names that look almost like references stay names', () => {
    for (const name of ['XFE1', 'A1048577', 'A0', 'Total', 'tax_rate', '\\path', 'A1.B', 'Rate'])
      expect(kinds(name)).toEqual(['name']);
    expect(kinds('Sheet1!Rate')).toEqual(['name']);
  });

  test('a function name that looks like a reference is a function', () => {
    expect(kinds('LOG10(100)')).toEqual(['function', 'open', 'number', 'close']);
    expect(kinds('_xlfn.CONCAT("a")')).toEqual(['function', 'open', 'string', 'close']);
  });
});

describe('tokenizeFormula: literals', () => {
  test('a string is text, never a reference', () => {
    expect(kinds('"A1"&B1')).toEqual(['string', 'operator', 'reference']);
    expect(tokens('"say ""A1"""')[0]!.text).toBe('"say ""A1"""');
    expect(kinds('INDIRECT("Sheet1!A1")')).toEqual(['function', 'open', 'string', 'close']);
  });

  test('numbers, booleans, and errors', () => {
    expect(kinds('1+2.5-.5*1E3/2.5e-2')).toEqual([
      'number',
      'operator',
      'number',
      'operator',
      'number',
      'operator',
      'number',
      'operator',
      'number',
    ]);
    expect(kinds('TRUE')).toEqual(['boolean']);
    expect(kinds('false')).toEqual(['boolean']);
    expect(kinds('#REF!+#N/A+#DIV/0!+#getting_data')).toEqual([
      'error',
      'operator',
      'error',
      'operator',
      'error',
      'operator',
      'error',
    ]);
    expect(kinds('Sheet1!#REF!')).toEqual(['error']);
    for (const e of [
      '#SPILL!',
      '#CALC!',
      '#FIELD!',
      '#BLOCKED!',
      '#CONNECT!',
      '#BUSY!',
      '#UNKNOWN!',
    ])
      expect(kinds(`${e}+#spill!`)).toEqual(['error', 'operator', 'error']);
  });

  test('tab, line feed, and carriage return are spaces', () => {
    expect(kinds('A1\t+\nB1\r\n')).toEqual([
      'reference',
      'space',
      'operator',
      'space',
      'reference',
      'space',
    ]);
  });

  test('whole-column ranges and A1 cells that look like R1C1 stay A1', () => {
    expect(kinds('R:R')).toEqual(['reference']);
    expect(kinds('C:C+R1+C2+RC1')).toEqual([
      'reference',
      'operator',
      'reference',
      'operator',
      'reference',
      'operator',
      'reference',
    ]);
    expect(kinds('ROUND(1,0)')).toEqual([
      'function',
      'open',
      'number',
      'separator',
      'number',
      'close',
    ]);
  });

  test('operators, arrays, spaces, and structured references', () => {
    expect(kinds('A1<>B1')).toEqual(['reference', 'operator', 'reference']);
    expect(kinds('A1>=1')).toEqual(['reference', 'operator', 'number']);
    expect(kinds('{1,2;3,4}')).toEqual([
      'arrayOpen',
      'number',
      'separator',
      'number',
      'separator',
      'number',
      'separator',
      'number',
      'arrayClose',
    ]);
    expect(kinds('A1:B2 B1:C3')).toEqual(['reference', 'space', 'reference']);
    expect(kinds('50%')).toEqual(['number', 'operator']);
    expect(kinds('SUM(Table1[Amount])')).toEqual(['function', 'open', 'structured', 'close']);
    expect(tokens("T[[#Headers],[A']B]]")[0]!.text).toBe("T[[#Headers],[A']B]]");
    expect(kinds('[@Price]*2')).toEqual(['structured', 'operator', 'number']);
    expect(kinds('A1:INDEX(B:B,2)')).toEqual([
      'reference',
      'operator',
      'function',
      'open',
      'reference',
      'separator',
      'number',
      'close',
    ]);
  });
});

describe('tokenizeFormula: text it cannot read', () => {
  const cases: [string, FormulaProblem][] = [
    ['', 'empty'],
    ['   ', 'empty'],
    ['"open', 'unterminated-string'],
    ["'Sheet1!A1", 'unterminated-sheet-name'],
    ["'Sheet1'A1", 'unterminated-sheet-name'],
    ['T[Col', 'unterminated-bracket'],
    ['SUM(A1', 'unbalanced'],
    ['A1)', 'unbalanced'],
    ['(1}', 'unbalanced'],
    ['1;2', 'unexpected-character'],
    ['A1 @ B1', 'unexpected-character'],
    ['A1!', 'unexpected-character'],
    ['1A', 'bad-number'],
    ['1E', 'bad-number'],
    ['$A1B', 'bad-name'],
    ['$', 'bad-name'],
    ['#BOGUS!', 'bad-error'],
    ['A1\u0000', 'invalid-character'],
    ['"a\u0001"', 'invalid-character'],
    ['"\ud800x"', 'invalid-character'],
    ["'S\u0007'!A1", 'invalid-character'],
    ['T[\udc00]', 'invalid-character'],
    ['"￾"', 'invalid-character'],
    ['R1C1', 'r1c1-reference'],
    ['RC[1]', 'r1c1-reference'],
    ['R[1]C[-1]', 'r1c1-reference'],
    ['S!R2C3+1', 'r1c1-reference'],
    ['SUM(R[-2]C:R[-1]C)', 'r1c1-reference'],
    ['A1.5', 'bad-name'],
    ['?', 'unexpected-character'],
  ];
  for (const [text, reason] of cases)
    test(`${JSON.stringify(text)}: ${reason}`, () => expect(problem(text)).toBe(reason));

  test('reports the offset where it stopped', () => {
    expect(tokenizeFormula('1+"abc')).toEqual({ ok: false, problem: 'unterminated-string', at: 2 });
  });
});

describe('tokenizeFormula: hostile input', () => {
  test('caps the length', () => {
    expect(problem('1'.repeat(MAX_FORMULA_LENGTH + 1))).toBe('too-long');
  });

  test('caps the token count', () => {
    expect(problem('1+'.repeat(MAX_FORMULA_TOKENS / 2 + 1) + '1')).toBe('too-many-tokens');
  });

  test('caps the nesting depth', () => {
    const deep = `${'('.repeat(MAX_FORMULA_DEPTH + 1)}1${')'.repeat(MAX_FORMULA_DEPTH + 1)}`;
    expect(problem(deep)).toBe('too-deep');
    expect(problem(`T${'['.repeat(MAX_FORMULA_DEPTH + 1)}`)).toBe('too-deep');
  });

  test('runs in linear time on long runs that a backtracking reader would struggle with', () => {
    const inputs = [
      `'${"''".repeat(4000)}`,
      `"${'""'.repeat(4000)}`,
      'A'.repeat(MAX_FORMULA_LENGTH),
      `T[${"'[".repeat(4000)}`,
      `${'$'.repeat(MAX_FORMULA_LENGTH)}`,
      `${'1'.repeat(MAX_FORMULA_LENGTH - 1)}:`,
    ];
    const start = performance.now();
    for (const text of inputs) tokenizeFormula(text);
    expect(performance.now() - start).toBeLessThan(500);
  });

  test('never throws on any short input', () => {
    const alphabet = [
      'A',
      '1',
      '$',
      ':',
      '!',
      "'",
      '"',
      '[',
      ']',
      '(',
      ')',
      '{',
      '}',
      '#',
      ',',
      ';',
      ' ',
      '.',
      'é',
    ];
    let seed = 7;
    for (let n = 0; n < 3000; n += 1) {
      let text = '';
      for (let k = 0; k < 8; k += 1) {
        seed = (seed * 1103515245 + 12345) % 2147483648;
        text += alphabet[seed % alphabet.length];
      }
      expect(() => tokenizeFormula(text)).not.toThrow();
    }
  });
});

describe('readFormulaInput: the grammar of user input', () => {
  const problemOf = (text: string) => {
    const r = readFormulaInput(text);
    return r.ok ? null : r.problem;
  };

  test('refuses an operator or a group with no operand', () => {
    for (const text of ['A1+', '+', '()', '1+*2', '*1', 'SUM(1+)', 'A1 +', '{}', '{1,}', '%'])
      expect([text, problemOf(text)]).toEqual([text, 'missing-operand']);
    expect(problemOf('A1+')).toBe('missing-operand');
    expect(problemOf('()')).toBe('missing-operand');
    expect(problemOf('1+*2')).toBe('missing-operand');
  });

  test('refuses two operands with no operator, and a broken number', () => {
    expect(problemOf('1 2')).toBe('missing-operator');
    expect(problemOf('A1"x"')).toBe('missing-operator');
    expect(problemOf('(1)(2)')).toBe('missing-operator');
    expect(problemOf('A1.5')).toBe('bad-name');
  });

  test('refuses the range operator between operands that are not references', () => {
    for (const text of [
      '"a":"b"',
      '1.5:2',
      'TRUE:FALSE',
      '{1,2}:A1',
      'A1:1',
      'A1:-B1',
      '#N/A:A1',
      'SUM(0:3)',
      'SUM(1:1048577)',
      '0:3',
      'A1:A',
      'A:A1',
      'A1:XFE1',
      'A1:A1048577',
      'S!A1:XFE1',
    ])
      expect([text, problemOf(text)]).toEqual([text, 'bad-range']);
    // A cell before a sheet prefix is not a sheet range: refused at the colon.
    const sheetRange = readFormulaInput('A1:Data!B2');
    expect(sheetRange).toEqual({ ok: false, problem: 'bad-range', at: 2 });
    expect(readFormulaInput('Jan:Feb!B2').ok).toBe(true);
    // A word past column XFD or row 1048576 is no cell, so it starts a sheet range.
    for (const text of ['SUM(ZZZ1:Feb!A1)', 'SUM(XFE1:Feb!A1)', 'SUM(A1048577:Feb!A1)', 'ZZZ1!A1'])
      expect([text, readFormulaInput(text).ok]).toEqual([text, true]);
    expect(readFormulaInput('SUM(A1:Feb!A1)')).toMatchObject({ ok: false, problem: 'bad-range' });
    expect(readFormulaInput('SUM(XFD1048576:Feb!A1)')).toMatchObject({
      ok: false,
      problem: 'bad-range',
    });
    // A `$` column end with no row is not a name either.
    expect(problemOf('$A$1:$B')).toBe('bad-name');
    // The tokenizer refuses a row range outside the grid on its own.
    const r = tokenizeFormula('1:1048577');
    expect(r.ok ? null : r.problem).toBe('bad-range');
  });

  test('accepts the range operator between references, names, groups, and calls', () => {
    for (const text of [
      'A1:B2',
      '1:3',
      '$1:$3',
      'A:C',
      'Sheet!A1:B2',
      'OFFSET(A1,1,1):A1',
      'A1:INDEX(B1:B5,2)',
      'Name1:Name2',
      '(A1):B2',
      '#REF!:A1',
      'A1: B2',
      'SUM(1:1048576)',
    ])
      expect([text, problemOf(text)]).toEqual([text, null]);
  });

  test('refuses a separator outside a call, a group, or an array', () => {
    expect(problemOf('A1,B1')).toBe('misplaced-separator');
    expect(problemOf('(1,)')).toBe('missing-operand');
  });

  test('accepts unary signs, percent, empty call arguments, unions, and intersections', () => {
    for (const text of [
      '-1',
      '+-A1',
      '1--1',
      '50%',
      '50%*2',
      'NOW()',
      'IF(A1,,1)',
      'ROUND(1,)',
      'SUM((A1,B1))',
      'A1:B2 B1:C3',
      'SUM( A1 , B1 )',
      '{1,2;3,4}',
      '{-1,"a"}',
      'A1:INDEX(B:B,2)',
      'Table1[Amount]*2',
      '#REF!+1',
      'S!A1=1',
    ])
      expect([text, readFormulaInput(text).ok]).toEqual([text, true]);
  });

  test('file text is never checked for grammar', () => {
    expect(normalizeFormula('A1+').ok).toBe(true);
    expect(moveFormula('1 2', { rows: 1, cols: 0 }).ok).toBe(true);
  });
});

describe('formula text', () => {
  test('normalizes what a user enters, and keeps strings and names', () => {
    expect(normalizeFormula('sum(a1:b2)+true+"x a1"+Rate+#n/a')).toEqual({
      ok: true,
      text: 'SUM(A1:B2)+TRUE+"x a1"+Rate+#N/A',
    });
    expect(normalizeFormula("'my sheet'!a1")).toEqual({ ok: true, text: "'my sheet'!A1" });
  });

  test('moves relative references and keeps absolute ones', () => {
    expect(moveFormula('A1+$B$1+C$1+$D1+SUM(E1:F2)', { rows: 2, cols: 1 })).toEqual({
      ok: true,
      text: 'B3+$B$1+D$1+$D3+SUM(F3:G4)',
    });
    expect(moveFormula('SUM(A:A)+SUM(1:1)+Sheet2!A1', { rows: 1, cols: 1 })).toEqual({
      ok: true,
      text: 'SUM(B:B)+SUM(2:2)+Sheet2!B2',
    });
    expect(moveFormula('"A1"&A1', { rows: 1, cols: 0 })).toEqual({ ok: true, text: '"A1"&A2' });
  });

  test('a reference moved out of the grid fails', () => {
    expect(moveFormula('A1', { rows: -1, cols: 0 })).toEqual({
      ok: false,
      problem: 'out-of-grid',
      at: 0,
    });
    expect(moveFormula('XFD1', { rows: 0, cols: 1 }).ok).toBe(false);
  });
});

describe('formula rewrite: function names', () => {
  test('a rewrite upper-cases a function name and keeps its _xlfn. and _xlws. prefixes', () => {
    const moved = moveFormula('_xlfn.xlookup(1,A1:A2,B1:B2)+_xlfn._xlws.sort(A1:A2)', {
      rows: 1,
      cols: 0,
    });
    expect(moved).toEqual({
      ok: true,
      text: '_xlfn.XLOOKUP(1,A2:A3,B2:B3)+_xlfn._xlws.SORT(A2:A3)',
    });
    expect(normalizeFormula('sum(A1)')).toEqual({ ok: true, text: 'SUM(A1)' });
  });
});
