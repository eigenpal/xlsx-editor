// The text that cell A1 of each opening case of the hang test must show, or why a case does not
// check it. A workbook can open and still lose its data, so `cases.test.ts` keeps this list whole.

/** The cases whose cell A1 holds the text `first`, which a lost sheet would not show. */
const FIRST_IN_A1 = new Set([
  'relationship-target-traversal',
  'names-object-keys',
  'defined-names-many',
  'macro-part',
  'dimension-whole-sheet',
  'dimension-garbage',
  'col-range-huge',
  'col-overlap-many',
  'axis-default-zero',
  'merge-whole-sheet',
  'merge-overlap-many',
  'merge-small-many',
  'merge-hostile-reference',
  'style-xml-deep',
  'view-hostile',
  'conditional-many',
  'drawing-anchors-hostile',
  'sheets-many',
  'external-references',
  'network-blocked',
]);

/** Why an opening case does not check the text of cell A1. */
export const NO_FIRST_CELL: ReadonlyMap<string, string> = new Map([
  ['xml-entity-expansion', 'A1 holds text that the entity rules decide'],
  ['sst-count-huge', 'A1 holds a shared string of generated text'],
  ['sst-index-huge', 'A1 names a shared string that does not exist'],
  ['row-number-hostile', 'A1 holds a one-digit number, which other cells can also show'],
  ['row-last-of-grid', 'the sheet has no cell A1'],
  ['cell-reference-hostile', 'A1 holds a one-digit number, which other cells can also show'],
  ['row-spans-huge', 'A1 holds a one-digit number, which other cells can also show'],
  ['everything-hidden', 'A1 is hidden'],
  ['rows-alternate-hidden', 'the sheet has no cell A1'],
  ['style-index-hostile', 'A1 holds a one-digit number, which other cells can also show'],
  ['style-many-formats', 'A1 holds a one-digit number, which other cells can also show'],
  ['number-format-hostile', 'the hostile number format decides the text of A1'],
  ['cell-values-hostile', 'the text of A1 depends on the number format'],
  ['sheet-dense', 'A1 holds the number 0, which other cells can also show'],
  ['text-long-cell', 'A1 holds 10 MB of random text, which the live region shortens'],
  ['rich-string-many-rows', 'A1 holds 20000 runs of text, which the live region shortens'],
]);

/** The text that A1 of the case `id` must show, when the table above knows it. */
export const FIRST_CELL: ReadonlyMap<string, string> = new Map([
  ...[...FIRST_IN_A1].map((id): [string, string] => [id, 'first']),
  ['axis-size-hostile', 'a'],
  ['style-chain-cycle', 'chain'],
  ['font-size-hostile', 'huge'],
]);

/** The text that the last cell of the case `id` must show, so a lost row fails the case. */
export const LAST_CELL: ReadonlyMap<string, string> = new Map([
  ['sheet-dense', '474981'],
  ['sheets-many', 'third'],
]);

/** The large cases whose saved copy must keep at least half of the XML of the opened file. */
export const SAVE_KEEPS_CONTENT: ReadonlySet<string> = new Set([
  'sheet-dense',
  'wrapped-rows-many',
  'style-many-formats',
  'merge-small-many',
  'sheets-many',
]);

/**
 * The cases whose edit of cell A1 the store must refuse, with the refusal code and why. Any other
 * refusal, or a refusal with another code, fails.
 */
export const EDIT_REFUSED: ReadonlyMap<string, { readonly code: string; readonly why: string }> =
  new Map([
    [
      'xml-entity-expansion',
      {
        code: 'unsupported-package',
        why: 'the shared string table declares a DTD, so it stays as bytes',
      },
    ],
    [
      'row-number-hostile',
      { code: 'invalid-sheet', why: 'the sheet has a row number outside the grid' },
    ],
    [
      'cell-reference-hostile',
      { code: 'invalid-sheet', why: 'the sheet has a cell reference outside the grid' },
    ],
  ]);
