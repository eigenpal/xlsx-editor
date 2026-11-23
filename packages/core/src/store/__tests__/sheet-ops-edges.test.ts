// Edge cases of the sheet ops of design D1 of the change `workbook-structure-editing`: a sheet
// range prefix whose start reads as a cell reference, rows inside `mc:AlternateContent`, and a
// content type override that names a missing worksheet part.

import { describe, expect, test } from 'bun:test';
import { c, row } from '../../../../../scripts/fixtures/builder.ts';
import { attr } from '../package/query.ts';
import { parseXmlPart, type Element } from '../package/tree.ts';
import { CONTENT_TYPES_PART, xmlPartOf, type XlsxPackage } from '../package/package.ts';
import { qualifyFormula } from '../formula/format.ts';
import { rewriteSheetPrefixes } from '../structure/sheet-refs.ts';
import { touchedPartViolation } from '../store/part-checks.ts';
import { WorkbookDocumentStore } from '../store/document-store.ts';
import type { WorkbookOp } from '../store/ops.ts';
import { formulaOf, open, storeOf } from './structure-helpers.ts';

const rejection = (s: WorkbookDocumentStore, ops: readonly WorkbookOp[]) => {
  const before = s.package;
  const r = s.transact(ops);
  expect(s.package).toBe(before);
  return r.ok ? null : { code: r.rejection.code, subject: r.rejection.subject };
};

describe('a sheet range prefix whose start reads as a cell reference', () => {
  const book = () =>
    storeOf({
      sheets: [
        { name: 'S', rows: row(1, [c('A1', 1), c('B2', 2)]) },
        { name: 'T', rows: row(1, [c('A1', 1, { f: 'SUM(A1:S!B2)' })]) },
      ],
    });

  test('a rename of the sheet that the range names is refused, and the formula keeps its text', () => {
    const s = book();
    expect(rejection(s, [{ op: 'renameSheet', sheet: 'S', name: 'New' }])).toEqual({
      code: 'structure-blocked',
      subject: 'T!A1',
    });
    expect(formulaOf(s.package, 'T', 'A1')).toBe('SUM(A1:S!B2)');
  });

  test('a delete of the sheet is refused, as for every sheet range that ends at it', () => {
    const s = book();
    expect(rejection(s, [{ op: 'deleteSheet', sheet: 'S' }])?.code).toBe('sheet-range-blocked');
  });

  test('the prefix rewrite refuses a range with a cell or an R1C1 name at either end', () => {
    for (const text of ['A1:S!B2', '$B$2:S!B2', 'R1C1:S!B2', 'S:A1!B2'])
      expect(rewriteSheetPrefixes(text, { kind: 'rename', sheet: 'S', name: 'New' }).ok).toBe(
        false
      );
  });

  test('a range that names a sheet that is not in the workbook is refused', () => {
    const sheets = new Set(['S', 'T']);
    const edit = { kind: 'rename', sheet: 'S', name: 'New', sheets } as const;
    expect(rewriteSheetPrefixes('Nope:S!B2', edit).ok).toBe(false);
    expect(rewriteSheetPrefixes('S:T!B2', edit)).toEqual({
      ok: true,
      text: 'New:T!B2',
      changed: true,
    });
  });

  test('a quoted range keeps its meaning, so a rename rewrites it', () => {
    const edit = { kind: 'rename', sheet: 'S', name: 'New', sheets: new Set(['A1', 'S']) } as const;
    expect(rewriteSheetPrefixes("'A1:S'!B2", edit)).toEqual({
      ok: true,
      text: "'A1:New'!B2",
      changed: true,
    });
  });

  test('qualifyFormula refuses the same text', () => {
    expect(qualifyFormula('SUM(A1:S!B2)', 'Data!')).toEqual({
      ok: false,
      problem: 'ambiguous-sheet-range',
      at: 4,
    });
    expect(qualifyFormula("SUM('A1:S'!B2)", 'Data!').ok).toBe(true);
  });
});

describe('rows inside mc:AlternateContent', () => {
  const wrapped =
    '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">' +
    '<mc:Choice xmlns:x14="urn:example:x14" Requires="x14"><row r="2"><c r="A2"><f>S!A1</f><v>1</v></c></row></mc:Choice>' +
    '<mc:Fallback><row r="2"><c r="A2"><f>S!A1</f><v>1</v></c></row></mc:Fallback>' +
    '</mc:AlternateContent>';
  const book = (rows = wrapped) =>
    storeOf({
      sheets: [
        { name: 'S', rows: row(1, [c('A1', 1)]) },
        { name: 'T', rows: row(1, [c('A1', 1, { f: 'S!A1' })]) + rows },
      ],
    });

  test('a rename and a delete of the sheet that the wrapped cells name are refused', () => {
    for (const op of [
      { op: 'renameSheet', sheet: 'S', name: 'New' },
      { op: 'deleteSheet', sheet: 'S' },
    ] as const) {
      const s = book();
      const r = rejection(s, [op]);
      expect(r?.code).toBe('structure-blocked');
      expect(r?.subject).toContain('mc:AlternateContent');
    }
  });

  test('wrapped cells that do not name the sheet do not refuse a rename', () => {
    const s = book(wrapped.replaceAll('S!A1', 'B1'));
    expect(s.transact([{ op: 'renameSheet', sheet: 'S', name: 'New' }]).ok).toBe(true);
    expect(formulaOf(s.package, 'T', 'A1')).toBe('New!A1');
  });
});

describe('a content type override that names a missing worksheet part', () => {
  const OVERRIDE =
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/' +
    'vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>';

  /** The package with one more override of `/xl/worksheets/sheet2.xml`. */
  function withDangling(pkg: XlsxPackage): XlsxPackage {
    const ct = xmlPartOf(pkg, CONTENT_TYPES_PART)!;
    const extra = parseXmlPart(CONTENT_TYPES_PART, OVERRIDE);
    if (!extra.ok) throw new Error('bad override');
    const ov = { ...(extra.part.root.children[0] as Element), id: `n${ct.nextId}` };
    const parts = new Map(pkg.parts);
    parts.set(CONTENT_TYPES_PART, {
      kind: 'xml',
      partName: CONTENT_TYPES_PART,
      xml: {
        ...ct,
        root: { ...ct.root, children: [...ct.root.children, ov] },
        nextId: ct.nextId + 1,
      },
    });
    return { parts };
  }

  const overrides = (pkg: XlsxPackage) =>
    xmlPartOf(pkg, CONTENT_TYPES_PART)!
      .root.children.filter((n): n is Element => n.type === 'element' && n.local === 'Override')
      .map((n) => attr(n, 'PartName'));

  test('addSheet skips the part name that the override names', () => {
    const pkg = withDangling(open({ sheets: [{ name: 'S', rows: row(1, [c('A1', 1)]) }] }));
    const s = new WorkbookDocumentStore(pkg);
    expect(s.transact([{ op: 'addSheet', name: 'N', index: 1 }]).ok).toBe(true);
    const names = overrides(s.package);
    expect(names.filter((n) => n === '/xl/worksheets/sheet2.xml')).toHaveLength(1);
    expect(names).toContain('/xl/worksheets/sheet3.xml');
  });

  test('the part check refuses a transaction that adds a second override for one part', () => {
    const before = open({ sheets: [{ name: 'S', rows: row(1, [c('A1', 1)]) }] });
    const once = withDangling(before);
    const twice = withDangling(once);
    expect(touchedPartViolation(once, twice, [CONTENT_TYPES_PART])).toContain(
      'two content type overrides'
    );
    // A duplicate that the input already had does not refuse an edit.
    expect(touchedPartViolation(twice, twice, [CONTENT_TYPES_PART])).toBeUndefined();
  });
});
