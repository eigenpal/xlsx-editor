// The children of a cell that a value edit does not replace (`buildCell`, ECMA-376 Part 1, 18.3.1.4): a value
// edit, a clear, a paste, and the single-cell op keep `x:extLst`, an element in another
// namespace, and content that the schema does not allow, in their places, through save and
// reopen. Value content in `mc:AlternateContent` is replaced with the value. Other markup
// compatibility content, which the edit can neither keep nor replace, refuses the edit.

import { describe, expect, test } from 'bun:test';
import { row } from '../../../../../scripts/fixtures/builder.ts';
import { readXlsx, xmlPartOf, type XlsxPackage } from '../package/package.ts';
import { serializeXmlPart } from '../package-write/serialize.ts';
import { saveXlsx } from '../package-write/save.ts';
import type { WorkbookDocumentStore } from '../store/document-store.ts';
import type { CellInput, WorkbookOp } from '../store/ops.ts';
import { storeOf } from './structure-helpers.ts';
import { firstValueLost } from '../store/merge-cells.ts';

const SHEET = '/xl/worksheets/sheet1.xml';
const E = 'urn:example:extension';
const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
// `mc:Choice/@Requires` names a prefix, which must be declared (ECMA-376 Part 3, 7.6).
const X14 = 'urn:example:x14';

/** A store whose sheet `S` holds `cells` in row 2, with a value in A1 and B3 around them. */
const sheetWith = (cells: string) =>
  storeOf({
    sheets: [
      {
        name: 'S',
        rows:
          row(1, ['<c r="A1"><v>1</v></c>']) +
          `<row r="2">${cells}</row>` +
          row(3, ['<c r="B3"><v>3</v></c>']),
      },
    ],
  });

const set = (ref: string, value: CellInput): WorkbookOp => ({
  op: 'setCellValue',
  sheet: 'S',
  ref,
  value,
});

/** The XML of the cell `ref` in the sheet part, from `<c` to its end. */
function cellXml(pkg: XlsxPackage, ref: string): string | undefined {
  const text = serializeXmlPart(xmlPartOf(pkg, SHEET)!);
  const at = text.indexOf(`<c r="${ref}"`);
  if (at < 0) return undefined;
  const close = text.indexOf('</c>', at);
  const empty = text.indexOf('/>', at);
  const end =
    close < 0 || (empty >= 0 && empty < text.indexOf('>', at) + 1) ? empty + 2 : close + 4;
  return text.slice(at, end);
}

/** The cell after a save and a reopen of the store's package. */
function reopened(s: WorkbookDocumentStore, ref: string): string | undefined {
  const read = readXlsx(saveXlsx(s.package));
  if (!read.ok) throw new Error('reopen');
  return cellXml(read.pkg, ref);
}

const FOREIGN = `<e:x xmlns:e="${E}">keep</e:x>`;

describe('a value edit keeps the children that it does not replace', () => {
  test('keeps an element in another namespace through save and reopen', () => {
    const s = sheetWith(`<c r="A2"><v>2</v>${FOREIGN}</c>`);
    expect(s.transact([set('A2', { type: 'number', value: 7 })]).ok).toBe(true);
    for (const xml of [cellXml(s.package, 'A2'), reopened(s, 'A2')]) {
      expect(xml).toContain('<v>7</v>');
      expect(xml).toMatch(/<v>7<\/v><e:x xmlns:e="urn:example:extension">keep<\/e:x><\/c>$/);
    }
  });

  test('puts the new value where the old one was, and keeps each other child in its place', () => {
    const s = sheetWith(
      `<c r="A2"><e:a xmlns:e="${E}"/><f>1+1</f><v>2</v><e:b xmlns:e="${E}"/><zz/>` +
        '<extLst><ext uri="{0}"/></extLst></c>'
    );
    expect(s.transact([set('A2', { type: 'string', value: 'text' })]).ok).toBe(true);
    const xml = reopened(s, 'A2')!;
    expect(xml.replace(/ xmlns:e="[^"]*"/g, '')).toMatch(
      /^<c r="A2" t="s"><e:a\/><v>\d+<\/v><e:b\/><zz\/><extLst><ext uri="\{0\}"\/><\/extLst><\/c>$/
    );
  });

  test('puts a value before x:extLst when the cell had none', () => {
    const s = sheetWith(`<c r="A2" s="0">${FOREIGN}<extLst><ext uri="{0}"/></extLst></c>`);
    expect(s.transact([set('A2', { type: 'boolean', value: true })]).ok).toBe(true);
    expect(reopened(s, 'A2')).toMatch(/<\/e:x><v>1<\/v><extLst>/);
  });

  test('keeps the children in a run of edits, as one edit does', () => {
    const one = sheetWith(`<c r="A2"><v>2</v>${FOREIGN}</c><c r="B2"><v>5</v></c>`);
    const run = sheetWith(`<c r="A2"><v>2</v>${FOREIGN}</c><c r="B2"><v>5</v></c>`);
    expect(one.transact([set('A2', { type: 'number', value: 7 })]).ok).toBe(true);
    expect(
      run.transact([
        set('A2', { type: 'number', value: 7 }),
        set('B2', { type: 'number', value: 8 }),
      ]).ok
    ).toBe(true);
    expect(cellXml(run.package, 'A2')).toBe(cellXml(one.package, 'A2'));
    expect(reopened(run, 'A2')).toContain('keep</e:x>');
  });

  test('keeps a cell that a clear empties when it holds another child', () => {
    const s = sheetWith(`<c r="A2"><v>2</v>${FOREIGN}</c>`);
    expect(s.transact([{ op: 'clearValues', sheet: 'S', ref: 'A1:B2' }]).ok).toBe(true);
    expect(reopened(s, 'A2')).toMatch(
      /^<c r="A2"><e:x xmlns:e="urn:example:extension">keep<\/e:x><\/c>$/
    );
    // A cell with no other child goes, as before.
    expect(cellXml(s.package, 'A1')).toBeUndefined();
  });

  test('drops white space between the children, so a cell with nothing else goes', () => {
    const s = sheetWith('<c r="A2"> <v>2</v> </c>');
    expect(s.transact([set('A2', { type: 'number', value: 7 })]).ok).toBe(true);
    expect(cellXml(s.package, 'A2')).toBe('<c r="A2"><v>7</v></c>');
    expect(s.transact([{ op: 'clearValues', sheet: 'S', ref: 'A2' }]).ok).toBe(true);
    expect(cellXml(s.package, 'A2')).toBeUndefined();
  });

  test('keeps the children through a paste of a value and a cell format', () => {
    const s = sheetWith(`<c r="A2"><v>2</v>${FOREIGN}</c><c r="B2"><v>5</v></c>`);
    const paste: WorkbookOp[] = [
      set('A2', { type: 'number', value: 9 }),
      set('B2', { type: 'number', value: 9 }),
      { op: 'setCellFormatIndex', sheet: 'S', ref: 'A2', index: 0 },
      { op: 'setCellFormatIndex', sheet: 'S', ref: 'B2', index: 0 },
    ];
    expect(s.transact(paste).ok).toBe(true);
    expect(reopened(s, 'A2')).toMatch(
      /<v>9<\/v><e:x xmlns:e="urn:example:extension">keep<\/e:x><\/c>$/
    );
  });

  test('restores the children on undo', () => {
    const s = sheetWith(`<c r="A2"><v>2</v>${FOREIGN}</c>`);
    const before = cellXml(s.package, 'A2');
    expect(s.transact([set('A2', { type: 'number', value: 7 })]).ok).toBe(true);
    expect(s.undo().ok).toBe(true);
    expect(cellXml(s.package, 'A2')).toBe(before);
  });
});

describe('markup compatibility content in a cell', () => {
  const VALUE_ONLY =
    `<mc:AlternateContent xmlns:mc="${MC}" xmlns:x14="${X14}"><mc:Choice Requires="x14"><is><t>2</t></is></mc:Choice>` +
    '<mc:Fallback><v>2</v></mc:Fallback></mc:AlternateContent>';
  const OTHER =
    `<mc:AlternateContent xmlns:mc="${MC}" xmlns:x14="${X14}"><mc:Choice Requires="x14">${FOREIGN}</mc:Choice>` +
    '</mc:AlternateContent>';

  test('replaces alternate content that holds only value content', () => {
    const s = sheetWith(`<c r="A2">${VALUE_ONLY}${FOREIGN}</c>`);
    expect(s.transact([set('A2', { type: 'number', value: 7 })]).ok).toBe(true);
    expect(reopened(s, 'A2')).toMatch(
      /^<c r="A2"><v>7<\/v><e:x xmlns:e="urn:example:extension">keep<\/e:x><\/c>$/
    );
  });

  test('a clear replaces alternate content that holds the only value of the cell', () => {
    const wrapped =
      `<mc:AlternateContent xmlns:mc="${MC}" xmlns:x14="${X14}"><mc:Choice Requires="x14"><v>5</v></mc:Choice>` +
      '</mc:AlternateContent>';
    const s = sheetWith(`<c r="A2">${wrapped}${FOREIGN}</c>`);
    expect(s.transact([{ op: 'clearValues', sheet: 'S', ref: 'A2' }]).ok).toBe(true);
    expect(reopened(s, 'A2')).toMatch(
      /^<c r="A2"><e:x xmlns:e="urn:example:extension">keep<\/e:x><\/c>$/
    );
  });

  test('a merge names a cell whose only value is wrapped, as the clear that it asks for', () => {
    const wrapped =
      `<mc:AlternateContent xmlns:mc="${MC}" xmlns:x14="${X14}"><mc:Choice Requires="x14"><v>5</v></mc:Choice>` +
      '</mc:AlternateContent>';
    const s = sheetWith(`<c r="B2">${wrapped}</c>`);
    // A1 is the top-left cell of A1:B2, which the merge keeps.
    expect(firstValueLost(xmlPartOf(s.package, SHEET)!, [0, 0, 1, 1])).toBe('B2');
    const r = s.transact([{ op: 'mergeCells', sheet: 'S', ref: 'A1:B2' }]);
    expect(r.ok ? null : r.rejection.code).toBe('merge-drops-values');
  });

  test('puts the value before x:extLst when the extLst comes before the replaced wrapper', () => {
    const s = sheetWith(`<c r="A2"><extLst><ext uri="{0}"/></extLst>${VALUE_ONLY}</c>`);
    expect(s.transact([set('A2', { type: 'number', value: 3 })]).ok).toBe(true);
    expect(reopened(s, 'A2')).toMatch(
      /^<c r="A2"><v>3<\/v><extLst><ext uri="\{0\}"\/><\/extLst><\/c>$/
    );
  });

  test('a range clear on a protected sheet refuses a locked cell with a wrapped value', () => {
    const wrapped =
      `<mc:AlternateContent xmlns:mc="${MC}" xmlns:x14="${X14}"><mc:Choice Requires="x14"><v>5</v></mc:Choice>` +
      '</mc:AlternateContent>';
    const s = storeOf({
      sheets: [
        {
          name: 'S',
          rows: `<row r="1"><c r="A1">${wrapped}</c></row>`,
          beforeMerges: '<sheetProtection sheet="1"/>',
        },
      ],
    });
    const before = s.package;
    const r = s.transact([{ op: 'clearValues', sheet: 'S', ref: 'A1:B2' }]);
    expect(r.ok ? null : r.rejection.code).toBe('locked');
    expect(s.package).toBe(before);
  });

  test('finds a wrapped value after ten other children, for a clear and a locked cell', () => {
    const wrapped =
      `<mc:AlternateContent xmlns:mc="${MC}" xmlns:x14="${X14}"><mc:Choice Requires="x14"><v>5</v></mc:Choice>` +
      '</mc:AlternateContent>';
    const ten = FOREIGN.repeat(10);
    const s = sheetWith(`<c r="A2">${ten}${wrapped}</c>`);
    expect(s.transact([{ op: 'clearValues', sheet: 'S', ref: 'A2' }]).ok).toBe(true);
    expect(reopened(s, 'A2')).not.toContain('<v>5</v>');
    const locked = storeOf({
      sheets: [
        {
          name: 'S',
          rows: `<row r="1"><c r="A1">${ten}${wrapped}</c></row>`,
          beforeMerges: '<sheetProtection sheet="1"/>',
        },
      ],
    });
    const r = locked.transact([{ op: 'clearValues', sheet: 'S', ref: 'A1:B2' }]);
    expect(r.ok ? null : r.rejection.code).toBe('locked');
  });

  test('refuses an element in another namespace that holds value content', () => {
    const s = sheetWith(`<c r="A2"><v>2</v><e:w xmlns:e="${E}"><v>3</v></e:w></c>`);
    const r = s.transact([set('A2', { type: 'number', value: 7 })]);
    expect(r.ok ? null : [r.rejection.code, r.rejection.subject]).toEqual([
      'cell-alternate-content',
      'S!A2',
    ]);
  });

  test('refuses a wrapper with a branch that holds no value, or text', () => {
    for (const branches of [
      '<mc:Choice Requires="x14"/>',
      '<mc:Choice Requires="x14">note</mc:Choice>',
      '<mc:Choice Requires="x14"><v>2</v>note</mc:Choice>',
      '<mc:Choice Requires="x14"><v>2</v></mc:Choice><mc:Fallback/>',
    ]) {
      const s = sheetWith(
        `<c r="A2"><v>2</v><mc:AlternateContent xmlns:mc="${MC}" xmlns:x14="${X14}">${branches}</mc:AlternateContent></c>`
      );
      const r = s.transact([set('A2', { type: 'number', value: 7 })]);
      expect(r.ok ? 'ok' : r.rejection.code).toBe('cell-alternate-content');
    }
  });

  test('keeps a wrapper that holds only white space', () => {
    const s = sheetWith(
      `<c r="A2"><v>2</v><mc:AlternateContent xmlns:mc="${MC}" xmlns:x14="${X14}">\n  </mc:AlternateContent></c>`
    );
    expect(s.transact([set('A2', { type: 'number', value: 7 })]).ok).toBe(true);
    expect(reopened(s, 'A2')).toContain('<v>7</v><mc:AlternateContent');
  });

  test('keeps an empty wrapper', () => {
    const s = sheetWith(`<c r="A2"><v>2</v><mc:AlternateContent xmlns:mc="${MC}"/></c>`);
    expect(s.transact([set('A2', { type: 'number', value: 7 })]).ok).toBe(true);
    expect(reopened(s, 'A2')).toContain('<v>7</v><mc:AlternateContent');
  });

  const refused = (s: WorkbookDocumentStore, ops: WorkbookOp[]) => {
    const before = s.package;
    const r = s.transact(ops);
    expect(s.package).toBe(before);
    return r.ok
      ? null
      : {
          code: r.rejection.code as string,
          subject: r.rejection.subject,
          detail: r.rejection.detail,
        };
  };

  test('refuses an edit, a run, a paste, and a clear of a cell with other alternate content', () => {
    const s = sheetWith(`<c r="A2"><v>2</v>${OTHER}</c><c r="B2"><v>5</v></c>`);
    const reason = {
      code: 'cell-alternate-content',
      subject: 'S!A2',
      detail: 'A2 has the child mc:AlternateContent, which the edit can neither keep nor replace',
    };
    expect(refused(s, [set('A2', { type: 'number', value: 7 })])).toEqual(reason);
    expect(
      refused(s, [set('B2', { type: 'number', value: 8 }), set('A2', { type: 'number', value: 7 })])
    ).toEqual(reason);
    expect(
      refused(s, [
        set('A2', { type: 'number', value: 7 }),
        { op: 'setCellFormatIndex', sheet: 'S', ref: 'A2', index: 0 },
      ])
    ).toEqual(reason);
    expect(refused(s, [{ op: 'clearValues', sheet: 'S', ref: 'A2:B2' }])).toEqual(reason);
    // A cell next to it edits as before.
    expect(s.transact([set('B2', { type: 'number', value: 8 })]).ok).toBe(true);
  });

  test('a clear leaves a cell with other alternate content and no value as it is', () => {
    // The clear reads only a cell that holds a value in some form: it does not refuse a cell
    // that it leaves as it is. A value edit of that cell is still refused.
    const s = sheetWith(`<c r="A2">${OTHER}</c><c r="B2"><v>5</v></c>`);
    const before = cellXml(s.package, 'A2');
    expect(s.transact([{ op: 'clearValues', sheet: 'S', ref: 'A2:B2' }]).ok).toBe(true);
    expect(cellXml(s.package, 'A2')).toBe(before);
    expect(cellXml(s.package, 'B2')).toBeUndefined();
    const r = s.transact([set('A2', { type: 'number', value: 7 })]);
    expect(r.ok ? 'ok' : r.rejection.code).toBe('cell-alternate-content');
  });

  test('refuses an edit and a clear of a formula in alternate content, and keeps it', () => {
    // The calc chain, stale value, array, and shared formula checks read only a direct `x:f`, so
    // a wrapped formula is never replaced: the edit and the clear refuse its cell.
    const formula = (f: string) =>
      `<mc:AlternateContent xmlns:mc="${MC}" xmlns:x14="${X14}"><mc:Choice Requires="x14">${f}<v>2</v></mc:Choice>` +
      '<mc:Fallback><v>2</v></mc:Fallback></mc:AlternateContent>';
    for (const f of [
      '<f>1+1</f>',
      '<f t="array" ref="A2">1+1</f>',
      '<f t="shared" si="0" ref="A2:B2">1+1</f>',
    ]) {
      const s = sheetWith(`<c r="A2">${formula(f)}</c><c r="B2"><v>5</v></c>`);
      const reason = {
        code: 'cell-alternate-content',
        subject: 'S!A2',
        detail: 'A2 has the child mc:AlternateContent, which the edit can neither keep nor replace',
      };
      expect(refused(s, [set('A2', { type: 'number', value: 7 })])).toEqual(reason);
      expect(refused(s, [set('A2', { type: 'empty' })])).toEqual(reason);
      expect(refused(s, [{ op: 'clearValues', sheet: 'S', ref: 'A2:B2' }])).toEqual(reason);
      expect(reopened(s, 'A2')).toContain(f);
    }
  });
});

describe('the nesting limit of the cell child checks', () => {
  /** `n` wrappers, each in the `mc:Choice` of the one before, around `inner`. */
  const nested = (n: number, inner: string): string =>
    n === 0
      ? inner
      : `<mc:AlternateContent xmlns:mc="${MC}" xmlns:x14="${X14}"><mc:Choice Requires="x14">` +
        `${nested(n - 1, inner)}</mc:Choice></mc:AlternateContent>`;
  /** A chain of `n` elements in another namespace, each in the one before, around `inner`. */
  const chain = (n: number, inner = ''): string =>
    n === 0 ? inner : `<e:k xmlns:e="${E}">${chain(n - 1, inner)}</e:k>`;
  const code = (s: WorkbookDocumentStore) => {
    const r = s.transact([set('A2', { type: 'number', value: 7 })]);
    return r.ok ? 'ok' : r.rejection.code;
  };

  test('replaces nine nested value wrappers, and refuses ten', () => {
    const nine = sheetWith(`<c r="A2">${nested(9, '<v>2</v>')}</c>`);
    expect(code(nine)).toBe('ok');
    expect(reopened(nine, 'A2')).toBe('<c r="A2"><v>7</v></c>');
    expect(code(sheetWith(`<c r="A2">${nested(10, '<v>2</v>')}</c>`))).toBe(
      'cell-alternate-content'
    );
  });

  test('keeps a foreign subtree nine elements deep with no value, and refuses ten', () => {
    const nine = sheetWith(`<c r="A2"><v>2</v>${chain(9)}</c>`);
    expect(code(nine)).toBe('ok');
    const xml = reopened(nine, 'A2')!;
    expect(xml).toContain('<v>7</v><e:k');
    expect(xml.split('<e:k').length - 1).toBe(9);
    expect(code(sheetWith(`<c r="A2"><v>2</v>${chain(10)}</c>`))).toBe('cell-alternate-content');
  });
});
