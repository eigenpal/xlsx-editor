// An edit that builds an element again keeps the namespace declarations of the old element.
// Markup Compatibility needs each prefix that `mc:Ignorable` names to be declared
// in scope (ECMA-376 Part 3, 7.2), so a rebuilt element that drops a declaration saves a part
// that is not valid. The edit tests edit, save, open the saved bytes again, and read the scope of
// every `mc:Ignorable` value. The other tests show that the reader refuses, the serializer fails,
// and the transaction check refuses a part whose prefix lists name an undeclared prefix.

import { describe, expect, test } from 'bun:test';
import {
  buildXlsx,
  c,
  row,
  type FixtureWorkbook,
} from '../../../../../scripts/fixtures/builder.ts';
import { readXlsx, xmlPartOf, type XlsxPackage } from '../package/package.ts';
import { attr, child } from '../package/query.ts';
import { NS } from '../package/namespaces.ts';
import { parseXmlPart, type Element, type Node } from '../package/tree.ts';
import { saveXlsx } from '../package-write/save.ts';
import { serializeXmlPart } from '../package-write/serialize.ts';
import { cellKey, sheetIndex } from '../workbook/sheet-index.ts';
import { styleLists, workbookIndex } from '../workbook/workbook-index.ts';
import { WorkbookDocumentStore } from '../store/document-store.ts';
import { touchedPartViolation } from '../store/part-checks.ts';
import { applyHistory } from './history-helpers.ts';
import { withNamespacesOf, withParentDeclarations } from '../store/tree-edit.ts';
import type { WorkbookOp } from '../store/ops.ts';

const SHEET = '/xl/worksheets/sheet1.xml';
const STYLES = '/xl/styles.xml';
const EXT = 'urn:example:ignorable';
/** The declarations of the `mc` prefix and of an ignorable prefix `e`. */
const DECLS = `xmlns:mc="${NS.mc}" xmlns:e="${EXT}"`;

function open(wb: FixtureWorkbook): XlsxPackage {
  const read = readXlsx(buildXlsx(wb));
  if (!read.ok) throw new Error(JSON.stringify(read.error));
  return read.pkg;
}

/** The package with one XML part rewritten by a text replacement. */
function rewrite(pkg: XlsxPackage, partName: string, from: string, to: string): XlsxPackage {
  const xml = serializeXmlPart(xmlPartOf(pkg, partName)!);
  expect(xml).toContain(from);
  const parsed = parseXmlPart(partName, xml.replace(from, to));
  if (!parsed.ok) throw new Error(parsed.reason);
  const parts = new Map(pkg.parts);
  parts.set(partName, { kind: 'xml', partName, xml: parsed.part });
  return { parts };
}

/** Commit the ops, save, and open the saved bytes again. */
function editSaveReopen(pkg: XlsxPackage, ops: WorkbookOp[]): XlsxPackage {
  const store = new WorkbookDocumentStore(pkg);
  const r = store.transact(ops);
  if (!r.ok) throw new Error(JSON.stringify(r.rejection));
  const read = readXlsx(saveXlsx(store.package));
  if (!read.ok) throw new Error(JSON.stringify(read.error));
  return read.pkg;
}

/**
 * Each prefix that an `mc:Ignorable` value names and that its scope does not declare, as
 * `element: prefix`. A valid part gives an empty list.
 */
function undeclaredIgnorable(root: Element): string[] {
  const out: string[] = [];
  const walk = (el: Element, parent: ReadonlyMap<string, string>) => {
    const scope = new Map(parent);
    for (const d of el.nsDecls) scope.set(d.prefix, d.uri);
    const value = el.attrs.find((a) => a.ns === NS.mc && a.local === 'Ignorable')?.value;
    for (const prefix of value?.split(/\s+/).filter(Boolean) ?? [])
      if (!scope.has(prefix)) out.push(`${el.local}: ${prefix}`);
    for (const n of el.children) if (n.type === 'element') walk(n, scope);
  };
  walk(root, new Map());
  return out;
}

const ignorable = (el: Element | undefined) =>
  el?.attrs.find((a) => a.ns === NS.mc && a.local === 'Ignorable')?.value;
const declared = (el: Element | undefined, prefix: string) =>
  el?.nsDecls.find((d) => d.prefix === prefix)?.uri;

function cellAt(pkg: XlsxPackage, row0: number, col0: number): Element | undefined {
  return sheetIndex(xmlPartOf(pkg, SHEET)!).cells.get(cellKey(row0, col0))?.el;
}

const setNumber = (ref: string, value: number): WorkbookOp => ({
  op: 'setCellValue',
  sheet: 'S',
  ref,
  value: { type: 'number', value },
});

describe('a rebuilt cell keeps its namespace declarations', () => {
  test('a prefix that only mc:Ignorable names stays declared on the cell', () => {
    const pkg = open({
      sheets: [{ name: 'S', rows: row(1, [`<c r="A1" ${DECLS} mc:Ignorable="e"><v>1</v></c>`]) }],
    });
    const saved = editSaveReopen(pkg, [setNumber('A1', 2)]);
    const a1 = cellAt(saved, 0, 0);
    expect(child(a1, 'x:v')?.children[0]).toMatchObject({ value: '2' });
    expect(ignorable(a1)).toBe('e');
    expect(declared(a1, 'e')).toBe(EXT);
    expect(undeclaredIgnorable(xmlPartOf(saved, SHEET)!.root)).toEqual([]);
    expect(sheetIndex(xmlPartOf(saved, SHEET)!).problems).toEqual([]);
  });

  test('a cell with an attribute in the ignorable namespace keeps it and its declaration', () => {
    const pkg = open({
      sheets: [
        {
          name: 'S',
          rows: row(1, [`<c r="A1" ${DECLS} mc:Ignorable="e" e:x="1"><v>1</v></c>`]),
        },
      ],
    });
    const saved = editSaveReopen(pkg, [setNumber('A1', 2)]);
    const a1 = cellAt(saved, 0, 0)!;
    expect(a1.attrs.find((a) => a.ns === EXT && a.local === 'x')?.value).toBe('1');
    expect(declared(a1, 'e')).toBe(EXT);
    expect(undeclaredIgnorable(xmlPartOf(saved, SHEET)!.root)).toEqual([]);
  });

  test('an edit that removes the last use of the namespace keeps the declaration', () => {
    const pkg = open({
      sheets: [
        {
          name: 'S',
          rows: row(1, [`<c r="A1" ${DECLS} mc:Ignorable="e"><v e:mark="1">1</v></c>`]),
        },
      ],
    });
    const saved = editSaveReopen(pkg, [setNumber('A1', 2)]);
    const a1 = cellAt(saved, 0, 0)!;
    // The value edit replaces `x:v`, the last node that uses the namespace. This test checks
    // the declaration, not that removal.
    const v = child(a1, 'x:v')!;
    expect(v.attrs.some((a) => a.ns === EXT)).toBe(false);
    expect(ignorable(a1)).toBe('e');
    expect(declared(a1, 'e')).toBe(EXT);
    expect(undeclaredIgnorable(xmlPartOf(saved, SHEET)!.root)).toEqual([]);
  });

  test('a prefix declared on the row stays in scope of the rebuilt cell', () => {
    const pkg = open({
      sheets: [
        {
          name: 'S',
          rows: row(1, [`<c r="A1" mc:Ignorable="e"><v>1</v></c>`, c('B1', 5)], DECLS),
        },
      ],
    });
    const saved = editSaveReopen(pkg, [setNumber('A1', 2), setNumber('B1', 6)]);
    const a1 = cellAt(saved, 0, 0)!;
    expect(ignorable(a1)).toBe('e');
    expect(a1.nsDecls).toEqual([]);
    expect(undeclaredIgnorable(xmlPartOf(saved, SHEET)!.root)).toEqual([]);
  });

  test('a cleared cell keeps the declarations when its format keeps it', () => {
    const pkg = open({
      sheets: [
        {
          name: 'S',
          rows: row(1, [`<c r="A1" s="1" ${DECLS} mc:Ignorable="e"><v>1</v></c>`]),
        },
      ],
      styles: { cellXfs: ['<xf numFmtId="2" fontId="0" fillId="0" borderId="0" xfId="0"/>'] },
    });
    const saved = editSaveReopen(pkg, [{ op: 'clearValues', sheet: 'S', ref: 'A1' }]);
    const a1 = cellAt(saved, 0, 0)!;
    expect(child(a1, 'x:v')).toBeUndefined();
    expect(declared(a1, 'e')).toBe(EXT);
    expect(undeclaredIgnorable(xmlPartOf(saved, SHEET)!.root)).toEqual([]);
  });

  test('a cell written with a prefix keeps the prefix', () => {
    const pkg = open({
      sheets: [
        {
          name: 'S',
          rows: row(1, [
            `<x:c r="A1" xmlns:x="${NS.x}" ${DECLS} mc:Ignorable="e"><x:v>1</x:v></x:c>`,
          ]),
        },
      ],
    });
    const saved = editSaveReopen(pkg, [setNumber('A1', 2)]);
    const a1 = cellAt(saved, 0, 0)!;
    expect(a1.prefix).toBe('x');
    expect(declared(a1, 'x')).toBe(NS.x);
    expect(undeclaredIgnorable(xmlPartOf(saved, SHEET)!.root)).toEqual([]);
  });
});

describe('the other rebuilt elements keep their namespace declarations', () => {
  test('a shared formula child that the edit unshares keeps the declarations of x:f', () => {
    const pkg = open({
      sheets: [
        {
          name: 'S',
          rows:
            row(1, [`<c r="A1"><f t="shared" ref="A1:A2" si="0">B1+1</f><v>1</v></c>`]) +
            row(2, [`<c r="A2"><f t="shared" si="0" ${DECLS} mc:Ignorable="e"/><v>1</v></c>`]),
        },
      ],
    });
    const saved = editSaveReopen(pkg, [setNumber('A1', 7)]);
    const f = child(cellAt(saved, 1, 0), 'x:f')!;
    expect(f.children[0]).toMatchObject({ value: 'B2+1' });
    expect(attr(f, 't')).toBeUndefined();
    expect(ignorable(f)).toBe('e');
    expect(declared(f, 'e')).toBe(EXT);
    expect(undeclaredIgnorable(xmlPartOf(saved, SHEET)!.root)).toEqual([]);
  });

  test('a format edit keeps the declarations of the cell format, font, and border it derives from', () => {
    const pkg = open({
      sheets: [{ name: 'S', rows: row(1, [c('A1', 1, { s: 1 })]) }],
      styles: {
        fonts: [`<font ${DECLS} mc:Ignorable="e"><sz val="12"/><name val="Arial"/></font>`],
        borders: [
          `<border ${DECLS} mc:Ignorable="e"><left/><right/><top/><bottom/><diagonal/></border>`,
        ],
        cellXfs: [
          `<xf numFmtId="0" fontId="1" fillId="0" borderId="1" xfId="0" ${DECLS} mc:Ignorable="e"/>`,
        ],
      },
    });
    const saved = editSaveReopen(pkg, [
      {
        op: 'applyFormat',
        sheet: 'S',
        ref: 'A1',
        patch: { bold: true, borders: { edges: 'all' } },
      },
    ]);
    const lists = styleLists(saved, workbookIndex(saved)!);
    const xf = lists.cellXfs.at(-1)!;
    expect(Number(attr(cellAt(saved, 0, 0), 's'))).toBe(lists.cellXfs.length - 1);
    expect(declared(xf, 'e')).toBe(EXT);
    expect(declared(lists.fonts[Number(attr(xf, 'fontId'))], 'e')).toBe(EXT);
    expect(declared(lists.borders[Number(attr(xf, 'borderId'))], 'e')).toBe(EXT);
    expect(undeclaredIgnorable(xmlPartOf(saved, STYLES)!.root)).toEqual([]);
  });

  test('a clear of formats on a protected sheet keeps the declarations of cell format 0', () => {
    const base = open({
      sheets: [
        {
          name: 'S',
          rows: row(1, [c('A1', 1, { s: 1 })]),
          beforeMerges: '<sheetProtection sheet="1" formatCells="0"/>',
        },
      ],
      styles: {
        cellXfs: [
          '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyProtection="1">' +
            '<protection locked="0"/></xf>',
        ],
      },
    });
    const zero = '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>';
    const pkg = rewrite(
      base,
      STYLES,
      zero,
      `<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" ${DECLS} mc:Ignorable="e"/>`
    );
    const saved = editSaveReopen(pkg, [{ op: 'clearFormats', sheet: 'S', ref: 'A1' }]);
    const lists = styleLists(saved, workbookIndex(saved)!);
    const xf = lists.cellXfs[Number(attr(cellAt(saved, 0, 0), 's'))]!;
    expect(child(xf, 'x:protection')).toBeDefined();
    expect(declared(xf, 'e')).toBe(EXT);
    expect(undeclaredIgnorable(xmlPartOf(saved, STYLES)!.root)).toEqual([]);
  });
});

describe('an element that moves to another parent keeps the scope of its prefixes', () => {
  test('a column that a width edit moves out of a later x:cols takes its declarations', () => {
    const pkg = open({
      sheets: [
        {
          name: 'S',
          rows: row(1, [c('A1', 1)]),
          cols:
            '<cols><col min="1" max="1" width="9" customWidth="1"/></cols>' +
            `<cols ${DECLS}><col min="3" max="3" width="9" customWidth="1" mc:Ignorable="e"/></cols>`,
        },
      ],
    });
    const saved = editSaveReopen(pkg, [
      { op: 'setColumnWidth', sheet: 'S', first: 1, last: 1, width: 20 },
    ]);
    const root = xmlPartOf(saved, SHEET)!.root;
    const lists = root.children.filter((n) => n.type === 'element' && n.local === 'cols');
    expect(lists).toHaveLength(1);
    const moved = (lists[0] as Element).children[1] as Element;
    expect(attr(moved, 'min')).toBe('3');
    expect(ignorable(moved)).toBe('e');
    expect(declared(moved, 'e')).toBe(EXT);
    expect(undeclaredIgnorable(root)).toEqual([]);
  });

  test('a moved column keeps a prefix of the root that the first x:cols declares again', () => {
    const OTHER = 'urn:example:other';
    const built = open({
      sheets: [
        {
          name: 'S',
          rows: row(1, [c('A1', 1)]),
          cols:
            `<cols xmlns:e="${OTHER}"><col min="1" max="1" width="9" customWidth="1"/></cols>` +
            `<cols ${DECLS}><col min="3" max="3" width="9" customWidth="1" mc:Ignorable="e" e:x="1"/></cols>`,
        },
      ],
    });
    // The fixture builder declares no extra prefix on the worksheet root: move the declarations
    // of `mc` and `e` from the later list to the root.
    const xml = serializeXmlPart(xmlPartOf(built, SHEET)!);
    const later = xml.match(/<cols xmlns:mc="[^"]*" xmlns:e="[^"]*">/)![0];
    const text = xml.replace(later, '<cols>').replace('<worksheet ', `<worksheet ${DECLS} `);
    const reparsed = parseXmlPart(SHEET, text);
    if (!reparsed.ok) throw new Error(reparsed.reason);
    const pkg = {
      parts: new Map(built.parts).set(SHEET, { kind: 'xml', partName: SHEET, xml: reparsed.part }),
    } as XlsxPackage;

    const saved = editSaveReopen(pkg, [
      { op: 'setColumnWidth', sheet: 'S', first: 1, last: 1, width: 20 },
    ]);
    const root = xmlPartOf(saved, SHEET)!.root;
    const list = root.children.find(
      (n): n is Element => n.type === 'element' && n.local === 'cols'
    );
    const moved = list!.children[1] as Element;
    expect(attr(moved, 'min')).toBe('3');
    expect(ignorable(moved)).toBe('e');
    expect(declared(moved, 'e')).toBe(EXT);
    expect(moved.attrs.find((a) => a.local === 'x')?.ns).toBe(EXT);
  });

  test('a column width edit is refused when a later x:cols has an attribute', () => {
    const pkg = open({
      sheets: [
        {
          name: 'S',
          rows: row(1, [c('A1', 1)]),
          cols:
            '<cols><col min="1" max="1" width="9" customWidth="1"/></cols>' +
            `<cols ${DECLS} mc:Ignorable="e"><col min="3" max="3" width="9" e:x="1"/></cols>`,
        },
      ],
    });
    const store = new WorkbookDocumentStore(pkg);
    const r = store.transact([{ op: 'setColumnWidth', sheet: 'S', first: 1, last: 1, width: 20 }]);
    expect(r.ok ? null : r.rejection).toMatchObject({
      code: 'invalid-sheet',
      detail:
        'a column width edit cannot keep the attribute mc:Ignorable of an x:cols after the first',
    });
    expect(store.package).toBe(pkg);
  });
});

describe('a workbook whose worksheet names an undeclared prefix', () => {
  test('is refused on read', () => {
    const bytes = buildXlsx({
      sheets: [{ name: 'S', rows: row(1, [`<c r="A1" xmlns:mc="${NS.mc}" mc:Ignorable="e"/>`]) }],
    });
    const read = readXlsx(bytes);
    expect(read.ok ? null : read.error).toMatchObject({ reason: 'xml', detail: 'unbound-prefix' });
  });

  test('keeps an optional part with an undeclared prefix as bytes, and saves those bytes', () => {
    const pkg = open({ sheets: [{ name: 'S', rows: row(1, [c('A1', 1)]) }] });
    const styles = serializeXmlPart(xmlPartOf(pkg, STYLES)!);
    const bad = styles.replace('<styleSheet ', `<styleSheet xmlns:mc="${NS.mc}" mc:Ignorable="e" `);
    expect(bad).not.toBe(styles);
    const zipped = saveXlsx({
      parts: new Map(pkg.parts).set(STYLES, {
        kind: 'binary',
        partName: STYLES,
        bytes: new TextEncoder().encode(bad),
      }),
    });
    const read = readXlsx(zipped);
    if (!read.ok) throw new Error(JSON.stringify(read.error));
    const part = read.pkg.parts.get(STYLES);
    expect(part?.kind).toBe('binary');
    const again = readXlsx(saveXlsx(read.pkg));
    if (!again.ok) throw new Error(JSON.stringify(again.error));
    const kept = again.pkg.parts.get(STYLES);
    expect(kept?.kind === 'binary' ? new TextDecoder().decode(kept.bytes) : null).toBe(bad);
  });
});

/** The root of a part parsed from `xml`, or the reason the reader refused it. */
function parsed(xml: string): Element | string {
  const r = parseXmlPart('/p.xml', xml);
  return r.ok ? r.part.root : r.reason;
}

/** The root of a part parsed from `xml`. Throws with the reason when the reader refuses it. */
function accepted(xml: string): Element {
  const r = parseXmlPart('/p.xml', xml);
  if (!r.ok) throw new Error(`${r.reason}: ${r.detail ?? ''}`);
  return r.part.root;
}

describe('prefixes named in attribute values', () => {
  const MC = `xmlns:mc="${NS.mc}"`;
  const XSI = `xmlns:xsi="${NS.xsi}"`;

  test('the reader refuses a prefix list that names an undeclared prefix', () => {
    expect(parsed(`<a ${MC} mc:Ignorable="e"/>`)).toBe('unbound-prefix');
    expect(parseXmlPart('/p.xml', `<a ${MC} mc:Ignorable="e"/>`)).toMatchObject({
      detail: 'a/@mc:Ignorable: the prefix e is not declared',
    });
    expect(parsed(`<a ${MC} mc:MustUnderstand="e"/>`)).toBe('unbound-prefix');
    expect(parsed(`<mc:Choice ${MC} Requires="e"/>`)).toBe('unbound-prefix');
    expect(parsed(`<a ${MC} mc:Ignorable="a:b"/>`)).toBe('bad-name');
    // `Requires` is a prefix list on `mc:Choice` only.
    accepted(`<a Requires="e"/>`);
  });

  test('the reader accepts a prefix that the element or an ancestor declares', () => {
    accepted(`<a ${DECLS} mc:Ignorable="e"/>`);
    accepted(`<a ${DECLS}><b mc:Ignorable=" e  mc "/></a>`);
  });

  test('the reader checks the prefix of each qualified name, and allows prefix:*', () => {
    accepted(`<a ${DECLS} mc:ProcessContent="e:*  e:x"/>`);
    expect(parsed(`<a ${DECLS} mc:PreserveElements="f:x"/>`)).toBe('unbound-prefix');
    expect(parsed(`<a ${DECLS} mc:PreserveAttributes="*"/>`)).toBe('bad-name');
    expect(parsed(`<a ${DECLS} mc:ProcessContent="e:x:y"/>`)).toBe('bad-name');
    accepted(`<a ${XSI} xsi:type="T"/>`);
    expect(parsed(`<a ${XSI} xsi:type="e:T"/>`)).toBe('unbound-prefix');
  });

  test('the serializer fails a part that lost the declaration of an ignorable prefix', () => {
    const root = accepted(`<a ${MC}><b xmlns:e="${EXT}" mc:Ignorable="e"/></a>`);
    const b = root.children[0] as Element;
    const lost = { ...root, children: [{ ...b, nsDecls: [] }] };
    expect(() => serializeXmlPart({ partName: '/p.xml', root: lost, nextId: 9 })).toThrow(
      'cannot serialize attribute mc:Ignorable of b: the prefix e is not declared'
    );
    expect(serializeXmlPart({ partName: '/p.xml', root, nextId: 9 })).toContain('xmlns:e=');
  });

  test('the serializer checks each prefix list and qualified name the reader checks', () => {
    const lose = (xml: string) => {
      const root = accepted(xml);
      const b = root.children[0] as Element;
      const part = { partName: '/p.xml', root: { ...root, children: [{ ...b, nsDecls: [] }] } };
      return () => serializeXmlPart({ ...part, nextId: 9 });
    };
    const E = `xmlns:e="${EXT}"`;
    expect(lose(`<a ${MC}><mc:Choice ${E} Requires="e"/></a>`)).toThrow('prefix e');
    expect(lose(`<a ${MC}><b ${E} mc:MustUnderstand="e"/></a>`)).toThrow('prefix e');
    expect(lose(`<a ${MC}><b ${E} mc:ProcessContent="e:*"/></a>`)).toThrow('prefix e');
    expect(lose(`<a ${XSI}><b ${E} xsi:type="e:T"/></a>`)).toThrow('prefix e');
  });
});

describe('withNamespacesOf', () => {
  const root = accepted(`<x:a xmlns:x="${NS.x}" ${DECLS}><x:b/><e:c/></x:a>`);
  const [b, c] = root.children as Element[];

  test('copies the declarations and the prefix of an element in the same namespace', () => {
    const out = withNamespacesOf(root, b!);
    expect(out.nsDecls).toBe(root.nsDecls);
    expect(out.prefix).toBe('x');
    expect(out.id).toBe(b!.id);
  });

  test('copies the declarations but not the prefix of an element in another namespace', () => {
    const out = withNamespacesOf(root, c!);
    expect(out.nsDecls).toBe(root.nsDecls);
    expect(out.prefix).toBe('e');
  });

  test('returns the element as it is with no source, or a source with nothing to copy', () => {
    expect(withNamespacesOf(undefined, b!)).toBe(b);
    expect(withNamespacesOf(null, b!)).toBe(b);
    const plain = accepted('<a/>');
    expect(withNamespacesOf(plain, b!)).toBe(b);
  });
});

describe('withParentDeclarations', () => {
  test('adds the declarations of the parent, and a declaration of the element wins', () => {
    const root = accepted(`<a ${DECLS}><b xmlns:e="urn:own"/><c/></a>`);
    const [b, c] = root.children as Element[];
    const own = withParentDeclarations(b!, root);
    expect(own.nsDecls.filter((d) => d.prefix === 'e').map((d) => d.uri)).toEqual(['urn:own']);
    expect(declared(own, 'mc')).toBe(NS.mc);
    expect(declared(withParentDeclarations(c!, root), 'e')).toBe(EXT);
    expect(withParentDeclarations(c!, accepted('<a/>'))).toBe(c);
  });
});

describe('the transaction check refuses an undeclared prefix in touched content', () => {
  // No op drops a declaration, so each test builds the result such an op would leave.
  /** The package with the root of one XML part replaced. */
  const withRoot = (pkg: XlsxPackage, partName: string, root: Element): XlsxPackage => {
    const xml = { ...xmlPartOf(pkg, partName)!, root };
    return { parts: new Map(pkg.parts).set(partName, { kind: 'xml', partName, xml }) };
  };
  /** `root` with its child `from` replaced by `to`. */
  const swap = (root: Element, from: Node, to: Node): Element => ({
    ...root,
    children: root.children.map((n) => (n === from ? to : n)),
  });
  const IGNORE_E = { ns: NS.mc, local: 'Ignorable', prefix: 'mc', value: 'e' };

  test('a changed row whose cell lost the declaration of an ignorable prefix', () => {
    const pkg = open({
      sheets: [{ name: 'S', rows: row(1, [`<c r="A1" ${DECLS} mc:Ignorable="e"><v>1</v></c>`]) }],
    });
    const root = xmlPartOf(pkg, SHEET)!.root;
    const sheetData = child(root, 'x:sheetData')!;
    const r1 = sheetData.children[0] as Element;
    const lost = { ...r1, children: [{ ...(r1.children[0] as Element), nsDecls: [] }] };
    const after = withRoot(pkg, SHEET, swap(root, sheetData, { ...sheetData, children: [lost] }));
    expect(touchedPartViolation(pkg, after, [SHEET])).toBe(
      `${SHEET}: c/@mc:Ignorable: the prefix e is not declared`
    );
    expect(touchedPartViolation(pkg, pkg, [SHEET])).toBeUndefined();
    // The store maps the violation to `invalid-part` and commits nothing.
    const r = applyHistory(pkg, [{ op: 'restoreRow', sheet: 'S', row: 1, node: lost }]);
    expect(r.ok ? null : r.rejection.code).toBe('invalid-part');
  });

  test('a changed child of the worksheet root, such as x:cols', () => {
    const pkg = open({
      sheets: [
        {
          name: 'S',
          rows: row(1, [c('A1', 1)]),
          cols: '<cols><col min="1" max="1" width="9" customWidth="1"/></cols>',
        },
      ],
    });
    const root = xmlPartOf(pkg, SHEET)!.root;
    const cols = child(root, 'x:cols')!;
    const col = cols.children[0] as Element;
    const marked = { ...cols, children: [{ ...col, attrs: [...col.attrs, IGNORE_E] }] };
    const after = withRoot(pkg, SHEET, swap(root, cols, marked));
    expect(touchedPartViolation(pkg, after, [SHEET])).toBe(
      `${SHEET}: col/@mc:Ignorable: the prefix e is not declared`
    );
  });

  test('a new cell format record', () => {
    const pkg = open({ sheets: [{ name: 'S', rows: row(1, [c('A1', 1)]) }] });
    const root = xmlPartOf(pkg, STYLES)!.root;
    const cellXfs = child(root, 'x:cellXfs')!;
    const xf = cellXfs.children.find((n): n is Element => n.type === 'element')!;
    const added = { ...xf, attrs: [...xf.attrs, IGNORE_E] };
    const grown = { ...cellXfs, children: [...cellXfs.children, added] };
    const after = withRoot(pkg, STYLES, swap(root, cellXfs, grown));
    expect(touchedPartViolation(pkg, after, [STYLES])).toBe(
      `${STYLES}: xf/@mc:Ignorable: the prefix e is not declared`
    );
  });

  test('a changed cell format list with an attribute of its own', () => {
    const pkg = open({ sheets: [{ name: 'S', rows: row(1, [c('A1', 1)]) }] });
    const root = xmlPartOf(pkg, STYLES)!.root;
    const cellXfs = child(root, 'x:cellXfs')!;
    const marked = { ...cellXfs, attrs: [...cellXfs.attrs, IGNORE_E] };
    const after = withRoot(pkg, STYLES, swap(root, cellXfs, marked));
    expect(touchedPartViolation(pkg, after, [STYLES])).toBe(
      `${STYLES}: cellXfs/@mc:Ignorable: the prefix e is not declared`
    );
  });

  test('a new shared string item', () => {
    const pkg = open({
      sheets: [{ name: 'S', rows: row(1, [c('A1', 0, { t: 's' })]) }],
      strings: ['text'],
    });
    const sst = workbookIndex(pkg)!.sharedStringsPart!;
    const root = xmlPartOf(pkg, sst)!.root;
    const si = root.children.find((n): n is Element => n.type === 'element')!;
    const grown = { ...root, children: [...root.children, { ...si, attrs: [IGNORE_E] }] };
    expect(touchedPartViolation(pkg, withRoot(pkg, sst, grown), [sst])).toBe(
      `${sst}: si/@mc:Ignorable: the prefix e is not declared`
    );
  });
});
