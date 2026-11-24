// The headings op: `setShowHeadings` writes `x:sheetView/@showRowColHeaders` of the first sheet
// view (`CT_SheetView`, ECMA-376 Part 1, 18.3.1.87). A hide writes `0`, a show removes the
// attribute (default true), and a sheet with no `x:sheetViews` gets one in its schema place in
// `CT_Worksheet` (18.3.1.99). One undo gives back the worksheet part byte for byte, and a saved
// and read package keeps the flag.

import { describe, expect, test } from 'bun:test';
import { c, row, type FixtureSheet } from '../../../../../scripts/fixtures/builder.ts';
import { attr, child, children } from '../package/query.ts';
import { readXlsx, xmlPartOf, type XlsxPackage } from '../package/package.ts';
import { serializeXmlPart } from '../package-write/serialize.ts';
import { saveXlsx } from '../package-write/save.ts';
import { sheetByName, workbookIndex } from '../workbook/workbook-index.ts';
import { WorkbookDocumentStore } from '../store/document-store.ts';
import { withPart } from '../store/cells.ts';
import { IdSource, makeElement, smlElement } from '../store/tree-edit.ts';
import { NS } from '../package/namespaces.ts';
import type { Node } from '../package/tree.ts';
import type { WorkbookOp } from '../store/ops.ts';
import { firstSheetView, viewShowsHeadings } from '../store/sheet-view.ts';
import { open, sheetRoot } from './structure-helpers.ts';

const partName = (pkg: XlsxPackage) => sheetByName(workbookIndex(pkg)!, 'S')!.partName;
const partText = (pkg: XlsxPackage) => serializeXmlPart(xmlPartOf(pkg, partName(pkg))!);
const views = (pkg: XlsxPackage) =>
  children(child(sheetRoot(pkg, 'S'), 'x:sheetViews'), 'x:sheetView');
const flag = (pkg: XlsxPackage) => attr(views(pkg)[0], 'showRowColHeaders');

/** The local names of the worksheet children, in order. */
const order = (pkg: XlsxPackage) =>
  sheetRoot(pkg, 'S').children.flatMap((n) => (n.type === 'element' ? [n.local] : []));

const pkgOf = (sheet: Partial<FixtureSheet> = {}) =>
  open({ sheets: [{ name: 'S', rows: row(1, [c('A1', 1)]), ...sheet }] });

/** The package with the `x:sheetViews` of sheet `S` removed. */
function withoutViews(pkg: XlsxPackage): XlsxPackage {
  const part = xmlPartOf(pkg, partName(pkg))!;
  const root = {
    ...part.root,
    children: part.root.children.filter((n) => n.type !== 'element' || n.local !== 'sheetViews'),
  };
  return withPart(pkg, { kind: 'xml', partName: part.partName, xml: { ...part, root } });
}

/** The package with the children of the worksheet root of sheet `S` replaced by `f`. */
function withRoot(pkg: XlsxPackage, f: (kids: readonly Node[], ids: IdSource) => Node[]) {
  const part = xmlPartOf(pkg, partName(pkg))!;
  const ids = new IdSource(part);
  const root = { ...part.root, children: f(part.root.children, ids) };
  return withPart(pkg, {
    kind: 'xml',
    partName: part.partName,
    xml: { ...part, root, nextId: ids.next },
  });
}

/** The package with the `x:sheetViews` of sheet `S` holding `kids` in place of its views. */
const viewsHolding = (pkg: XlsxPackage, kids: (ids: IdSource) => Node[]) =>
  withRoot(pkg, (nodes, ids) =>
    nodes.map((n) =>
      n.type === 'element' && n.local === 'sheetViews' ? { ...n, children: kids(ids) } : n
    )
  );

const headings = (show: boolean): WorkbookOp => ({ op: 'setShowHeadings', sheet: 'S', show });

/** Commit, then undo gives back the part byte for byte, and redo the edit. */
function commit(s: WorkbookDocumentStore, op: WorkbookOp): XlsxPackage {
  const before = partText(s.package);
  const r = s.transact([op]);
  if (!r.ok) throw new Error(JSON.stringify(r.rejection));
  const after = partText(s.package);
  expect(s.undo().ok).toBe(true);
  expect(partText(s.package)).toBe(before);
  expect(s.redo().ok).toBe(true);
  expect(partText(s.package)).toBe(after);
  return s.package;
}

describe('setShowHeadings', () => {
  test('a hide writes showRowColHeaders="0" on the first sheet view', () => {
    const s = new WorkbookDocumentStore(pkgOf({ headings: true }));
    const pkg = commit(s, headings(false));
    expect(flag(pkg)).toBe('0');
    expect(viewShowsHeadings(firstSheetView(sheetRoot(pkg, 'S')))).toBe(false);
    // The other attributes of the view stay.
    expect(attr(views(pkg)[0], 'workbookViewId')).toBe('0');
    expect(attr(views(pkg)[0], 'showGridLines')).toBe('0');
  });

  test('a show removes the attribute, whose default is true', () => {
    const s = new WorkbookDocumentStore(pkgOf());
    expect(flag(s.package)).toBe('0');
    const pkg = commit(s, headings(true));
    expect(flag(pkg)).toBeUndefined();
    expect(viewShowsHeadings(firstSheetView(sheetRoot(pkg, 'S')))).toBe(true);
  });

  test('only the first sheet view changes', () => {
    const base = pkgOf({ headings: true });
    const part = xmlPartOf(base, partName(base))!;
    const list = child(part.root, 'x:sheetViews')!;
    const second = {
      ...views(base)[0]!,
      id: 'second',
      attrs: [{ ns: '', local: 'workbookViewId', value: '1' }],
    };
    const two = { ...list, children: [...list.children, second] };
    const root = { ...part.root, children: part.root.children.map((n) => (n === list ? two : n)) };
    const s = new WorkbookDocumentStore(
      withPart(base, { kind: 'xml', partName: part.partName, xml: { ...part, root } })
    );
    const pkg = commit(s, headings(false));
    expect(views(pkg).map((v) => attr(v, 'showRowColHeaders'))).toEqual(['0', undefined]);
  });

  test('a sheet with no x:sheetViews gets one after x:dimension, before x:sheetFormatPr', () => {
    const s = new WorkbookDocumentStore(withoutViews(pkgOf({ dimension: 'A1', sheetPr: '' })));
    expect(order(s.package)).not.toContain('sheetViews');
    const pkg = commit(s, headings(false));
    const names = order(pkg);
    expect(names.indexOf('sheetViews')).toBe(names.indexOf('dimension') + 1);
    expect(names.indexOf('sheetViews')).toBe(names.indexOf('sheetFormatPr') - 1);
    expect(views(pkg).map((v) => v.attrs.map((a) => `${a.local}=${a.value}`))).toEqual([
      ['showRowColHeaders=0', 'workbookViewId=0'],
    ]);
  });

  test('an x:sheetViews with no x:sheetView gets one as its first child, before x:extLst', () => {
    const cases: [(ids: IdSource) => Node[], string[]][] = [
      [() => [], ['sheetView']],
      [(ids) => [smlElement(ids, 'extLst', 'sheetViews')], ['sheetView', 'extLst']],
    ];
    for (const [kids, expected] of cases) {
      const s = new WorkbookDocumentStore(viewsHolding(pkgOf(), kids));
      const pkg = commit(s, headings(false));
      const list = child(sheetRoot(pkg, 'S'), 'x:sheetViews')!;
      const names = list.children.flatMap((n) => (n.type === 'element' ? [n.local] : []));
      expect(names).toEqual(expected);
      expect(flag(pkg)).toBe('0');
    }
  });

  test('a hide with no x:sheetViews and markup compatibility content before x:sheetData is refused', () => {
    const base = withoutViews(pkgOf({ dimension: 'A1' }));
    const wrapped = withRoot(base, (nodes, ids) => [
      makeElement(ids, NS.mc, 'AlternateContent', 'worksheet'),
      ...nodes,
    ]);
    const s = new WorkbookDocumentStore(wrapped);
    const before = s.package;
    const r = s.transact([headings(false)]);
    expect(r.ok ? null : r.rejection.code).toBe('unsupported-package');
    expect(s.package).toBe(before);
    // With its own x:sheetViews, the sheet writes the flag there.
    const kept = new WorkbookDocumentStore(
      withRoot(pkgOf(), (nodes, ids) => [
        makeElement(ids, NS.mc, 'AlternateContent', 'worksheet'),
        ...nodes,
      ])
    );
    expect(flag(commit(kept, headings(true)))).toBeUndefined();
  });

  test('a sheet with no x:sheetViews shows its headings, so a show changes nothing', () => {
    const s = new WorkbookDocumentStore(withoutViews(pkgOf()));
    const before = s.package;
    expect(s.transact([headings(true)]).ok).toBe(true);
    expect(s.package).toBe(before);
    expect(s.canUndo).toBe(false);
  });

  test('a flag that has the asked state changes nothing and makes no undo step', () => {
    const s = new WorkbookDocumentStore(pkgOf());
    const before = s.package;
    expect(s.transact([headings(false)]).ok).toBe(true);
    expect(s.package).toBe(before);
    expect(s.revision).toBe(0);
    expect(s.canUndo).toBe(false);
  });

  test('an unknown sheet is refused with unknown-sheet', () => {
    const s = new WorkbookDocumentStore(pkgOf());
    const r = s.transact([{ op: 'setShowHeadings', sheet: 'Nope', show: true }]);
    expect(r.ok ? null : r.rejection.code).toBe('unknown-sheet');
  });

  test('sheet protection and workbook protection do not refuse it', () => {
    const s = new WorkbookDocumentStore(
      open({
        sheets: [
          { name: 'S', rows: row(1, [c('A1', 1)]), beforeMerges: '<sheetProtection sheet="1"/>' },
        ],
        workbookProtection: 'lockStructure="1" lockWindows="1"',
      })
    );
    const pkg = commit(s, headings(true));
    expect(flag(pkg)).toBeUndefined();
  });

  test('a saved and read package keeps the flag', () => {
    const s = new WorkbookDocumentStore(pkgOf({ headings: true }));
    expect(s.transact([headings(false)]).ok).toBe(true);
    const read = readXlsx(saveXlsx(s.package));
    if (!read.ok) throw new Error(JSON.stringify(read.error));
    expect(flag(read.pkg)).toBe('0');
  });
});
