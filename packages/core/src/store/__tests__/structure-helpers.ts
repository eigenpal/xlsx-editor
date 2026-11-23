// Helpers of the structural edit tests: open a built workbook in a store, and read cells,
// formulas, and elements of a sheet by name.

import { buildXlsx, type FixtureWorkbook } from '../../../../../scripts/fixtures/builder.ts';
import { readXlsx, xmlPartOf, type XlsxPackage } from '../package/package.ts';
import { attr, child, children, textOf } from '../package/query.ts';
import type { Element } from '../package/tree.ts';
import { parseCellRef } from '../cell-ref.ts';
import { cellKey, sheetIndex } from '../workbook/sheet-index.ts';
import { sheetByName, workbookIndex } from '../workbook/workbook-index.ts';
import { WorkbookDocumentStore } from '../store/document-store.ts';
import type { StructuralOpName, WorkbookOp } from '../store/ops.ts';

export function open(wb: FixtureWorkbook): XlsxPackage {
  const read = readXlsx(buildXlsx(wb));
  if (!read.ok) throw new Error(JSON.stringify(read.error));
  return read.pkg;
}

export const storeOf = (wb: FixtureWorkbook) => new WorkbookDocumentStore(open(wb));

export function sheetRoot(pkg: XlsxPackage, sheet: string): Element {
  const entry = sheetByName(workbookIndex(pkg)!, sheet)!;
  return xmlPartOf(pkg, entry.partName)!.root;
}

export function cell(pkg: XlsxPackage, sheet: string, ref: string): Element | undefined {
  const entry = sheetByName(workbookIndex(pkg)!, sheet)!;
  const a = parseCellRef(ref)!;
  return sheetIndex(xmlPartOf(pkg, entry.partName)!).cells.get(cellKey(a.row, a.col))?.el;
}

/** The formula text of a cell, or undefined. */
export const formulaOf = (pkg: XlsxPackage, sheet: string, ref: string) => {
  const f = child(cell(pkg, sheet, ref), 'x:f');
  return f ? textOf(f) : undefined;
};

/** The value text of a cell, or undefined. */
export const valueOf = (pkg: XlsxPackage, sheet: string, ref: string) =>
  textOf(child(cell(pkg, sheet, ref), 'x:v')) || undefined;

/** The `@ref` of each child of a container element of a sheet. */
export const refsOf = (pkg: XlsxPackage, sheet: string, container: string, item: string) =>
  children(child(sheetRoot(pkg, sheet), container), item).map((el) => attr(el, 'ref'));

export const op = (
  name: StructuralOpName,
  sheet: string,
  start: number,
  count = 1
): WorkbookOp => ({ op: name, sheet, start, count });

/** The defined names of the workbook, by name. */
export function definedNames(pkg: XlsxPackage): Record<string, string> {
  const root = xmlPartOf(pkg, workbookIndex(pkg)!.mainPart)!.root;
  const out: Record<string, string> = {};
  for (const el of children(child(root, 'x:definedNames'), 'x:definedName'))
    out[attr(el, 'name') ?? ''] = textOf(el);
  return out;
}
