// Save a workbook: the save rules of design D2, then the package writer.
//
// The rules: `x:sst/@count` and `@uniqueCount` (`CT_Sst`, sml.xsd, both optional
// `xsd:unsignedInt`), and the `count` attribute of each collection of the styles part (such as
// `CT_CellXfs` and `CT_Fonts`), are written from the content that is saved. A file states them as
// hints, and an edit that appends a shared string or a cell format makes them stale. A part whose
// counts are already right is not touched, so its fingerprint does not change.

import type { Element, XmlPart } from '../package/tree.ts';
import {
  canonicalPackageFingerprint,
  packageFingerprints,
  readXlsx,
  xmlPartOf,
  type PackagePart,
  type XlsxPackage,
} from '../package/package.ts';
import { writeXlsx, writeXlsxSteps } from './write.ts';
import { attr, child } from '../package/query.ts';
import { withAttr } from '../store/tree-edit.ts';
import { workbookIndex } from '../workbook/workbook-index.ts';

/** The element with `count` set to `value`, when it states another count. */
function counted(el: Element, name: string, value: number): Element {
  const raw = attr(el, name);
  return raw === undefined || Number(raw) === value ? el : withAttr(el, name, String(value));
}

function elementChildren(el: Element): number {
  let n = 0;
  for (const c of el.children) if (c.type === 'element' && c.local !== 'extLst') n += 1;
  return n;
}

function replaced(part: XmlPart, root: Element): PackagePart {
  return { kind: 'xml', partName: part.partName, xml: { ...part, root } };
}

/** The package with the save rules applied. The same package object when nothing changes. */
export function applySaveRules(pkg: XlsxPackage): XlsxPackage {
  const index = workbookIndex(pkg);
  if (!index) return pkg;
  const updates: PackagePart[] = [];
  const sst = xmlPartOf(pkg, index.sharedStringsPart);
  if (sst && sst.root.kind === 'sst') {
    let references = 0;
    // Every `t="s"` cell element counts, also one that the index leaves out for a problem.
    for (const sheet of index.sheets) {
      const part = xmlPartOf(pkg, sheet.partName);
      const sheetData =
        part && sheet.kind === 'worksheet' ? child(part.root, 'x:sheetData') : undefined;
      for (const row of sheetData?.children ?? []) {
        if (row.type !== 'element' || row.kind !== 'row') continue;
        for (const cell of row.children)
          if (cell.type === 'element' && cell.kind === 'cell' && attr(cell, 't') === 's')
            references += 1;
      }
    }
    let root = counted(sst.root, 'count', references);
    root = counted(root, 'uniqueCount', elementChildren(sst.root));
    if (root !== sst.root) updates.push(replaced(sst, root));
  }
  const styles = xmlPartOf(pkg, index.stylesPart);
  if (styles && styles.root.kind === 'styleSheet') {
    let changed = false;
    const children = styles.root.children.map((node) => {
      if (node.type !== 'element') return node;
      const next = counted(node, 'count', elementChildren(node));
      if (next !== node) changed = true;
      return next;
    });
    if (changed) updates.push(replaced(styles, { ...styles.root, children }));
  }
  if (updates.length === 0) return pkg;
  const parts = new Map(pkg.parts);
  for (const part of updates) parts.set(part.partName, part);
  return { parts };
}

/** Save a workbook to `.xlsx` bytes, with the save rules applied. */
export function saveXlsx(pkg: XlsxPackage): Uint8Array<ArrayBuffer> {
  return writeXlsx(applySaveRules(pkg));
}

/** {@link saveXlsx}, yielding between units of work. The bytes are the same either way. */
export function saveXlsxSteps(pkg: XlsxPackage): Generator<void, Uint8Array<ArrayBuffer>> {
  return writeXlsxSteps(applySaveRules(pkg));
}

/**
 * The fidelity oracles. The canonical fingerprint is what a package holds after the
 * save rules; the semantic digest is the fingerprint of the package that a save reopens to.
 * Equal values prove that save loses nothing the model carries.
 */
export function workbookOracles(
  pkg: XlsxPackage,
  write: (pkg: XlsxPackage) => Uint8Array = saveXlsx
): { fingerprint: string; digest: string; changedParts: readonly string[] } {
  const ruled = applySaveRules(pkg);
  const reopened = readXlsx(write(pkg));
  if (!reopened.ok)
    throw new Error(`save produced an unreadable package: ${reopened.error.reason}`);
  const before = packageFingerprints(ruled);
  const after = packageFingerprints(reopened.pkg);
  const changedParts = [...new Set([...before.keys(), ...after.keys()])].filter(
    (name) => before.get(name) !== after.get(name)
  );
  return {
    fingerprint: canonicalPackageFingerprint(ruled),
    digest: canonicalPackageFingerprint(reopened.pkg),
    changedParts,
  };
}
