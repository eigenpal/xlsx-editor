// The styles part of a format edit: find it, create it when the workbook has none, and put back
// an earlier revision of it on undo (the `restoreStyles` op).
//
// A new styles part gets its workbook relationship and its content type override, as a new
// shared string table does (`shared-strings.ts`). Undo of the edit that created it removes the
// part, the relationship, and the override again, so one undo restores the package exactly.
// A new part is an empty `x:styleSheet` (`CT_Stylesheet`, sml.xsd): each child of its sequence
// is optional. The relationship and the override follow ECMA-376 Part 2.

import type { Element, XmlPart } from '../package/tree.ts';
import {
  CONTENT_TYPES_PART,
  equivalentPartName,
  relationshipsOf,
  relsPartName,
  xmlPartOf,
  type XlsxPackage,
} from '../package/package.ts';
import { CONTENT_TYPE, NS, REL } from '../package/namespaces.ts';
import { attr } from '../package/query.ts';
import { workbookIndex, type WorkbookIndex } from '../workbook/workbook-index.ts';
import { IdSource, makeElement, plainAttr, smlElement } from './tree-edit.ts';
import { relativeTarget } from './shared-strings.ts';
import { withPart } from './cells.ts';
import { refuse, type OpRejection, type OpResult, type HistoryOp } from './ops.ts';

const STYLES_PART = '/xl/styles.xml';

/**
 * The refusal of a format edit when the workbook has a styles part that the editor cannot read,
 * such as a part past a read limit that stays as bytes. `applyFormat` and `setCellFormatIndex`
 * with an index other than 0 are refused. Format 0 stays allowed, so `clearFormats` and a paste
 * still work. A saved cell then never names a format that the unread part may not hold.
 */
export function unreadStyles(pkg: XlsxPackage, index: WorkbookIndex): OpRejection | undefined {
  if (index.stylesPart === undefined) return undefined;
  const part = xmlPartOf(pkg, index.stylesPart);
  if (part && part.root.kind === 'styleSheet') return undefined;
  return {
    code: 'unsupported-package',
    detail: `${index.stylesPart} is not a styles part this change can read, so a cell can get only the default format 0`,
  };
}

/** The styles part that a format edit starts from: the workbook's, or a new empty one. */
export function stylesOf(
  pkg: XlsxPackage,
  index: WorkbookIndex
): { readonly part: XmlPart; readonly created: boolean } | OpRejection {
  if (index.stylesPart !== undefined) {
    const unread = unreadStyles(pkg, index);
    if (unread) return unread;
    return { part: xmlPartOf(pkg, index.stylesPart)!, created: false };
  }
  let partName = STYLES_PART;
  for (let n = 2; equivalentPartName(pkg, partName); n += 1) partName = `/xl/styles${n}.xml`;
  const ids = new IdSource({ partName, root: null as never, nextId: 1 });
  const root: Element = {
    ...smlElement(ids, 'styleSheet', null),
    nsDecls: [{ prefix: '', uri: NS.x }],
  };
  return { part: { partName, root, nextId: ids.next }, created: true };
}

/** The package with `part` as its styles part, adding the relationship and the override. */
function withStyles(
  pkg: XlsxPackage,
  index: WorkbookIndex,
  part: XmlPart
): XlsxPackage | OpRejection {
  if (index.stylesPart !== undefined)
    return withPart(pkg, {
      kind: 'xml',
      partName: index.stylesPart,
      xml: { ...part, partName: index.stylesPart },
    });
  const relsName = relsPartName(index.mainPart);
  const rels = xmlPartOf(pkg, relsName);
  const ct = xmlPartOf(pkg, CONTENT_TYPES_PART);
  if (!rels || !ct)
    return {
      code: 'unsupported-package',
      detail: 'a new styles part needs the workbook relationships and the content types',
    };
  if (equivalentPartName(pkg, part.partName))
    return { code: 'bad-value', detail: `${part.partName} is already a part of the package` };
  let next = withPart(pkg, { kind: 'xml', partName: part.partName, xml: part });
  const used = new Set(relationshipsOf(next, index.mainPart).map((r) => r.id));
  let id = 'rIdStyles';
  for (let n = 2; used.has(id); n += 1) id = `rIdStyles${n}`;
  const relIds = new IdSource(rels);
  const rel = makeElement(relIds, NS.rel, 'Relationship', null, [
    plainAttr('Id', id),
    plainAttr('Type', REL.styles),
    plainAttr('Target', relativeTarget(index.mainPart, part.partName)),
  ]);
  next = withPart(next, {
    kind: 'xml',
    partName: relsName,
    xml: {
      partName: relsName,
      root: { ...rels.root, children: [...rels.root.children, rel] },
      nextId: relIds.next,
    },
  });
  const ctIds = new IdSource(ct);
  const override = makeElement(ctIds, NS.ct, 'Override', null, [
    plainAttr('PartName', part.partName),
    plainAttr('ContentType', CONTENT_TYPE.styles),
  ]);
  return withPart(next, {
    kind: 'xml',
    partName: CONTENT_TYPES_PART,
    xml: {
      partName: CONTENT_TYPES_PART,
      root: { ...ct.root, children: [...ct.root.children, override] },
      nextId: ctIds.next,
    },
  });
}

/** The package without its styles part, its workbook relationship, and its override. */
function withoutStyles(pkg: XlsxPackage, index: WorkbookIndex, partName: string): XlsxPackage {
  const parts = new Map(pkg.parts);
  parts.delete(partName);
  const relsName = relsPartName(index.mainPart);
  const rels = xmlPartOf(pkg, relsName);
  const ids = new Set(
    relationshipsOf(pkg, index.mainPart)
      .filter((r) => r.type === REL.styles && !r.external && r.target === partName)
      .map((r) => r.id)
  );
  if (rels) {
    const kids = rels.root.children.filter(
      (n) => !(n.type === 'element' && n.local === 'Relationship' && ids.has(attr(n, 'Id') ?? ''))
    );
    parts.set(relsName, {
      kind: 'xml',
      partName: relsName,
      xml: { ...rels, root: { ...rels.root, children: kids } },
    });
  }
  const ct = xmlPartOf(pkg, CONTENT_TYPES_PART);
  if (ct) {
    const fold = partName.toLowerCase();
    const kids = ct.root.children.filter(
      (n) =>
        !(
          n.type === 'element' &&
          n.local === 'Override' &&
          (attr(n, 'PartName') ?? '').toLowerCase() === fold
        )
    );
    parts.set(CONTENT_TYPES_PART, {
      kind: 'xml',
      partName: CONTENT_TYPES_PART,
      xml: { ...ct, root: { ...ct.root, children: kids } },
    });
  }
  return { parts };
}

/**
 * Commit a styles part that a format edit built. Returns the package and the op that puts the
 * styles part back as it was: the old part, or `null` to remove a part this edit created.
 */
export function commitStyles(
  pkg: XlsxPackage,
  index: WorkbookIndex,
  part: XmlPart
): { readonly pkg: XlsxPackage; readonly inverse: HistoryOp } | OpRejection {
  const old = index.stylesPart === undefined ? null : (xmlPartOf(pkg, index.stylesPart) ?? null);
  const next = withStyles(pkg, index, part);
  if ('code' in next) return next;
  return { pkg: next, inverse: { op: 'restoreStyles', part: old } };
}

/**
 * The `restoreStyles` op: put back a styles part of an earlier revision, or remove the styles part
 * that a later edit created (`null`). The inverse of a format edit; its own inverse is the
 * styles part it replaced.
 */
export function restoreStyles(pkg: XlsxPackage, part: XmlPart | null): OpResult {
  const index = workbookIndex(pkg);
  if (!index) return refuse('unsupported-package', 'the package has no workbook part');
  const current =
    index.stylesPart === undefined ? null : (xmlPartOf(pkg, index.stylesPart) ?? null);
  const inverse: HistoryOp = { op: 'restoreStyles', part: current };
  const done = (next: XlsxPackage): OpResult => ({
    ok: true,
    pkg: next,
    inverse: [inverse],
    value: false,
    formula: false,
  });
  if (part === null)
    return done(index.stylesPart === undefined ? pkg : withoutStyles(pkg, index, index.stylesPart));
  if (part.root.ns !== NS.x || part.root.kind !== 'styleSheet')
    return refuse('bad-value', 'restoreStyles needs a styles part');
  if (index.stylesPart !== undefined && !current)
    return refuse('unsupported-package', `${index.stylesPart} is not a styles part`);
  const next = withStyles(pkg, index, part);
  if ('code' in next) return { ok: false, rejection: next };
  return done(next);
}
