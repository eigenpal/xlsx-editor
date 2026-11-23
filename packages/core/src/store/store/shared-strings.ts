// The shared string table for the store ops (design D2): reuse an item with the same text, or
// append one; create the table, its relationship, and its content type when the workbook has
// none. Existing items are never removed or reordered.
//
// The table is `x:sst` (`CT_Sst`, sml.xsd), a sequence of `x:si` items (`CT_Rst`). An item holds
// one `x:t` (`ST_Xstring`), or `x:r` runs (`CT_RElt`), and optional phonetic runs (`x:rPh`).

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
import { workbookIndex, type WorkbookIndex } from '../workbook/workbook-index.ts';
import { encodeXstring } from '../workbook/xstring.ts';
import { IdSource, makeElement, plainAttr, smlElement, textNode, withChild } from './tree-edit.ts';
import { withPart } from './cells.ts';
import type { HistoryOp, OpRejection } from './ops.ts';

const SST_PART = '/xl/sharedStrings.xml';

/** The text of a plain `x:si` (one `x:t`, no runs), or undefined for a rich or phonetic item. */
export function plainText(si: Element): string | undefined {
  if (si.children.length !== 1) return undefined;
  const t = si.children[0]!;
  if (t.type !== 'element' || t.kind !== 'stringText') return undefined;
  let out = '';
  for (const n of t.children) {
    if (n.type !== 'text') return undefined;
    out += n.value;
  }
  return out;
}

const plainIndex = new WeakMap<XmlPart, Map<string, number>>();

/**
 * The index of a shared string with this text, appending one when none exists. A shared
 * string table that does not exist is created with its relationship and content type.
 * Existing items are never removed or reordered: saving never compacts the table (design D2).
 */
export function sharedStringIndex(
  pkg: XlsxPackage,
  index: WorkbookIndex,
  text: string
): { pkg: XlsxPackage; index: number } | OpRejection {
  const r = sharedStringIndices(pkg, index, [text]);
  return 'code' in r ? r : { pkg: r.pkg, index: r.indices.get(text)! };
}

/**
 * The index of a shared string for each text, as {@link sharedStringIndex} gives it, with every
 * missing text appended in one rebuild of the table, in the order of `texts`. A batch of cell
 * writes costs the table once, not once per new string.
 */
export function sharedStringIndices(
  pkg: XlsxPackage,
  index: WorkbookIndex,
  texts: readonly string[]
): { pkg: XlsxPackage; indices: ReadonlyMap<string, number> } | OpRejection {
  let next = pkg;
  let partName = index.sharedStringsPart;
  let part = xmlPartOf(pkg, partName);
  if (partName !== undefined && (!part || part.root.kind !== 'sst'))
    return {
      code: 'unsupported-package',
      detail: `${partName} is not a shared string table this change can read, so it cannot add a string`,
    };
  if (!part) {
    const created = createSharedStrings(pkg, index);
    if ('code' in created) return created;
    next = created.pkg;
    partName = created.partName;
    part = xmlPartOf(next, partName)!;
  }
  let map = plainIndex.get(part);
  if (!map) {
    map = new Map();
    let i = 0;
    for (const item of part.root.children) {
      if (item.type !== 'element' || item.kind !== 'stringItem') continue;
      const plain = plainText(item);
      if (plain !== undefined && !map.has(plain)) map.set(plain, i);
      i += 1;
    }
    plainIndex.set(part, map);
  }
  const indices = new Map<string, number>();
  const ids = new IdSource(part);
  const added: Element[] = [];
  let newMap: Map<string, number> | undefined;
  let count = -1;
  for (const text of texts) {
    if (indices.has(text)) continue;
    const encoded = encodeXstring(text);
    const hit = (newMap ?? map).get(encoded);
    if (hit !== undefined) {
      indices.set(text, hit);
      continue;
    }
    if (!newMap) {
      newMap = new Map(map);
      count = countItems(part.root);
    }
    const preserve = encoded !== encoded.trim();
    const t = smlElement(
      ids,
      't',
      'stringItem',
      preserve ? [{ ns: NS.xml, local: 'space', prefix: 'xml', value: 'preserve' }] : [],
      [textNode(ids, encoded)]
    );
    added.push(smlElement(ids, 'si', 'sst', [], [t]));
    newMap.set(encoded, count);
    indices.set(text, count);
    count += 1;
  }
  if (!newMap) return { pkg: next, indices };
  // `x:extLst` is the last child of `x:sst`; new items go before it.
  const kids = part.root.children;
  const extAt = kids.findIndex((n) => n.type === 'element' && n.local === 'extLst');
  const at = extAt < 0 ? kids.length : extAt;
  const root: Element = {
    ...part.root,
    children: [...kids.slice(0, at), ...added, ...kids.slice(at)],
  };
  const newPart: XmlPart = { partName: part.partName, root, nextId: ids.next };
  plainIndex.set(newPart, newMap);
  return {
    pkg: withPart(next, { kind: 'xml', partName: newPart.partName, xml: newPart }),
    indices,
  };
}

export function countItems(root: Element): number {
  let n = 0;
  for (const c of root.children) if (c.type === 'element' && c.kind === 'stringItem') n += 1;
  return n;
}

/** Add an empty shared string table, its workbook relationship, and its content type. */
export function createSharedStrings(
  pkg: XlsxPackage,
  index: WorkbookIndex
): { pkg: XlsxPackage; partName: string } | OpRejection {
  const relsName = relsPartName(index.mainPart);
  const rels = xmlPartOf(pkg, relsName);
  const ct = xmlPartOf(pkg, CONTENT_TYPES_PART);
  if (!rels || !ct)
    return {
      code: 'unsupported-package',
      detail: 'a new shared string table needs the workbook relationships and the content types',
    };
  let partName = SST_PART;
  for (let n = 2; equivalentPartName(pkg, partName); n += 1) partName = `/xl/sharedStrings${n}.xml`;
  const sstIds = new IdSource({ partName, root: null as never, nextId: 1 });
  const root: Element = {
    ...smlElement(sstIds, 'sst', null),
    nsDecls: [{ prefix: '', uri: NS.x }],
  };
  let next = withPart(pkg, {
    kind: 'xml',
    partName,
    xml: { partName, root, nextId: sstIds.next },
  });
  // The workbook relationship.
  {
    const used = new Set(relationshipsOf(next, index.mainPart).map((r) => r.id));
    let id = 'rIdSst';
    for (let n = 2; used.has(id); n += 1) id = `rIdSst${n}`;
    const ids = new IdSource(rels);
    const target = relativeTarget(index.mainPart, partName);
    const rel = makeElement(ids, NS.rel, 'Relationship', null, [
      plainAttr('Id', id),
      plainAttr('Type', REL.sharedStrings),
      plainAttr('Target', target),
    ]);
    const relsRoot = withChild(rels.root, rels.root.children.length, rel, true);
    next = withPart(next, {
      kind: 'xml',
      partName: relsName,
      xml: { partName: relsName, root: relsRoot, nextId: ids.next },
    });
  }
  // The content type override.
  {
    const ids = new IdSource(ct);
    const override = makeElement(ids, NS.ct, 'Override', null, [
      plainAttr('PartName', partName),
      plainAttr('ContentType', CONTENT_TYPE.sharedStrings),
    ]);
    const ctRoot = withChild(ct.root, ct.root.children.length, override, true);
    next = withPart(next, {
      kind: 'xml',
      partName: CONTENT_TYPES_PART,
      xml: { partName: CONTENT_TYPES_PART, root: ctRoot, nextId: ids.next },
    });
  }
  return { pkg: next, partName };
}

/**
 * The op that undoes the creation of a shared string table: it removes the table and puts back
 * the workbook relationships and the content types as they were, as the undo of a format edit
 * removes a styles part that it created. Null when `before` already had a table, or when no
 * table was created.
 */
export function sharedStringsUndo(
  before: XlsxPackage,
  index: WorkbookIndex,
  after: XlsxPackage
): HistoryOp | null {
  if (index.sharedStringsPart !== undefined || after === before) return null;
  const created = workbookIndex(after)?.sharedStringsPart;
  if (created === undefined) return null;
  const rels = relsPartName(index.mainPart);
  return {
    op: 'restoreParts',
    parts: [
      { partName: created, part: null },
      { partName: rels, part: before.parts.get(rels) ?? null },
      { partName: CONTENT_TYPES_PART, part: before.parts.get(CONTENT_TYPES_PART) ?? null },
    ],
  };
}

/** The relationship target of `to` from a part in the same folder tree as `from`. */
export function relativeTarget(from: string, to: string): string {
  const fromDir = from.slice(0, from.lastIndexOf('/') + 1);
  return to.startsWith(fromDir) ? to.slice(fromDir.length) : to;
}
