// The hyperlink relationships of a worksheet: the ones that a structural edit deletes (design
// D4), and the places in the workbook that their targets name (`moveHyperlinkTargets`).
//
// A hyperlink to an outside target holds its address in a relationship of the worksheet
// (`x:hyperlink/@r:id`, `CT_Hyperlink`, sml.xsd), an external relationship of ECMA-376 Part 2. When
// a deletion removes the hyperlink with its cells, its relationship goes too, unless another
// hyperlink of the sheet still names it, so a saved file holds no relationship that nothing uses.
// Undo puts the relationship part back with the worksheet (`restoreParts`).

import type { Element, Node, XmlPart } from '../package/tree.ts';
import { relsPartName, xmlPartOf, type XlsxPackage } from '../package/package.ts';
import { attr, child, children } from '../package/query.ts';
import { withAttr } from '../store/tree-edit.ts';
import type { OpRejection } from '../store/ops.ts';
import { NS, REL } from '../package/namespaces.ts';

/** The `@r:id` of each hyperlink of a worksheet root. */
function hyperlinkIds(root: Element): Set<string> {
  const out = new Set<string>();
  for (const el of children(child(root, 'x:hyperlinks'), 'x:hyperlink')) {
    const id = attr(el, 'r:id');
    if (id) out.add(id);
  }
  return out;
}

/**
 * The relationship part of a worksheet without the hyperlink relationships that the edit left
 * unused, or `moved` (the part after its targets moved) when nothing more changes, or
 * undefined when nothing changes.
 */
export function withoutDeletedHyperlinks(
  pkg: XlsxPackage,
  before: XmlPart,
  after: XmlPart,
  moved?: XmlPart
): XmlPart | undefined {
  if (before.root === after.root) return moved;
  const kept = hyperlinkIds(after.root);
  const gone = [...hyperlinkIds(before.root)].filter((id) => !kept.has(id));
  if (gone.length === 0) return moved;
  const rels = moved ?? xmlPartOf(pkg, relsPartName(before.partName));
  if (!rels) return undefined;
  const drop = new Set(gone);
  const kids = rels.root.children.filter(
    (n) =>
      !(
        n.type === 'element' &&
        n.ns === NS.rel &&
        n.local === 'Relationship' &&
        attr(n, 'Type') === REL.hyperlink &&
        drop.has(attr(n, 'Id') ?? '')
      )
  );
  if (kids.length === rels.root.children.length) return moved;
  return { ...rels, root: { ...rels.root, children: kids } };
}

/** Formula text rewritten by an edit, or why the edit cannot keep it valid. */
export type TargetRewrite =
  | { readonly ok: true; readonly text: string; readonly changed: boolean }
  | { readonly ok: false; readonly reason: string; readonly range?: true };

/** A worksheet and its relationship part after the hyperlink targets moved. */
export interface MovedTargets {
  readonly part: XmlPart;
  readonly rels?: XmlPart;
}

/** A character outside the fragment set of RFC 3986, section 3.5. */
const FRAGMENT_OUTSIDE = /[^A-Za-z0-9\-._~!$&'()*+,;=:@/?]/;

const isHyperlinkRel = (n: Node): n is Element =>
  n.type === 'element' &&
  n.ns === NS.rel &&
  n.local === 'Relationship' &&
  attr(n, 'Type') === REL.hyperlink;

/**
 * Move the place in the workbook that a hyperlink relationship of a worksheet names (design D1
 * and D4). Such a target is a fragment, `#Data!A1`: the same text as `x:hyperlink/@location`
 * after the `#`, so the edit rewrites it as it rewrites a location. A target with a `%` escape
 * that can name the sheet is refused, because the editor does not decode it.
 *
 * A fragment holds only the characters of RFC 3986, 3.5: not a space, `#`, `%`, or a character
 * outside ASCII. When the rewritten text holds another character (a delete writes `#REF!`, a
 * rename can write `'Q 1'!A1`), each hyperlink that uses the relationship takes the text as its
 * `@location`, and the relationship goes: the result of a hyperlink that held the location. A
 * hyperlink with both `@r:id` and `@location` is refused then, because it has no place for the
 * text. A relationship that no hyperlink uses goes too.
 */
export function moveHyperlinkTargets(
  pkg: XlsxPackage,
  part: XmlPart,
  name: string,
  rewrite: (text: string) => TargetRewrite
): MovedTargets | OpRejection {
  const subject = `${name}!x:hyperlink`;
  const refusal = (r: { reason: string; range?: true }): OpRejection =>
    r.range
      ? { code: 'sheet-range-blocked', subject, detail: `${subject}: the target ${r.reason}` }
      : { code: 'structure-blocked', subject, detail: `${subject}: the target ${r.reason}` };
  const rels = xmlPartOf(pkg, relsPartName(part.partName));
  if (!rels) return { part };
  const kept: Node[] = [];
  const located = new Map<string, string>();
  let changed = false;
  for (const n of rels.root.children) {
    const target = isHyperlinkRel(n) ? attr(n, 'Target') : undefined;
    if (n.type !== 'element' || target === undefined || !target.startsWith('#')) {
      kept.push(n);
      continue;
    }
    const text = target.slice(1);
    let plain = text;
    if (text.includes('%')) {
      try {
        plain = decodeURIComponent(text);
      } catch {
        plain = '';
      }
      const read = rewrite(plain);
      if (plain === '' || !read.ok || read.changed)
        return refusal({ reason: 'has a % escape, and the editor does not decode it' });
      kept.push(n);
      continue;
    }
    const r = rewrite(plain);
    if (!r.ok) return refusal(r);
    if (!r.changed) {
      kept.push(n);
      continue;
    }
    changed = true;
    if (FRAGMENT_OUTSIDE.test(r.text)) located.set(attr(n, 'Id') ?? '', r.text);
    else kept.push(withAttr(n, 'Target', `#${r.text}`));
  }
  if (!changed) return { part };
  const nextRels: XmlPart = { ...rels, root: { ...rels.root, children: kept } };
  if (located.size === 0) return { part, rels: nextRels };
  const sheetKids = part.root.children.slice();
  for (let k = 0; k < sheetKids.length; k += 1) {
    const list = sheetKids[k]!;
    if (list.type !== 'element' || list.ns !== NS.x || list.local !== 'hyperlinks') continue;
    const links = list.children.map((el) => {
      const id = el.type === 'element' ? attr(el, 'r:id') : undefined;
      const text = id === undefined ? undefined : located.get(id);
      if (el.type !== 'element' || text === undefined) return el;
      if (attr(el, 'location') !== undefined) return undefined;
      const noId = { ...el, attrs: el.attrs.filter((a) => !(a.ns === NS.r && a.local === 'id')) };
      return withAttr(noId, 'location', text);
    });
    if (links.includes(undefined))
      return refusal({ reason: 'and the location of a hyperlink cannot both be kept' });
    sheetKids[k] = { ...list, children: links as Node[] };
  }
  return { part: { ...part, root: { ...part.root, children: sheetKids } }, rels: nextRels };
}
