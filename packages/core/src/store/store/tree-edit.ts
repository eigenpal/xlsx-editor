// Persistent (structurally shared) node construction and replacement for the store ops.
//
// Same approach as pptx-editor/packages/core/src/store/store/tree-edit.ts at 1943d2c9
// (EigenPal): an edit rebuilds only the path from the root to the changed node, and every
// untouched subtree is shared by reference with the previous revision. Here the path is
// addressed by position, because the SpreadsheetML paths the ops touch are short and fixed
// (worksheet, sheetData, row, cell). `withParentDeclarations` follows `withParentBindings` in
// docx-editor/packages/core/src/store/store/tree-op-fragment.ts at 1d092ab8.

import type { Attr, Element, Node, XmlPart } from '../package/tree.ts';
import { deriveKind, type NodeKind } from '../package/node-kinds.ts';
import { NS } from '../package/namespaces.ts';

/** Mints node ids for one part, continuing its counter. */
export class IdSource {
  #next: number;
  constructor(part: XmlPart) {
    this.#next = part.nextId;
  }
  id(): string {
    return `n${this.#next++}`;
  }
  get next(): number {
    return this.#next;
  }
}

/** A new element in `ns`, typed under `parentKind` by the reader's own rule. */
export function makeElement(
  ids: IdSource,
  ns: string,
  local: string,
  parentKind: NodeKind | null,
  attrs: readonly Attr[] = [],
  children: readonly Node[] = []
): Element {
  return {
    type: 'element',
    id: ids.id(),
    kind: deriveKind(ns, local, parentKind).kind,
    ns,
    local,
    nsDecls: [],
    attrs,
    children,
  };
}

/** A new SpreadsheetML element. */
export const smlElement = (
  ids: IdSource,
  local: string,
  parentKind: NodeKind | null,
  attrs: readonly Attr[] = [],
  children: readonly Node[] = []
): Element => makeElement(ids, NS.x, local, parentKind, attrs, children);

/**
 * `el` with its namespace declarations replaced by those of `source`, the element that it
 * rebuilds or derives from, and with the authored prefix of `source` when both are in one
 * namespace. An op that builds an element again from the attributes of an old one keeps every
 * declaration that the file put on it. So a prefix that `mc:Ignorable` names (ECMA-376 Part 3, 7.2) stays
 * declared, also when the edit removed the last node that used it. The serializer fails a
 * part where such a prefix is not declared (`prefixValueError`), as DOCX Editor's serializer
 * does; this helper keeps the edit from reaching that refusal. With no `source`, `el` comes
 * back as it is.
 */
export function withNamespacesOf(source: Element | null | undefined, el: Element): Element {
  if (!source || (source.nsDecls.length === 0 && source.prefix === undefined)) return el;
  const out: Element = { ...el, nsDecls: source.nsDecls };
  return source.prefix !== undefined && source.ns === el.ns
    ? { ...out, prefix: source.prefix }
    : out;
}

/**
 * `el`, which moves out of `parent`, with the namespace declarations of `parent` that `el` does
 * not make itself. A declaration on `el` wins over one on `parent` for the same prefix. The
 * prefixes of `el` and of its attribute values (`mc:Ignorable`) then mean the same under its new
 * parent, unless the new parent declares one of those prefixes again with another URI. Only the
 * declarations of `parent` itself are added: a move between two children of one element keeps
 * the declarations above them. The same rule as `withParentBindings` in
 * docx-editor/packages/core/src/store/store/tree-op-fragment.ts at 1d092ab8, with this
 * repository's `Element` in place of `OoxmlNode`.
 */
export function withParentDeclarations(el: Element, parent: Element): Element {
  const own = new Set(el.nsDecls.map((d) => d.prefix));
  const added = parent.nsDecls.filter((d) => !own.has(d.prefix));
  return added.length === 0 ? el : { ...el, nsDecls: [...added, ...el.nsDecls] };
}

/** An unprefixed attribute. */
export const plainAttr = (local: string, value: string): Attr => ({ ns: '', local, value });

/** A text node. */
export const textNode = (ids: IdSource, value: string): Node => ({
  type: 'text',
  id: ids.id(),
  value,
});

/** The element with the attribute set (or removed when `value` is undefined), in place order. */
export function withAttr(el: Element, local: string, value: string | undefined): Element {
  const at = el.attrs.findIndex((a) => a.ns === '' && a.local === local);
  if (value === undefined) {
    if (at < 0) return el;
    return { ...el, attrs: el.attrs.filter((_, i) => i !== at) };
  }
  const attrs = el.attrs.slice();
  if (at < 0) attrs.push(plainAttr(local, value));
  else attrs[at] = { ...attrs[at]!, value };
  return { ...el, attrs };
}

/** The element with child `index` replaced, inserted (`insert`), or removed (`null`). */
export function withChild(el: Element, index: number, child: Node | null, insert = false): Element {
  const children = el.children.slice();
  if (insert) {
    if (child) children.splice(index, 0, child);
  } else if (child) children[index] = child;
  else children.splice(index, 1);
  return { ...el, children };
}

/** The index of `child` among `el`'s children, by identity, or -1. */
export function childIndex(el: Element, child: Node): number {
  return el.children.indexOf(child);
}
