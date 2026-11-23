// Taken through pptx-editor/packages/core/src/store/package/tree.ts at 1943d2c9 (EigenPal).
// The canonical ordered tree for one XML part: typed kinds where layout needs them, a lossless
// generic node for everything else.
//
// Differences: the typed kinds are the SpreadsheetML kinds of `node-kinds.ts`, text is
// significant inside `x:t`, and a part whose attribute value names an undeclared prefix
// (`prefixValueError`: `mc:Ignorable` and the other prefix and name lists) is refused, as DOCX
// Editor's reader refuses it. `XmlPartRejection` and `XmlPartResult` rename the source's
// `TreeRejection` and `TreeResult`, because exported names do not name engine internals.
//
// Same contract as docx-editor/packages/core/src/store/package/ooxml-tree.ts, reduced to the
// needs of one format: names are EXPANDED (namespace URI + local name), prefixes and
// namespace declarations are retained as fidelity evidence, attribute records are arrays (so a
// file-supplied `__proto__` is inert by construction), and every element and text node has a
// stable id minted deterministically in document order. The same bytes always produce the same
// ids, and an edit never renumbers untouched nodes.

import {
  checkXmlSize,
  fastReadXmlSteps,
  readXml,
  readXmlBounded,
  type XmlNode,
  type XmlLimits,
  type XmlRejection,
  type XmlResult,
} from '@eigenpal/ooxml-core';
import { NS } from './namespaces.ts';
import { deriveKind } from './node-kinds.ts';
import {
  isNcName,
  namespaceBindingError,
  prefixValueError,
  qualifiedNameError,
} from './xml-names.ts';

export { kindOf } from './node-kinds.ts';

/** A node's stable identity within its part. */
export type NodeId = string;

export type { NodeKind } from './node-kinds.ts';
import type { NodeKind } from './node-kinds.ts';

/** One `xmlns` declaration authored on an element. `prefix` is '' for the default namespace. */
export interface NamespaceDecl {
  readonly prefix: string;
  readonly uri: string;
}

/** One attribute, carried verbatim. */
export interface Attr {
  readonly ns: string;
  readonly local: string;
  readonly prefix?: string;
  readonly value: string;
}

/** An element node. Immutable; edits produce new nodes along the changed path. */
export interface Element {
  readonly type: 'element';
  readonly id: NodeId;
  readonly kind: NodeKind;
  readonly ns: string;
  readonly local: string;
  readonly prefix?: string;
  readonly nsDecls: readonly NamespaceDecl[];
  readonly attrs: readonly Attr[];
  readonly children: readonly Node[];
}

/** A text node. Only retained where it is significant (see {@link isSignificantText}). */
export interface TextNode {
  readonly type: 'text';
  readonly id: NodeId;
  readonly value: string;
}

export type Node = Element | TextNode;

/** One parsed XML part. `nextId` is the id counter for nodes created by later edits. */
export interface XmlPart {
  readonly partName: string;
  readonly root: Element;
  readonly nextId: number;
}

/** Why a part was refused at the trust boundary. */
export type XmlPartRejection = XmlRejection | 'no-root' | 'unbound-prefix' | 'bad-name';

export type XmlPartResult =
  | { readonly ok: true; readonly part: XmlPart }
  | { readonly ok: false; readonly reason: XmlPartRejection; readonly detail?: string };

/** Text content is significant inside `x:t` and in leaf elements that hold only text. */
function isSignificantText(parentKind: NodeKind, siblings: readonly XmlNode[]): boolean {
  if (parentKind === 'stringText') return true;
  return siblings.every((s) => s.type === 'text');
}

class TreeBuildError extends Error {
  constructor(
    readonly reason: XmlPartRejection,
    detail: string
  ) {
    super(detail);
  }
}

/** Parse one XML part into the canonical tree. Never throws: the bytes are untrusted. */
export function parseXmlPart(partName: string, xml: string, limits?: XmlLimits): XmlPartResult {
  const steps = parseReadSteps(partName, readXml(xml, limits), false);
  for (;;) {
    const next = steps.next();
    if (next.done) return next.value;
  }
}

/** A part shorter than this reads in one step: the steps pay off only on long parts. */
const STEPPED_PART_CHARS = 2_000_000;
/** Characters that the one-pass reader reads between yields. */
const READ_STEP_CHARS = 256 * 1024;
/** An element with at least this many children builds child by child, between yields. */
const WIDE_ELEMENT_CHILDREN = 64;
/** A subtree of this many nodes takes steps of its own, at any depth. */
const LARGE_SUBTREE_NODES = 256;

/**
 * {@link parseXmlPart}, yielding between units of work on a long part. The tree, the node ids,
 * and every refusal are the same. A part outside the plain subset of the one-pass reader takes
 * the bounded read in one step.
 */
export function* parseXmlPartSteps(
  partName: string,
  xml: string,
  limits?: XmlLimits
): Generator<void, XmlPartResult> {
  if (xml.length < STEPPED_PART_CHARS) return parseXmlPart(partName, xml, limits);
  const sized = checkXmlSize(xml, limits);
  if (!sized.ok) return { ok: false, reason: sized.reason };
  const fast = yield* fastReadXmlSteps(xml, sized.maxElements, READ_STEP_CHARS);
  const read: XmlResult = fast !== null ? { ok: true, nodes: fast } : readXmlBounded(xml, limits);
  yield;
  return yield* parseReadSteps(partName, read, true);
}

/** The tree of one read part. `stepped` yields while it builds a wide element. */
function* parseReadSteps(
  partName: string,
  read: XmlResult,
  stepped: boolean
): Generator<void, XmlPartResult> {
  if (!read.ok) return { ok: false, reason: read.reason };
  const roots = read.nodes.filter((n) => n.type === 'element');
  if (roots.length !== 1 || roots[0].type !== 'element') return { ok: false, reason: 'no-root' };
  const counter = { next: 1 };
  const scope = new Map([['xml', NS.xml]]);
  try {
    const root = stepped
      ? yield* buildSteps(roots[0], scope, null, counter)
      : build(roots[0], scope, null, counter);
    return { ok: true, part: { partName, root, nextId: counter.next } };
  } catch (e) {
    if (e instanceof TreeBuildError) return { ok: false, reason: e.reason, detail: e.message };
    return { ok: false, reason: 'parse-error' };
  }
}

function splitName(raw: string): { prefix: string; local: string } {
  const colon = raw.indexOf(':');
  if (colon < 0) return { prefix: '', local: raw };
  return { prefix: raw.slice(0, colon), local: raw.slice(colon + 1) };
}

type XmlElementNode = Extract<XmlNode, { type: 'element' }>;

/** An element with its checked names and its id, before its children are built. */
interface OpenedElement {
  readonly head: Omit<Element, 'children'>;
  readonly scope: ReadonlyMap<string, string>;
  readonly childContext: NodeKind | null;
  readonly keepText: boolean;
}

function build(
  node: XmlElementNode,
  scope: ReadonlyMap<string, string>,
  parentKind: NodeKind | null,
  counter: { next: number }
): Element {
  const opened = openElement(node, scope, parentKind, counter);
  const children: Node[] = [];
  for (const child of node.children) {
    if (child.type === 'text') pushText(children, child.value, opened.keepText, counter);
    else children.push(build(child, opened.scope, opened.childContext, counter));
  }
  return { ...opened.head, children };
}

/** Whether the subtree of `node` is large, so that its build needs steps of its own. */
function holdsManyNodes(node: XmlElementNode): boolean {
  const open = [node];
  let seen = 0;
  for (let at = open.pop(); at; at = open.pop()) {
    for (const child of at.children) {
      seen += 1;
      if (seen >= LARGE_SUBTREE_NODES) return true;
      if (child.type !== 'text') open.push(child);
    }
  }
  return false;
}

/** {@link build}, yielding between the children of a wide element. Node ids are the same. */
function* buildSteps(
  node: XmlElementNode,
  scope: ReadonlyMap<string, string>,
  parentKind: NodeKind | null,
  counter: { next: number }
): Generator<void, Element> {
  const opened = openElement(node, scope, parentKind, counter);
  const children: Node[] = [];
  let built = 0;
  for (const child of node.children) {
    if (child.type === 'text') pushText(children, child.value, opened.keepText, counter);
    else if (holdsManyNodes(child))
      children.push(yield* buildSteps(child, opened.scope, opened.childContext, counter));
    else children.push(build(child, opened.scope, opened.childContext, counter));
    built += 1;
    if (built % WIDE_ELEMENT_CHILDREN === 0) yield;
  }
  return { ...opened.head, children };
}

/** Adds a text node, unless it is white space where text is not significant. */
function pushText(out: Node[], value: string, keep: boolean, counter: { next: number }): void {
  if (!keep && value.trim() === '') return;
  out.push({ type: 'text', id: `n${counter.next++}`, value });
}

function openElement(
  node: XmlElementNode,
  scope: ReadonlyMap<string, string>,
  parentKind: NodeKind | null,
  counter: { next: number }
): OpenedElement {
  const nsDecls: NamespaceDecl[] = [];
  let local = scope;
  const rawAttrs: [string, string][] = [];
  for (const name of Object.keys(node.attributes)) {
    const value = node.attributes[name];
    if (name === 'xmlns' || name.startsWith('xmlns:')) {
      const prefix = name === 'xmlns' ? '' : name.slice(6);
      if (prefix !== '' && !isNcName(prefix)) throw new TreeBuildError('bad-name', name);
      // The same binding rules as an op and the serializer: a deck that opens can be saved.
      const bindingError = namespaceBindingError(prefix, value);
      if (bindingError) throw new TreeBuildError('bad-name', `${name}: ${bindingError.detail}`);
      nsDecls.push({ prefix, uri: value });
      if (local === scope) local = new Map(scope);
      (local as Map<string, string>).set(prefix, value);
    } else rawAttrs.push([name, value]);
  }
  const { prefix, local: localName } = splitName(node.name);
  if (!isNcName(localName) || (prefix && !isNcName(prefix)))
    throw new TreeBuildError('bad-name', node.name);
  const ns = local.get(prefix);
  if (ns === undefined && prefix !== '') throw new TreeBuildError('unbound-prefix', node.name);
  const elementError = qualifiedNameError(ns ?? '', localName, prefix || undefined, false);
  if (elementError) throw new TreeBuildError('bad-name', `${node.name}: ${elementError.detail}`);
  const expanded = new Set<string>();
  const attrs: Attr[] = rawAttrs.map(([name, value]) => {
    const split = splitName(name);
    if (!isNcName(split.local) || (split.prefix && !isNcName(split.prefix)))
      throw new TreeBuildError('bad-name', name);
    const attrNs = split.prefix === '' ? '' : local.get(split.prefix);
    if (attrNs === undefined) throw new TreeBuildError('unbound-prefix', name);
    const attrError = qualifiedNameError(attrNs, split.local, split.prefix || undefined, true);
    if (attrError) throw new TreeBuildError('bad-name', `${name}: ${attrError.detail}`);
    // Two prefixes bound to one URI can spell the same expanded attribute twice.
    const key = `{${attrNs}}${split.local}`;
    if (expanded.has(key)) throw new TreeBuildError('bad-name', `${name}: repeated attribute`);
    expanded.add(key);
    if (split.prefix === '') return { ns: '', local: split.local, value };
    return { ns: attrNs, local: split.local, prefix: split.prefix, value };
  });
  for (const a of attrs) {
    const valueError = prefixValueError(a, ns ?? '', localName, local);
    if (valueError) {
      const name = a.prefix ? `${a.prefix}:${a.local}` : a.local;
      throw new TreeBuildError(valueError.reason, `${node.name}/@${name}: ${valueError.detail}`);
    }
  }
  const { kind, childContext } = deriveKind(ns ?? '', localName, parentKind);
  const id = `n${counter.next++}`;
  return {
    head: {
      type: 'element',
      id,
      kind,
      ns: ns ?? '',
      local: localName,
      ...(prefix ? { prefix } : {}),
      nsDecls,
      attrs,
    },
    scope: local,
    childContext,
    keepText: isSignificantText(kind, node.children),
  };
}
