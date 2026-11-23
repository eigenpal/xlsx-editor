// Which nodes a whole-subtree memo keeps an entry for. A copy of DOCX Editor's
// `packages/core/src/store/package/subtree-memo-policy.ts` at 6c11927b. Differences: a node of
// this tree is a text node by `type: 'text'`, not by `kind: 'textValue'`; and the aggregate memo
// (`createSubtreeAggregateMemo`, `LARGE_SUBTREE_ANSWER`) is not copied, because no memo here keeps
// an answer that grows with the subtree.
//
// Subtree memos (`WeakMap<node, answer>`) let an edit reuse every untouched subtree, because
// an edit shares those nodes by reference. Most nodes of a worksheet are leaves or hold only
// leaves: `x:v`, `x:f`, and `x:t`. One entry per such node is most of a memo's size, and a major
// GC visits every entry of every live weak collection.
//
// A node whose children are all leaves is answered again from those children, at most
// `WIDE_SUBTREE_CHILDREN - 1` of them, with no deeper walk. Its parent has a grandchild
// element, so the parent keeps its entry and an untouched subtree still answers in one read.

import type { Node } from './tree.ts';

/** A node with this many children keeps its entry even when they are all leaves. */
export const WIDE_SUBTREE_CHILDREN = 16;

/** True for an element grandchild or a wide child. Bounds the policy scan as well. */
function hasElementGrandchild(node: Node): boolean {
  if (node.type === 'text') return false;
  for (const child of node.children) {
    if (child.type === 'text') continue;
    if (child.children.length >= WIDE_SUBTREE_CHILDREN) return true;
    for (const grandchild of child.children) if (grandchild.type !== 'text') return true;
  }
  return false;
}

/**
 * Whether a whole-subtree memo keeps an entry for `node`: a node that is wide, or that
 * has an element below its children, or has a wide child. Every other node is answered from its
 * children, which are leaves.
 */
export function keepsSubtreeMemo(node: Node): boolean {
  if (node.type === 'text') return false;
  return node.children.length >= WIDE_SUBTREE_CHILDREN || hasElementGrandchild(node);
}
