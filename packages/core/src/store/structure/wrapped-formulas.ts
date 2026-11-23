// Formula text that a cell or the defined name list keeps in a branch of `mc:AlternateContent`,
// under a structural edit (`structure.ts`, `sheet-cells.ts`) or a sheet op (`sheet-ops.ts`).
//
// A writer puts one formula in several branches (ECMA-376 Part 3, 7.5): each `mc:Choice` is
// for a reader that understands the namespaces that it requires, and `mc:Fallback` is for any
// other reader. Each edit runs its rewrite of formula text on every copy, and writes none of
// them. When the rewrite would change a copy, or cannot read it, the edit is refused, and the
// refusal names the cell or the defined name. When every copy stays as it is, the edit goes on.
//
// Refuse, not rewrite. PPTX Editor keeps a slide timing inside `mc:AlternateContent` read-only,
// because each branch would need the same edit
// (`packages/core/src/automation/commands/animation-model.ts` at d4502552). Its rule for the
// copies of a shape (`alternate-copies.ts` at d4502552) writes the same attributes into every
// copy only for a transform, and removes the other copies for any other edit. A transform is a
// fixed set of attributes with one meaning in every branch. A formula copy is text in a grammar
// that an `mc:Choice` can extend with functions or syntax that only a reader of its namespaces
// understands, so a move of its references cannot be checked as a transform can, and a removal
// would lose a formula. So these edits follow the read-only rule, as the value edit of such a
// cell does (`cellChildBlocker` in `store/cells.ts`).

import type { Element, Node } from '../package/tree.ts';
import { NS } from '../package/namespaces.ts';
import { textOf } from '../package/query.ts';
import { wrappedFormulas } from '../workbook/cell-formula.ts';

/** A rewrite of formula text: whether it changes the text, or why it cannot rewrite it. */
export type TextRewrite = (
  text: string
) =>
  | { readonly ok: true; readonly changed: boolean }
  | { readonly ok: false; readonly reason: string; readonly range?: true };

/**
 * Why an edit cannot keep a wrapped copy of formula text. `range` marks a sheet range that a
 * sheet move or delete would change (`sheet-range-blocked`).
 */
export interface WrappedProblem {
  readonly reason: string;
  readonly range?: true;
}

/** What a wrapped copy of formula text belongs to: a cell formula or a defined name. */
export type WrappedHolder = 'a formula' | 'a defined name';

/**
 * Why the edit cannot keep one wrapped copy of formula text, or undefined. `holder` names what
 * the copy belongs to in the reason.
 */
export function wrappedTextProblem(
  text: string,
  rewrite: TextRewrite,
  holder: WrappedHolder = 'a formula'
): WrappedProblem | undefined {
  if (text === '') return undefined;
  const r = rewrite(text);
  if (!r.ok) {
    const reason = `${holder} inside mc:AlternateContent (${r.reason})`;
    return r.range ? { reason, range: true } : { reason };
  }
  return r.changed
    ? {
        reason: `the edit changes ${holder} inside mc:AlternateContent, and the editor does not rewrite each branch`,
      }
    : undefined;
}

/** Why the edit cannot keep a formula that a cell holds in `mc:AlternateContent`, or undefined. */
export function wrappedCellProblem(
  cell: Element,
  rewrite: TextRewrite
): WrappedProblem | undefined {
  for (const f of wrappedFormulas(cell)) {
    const problem = wrappedTextProblem(textOf(f), rewrite);
    if (problem) return problem;
  }
  return undefined;
}

const isElement = (n: Node, ns: string, local: string): n is Element =>
  n.type === 'element' && n.ns === ns && n.local === local;

/**
 * The `x:definedName` elements in the branches of an `mc:AlternateContent` child of
 * `x:definedNames` or of `x:workbook`, nested wrappers and an `x:definedNames` in a branch
 * included. The edits read the set, so the order is not fixed. The typed name list skips them,
 * so the edits read them here. It uses an explicit stack, so the parser's element count cap
 * bounds it.
 */
export function wrappedDefinedNames(alternate: Element): readonly Element[] {
  const out: Element[] = [];
  const stack: Element[] = [alternate];
  while (stack.length > 0) {
    const el = stack.pop()!;
    const inWrapper = isElement(el, NS.mc, 'AlternateContent');
    for (const k of el.children) {
      if (k.type !== 'element') continue;
      if (inWrapper) {
        if (k.ns === NS.mc && (k.local === 'Choice' || k.local === 'Fallback')) stack.push(k);
      } else if (isElement(k, NS.x, 'definedName')) {
        out.push(k);
      } else if (isElement(k, NS.mc, 'AlternateContent') || isElement(k, NS.x, 'definedNames'))
        stack.push(k);
    }
  }
  return out;
}

/** Whether a node is an `mc:AlternateContent` element. */
export const isAlternateContent = (n: Node): n is Element =>
  isElement(n, NS.mc, 'AlternateContent');
