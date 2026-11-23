// The formula cell rule of design D3. The grid stale mark and the formula bar read it
// for each cell. The stale flag, the stale notice, the calculation chain drop, undo, and clear
// read it through the scan of `calc.ts`, which also counts a cell out of place inside another
// cell. The array range checks read a wrapped array or data table master too
// ({@link wrappedArrayRefs}). The copy and the cut refuse a wrapped formula through
// `cellHasFormula`, and the formula bar shows a wrapped shared child with the text of its
// direct master. The structural edits and the sheet ops read every wrapped `x:f`
// ({@link wrappedFormulas}) and refuse an edit that would change one
// (`structure/wrapped-formulas.ts`). The sheet index reads only a direct `x:f`.

import type { Element, Node } from '../package/tree.ts';
import { NS } from '../package/namespaces.ts';
import { attr } from '../package/query.ts';

/** The branches of a cell's wrappers that hold its formula and its cached value. */
interface WrappedBranches {
  readonly formula: Element | undefined;
  readonly value: Element | undefined;
  /** Whether any branch holds a formula range master ({@link rangeMasterBlocker}). */
  readonly master: boolean;
  /** The `@ref` of the first formula range master in a branch, in branch order. */
  readonly masterRef: string | undefined;
  /** The `@si` of each shared formula child (`t="shared"` with no `@ref`) in a branch. */
  readonly sharedChildren: readonly string[];
  /** The `@ref` of each array or data table master in a branch. */
  readonly arrayRefs: readonly string[];
  /** Every `x:f` child of every branch, in branch order. */
  readonly formulas: readonly Element[];
}

const branches = new WeakMap<Element, WrappedBranches>();
const NONE: WrappedBranches = {
  formula: undefined,
  value: undefined,
  master: false,
  masterRef: undefined,
  sharedChildren: [],
  arrayRefs: [],
  formulas: [],
};

const isElement = (n: Node, ns: string, local: string): n is Element =>
  n.type === 'element' && n.ns === ns && n.local === local;

/**
 * Whether an `x:f` is the master of a formula range: an `@ref` with a `t` other than `normal`
 * (`ST_CellFormulaType`, Part 1, 18.18.6: a shared, array, or data table formula). In
 * `CT_CellFormula` (sml.xsd), `@t` has the default `normal`. The direct checks of a value edit and
 * this wrapped check read the same test.
 */
function isFormulaRangeMaster(f: Element | undefined): boolean {
  const t = attr(f, 't');
  return attr(f, 'ref') !== undefined && t !== undefined && t !== 'normal';
}

/**
 * The branches of the markup-compatibility wrapper children of a cell (`mc:AlternateContent`)
 * that hold the cell's formula and its cached value: the first branch, in order, with an `x:f`
 * child, and the first with an `x:v` child and no `x:f` child. It reads the wrappers in document
 * order. In each wrapper, `mc:Fallback` comes first, as the branch that a consumer takes when it
 * understands no `mc:Choice` (ECMA-376 Part 3, Markup Compatibility, `AlternateContent`), then
 * each `mc:Choice` in document order. The engine understands no `@Requires` namespace of a
 * cell, so PPTX Editor's rule (`cascade/alternate-content.ts`: the first understood Choice, else
 * the Fallback) also takes the Fallback here. A Choice is still read after it, so a formula that
 * only a Choice holds counts, and the cell is never shown as a constant. The fixture
 * `wrapped-formula` measures the value of a Fallback; the other branch cases have no reference PDF
 * (approximation `wrapped-formula-values`). A branch comes before
 * the wrappers nested in it. It reads only the children of each wrapper and branch, with an
 * explicit stack, and it is cached for each cell element that has a wrapper child, so a layout
 * pays it once for each revision.
 */
function wrappedBranches(cell: Element): WrappedBranches {
  if (!cell.children.some((n) => isElement(n, NS.mc, 'AlternateContent'))) return NONE;
  const known = branches.get(cell);
  if (known) return known;
  const top = cell.children.filter((n) => isElement(n, NS.mc, 'AlternateContent'));
  let formula: Element | undefined;
  let value: Element | undefined;
  let master = false;
  let masterRef: string | undefined;
  const sharedChildren: string[] = [];
  const arrayRefs: string[] = [];
  const formulas: Element[] = [];
  // Work items in reverse, so that the stack pops them in document order.
  const stack: [Element, boolean][] = top.reverse().map((w) => [w as Element, true]);
  while (stack.length > 0) {
    const [el, wrapper] = stack.pop()!;
    if (wrapper) {
      const kids = el.children.filter((n): n is Element => n.type === 'element' && n.ns === NS.mc);
      const ordered = [
        ...kids.filter((k) => k.local === 'Fallback'),
        ...kids.filter((k) => k.local === 'Choice'),
      ];
      for (let i = ordered.length - 1; i >= 0; i -= 1) stack.push([ordered[i]!, false]);
      continue;
    }
    const fs = el.children.filter((k): k is Element => isElement(k, NS.x, 'f'));
    const hasF = fs.length > 0;
    if (!formula && hasF) formula = el;
    for (const f of fs) formulas.push(f);
    const rangeMaster = fs.find(isFormulaRangeMaster);
    if (rangeMaster) {
      master = true;
      masterRef ??= attr(rangeMaster, 'ref');
    }
    for (const f of fs) {
      const si = attr(f, 'si');
      const t = attr(f, 't');
      const ref = attr(f, 'ref');
      if (t === 'shared' && ref === undefined && si !== undefined) sharedChildren.push(si);
      if ((t === 'array' || t === 'dataTable') && ref) arrayRefs.push(ref);
    }
    // A value next to another formula is that formula's result, so only a branch with no
    // formula gives a value to the formula of another branch.
    if (!value && !hasF && el.children.some((k) => isElement(k, NS.x, 'v'))) value = el;
    const nested = el.children.filter((k) => isElement(k, NS.mc, 'AlternateContent'));
    for (let i = nested.length - 1; i >= 0; i -= 1) stack.push([nested[i] as Element, true]);
  }
  const found = { formula, value, master, masterRef, sharedChildren, arrayRefs, formulas };
  branches.set(cell, found);
  return found;
}

/** The words for the master kind in a refusal detail. A file can store any `@t` text. */
const MASTER_KIND: Readonly<Record<string, string>> = Object.assign(Object.create(null), {
  array: 'an array formula',
  dataTable: 'a data table formula',
});

/**
 * Why an edit that replaces a cell must be refused to keep a formula range whole, or undefined.
 * The cell holds an array or data table master in a direct `x:f`, or any formula range master
 * in an `mc:AlternateContent` branch. A direct shared master is not a reason: the edit unshares
 * it. The value edit, its batch, and the clear read this one test. They refuse a cell with a
 * formula in a branch for its child first (`cellChildBlocker`); the branch test keeps the range
 * whole also without that check.
 */
export function rangeMasterBlocker(cell: Element): string | undefined {
  if (wrappedBranches(cell).master)
    return 'holds the master of a formula range inside mc:AlternateContent';
  const f = cell.children.find((n): n is Element => isElement(n, NS.x, 'f'));
  const t = attr(f, 't');
  if (!isFormulaRangeMaster(f) || t === 'shared') return undefined;
  const kind = MASTER_KIND[t!] ?? 'a formula range';
  return `holds the master of ${kind}, which the editor cannot edit`;
}

/**
 * The `@ref` of the formula range master that a cell holds, in an `mc:AlternateContent` branch
 * first, as {@link rangeMasterBlocker} reads it, else in a direct `x:f`, or undefined. A refusal
 * names this range.
 */
export function rangeMasterRef(cell: Element): string | undefined {
  const wrapped = wrappedBranches(cell).masterRef;
  if (wrapped !== undefined) return wrapped;
  const f = cell.children.find((n): n is Element => isElement(n, NS.x, 'f'));
  return isFormulaRangeMaster(f) ? attr(f, 'ref') : undefined;
}

/**
 * The `@si` of each shared formula child in a branch of an `mc:AlternateContent` child of a cell.
 * The unshare of a master reads only a direct `x:f`, so it cannot rewrite such a child.
 */
export function wrappedSharedChildren(cell: Element): readonly string[] {
  return wrappedBranches(cell).sharedChildren;
}

/**
 * The `@ref` of each array or data table master in a branch of an `mc:AlternateContent` child of
 * a cell. The array range checks read them with the direct masters, so an edit, a clear, or a
 * copy inside such a range is refused as for a direct master.
 */
export function wrappedArrayRefs(cell: Element): readonly string[] {
  return wrappedBranches(cell).arrayRefs;
}

/**
 * Whether a cell holds a formula: an `x:f` child (`CT_Cell`, Part 1, 18.3.1.4), or an `x:f` child
 * of a branch of a markup-compatibility wrapper child ({@link wrappedFormula}).
 */
export function cellHasFormula(cell: Element): boolean {
  for (const n of cell.children) if (isElement(n, NS.x, 'f')) return true;
  return wrappedBranches(cell).formula !== undefined;
}

/**
 * Every `x:f` child of every branch of the `mc:AlternateContent` children of a cell, nested
 * wrappers included, in the branch order of {@link wrappedBranches}. Each one is a copy of the
 * cell's formula for one kind of reader, so an edit that changes the formula changes each copy.
 */
export function wrappedFormulas(cell: Element): readonly Element[] {
  return wrappedBranches(cell).formulas;
}

/** The `x:f` of the first wrapper branch that holds one, for the formula bar. */
export function wrappedFormula(cell: Element): Element | undefined {
  return wrappedBranches(cell).formula?.children.find((n): n is Element => isElement(n, NS.x, 'f'));
}

/**
 * The cached value of a wrapped formula: the `x:v` of the branch of the formula, else of the
 * first branch with an `x:v` and no `x:f`. A value never comes from the branch of another formula.
 */
export function wrappedValue(cell: Element): Element | undefined {
  const { formula, value } = wrappedBranches(cell);
  const isV = (n: Node): n is Element => isElement(n, NS.x, 'v');
  return formula?.children.find(isV) ?? value?.children.find(isV);
}
