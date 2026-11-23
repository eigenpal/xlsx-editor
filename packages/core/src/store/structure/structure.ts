// The structural ops of design D4: insert and delete rows and columns of one worksheet.
//
// One op moves every reference that design D4 lists, in every part, or refuses the whole edit
// with a typed reason that names the element, the part, or the cell; it never writes a partly
// moved workbook. The order is: check the arguments and the blockers (`blockers.ts`), then move
// the cells and formulas of every sheet (`sheet-cells.ts`), the elements of every sheet
// (`sheet-elements.ts`), the drawing anchors of the edited sheet (`drawing.ts`), and the
// defined names of the workbook. A formula or a defined name inside `mc:AlternateContent` that
// the edit would change refuses it (`wrapped-formulas.ts`). A hyperlink relationship target that
// is a fragment (`#Sheet1!A1`) moves as a hyperlink location does, and a hyperlink that a
// deletion removes takes its relationship with it (`hyperlink-rels.ts`).
//
// Undo: a structural edit rewrites references across the workbook, so its inverse puts back
// each part that it changed (`restoreParts`), as the inverse of a format edit puts back the
// styles part (`restoreStyles`). The redo entry is the inverse of that. The side effects of
// designs D3 and D5 (the stale flag and the dropped calculation chain) are not in the inverse:
// `applyOps` adds them again, so undo never brings the stale chain back or clears the flag.

import type { Node, XmlPart } from '../package/tree.ts';
import { attr, textOf } from '../package/query.ts';
import { xmlPartOf, type PackagePart, type XlsxPackage } from '../package/package.ts';
import { sheetIndex } from '../workbook/sheet-index.ts';
import { sheetNameKey, sheetsByKey, workbookIndex } from '../workbook/workbook-index.ts';
import { withFullCalcOnLoad, workbookHasFormula } from '../store/calc.ts';
import { sheetTarget } from '../store/cells.ts';
import { IdSource, textNode } from '../store/tree-edit.ts';
import {
  refuse,
  type OpRejection,
  type OpResult,
  type StructuralOpName,
  type WorkbookOp,
} from '../store/ops.ts';
import { axisLimit, type AxisEdit } from './axis.ts';
import { moveReferences, type ReferenceContext } from './references.ts';
import { blocked, moveSheetCells } from './sheet-cells.ts';
import { moveSheetElements } from './sheet-elements.ts';
import { drawingBlocker, moveDrawing } from './drawing.ts';
import { moveHyperlinkTargets, withoutDeletedHyperlinks } from './hyperlink-rels.ts';
import { longSheetNameBlocker, sheetBlocker, sheetMentions, workbookBlocker } from './blockers.ts';
import { isAlternateContent, wrappedDefinedNames, wrappedTextProblem } from './wrapped-formulas.ts';

/** The structural ops. */
export type StructuralOp = Extract<WorkbookOp, { readonly op: StructuralOpName }>;

/** One part to put back: an earlier revision of it, or `null` to remove it. */
export interface PartRestore {
  readonly partName: string;
  readonly part: PackagePart | null;
}

const isRefusal = (r: unknown): r is OpRejection =>
  typeof r === 'object' && r !== null && 'code' in r;

/** The axis edit of an op, or the refusal of its arguments. */
function axisEditOf(op: StructuralOp): AxisEdit | OpRejection {
  const axis = op.op === 'insertRows' || op.op === 'deleteRows' ? 'rows' : 'cols';
  const kind = op.op === 'insertRows' || op.op === 'insertColumns' ? 'insert' : 'delete';
  const limit = axisLimit(axis);
  const { start, count } = op;
  const fits =
    Number.isInteger(start) &&
    Number.isInteger(count) &&
    start >= 1 &&
    start <= limit &&
    count >= 1 &&
    count <= limit &&
    (kind === 'insert' || start + count - 1 <= limit);
  if (!fits)
    return {
      code: 'bad-reference',
      detail: `${op.op} ${start} count ${count} is outside the ${axis} of the grid`,
    };
  return { axis, kind, at: start - 1, count };
}

/** The workbook part with the references of its defined names moved, or the refusal. */
function moveDefinedNames(
  part: XmlPart,
  ctx: Omit<ReferenceContext, 'host'>
): XmlPart | OpRejection {
  const ids = new IdSource(part);
  const children: Node[] = [];
  let changed = false;
  // A defined name has no host sheet: a reference without a sheet is not certain.
  const move = (text: string) => moveReferences(text, { ...ctx, host: undefined });
  /** The refusal for a defined name in the branches of a wrapper, or undefined. */
  const wrappedRefusal = (wrapper: Node): OpRejection | undefined => {
    if (!isAlternateContent(wrapper)) return undefined;
    for (const name of wrappedDefinedNames(wrapper)) {
      const problem = wrappedTextProblem(textOf(name), move, 'a defined name');
      if (problem)
        return blocked(`x:definedName ${attr(name, 'name') ?? ''}`, problem.reason, {
          kind: 'defined-name',
          name: attr(name, 'name') ?? '',
        });
    }
    return undefined;
  };
  for (const node of part.root.children) {
    if (node.type !== 'element' || node.kind !== 'definedNames') {
      // A defined name list inside a wrapper of the workbook element.
      const inside = wrappedRefusal(node);
      if (inside) return inside;
      children.push(node);
      continue;
    }
    const names: Node[] = [];
    for (const el of node.children) {
      const inside = wrappedRefusal(el);
      if (inside) return inside;
      const text = el.type === 'element' && el.kind === 'definedName' ? textOf(el) : '';
      if (el.type !== 'element' || text === '') {
        names.push(el);
        continue;
      }
      const m = move(text);
      if (!m.ok)
        return blocked(`x:definedName ${attr(el, 'name') ?? ''}`, m.reason, {
          kind: 'defined-name',
          name: attr(el, 'name') ?? '',
        });
      names.push(m.changed ? { ...el, children: [textNode(ids, m.text)] } : el);
      changed ||= m.changed;
    }
    children.push({ ...node, children: names });
  }
  if (!changed) return part;
  return { partName: part.partName, root: { ...part.root, children }, nextId: ids.next };
}

/**
 * Why every structural edit of the worksheet `sheet` is refused, whatever rows or columns it
 * names, or undefined: the sheet is not a worksheet, a sheet has cells that the editor cannot
 * read, or a blocker holds a reference that the edit would move (`blockers.ts`).
 * The blockers do not depend on the edited rows or columns, so the editor reads this as the
 * state of its structural commands. An edit that passes can still be refused by its own rows
 * or columns, for example with `grid-full`. The drawing of the edited sheet refuses every edit
 * when it cannot be read or relates to a part that can refer to cells (`drawingBlocker`). A
 * hyperlink target that the edit moves can also
 * refuse it (`moveHyperlinkTargets`): a fragment target with a `%` escape, or a hyperlink with
 * both `@r:id` and `@location`, refuses only an edit that changes the place it names, so it
 * depends on the rows or columns and is not part of this check. `applyStructure` runs the same
 * check first.
 */
export function structureBlocker(pkg: XlsxPackage, sheet: string): OpRejection | undefined {
  const target = sheetTarget(pkg, sheet);
  if ('code' in target) return target;
  const index = target.index;
  const entry = index.sheets.find((s) => s.partName === target.part.partName)!;
  // Another sheet with the same name without case, which only a file can hold: a prefix such
  // as `DATA!A1` can name either sheet, so no reference can move with certainty.
  const twin = sheetsByKey(index, entry.name).find((s) => s !== entry);
  if (twin)
    return {
      code: 'sheet-name-ambiguous',
      subject: twin.name,
      detail: `the sheets ${JSON.stringify(entry.name)} and ${JSON.stringify(twin.name)} have the same name without case`,
    };
  const long = longSheetNameBlocker(index);
  if (long) return long;
  const needles = sheetMentions(
    entry.name,
    index.sheets.map((s) => s.name)
  );
  const wb = workbookBlocker(pkg, index, needles);
  if (wb) return wb;
  for (const other of index.sheets) {
    const part = other.kind === 'worksheet' ? xmlPartOf(pkg, other.partName) : undefined;
    if (!part) continue;
    const problem = sheetIndex(part).problems[0];
    if (problem)
      return {
        code: 'invalid-sheet',
        detail: `${other.name}: ${problem.code} at ${problem.at ?? '?'}`,
      };
    const blocker = sheetBlocker(
      pkg,
      part,
      other.name,
      part.partName === target.part.partName,
      needles
    );
    if (blocker) return blocker;
  }
  return drawingBlocker(pkg, target.part);
}

/** Apply one structural op. */
export function applyStructure(pkg: XlsxPackage, op: StructuralOp): OpResult {
  const target = sheetTarget(pkg, op.sheet);
  if ('code' in target) return { ok: false, rejection: target };
  const edit = axisEditOf(op);
  if (isRefusal(edit)) return { ok: false, rejection: edit };
  const blocker = structureBlocker(pkg, op.sheet);
  if (blocker) return { ok: false, rejection: blocker };
  const index = target.index;
  const entry = index.sheets.find((s) => s.partName === target.part.partName)!;
  const order = new Map<string, number>();
  for (const s of index.sheets)
    if (!order.has(sheetNameKey(s.name))) order.set(sheetNameKey(s.name), s.index);
  const base = { edit, sheet: sheetNameKey(entry.name), order };
  const parts = new Map(pkg.parts);
  for (const sheet of index.sheets) {
    const part = sheet.kind === 'worksheet' ? xmlPartOf(pkg, sheet.partName) : undefined;
    if (!part) continue;
    const isTarget = part.partName === target.part.partName;
    const ctx: ReferenceContext = { ...base, host: sheetNameKey(sheet.name) };
    const cells = moveSheetCells(part, sheet.name, ctx, isTarget);
    if (isRefusal(cells)) return { ok: false, rejection: cells };
    const elements = moveSheetElements(cells, sheet.name, ctx, isTarget);
    if (isRefusal(elements)) return { ok: false, rejection: elements };
    const links = moveHyperlinkTargets(pkg, elements, sheet.name, (t) => moveReferences(t, ctx));
    if (isRefusal(links)) return { ok: false, rejection: links };
    const moved = links.part;
    if (moved !== part)
      parts.set(part.partName, { kind: 'xml', partName: part.partName, xml: moved });
    const rels = isTarget ? withoutDeletedHyperlinks(pkg, part, moved, links.rels) : links.rels;
    if (rels) parts.set(rels.partName, { kind: 'xml', partName: rels.partName, xml: rels });
    if (isTarget) {
      const drawing = moveDrawing(pkg, part, edit);
      if (isRefusal(drawing)) return { ok: false, rejection: drawing };
      if (drawing && xmlPartOf(pkg, drawing.partName) !== drawing)
        parts.set(drawing.partName, { kind: 'xml', partName: drawing.partName, xml: drawing });
    }
  }
  const main = xmlPartOf(pkg, index.mainPart);
  if (main) {
    const names = moveDefinedNames(main, base);
    if (isRefusal(names)) return { ok: false, rejection: names };
    if (names !== main)
      parts.set(main.partName, { kind: 'xml', partName: main.partName, xml: names });
  }
  const next: XlsxPackage = { parts };
  const restore: PartRestore[] = [];
  for (const [name, part] of next.parts) {
    const old = pkg.parts.get(name);
    if (old !== part) restore.push({ partName: name, part: old ?? null });
  }
  if (restore.length === 0) return { ok: true, pkg, inverse: [], value: false, formula: false };
  return {
    ok: true,
    pkg: next,
    inverse: [{ op: 'restoreParts', parts: restore }],
    value: true,
    formula: workbookHasFormula(next),
  };
}

/**
 * The `restoreParts` op: put back parts of an earlier revision. Its inverse is the parts it
 * replaced. It is the undo and redo of a structural edit, so it reports no value or formula
 * change: the edit that it undoes or redoes already set the stale flag of design D3 and dropped
 * the calculation chain of design D5. A restored workbook part keeps the stale flag of the
 * revision it replaces, because undo never clears that flag.
 */
export function restoreParts(pkg: XlsxPackage, list: readonly PartRestore[]): OpResult {
  const parts = new Map(pkg.parts);
  const inverse: PartRestore[] = [];
  for (const { partName, part } of list) {
    if (part !== null && part.partName !== partName)
      return refuse('bad-value', `restoreParts names ${partName} for the part ${part.partName}`);
    inverse.push({ partName, part: pkg.parts.get(partName) ?? null });
    if (part) parts.set(partName, part);
    else parts.delete(partName);
  }
  const restored: XlsxPackage = { parts };
  const index = workbookIndex(restored);
  if (!index) return refuse('unsupported-package', 'the package has no workbook part');
  const stale = workbookIndex(pkg)?.fullCalcOnLoad === true && !index.fullCalcOnLoad;
  const next = (stale ? withFullCalcOnLoad(restored) : undefined) ?? restored;
  return {
    ok: true,
    pkg: next,
    inverse: [{ op: 'restoreParts', parts: inverse }],
    value: false,
    formula: false,
  };
}
