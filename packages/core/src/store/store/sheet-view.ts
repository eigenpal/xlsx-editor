// The headings flag of a sheet view: `x:sheetView/@showRowColHeaders` (`CT_SheetView`, ECMA-376
// Part 1, 18.3.1.87), an optional `xsd:boolean` whose default is `true` in sml.xsd. The editor
// reads and writes the first `x:sheetView` of `x:sheetViews` (`CT_SheetViews`, 18.3.1.88), as it
// reads the zoom and the scroll of a sheet from it (`sheet-opening.ts`).
//
// - Hide writes `showRowColHeaders="0"`. A worksheet with no `x:sheetViews` gets one in its
//   schema place in `CT_Worksheet` (18.3.1.99), after `x:sheetPr` and `x:dimension`, with one
//   `x:sheetView` whose required `@workbookViewId` is `0`, the first workbook view. An
//   `x:sheetViews` with no `x:sheetView` gets one as its first child. A worksheet with no
//   `x:sheetViews` and an element of another namespace before `x:sheetData`, such as
//   `mc:AlternateContent`, refuses a hide with `unsupported-package`: that element can hold the
//   sheet views or an element before them, so no place for a new `x:sheetViews` is certain.
// - Show removes the attribute, because its default is `true`. A worksheet with no `x:sheetView`
//   shows its headings, so a show there changes nothing.
// - A flag that already has the asked state changes nothing: no revision and no undo entry.
//
// Sheet protection (`x:sheetProtection`, 18.3.1.85) and workbook protection
// (`x:workbookProtection`, 18.2.29) name no view flag, so neither refuses the op.
//
// The inverse puts back the worksheet part (`restoreParts`), as the inverse of a merge does.

import type { Element, Node } from '../package/tree.ts';
import type { XlsxPackage } from '../package/package.ts';
import { NS } from '../package/namespaces.ts';
import { attr, child } from '../package/query.ts';
import { sheetTarget, withPart } from './cells.ts';
import { IdSource, plainAttr, smlElement, withAttr } from './tree-edit.ts';
import { refuse, unchanged, type OpResult, type WorkbookOp } from './ops.ts';

/** The headings op, as `ops.ts` declares it. */
export type ShowHeadingsOp = Extract<WorkbookOp, { readonly op: 'setShowHeadings' }>;

/** The children of `x:worksheet` that come before `x:sheetViews` in `CT_Worksheet`. */
const BEFORE_SHEET_VIEWS = new Set(['sheetPr', 'dimension']);

/** The first `x:sheetView` of a worksheet root, the view that the editor reads and writes. */
export function firstSheetView(root: Element | undefined): Element | undefined {
  return child(child(root, 'x:sheetViews'), 'x:sheetView');
}

/**
 * Whether a sheet view shows its row and column headings: `@showRowColHeaders` is absent (its
 * default is `true`), `1`, or `true`. A value that is not an `xsd:boolean` reads as false, as
 * the layout reads the gridlines flag.
 */
export function viewShowsHeadings(view: Element | undefined): boolean {
  const raw = attr(view, 'showRowColHeaders');
  return raw === undefined || raw === '1' || raw === 'true';
}

/** Apply `setShowHeadings`. */
export function applyShowHeadings(pkg: XlsxPackage, op: ShowHeadingsOp): OpResult {
  const target = sheetTarget(pkg, op.sheet);
  if ('code' in target) return { ok: false, rejection: target };
  const part = target.part;
  const root = part.root;
  const views = child(root, 'x:sheetViews');
  const view = child(views, 'x:sheetView');
  if (viewShowsHeadings(view) === op.show) return unchanged(pkg);
  let next: Element;
  let nextId = part.nextId;
  if (view && views) {
    // Show removes the attribute, and hide writes `0`.
    const changed = withAttr(view, 'showRowColHeaders', op.show ? undefined : '0');
    const list = { ...views, children: views.children.map((n) => (n === view ? changed : n)) };
    next = { ...root, children: root.children.map((n) => (n === views ? list : n)) };
  } else {
    // Only a hide reaches here: a worksheet with no `x:sheetView` shows its headings.
    const data = root.children.findIndex(
      (n) => n.type === 'element' && n.ns === NS.x && n.local === 'sheetData'
    );
    const lead = root.children.slice(0, data < 0 ? root.children.length : data);
    if (!views && lead.some((n) => n.type === 'element' && n.ns !== NS.x))
      return refuse(
        'unsupported-package',
        `${op.sheet}: content of another namespace before x:sheetData can hold the sheet views`
      );
    const ids = new IdSource(part);
    const made = smlElement(ids, 'sheetView', 'sheetViews', [
      plainAttr('showRowColHeaders', '0'),
      plainAttr('workbookViewId', '0'),
    ]);
    if (views) {
      const list = { ...views, children: [made, ...views.children] };
      next = { ...root, children: root.children.map((n) => (n === views ? list : n)) };
    } else {
      const list = smlElement(ids, 'sheetViews', root.kind, [], [made]);
      const at = root.children.findIndex(
        (n: Node) => n.type === 'element' && !(n.ns === NS.x && BEFORE_SHEET_VIEWS.has(n.local))
      );
      const children = root.children.slice();
      children.splice(at < 0 ? children.length : at, 0, list);
      next = { ...root, children };
    }
    nextId = ids.next;
  }
  const xml = { partName: part.partName, root: next, nextId };
  return {
    ok: true,
    pkg: withPart(pkg, { kind: 'xml', partName: part.partName, xml }),
    inverse: [
      {
        op: 'restoreParts',
        parts: [{ partName: part.partName, part: pkg.parts.get(part.partName)! }],
      },
    ],
    value: false,
    formula: false,
  };
}
