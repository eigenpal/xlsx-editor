// How a typed node kind is derived. One rule for the reader, for nodes an op creates, and for
// nodes an op moves, so a node has the same kind in the store as after save and reopen.
//
// Same shape as pptx-editor/packages/core/src/store/package/node-kinds.ts at 1943d2c9
// (EigenPal), with the SpreadsheetML kinds of design D2 and D4. A known element is typed only
// under the parent that the schema allows for it. Anywhere else it is `generic`, so a
// misplaced element never reaches typed code, and the rest of the part stays editable.

import { NS } from './namespaces.ts';

/**
 * The typed kinds. Everything not listed is `generic`, carried verbatim.
 *
 * A kind is a LABEL on the same node shape, assigned from the expanded name and the parent's
 * kind. Layout and editing read typed kinds; everything else only has to survive save.
 */
export type NodeKind =
  | 'generic'
  // Workbook part (`CT_Workbook`, sml.xsd).
  | 'workbook'
  | 'workbookPr'
  | 'bookViews'
  | 'workbookView'
  | 'sheets'
  | 'sheet'
  | 'definedNames'
  | 'definedName'
  | 'calcPr'
  // Worksheet part (`CT_Worksheet`, sml.xsd).
  | 'worksheet'
  | 'sheetPr'
  | 'dimension'
  | 'sheetViews'
  | 'sheetView'
  | 'pane'
  | 'selection'
  | 'sheetFormatPr'
  | 'cols'
  | 'col'
  | 'sheetData'
  | 'row'
  | 'cell'
  | 'formula'
  | 'value'
  | 'inlineString'
  | 'sheetCalcPr'
  | 'mergeCells'
  | 'mergeCell'
  | 'hyperlinks'
  | 'hyperlink'
  | 'printOptions'
  | 'pageMargins'
  | 'pageSetup'
  | 'drawing'
  // Shared string table (`CT_Sst`, sml.xsd; an item is a `CT_Rst`).
  | 'sst'
  | 'stringItem'
  | 'richRun'
  | 'runProperties'
  | 'stringText'
  // Styles part (`CT_Stylesheet`, sml.xsd).
  | 'styleSheet'
  | 'numFmts'
  | 'numFmt'
  | 'fonts'
  | 'font'
  | 'fills'
  | 'fill'
  | 'borders'
  | 'border'
  | 'cellStyleXfs'
  | 'cellXfs'
  | 'xf'
  | 'cellStyles'
  | 'cellStyle'
  | 'colors'
  // Spreadsheet drawing part (`CT_Drawing`, dml-spreadsheetDrawing.xsd; Part 1, 20.5).
  | 'wsDr'
  | 'twoCellAnchor'
  | 'oneCellAnchor'
  | 'absoluteAnchor'
  | 'anchorFrom'
  | 'anchorTo';

/**
 * The roots of the modeled parts, by local name in the SpreadsheetML namespace. A `Map`, never a
 * plain object: the local name comes from the file, and `constructor` must not find a member.
 */
const SML_ROOTS: ReadonlyMap<string, NodeKind> = new Map<string, NodeKind>([
  ['workbook', 'workbook'],
  ['worksheet', 'worksheet'],
  ['sst', 'sst'],
  ['styleSheet', 'styleSheet'],
]);

/**
 * `parentKind>local` to kind, for SpreadsheetML children. Each pair is a parent and a child
 * element of a sml.xsd content model, such as `row>c` (`CT_Row`) and `cell>is` (`CT_Cell`).
 */
const SML_CHILDREN: ReadonlyMap<string, NodeKind> = new Map<string, NodeKind>([
  ['workbook>workbookPr', 'workbookPr'],
  ['workbook>bookViews', 'bookViews'],
  ['bookViews>workbookView', 'workbookView'],
  ['workbook>sheets', 'sheets'],
  ['sheets>sheet', 'sheet'],
  ['workbook>definedNames', 'definedNames'],
  ['definedNames>definedName', 'definedName'],
  ['workbook>calcPr', 'calcPr'],
  ['worksheet>sheetPr', 'sheetPr'],
  ['worksheet>dimension', 'dimension'],
  ['worksheet>sheetViews', 'sheetViews'],
  ['sheetViews>sheetView', 'sheetView'],
  ['sheetView>pane', 'pane'],
  ['sheetView>selection', 'selection'],
  ['worksheet>sheetFormatPr', 'sheetFormatPr'],
  ['worksheet>cols', 'cols'],
  ['cols>col', 'col'],
  ['worksheet>sheetData', 'sheetData'],
  ['sheetData>row', 'row'],
  ['row>c', 'cell'],
  ['cell>f', 'formula'],
  ['cell>v', 'value'],
  ['cell>is', 'inlineString'],
  ['inlineString>t', 'stringText'],
  ['inlineString>r', 'richRun'],
  ['worksheet>sheetCalcPr', 'sheetCalcPr'],
  ['worksheet>mergeCells', 'mergeCells'],
  ['mergeCells>mergeCell', 'mergeCell'],
  ['worksheet>hyperlinks', 'hyperlinks'],
  ['hyperlinks>hyperlink', 'hyperlink'],
  ['worksheet>printOptions', 'printOptions'],
  ['worksheet>pageMargins', 'pageMargins'],
  ['worksheet>pageSetup', 'pageSetup'],
  ['worksheet>drawing', 'drawing'],
  ['sst>si', 'stringItem'],
  ['stringItem>t', 'stringText'],
  ['stringItem>r', 'richRun'],
  ['richRun>rPr', 'runProperties'],
  ['richRun>t', 'stringText'],
  ['styleSheet>numFmts', 'numFmts'],
  ['numFmts>numFmt', 'numFmt'],
  ['styleSheet>fonts', 'fonts'],
  ['fonts>font', 'font'],
  ['styleSheet>fills', 'fills'],
  ['fills>fill', 'fill'],
  ['styleSheet>borders', 'borders'],
  ['borders>border', 'border'],
  ['styleSheet>cellStyleXfs', 'cellStyleXfs'],
  ['cellStyleXfs>xf', 'xf'],
  ['styleSheet>cellXfs', 'cellXfs'],
  ['cellXfs>xf', 'xf'],
  ['styleSheet>cellStyles', 'cellStyles'],
  ['cellStyles>cellStyle', 'cellStyle'],
  ['styleSheet>colors', 'colors'],
]);

/**
 * `parentKind>local` to kind, for the anchors of dml-spreadsheetDrawing.xsd: `CT_TwoCellAnchor`
 * holds `xdr:from` and `xdr:to`, and `CT_OneCellAnchor` holds `xdr:from` (`CT_Marker`).
 */
const XDR_CHILDREN: ReadonlyMap<string, NodeKind> = new Map<string, NodeKind>([
  ['wsDr>twoCellAnchor', 'twoCellAnchor'],
  ['wsDr>oneCellAnchor', 'oneCellAnchor'],
  ['wsDr>absoluteAnchor', 'absoluteAnchor'],
  ['twoCellAnchor>from', 'anchorFrom'],
  ['twoCellAnchor>to', 'anchorTo'],
  ['oneCellAnchor>from', 'anchorFrom'],
]);

/** Assign a typed kind from the expanded name and the parent's kind. */
export function kindOf(ns: string, local: string, parent: NodeKind | null): NodeKind {
  if (ns === NS.x) {
    if (parent === null) return SML_ROOTS.get(local) ?? 'generic';
    return SML_CHILDREN.get(`${parent}>${local}`) ?? 'generic';
  }
  if (ns === NS.xdr) {
    if (parent === null) return local === 'wsDr' ? 'wsDr' : 'generic';
    return XDR_CHILDREN.get(`${parent}>${local}`) ?? 'generic';
  }
  return 'generic';
}

/**
 * Markup-compatibility wrappers (ECMA-376 Part 3, 7.5) are transparent for typing only in a
 * drawing: an anchor inside `mc:Fallback` is still an anchor of the drawing. Everywhere else the
 * content of a wrapper is `generic`, so a cell or a row inside `mc:Choice` is never a typed node
 * that no index reaches, and a cell in both `mc:Choice` and `mc:Fallback` is never two typed
 * cells. The wrapper itself is `generic`.
 */
export function isTransparentWrapper(ns: string, local: string): boolean {
  return (
    ns === NS.mc && (local === 'AlternateContent' || local === 'Choice' || local === 'Fallback')
  );
}

/** The kind of a node, and the kind context its children are typed under. */
export function deriveKind(
  ns: string,
  local: string,
  context: NodeKind | null
): { readonly kind: NodeKind; readonly childContext: NodeKind | null } {
  if (isTransparentWrapper(ns, local))
    return { kind: 'generic', childContext: context === 'wsDr' ? context : 'generic' };
  const kind = kindOf(ns, local, context);
  return { kind, childContext: kind };
}
