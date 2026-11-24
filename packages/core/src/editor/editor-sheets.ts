// The sheet tabs of the snapshot (`EditorSheet`), with the refusal of a rename or a delete of
// each sheet. A sheet of the opened file takes the refusal of the whole workbook
// (`workbookPartBlocker`). A sheet that an add op of the session made takes its own
// (`addedSheetBlocker`), so `sheetCommandState` and the store refuse the same sheets.

import type { WorkbookIndex, XlsxPackage } from '../store/index.ts';
import { workbookPartBlocker } from '../store/index.ts';
import { addedSheetBlocker, isAddedSheet, sameOpaqueContent } from '../store/structure/blockers.ts';
import { refusalOf, type EditRefusal } from './cell-input.ts';
import type { EditorSheet } from './editor-model.ts';

const sheetEditRefusals = new WeakMap<XlsxPackage, EditRefusal | null>();

/** The refusal of a rename or a delete of each sheet of the opened file, once for a package. */
export function fileSheetRefusal(pkg: XlsxPackage, index: WorkbookIndex): EditRefusal | null {
  let refusal = sheetEditRefusals.get(pkg);
  if (refusal === undefined) {
    const rejection = workbookPartBlocker(pkg, index);
    sheetEditRefusals.set(pkg, (refusal = rejection ? refusalOf(rejection) : null));
  }
  return refusal;
}

/**
 * The tabs of a workbook. `openedParts` holds the part names of the opened package; without it (at
 * the load), each sheet is a sheet of the file. `before` is the package and the tabs of the last
 * revision: when only worksheets that the scan did not read changed, as in a cell edit, the
 * tabs keep their reasons (`sameOpaqueContent`).
 */
export function editorSheets(
  pkg: XlsxPackage,
  index: WorkbookIndex,
  openedParts?: ReadonlySet<string>,
  before?: { readonly pkg: XlsxPackage; readonly sheets: readonly EditorSheet[] }
): EditorSheet[] {
  const file = fileSheetRefusal(pkg, index);
  const keep = !!before && sameOpaqueContent(before.pkg, pkg, index);
  return index.sheets.map((s) => {
    let editRefusal = file;
    // Macros block every sheet, also an added one, as the store refuses (`sheetEditBlocker`).
    if (file && file.about?.kind !== 'macros' && isAddedSheet(s, openedParts)) {
      const kept = keep ? before!.sheets[s.index]?.editRefusal : undefined;
      const added = kept === undefined ? addedSheetBlocker(pkg, index, s) : undefined;
      editRefusal = kept !== undefined ? kept : added ? refusalOf(added) : null;
    }
    const visible = s.state === 'visible';
    return {
      index: s.index,
      name: s.name,
      visible,
      worksheet: s.kind === 'worksheet',
      editRefusal,
    };
  });
}
