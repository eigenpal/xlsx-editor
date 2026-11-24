// The part of the snapshot that the whole workbook gives: its kind, whether it is
// macro-enabled, its protection, the refusal of a rename or a delete of any sheet, the parts past
// a limit, and the theme colors. `editor-model.ts` builds the rest of the snapshot.

import {
  contentTypeOf,
  readWorkbookProtection,
  type WorkbookIndex,
  type XlsxPackage,
} from '../store/index.ts';
import { isMacroEnabledType, isTemplateType } from '../store/package/main-part-type.ts';
import { workbookTheme } from '../cascade/index.ts';
import { fileSheetRefusal } from './editor-sheets.ts';
import { partsPastLimit } from './parts-past-limit.ts';
import type { EditorSnapshot, WorkbookKind } from './editor-model.ts';

/** The part of a snapshot that the whole workbook gives. */
export type WorkbookState = Pick<
  EditorSnapshot,
  | 'workbookKind'
  | 'macroEnabled'
  | 'workbookProtection'
  | 'sheetEditRefusal'
  | 'partsPastLimit'
  | 'themeColors'
>;

/** The workbook state with no workbook open. */
export const NO_WORKBOOK_STATE: WorkbookState = Object.freeze({
  workbookKind: null,
  macroEnabled: false,
  workbookProtection: null,
  sheetEditRefusal: null,
  partsPastLimit: [],
  themeColors: [],
});

/**
 * The snapshot state of the workbook: its kind, its protection, and the refusal of a rename or
 * a delete of each sheet of the opened file (`fileSheetRefusal`), read once for each package.
 */
export function workbookState(pkg: XlsxPackage, index: WorkbookIndex): WorkbookState {
  const refusal = fileSheetRefusal(pkg, index);
  return {
    workbookKind: workbookKindOf(pkg, index),
    macroEnabled: isMacroEnabledType(contentTypeOf(pkg, index.mainPart)),
    workbookProtection: readWorkbookProtection(pkg, index),
    sheetEditRefusal: refusal,
    partsPastLimit: partsPastLimit(pkg),
    // The theme part is cached by its root, so the list keeps its identity.
    themeColors: workbookTheme(pkg, index).palette,
  };
}

/** The kind of a package's main part. A macro-enabled template is a template. */
export function workbookKindOf(pkg: XlsxPackage, index: WorkbookIndex): WorkbookKind {
  return isTemplateType(contentTypeOf(pkg, index.mainPart)) ? 'template' : 'workbook';
}
