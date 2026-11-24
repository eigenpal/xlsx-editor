// The workbook that an editor holds open: the document store, which is the single source of
// truth, and what the editor reads from its revision.

import type { WorkbookDocumentStore, WorkbookIndex, XlsxPackage } from '../store/index.ts';
import type { PictureStore } from '../output/index.ts';
import type { EditorSheet } from './editor-model.ts';
import type { SheetView } from './surface-render.ts';

/** An open workbook. */
export interface Opened {
  /** The single source of truth. `pkg` and `index` follow its revision. */
  readonly store: WorkbookDocumentStore;
  readonly pkg: XlsxPackage;
  readonly index: WorkbookIndex;
  readonly sheets: readonly EditorSheet[];
  readonly pictures: PictureStore;
  readonly views: Map<number, SheetView>;
  /**
   * The part names of the opened package, as the store keeps them: a sheet whose part is not
   * among them is a sheet that an add op of the session made (`EditorSheet.editRefusal`).
   */
  readonly openedParts: ReadonlySet<string>;
}
