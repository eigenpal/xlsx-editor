// The package read of a load: it stops when the load is no longer wanted, and it gives the reason
// of a refused file as the failure of the editor. `readWorkbookForLoad` adds the workbook index
// and the opening sheet, which a load needs before it replaces the open workbook.

import { openLimitOf } from '../automation/open-diagnostics.ts';
import { readXlsxSteps, type XlsxPackage } from '../store/index.ts';
import { workbookIndex, type WorkbookIndex } from '../store/workbook/workbook-index.ts';
import { partNamesOf } from '../store/structure/blockers.ts';
import { editorSheets } from './editor-sheets.ts';
import {
  openingSheet,
  type EditorFailure,
  type EditorSheet,
  type EditorSnapshot,
} from './editor-model.ts';
import { runStepsInSlices } from './sliced-steps.ts';

/** `steps`, which end with null at the first yield after `wanted` turns false. */
export function* whileWanted<T>(
  steps: Generator<void, T>,
  wanted: () => boolean
): Generator<void, T | null> {
  for (;;) {
    const next = steps.next();
    if (next.done) return next.value;
    if (!wanted()) {
      steps.return(null as T);
      return null;
    }
    yield;
  }
}

/** The result of {@link readForLoad}. */
export type LoadRead =
  | { readonly ok: true; readonly pkg: XlsxPackage }
  | { readonly ok: false; readonly failure: EditorFailure };

// The count of package reads, for the test that a picked file is read once.
let reads = 0;

/** How many package reads `readForLoad` has started. For tests. @internal */
export function loadReadCount(): number {
  return reads;
}

/** Reads `bytes` in short tasks. Null when `wanted` turned false during the read. */
export async function readForLoad(
  bytes: Uint8Array,
  host: Node | null,
  wanted: () => boolean
): Promise<LoadRead | null> {
  reads += 1;
  const read = await runStepsInSlices(whileWanted(readXlsxSteps(bytes), wanted), host);
  if (!read || !wanted()) return null;
  if (read.ok) return read;
  const limit = limitOf(read.error);
  // The name of a refused part goes into the detail, for logs.
  const detail = read.error.reason === 'xml' ? `xml: ${read.error.partName}` : read.error.reason;
  const part = read.error.reason === 'xml' ? { part: read.error.partName } : {};
  return { ok: false, failure: { code: 'package', detail, ...part, ...(limit ? { limit } : {}) } };
}

/** The result of {@link readWorkbookForLoad}. `first` is -1 when no worksheet can show. */
export type WorkbookRead =
  | {
      readonly ok: true;
      readonly pkg: XlsxPackage;
      readonly index: WorkbookIndex;
      readonly sheets: EditorSheet[];
      readonly first: number;
      /** The part names of the package (`partNamesOf`): a sheet that an add makes is not one. */
      readonly openedParts: ReadonlySet<string>;
    }
  | { readonly ok: false; readonly failure: EditorFailure };

/** The error of a picked file that the browser could not read, for `readWorkbookForLoad`. */
export function readError(cause: unknown): { readonly error: string } {
  return { error: String(cause) };
}

/**
 * {@link readForLoad}, then the workbook index, the sheets, and the opening sheet. A
 * `readError` in place of the bytes fails with the code `read`.
 */
export async function readWorkbookForLoad(
  bytes: Uint8Array | { readonly error: string },
  host: Node | null,
  wanted: () => boolean
): Promise<WorkbookRead | null> {
  if ('error' in bytes) return { ok: false, failure: { code: 'read', detail: bytes.error } };
  const read = await readForLoad(bytes, host, wanted);
  if (!read?.ok) return read;
  const index = workbookIndex(read.pkg);
  if (!index) return { ok: false, failure: { code: 'package', detail: 'no workbook part' } };
  const sheets = editorSheets(read.pkg, index);
  const first = openingSheet(sheets, index.activeTab);
  return { ok: true, pkg: read.pkg, index, sheets, first, openedParts: partNamesOf(read.pkg) };
}

// The load of a picked file of each editor: it keeps the open workbook when the file does not
// open (`openPickedWorkbook`). A registry, so that the public instance gets no member for it.
const pickedLoads = new WeakMap<
  object,
  (bytes: Uint8Array | Promise<Uint8Array>) => Promise<EditorSnapshot>
>();

/** Register the picked-file load of `editor`. @internal */
export function registerPickedLoad(
  editor: object,
  load: (bytes: Uint8Array | Promise<Uint8Array>) => Promise<EditorSnapshot>
): void {
  pickedLoads.set(editor, load);
}

/** The picked-file load of `editor`, if the engine made it. @internal */
export function pickedLoadOf(
  editor: object
): ((bytes: Uint8Array | Promise<Uint8Array>) => Promise<EditorSnapshot>) | undefined {
  return pickedLoads.get(editor);
}

/**
 * The limit that `error` reports, with the names of DOCX Editor's `limit` field. The automation
 * host gives the same limit for a refused open.
 */
export const limitOf = openLimitOf;
