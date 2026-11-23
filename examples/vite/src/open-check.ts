// The demo's check that the editor can show a workbook, made before the load: a load that
// fails closes the open workbook. It copies the rule of `openingSheet` in
// packages/core/src/editor/editor-model.ts, which `XlsxEditorInstance.load` uses.
// `open-check.test.ts` keeps the two equal.

import { displayName } from '@xlsx-editor.dev/core/editor';
import { readXlsxSteps, workbookIndex, type readXlsx } from '@xlsx-editor.dev/core/store';
import { translatorFor } from '@xlsx-editor.dev/react';
import { runInTasks, whileWanted } from './steps-in-tasks.ts';

type WorkbookIndex = NonNullable<ReturnType<typeof workbookIndex>>;

/** True when the workbook has a visible worksheet, so the editor can open it. */
export const hasSheetToShow = (index: WorkbookIndex): boolean =>
  index.sheets.some((s) => s.kind === 'worksheet' && s.state === 'visible');

/** Why a file does not read, as `readXlsx` gives it. */
type PackageRejection = Extract<ReturnType<typeof readXlsx>, { ok: false }>['error'];

/** The editor's English strings, for the limit values that its own message names. */
const t = translatorFor(undefined);

/** The limit of a part that does not read, by the reason of the bounded XML reader. */
const PART_LIMITS = {
  'too-large': 'status.limits.xmlMaxBytes',
  'too-many-elements': 'status.limits.xmlMaxElements',
  'too-deep': 'status.limits.xmlMaxDepth',
} as const;

/** Why a file does not read, in words for people. */
export function readProblem(error: PackageRejection): string {
  switch (error.reason) {
    case 'zip':
      if (error.detail === 'too-large')
        return `the package passes the limit of ${t(
          error.limit === 'zip.maxRatio'
            ? 'status.limits.zipMaxRatio'
            : 'status.limits.zipMaxTotalBytes'
        )}`;
      if (error.detail === 'too-many-entries')
        return `the package passes the limit of ${t('status.limits.zipMaxEntries')}`;
      return 'the file is not a ZIP package';
    case 'no-content-types':
      return 'the package has no content types part';
    case 'not-a-workbook':
      return 'the package is not a workbook';
    case 'xml': {
      const detail = error.detail;
      // The file sets the name: show it as the editor's own messages do.
      const part = displayName(error.partName);
      return Object.hasOwn(PART_LIMITS, detail)
        ? `the part ${part} passes the limit of ${t(PART_LIMITS[detail as keyof typeof PART_LIMITS])}`
        : `the part ${part} is not valid XML`;
    }
  }
}

/**
 * The read of `data` before an open, in short tasks. It ends with null at the first yield after
 * `wanted` turns false, so a newer pick stops it, as the load of the editor stops its read.
 */
export const readForOpen = (data: Uint8Array, wanted: () => boolean) =>
  runInTasks(whileWanted(readXlsxSteps(data), wanted));

/** Why the editor cannot show the file of `read`, in words for people, or `null` when it can. */
export function whyNotOpen(read: ReturnType<typeof readXlsx>): string | null {
  if (!read.ok) return readProblem(read.error);
  const index = workbookIndex(read.pkg);
  if (!index) return 'the file has no workbook part';
  return hasSheetToShow(index) ? null : 'the file has no worksheet to show';
}
