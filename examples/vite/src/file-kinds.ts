// The file kinds that the demo opens and saves: a workbook, a template, and their macro-enabled
// forms. The editor gives the kind of the open package (`snapshot.workbookKind`) and whether it
// is macro-enabled (`snapshot.macroEnabled`). A save names the file with the extension of that
// kind, because a macro-enabled package saved with the extension of a plain one does not open
// cleanly. The React and Vue demos both use this module.

import type { EditorSnapshot } from '@xlsx-editor.dev/core/editor';

/** The kind of a file, by its package kind and whether it is macro-enabled. */
export type FileKind = 'workbook' | 'template' | 'macroWorkbook' | 'macroTemplate';

/** The extension, the media type, and the save dialog label of each file kind. */
export const FILE_KINDS: Readonly<
  Record<FileKind, { readonly ext: string; readonly type: string; readonly label: string }>
> = {
  workbook: {
    ext: '.xlsx',
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    label: 'Workbook',
  },
  template: {
    ext: '.xltx',
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.template',
    label: 'Workbook template',
  },
  macroWorkbook: {
    ext: '.xlsm',
    type: 'application/vnd.ms-excel.sheet.macroEnabled.12',
    label: 'Macro-enabled workbook',
  },
  macroTemplate: {
    ext: '.xltm',
    type: 'application/vnd.ms-excel.template.macroEnabled.12',
    label: 'Macro-enabled workbook template',
  },
};

/** The file kind of the open package, or null with no workbook open. */
export function fileKindOf(
  snapshot: Pick<EditorSnapshot, 'workbookKind' | 'macroEnabled'>
): FileKind | null {
  const { workbookKind: kind, macroEnabled } = snapshot;
  if (kind === null) return null;
  if (!macroEnabled) return kind;
  return kind === 'template' ? 'macroTemplate' : 'macroWorkbook';
}

/** The extensions of the files that the demo opens, for a file input. */
export const OPEN_ACCEPT = '.xlsx,.xltx,.xlsm,.xltm';

/** An extension of a file kind at the end of a name, as a pattern. */
export const KIND_EXTENSION = /\.(xlsx|xltx|xlsm|xltm)$/i;
