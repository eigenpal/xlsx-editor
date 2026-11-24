// The name of a workbook as the demo shows it. A picked file can name itself with control
// characters and bidirectional format controls, so each shown name goes through `displayName`.
// The React and Vue demos both show the open workbook's name with this module.

import { displayName } from '@xlsx-editor.dev/core/editor';
import { SAMPLE } from './sample.ts';

/** The name of a workbook in a message: the sample has a name for people, not its key. */
export const titleOf = (label: string): string =>
  label === SAMPLE ? 'the sample workbook' : displayName(label);

/** The name of the open workbook in the top bar, and its extension, which shows muted. */
export interface ShownName {
  readonly name: string;
  readonly ext: string;
}

/**
 * The name and the extension of the workbook `label` for the top bar. The sample shows as
 * "Sample workbook" with `.xlsx`. A picked file shows its name with the `.xlsx` or `.xltx`
 * extension apart; a name with another extension, or none, shows whole, with no extension.
 */
export function shownName(label: string): ShownName {
  if (label === SAMPLE) return { name: 'Sample workbook', ext: '.xlsx' };
  const shown = displayName(label);
  const ext = /\.(xlsx|xltx)$/i.exec(shown);
  if (!ext || ext.index === 0) return { name: shown, ext: '' };
  return { name: shown.slice(0, ext.index), ext: ext[0].toLowerCase() };
}

/**
 * The extension after the name in the top bar: the extension that Save writes for the package
 * `kind`, else, before the kind is known, the file's own `.xlsx` or `.xltx`, else none.
 */
export function shownExtension(shown: ShownName, kind: 'workbook' | 'template' | null): string {
  if (kind === 'template') return '.xltx';
  return kind === 'workbook' ? '.xlsx' : shown.ext;
}

/** The name of a workbook whose name the user cleared, when it had no name before. */
export const UNTITLED = 'Untitled workbook';

/** The accessible name of the workbook name field in the top bar. */
export const NAME_FIELD_LABEL = 'Workbook name';

/** The label of the File menu row that starts an edit of the workbook name. */
export const RENAME_ROW_LABEL = 'Rename workbook';

/**
 * The workbook name after an edit in the top bar, with no extension. The name loses control
 * characters, bidirectional format controls, the spaces at its ends, and an `.xlsx` or `.xltx`
 * extension that the user typed, because the extension shows apart. An empty name keeps the
 * name `before`, or with no name before, takes the untitled name.
 */
export function renamedTo(typed: string, before: string): string {
  const name = displayName(typed)
    .trim()
    .replace(/\.(xlsx|xltx)$/i, '')
    .trimEnd();
  return name || before.trim() || UNTITLED;
}
