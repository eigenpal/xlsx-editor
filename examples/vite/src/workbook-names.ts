// The name of a workbook as the demo shows it. A picked file can name itself with control
// characters and bidirectional format controls, so each shown name goes through `displayName`.

import { displayName } from '@xlsx-editor.dev/core/editor';
import { SAMPLE } from './sample.ts';

/** The name of a workbook in a message: the sample has a name for people, not its list key. */
export const titleOf = (label: string): string =>
  label === SAMPLE ? 'the sample workbook' : displayName(label);

/** The text of the workbook list for a picked file. */
export const listedName = (name: string): string => displayName(name);
