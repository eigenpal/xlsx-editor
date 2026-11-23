// The refusal code of a workbook that the engine does not open, for the corpus tasks
// (`fidelity:census` and `validate:xsd:corpus`).

import type { PackageRejection } from '../../packages/core/src/store/index.ts';

/** The refusal of an open, as a code with no file-derived text. */
export function refusalCode(error: PackageRejection): string {
  if (error.reason === 'zip')
    return /^[a-z-]{1,40}$/.test(error.detail) ? `zip:${error.detail}` : 'zip';
  if (error.reason === 'xml') return `xml:${error.detail}`;
  return error.reason;
}
