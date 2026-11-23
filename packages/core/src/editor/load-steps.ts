// The package read of a load: it stops when the load is no longer wanted, and it gives the reason
// of a refused file as the failure of the editor.

import { openLimitOf } from '../automation/open-diagnostics.ts';
import { readXlsxSteps, type XlsxPackage } from '../store/index.ts';
import type { EditorFailure } from './editor-model.ts';
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

/** Reads `bytes` in short tasks. Null when `wanted` turned false during the read. */
export async function readForLoad(
  bytes: Uint8Array,
  host: Node | null,
  wanted: () => boolean
): Promise<LoadRead | null> {
  const read = await runStepsInSlices(whileWanted(readXlsxSteps(bytes), wanted), host);
  if (!read || !wanted()) return null;
  if (read.ok) return read;
  const limit = limitOf(read.error);
  // The name of a refused part goes into the detail, for logs.
  const detail = read.error.reason === 'xml' ? `xml: ${read.error.partName}` : read.error.reason;
  const part = read.error.reason === 'xml' ? { part: read.error.partName } : {};
  return { ok: false, failure: { code: 'package', detail, ...part, ...(limit ? { limit } : {}) } };
}

/**
 * The limit that `error` reports, with the names of DOCX Editor's `limit` field. The automation
 * host gives the same limit for a refused open.
 */
export const limitOf = openLimitOf;
