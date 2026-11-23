// The package writer: one `.xlsx` from the parts of a package. Modeled XML parts re-emit
// normalized through `serialize.ts`; every other part is written byte for byte as read.

import { writeZipSteps } from '@eigenpal/ooxml-core';
import { CONTENT_TYPES_PART, type XlsxPackage } from '../package/package.ts';
import { serializeXmlPartSteps } from './serialize.ts';

/** Serialize the package to `.xlsx` bytes. Modeled parts re-emit normalized; others as read. */
export function writeXlsx(pkg: XlsxPackage): Uint8Array<ArrayBuffer> {
  const steps = writeXlsxSteps(pkg);
  for (;;) {
    const next = steps.next();
    if (next.done) return next.value;
  }
}

/**
 * {@link writeXlsx}, yielding inside a long part and between slices of a long ZIP entry. The
 * package is immutable, so the bytes describe `pkg` even when the workbook moves on between
 * steps. The bytes are the same either way.
 */
export function* writeXlsxSteps(pkg: XlsxPackage): Generator<void, Uint8Array<ArrayBuffer>> {
  const encoder = new TextEncoder();
  const entries = new Map<string, Uint8Array>();
  // [Content_Types].xml first. ECMA-376 Part 2 (the content types stream) does not require this
  // place, but some readers sniff the first entry. This order is an engine choice.
  const ordered = [...pkg.parts.values()].sort((a, b) =>
    a.partName === CONTENT_TYPES_PART ? -1 : b.partName === CONTENT_TYPES_PART ? 1 : 0
  );
  for (const part of ordered) {
    if (part.kind === 'xml') {
      const text = yield* serializeXmlPartSteps(part.xml);
      yield;
      entries.set(part.partName, encoder.encode(text));
      yield;
    } else entries.set(part.partName, part.bytes);
  }
  const zip = yield* writeZipSteps(entries);
  return zip.buffer instanceof ArrayBuffer ? (zip as Uint8Array<ArrayBuffer>) : new Uint8Array(zip);
}
