// The hyperlinks of a worksheet, as a sanitized projection (AGENTS.md, "Security"). It reads
// each `x:hyperlinks/x:hyperlink` (`CT_Hyperlink`, sml.xsd): `@ref` (`ST_Ref`, required), and
// the optional `@r:id`, `@location`, and `@display` (`ST_Xstring`). `@r:id` names a relationship
// of the worksheet part whose target is the external URL.
//
// The part keeps every target verbatim and inert. Only this projection reaches a sink: an
// external target goes through `sanitizeHref`, so `javascript:`, `data:`, `vbscript:`, and
// `file:` targets have no `href`. A target in the workbook (`@location`) is a place, never a
// URL. Nothing here fetches anything.

import { sanitizeHref } from '@eigenpal/ooxml-core';
import type { XmlPart } from '../package/tree.ts';
import { relationshipById, type XlsxPackage } from '../package/package.ts';
import { attr, child, children } from '../package/query.ts';

/** One hyperlink of a sheet. */
export interface HyperlinkProjection {
  /** The cell or range the hyperlink covers, as the file states it. */
  readonly ref: string;
  /** A sanitized external URL, when the target is one and the allowlist accepts it. */
  readonly href?: string;
  /** A place in the workbook (`Sheet2!A1`), when the hyperlink names one. */
  readonly location?: string;
  readonly display?: string;
}

/** The hyperlinks of a worksheet part, in document order. */
export function sheetHyperlinks(pkg: XlsxPackage, part: XmlPart): HyperlinkProjection[] {
  const out: HyperlinkProjection[] = [];
  for (const el of children(child(part.root, 'x:hyperlinks'), 'x:hyperlink')) {
    const ref = attr(el, 'ref');
    if (!ref) continue;
    const relId = attr(el, 'r:id');
    const rel = relId ? relationshipById(pkg, part.partName, relId) : undefined;
    const raw = rel?.external ? rel.externalTarget : undefined;
    const projected = raw === undefined ? undefined : sanitizeHref(raw);
    const location = attr(el, 'location');
    const display = attr(el, 'display');
    out.push({
      ref,
      ...(projected?.ok ? { href: projected.href } : {}),
      ...(location !== undefined ? { location } : {}),
      ...(display !== undefined ? { display } : {}),
    });
  }
  return out;
}
