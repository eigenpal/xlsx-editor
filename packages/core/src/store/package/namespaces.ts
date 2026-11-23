// Namespace URIs and the qualified-name shorthand the engine uses to address elements.
//
// Adapted from pptx-editor/packages/core/src/store/package/namespaces.ts at 1943d2c9
// (EigenPal): the same `q()` shorthand, with the SpreadsheetML and spreadsheet drawing
// namespaces in place of the PresentationML ones, and with `xsi` for the `xsi:type` check of
// `prefixValueError`.
//
// The tree stores EXPANDED names (namespace URI + local name). Prefixes are fidelity evidence
// only. Code addresses elements through `q('x:sheetData')`, which maps the conventional prefix
// to its URI once, so a lookup never depends on the prefix a file chose. `x` is the
// conventional prefix in this code; files usually bind SpreadsheetML as the default namespace.

export const NS = Object.freeze({
  x: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
  xdr: 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  mc: 'http://schemas.openxmlformats.org/markup-compatibility/2006',
  pic: 'http://schemas.openxmlformats.org/drawingml/2006/picture',
  c: 'http://schemas.openxmlformats.org/drawingml/2006/chart',
  rel: 'http://schemas.openxmlformats.org/package/2006/relationships',
  ct: 'http://schemas.openxmlformats.org/package/2006/content-types',
  xml: 'http://www.w3.org/XML/1998/namespace',
  xmlns: 'http://www.w3.org/2000/xmlns/',
  xsi: 'http://www.w3.org/2001/XMLSchema-instance',
});

/** The SpreadsheetML namespace of ECMA-376 Strict. The engine reads Transitional only. */
export const STRICT_SPREADSHEETML = 'http://purl.oclc.org/ooxml/spreadsheetml/main';

/** A conventional prefix the engine knows. */
export type KnownPrefix = keyof typeof NS;

/** An expanded element or attribute name. */
export interface QName {
  readonly ns: string;
  readonly local: string;
}

const cache = new Map<string, QName>();

/**
 * The expanded name for a conventional `prefix:local` spelling. Unprefixed means no namespace
 * (the usual case for attributes). Throws on an unknown prefix: this is for engine code, never
 * for file data.
 */
export function q(name: string): QName {
  const hit = cache.get(name);
  if (hit) return hit;
  const colon = name.indexOf(':');
  let out: QName;
  if (colon < 0) out = { ns: '', local: name };
  else {
    const prefix = name.slice(0, colon) as KnownPrefix;
    const ns = NS[prefix];
    if (ns === undefined) throw new Error(`unknown engine prefix ${prefix}`);
    out = { ns, local: name.slice(colon + 1) };
  }
  cache.set(name, out);
  return out;
}

/** The conventional prefix for a namespace URI, used when the engine creates nodes. */
export function conventionalPrefix(ns: string): string | undefined {
  // SpreadsheetML is written as the default namespace, as every writer in the corpus does.
  if (ns === NS.x) return undefined;
  for (const [prefix, uri] of Object.entries(NS)) if (uri === ns) return prefix;
  return undefined;
}

const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';

/** Relationship types the engine follows. */
export const REL = Object.freeze({
  officeDocument: `${R}officeDocument`,
  worksheet: `${R}worksheet`,
  chartsheet: `${R}chartsheet`,
  sharedStrings: `${R}sharedStrings`,
  styles: `${R}styles`,
  theme: `${R}theme`,
  calcChain: `${R}calcChain`,
  drawing: `${R}drawing`,
  image: `${R}image`,
  hyperlink: `${R}hyperlink`,
});

const SML = 'application/vnd.openxmlformats-officedocument.spreadsheetml.';

/** Content types the engine reads. */
export const CONTENT_TYPE = Object.freeze({
  workbook: `${SML}sheet.main+xml`,
  template: `${SML}template.main+xml`,
  worksheet: `${SML}worksheet+xml`,
  sharedStrings: `${SML}sharedStrings+xml`,
  styles: `${SML}styles+xml`,
  theme: 'application/vnd.openxmlformats-officedocument.theme+xml',
  drawing: 'application/vnd.openxmlformats-officedocument.drawing+xml',
  relationships: 'application/vnd.openxmlformats-package.relationships+xml',
});
