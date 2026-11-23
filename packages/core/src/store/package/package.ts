// The OPC package: bounded read, part classification, relationships, and save.
//
// Adapted from pptx-editor/packages/core/src/store/package/package.ts at 1943d2c9 (EigenPal).
// Differences: the parts the engine models are the SpreadsheetML ones (workbook, worksheets,
// shared strings, styles, theme, drawings, and relationships); there is no deferred read and no
// PresentationML repair.
//
// Pipeline: bytes → readZip (bounded, OPC names normalized) → [Content_Types].xml → each XML
// part the engine models through parseXmlPart (bounded, entity-free) → XlsxPackage. Every
// other part (media, charts, comments, macros, embeddings, XML the engine does not read) is
// carried as bytes and re-emitted byte-identical. Nothing here fetches anything: external
// relationship targets are recorded, never resolved.

import {
  asciiFold,
  normalizePartName,
  partNameKey,
  readZip,
  resolveInternalTarget,
  sha256FontBytes,
  validateExternalTarget,
  DEFAULT_ZIP_LIMITS,
  XML_HARD_MAX_BYTES,
  type ZipLimits,
} from '@eigenpal/ooxml-core';
import {
  parseXmlPart,
  parseXmlPartSteps,
  type Element,
  type XmlPart,
  type XmlPartRejection,
} from './tree.ts';
import { canonicalFingerprint, bytesFingerprint } from '../package-write/serialize.ts';
import { attr, children, elements } from './query.ts';
import { CONTENT_TYPE, NS, REL } from './namespaces.ts';

/** One part of the package. */
export type PackagePart =
  | { readonly kind: 'xml'; readonly partName: string; readonly xml: XmlPart }
  | { readonly kind: 'binary'; readonly partName: string; readonly bytes: Uint8Array };

/** The whole package. Immutable: an edit produces a new package that shares untouched parts. */
export interface XlsxPackage {
  readonly parts: ReadonlyMap<string, PackagePart>;
}

/**
 * Why a file was refused. The `reason` is one of these values:
 *
 * - `zip`: the bytes are not a zip archive, or the archive is damaged or crosses a limit.
 *   `detail` gives the zip reader's reason, such as `inflate-error`.
 * - `no-content-types`: the archive has no `[Content_Types].xml` part.
 * - `not-a-workbook`: the package has no workbook part of a known content type. `detail` gives
 *   the content type that was found, or what is missing.
 * - `workbook`: the workbook part passes a limit of the read. `limit` names it:
 *   `workbook.maxSheetNameLength` for a stored sheet name of more than `MAX_STORED_SHEET_NAME`
 *   UTF-16 code units.
 * - `xml`: the XML of the part `partName` is malformed or crosses a limit. `detail` gives the
 *   XML reader's reason.
 */
export type PackageRejection =
  | {
      readonly reason: 'zip';
      readonly detail: string;
      /** With the detail `too-large`, the size limit that the package passed. */
      readonly limit?: 'zip.maxTotalBytes' | 'zip.maxRatio';
    }
  | { readonly reason: 'no-content-types' }
  | { readonly reason: 'not-a-workbook'; readonly detail?: string }
  | {
      readonly reason: 'workbook';
      readonly detail: 'sheet-name-too-long';
      readonly limit: 'workbook.maxSheetNameLength';
    }
  | { readonly reason: 'xml'; readonly partName: string; readonly detail: XmlPartRejection };

export type PackageResult =
  | { readonly ok: true; readonly pkg: XlsxPackage }
  | { readonly ok: false; readonly error: PackageRejection };

export const CONTENT_TYPES_PART = '/[Content_Types].xml';
const decoder = new TextDecoder('utf-8', { fatal: false });

/** The main-part content types that ECMA-376 defines for a workbook. Others are refused. */
const WORKBOOK_TYPES = new Set([CONTENT_TYPE.workbook, CONTENT_TYPE.template].map(asciiFold));

/** Content types the engine models. A part that does not parse refuses the workbook. */
const REQUIRED_TYPES = new Set(
  [CONTENT_TYPE.workbook, CONTENT_TYPE.template, CONTENT_TYPE.worksheet].map(asciiFold)
);

/** Content types the engine models but can do without. A part that does not parse stays bytes. */
const OPTIONAL_TYPES = new Set(
  [CONTENT_TYPE.sharedStrings, CONTENT_TYPE.styles, CONTENT_TYPE.theme, CONTENT_TYPE.drawing].map(
    asciiFold
  )
);

/** The reasons of a part that passes an XML part limit. */
const LIMIT_REASONS: ReadonlySet<XmlPartRejection> = new Set([
  'too-large',
  'too-many-elements',
  'too-deep',
]);

/**
 * The XML part limit that each kept part passed, keyed by the part. An edit keeps the object of
 * each part that it does not change, so the mark stays until an edit replaces the part.
 */
const pastLimit = new WeakMap<PackagePart, XmlPartRejection>();

/**
 * The XML part limit that `part` passed when the read kept it as bytes, as the reason of the
 * bounded reader (`too-large`, `too-many-elements`, or `too-deep`), or undefined.
 */
export function partPastLimit(part: PackagePart): XmlPartRejection | undefined {
  return pastLimit.get(part);
}

/**
 * Whether a part that did not parse refuses the workbook. A shared string table past a limit
 * refuses it too: without the table, each string cell would show no text, as DOCX Editor refuses
 * a part past its limit. The part is found by its content type, also when no relationship uses
 * it: DOCX Editor reads each part with an XML content type, and refuses the file when one does not
 * read (`readOoxmlPackageSteps` in `packages/core/src/store/package/ooxml-package.ts`, commit
 * 13be4797).
 */
function refusesWorkbook(
  role: PartRole,
  reason: XmlPartRejection,
  type: string | undefined
): boolean {
  if (role === 'required') return true;
  const sst = asciiFold(type?.trim() ?? '') === asciiFold(CONTENT_TYPE.sharedStrings);
  return sst && LIMIT_REASONS.has(reason);
}

/**
 * How the reader treats a part.
 *
 * - `required`: parsed; a part that does not parse refuses the workbook.
 * - `optional`: parsed; a part that does not parse is kept as opaque bytes.
 * - `binary`: never parsed. Every XML part the engine does not model is re-emitted
 *   byte-identical instead of normalized.
 */
type PartRole = 'required' | 'optional' | 'binary';

/** The source part of a `.rels` part (`/` for the package), or undefined if the name is not one. */
function relsSource(partName: string): string | undefined {
  // OPC part names compare without ASCII case, so `/_RELS/.RELS` is the package relationships.
  const folded = asciiFold(partName);
  if (folded === '/_rels/.rels') return '/';
  const slash = folded.lastIndexOf('/_rels/');
  if (slash < 0 || !folded.endsWith('.rels')) return undefined;
  const name = partName.slice(slash + 7, -5);
  if (name.length === 0 || name.includes('/')) return undefined;
  return `${partName.slice(0, slash)}/${name}`;
}

function roleByType(contentType: string | undefined): PartRole {
  if (contentType === undefined) return 'binary';
  const type = asciiFold(contentType.trim());
  if (REQUIRED_TYPES.has(type)) return 'required';
  if (OPTIONAL_TYPES.has(type)) return 'optional';
  return 'binary';
}

function partRole(partName: string, types: ContentTypeIndex): PartRole {
  if (asciiFold(partName) === asciiFold(CONTENT_TYPES_PART)) return 'required';
  if (asciiFold(partName).endsWith('.rels')) {
    // Relationships are always parsed. Those of the package and of required parts must parse.
    const source = relsSource(partName);
    if (source === '/') return 'required';
    if (source === undefined) return 'optional';
    return roleByType(lookupContentType(types, source)) === 'required' ? 'required' : 'optional';
  }
  return roleByType(lookupContentType(types, partName));
}

/** Decode the bytes of an XML part: a UTF-16 byte order mark is honored, UTF-8 otherwise. */
export function decodeXml(bytes: Uint8Array): string {
  if (bytes[0] === 0xff && bytes[1] === 0xfe) return new TextDecoder('utf-16le').decode(bytes);
  if (bytes[0] === 0xfe && bytes[1] === 0xff) return new TextDecoder('utf-16be').decode(bytes);
  let text = decoder.decode(bytes);
  if (text.charCodeAt(0) === 0xfeff) text = text.slice(1);
  return text;
}

/** Read a `.xlsx` at the trust boundary. Never throws. */
export function readXlsx(bytes: Uint8Array, limits: ZipLimits = DEFAULT_ZIP_LIMITS): PackageResult {
  const steps = readXlsxSteps(bytes, limits);
  for (;;) {
    const next = steps.next();
    if (next.done) return next.value;
  }
}

/**
 * {@link readXlsx}, yielding between parts and inside a long part. The ZIP read and the content
 * types part come first, in one step each. A caller that stops between yields gives the page a
 * turn. The result is the same either way.
 */
export function* readXlsxSteps(
  bytes: Uint8Array,
  limits: ZipLimits = DEFAULT_ZIP_LIMITS
): Generator<void, PackageResult> {
  const zip = readZip(bytes, limits);
  if (!zip.ok) {
    const limit = zip.reason === 'too-large' ? zip.limit : undefined;
    return { ok: false, error: { reason: 'zip', detail: zip.reason, ...(limit ? { limit } : {}) } };
  }
  // The content types part is matched without ASCII case, as every OPC part name is.
  const ctName = [...zip.entries.keys()].find(
    (name) => asciiFold(name) === asciiFold(CONTENT_TYPES_PART)
  );
  const ctBytes = ctName === undefined ? undefined : zip.entries.get(ctName);
  if (!ctBytes) return { ok: false, error: { reason: 'no-content-types' } };
  const xmlLimits = { maxBytes: XML_HARD_MAX_BYTES };
  const ct = parseXmlPart(ctName!, decodeXml(ctBytes), xmlLimits);
  if (!ct.ok)
    return { ok: false, error: { reason: 'xml', partName: CONTENT_TYPES_PART, detail: ct.reason } };
  const types = contentTypeIndex(ct.part);
  const parts = new Map<string, PackagePart>();
  for (const [partName, data] of zip.entries) {
    yield;
    const role = partRole(partName, types);
    if (role === 'binary') {
      parts.set(partName, { kind: 'binary', partName, bytes: data });
      continue;
    }
    const tree =
      partName === ctName ? ct : yield* parseXmlPartSteps(partName, decodeXml(data), xmlLimits);
    if (tree.ok) parts.set(partName, { kind: 'xml', partName, xml: tree.part });
    else if (refusesWorkbook(role, tree.reason, lookupContentType(types, partName)))
      return { ok: false, error: { reason: 'xml', partName, detail: tree.reason } };
    else {
      const kept: PackagePart = { kind: 'binary', partName, bytes: data };
      if (LIMIT_REASONS.has(tree.reason)) pastLimit.set(kept, tree.reason);
      parts.set(partName, kept);
    }
  }
  const pkg: XlsxPackage = { parts };
  const main = mainWorkbookPart(pkg);
  if (!main) return { ok: false, error: { reason: 'not-a-workbook', detail: 'no main part' } };
  const mainType = contentTypeOf(pkg, main);
  if (!mainType || !WORKBOOK_TYPES.has(asciiFold(mainType.trim())))
    return {
      ok: false,
      error: { reason: 'not-a-workbook', detail: mainType ?? 'no content type' },
    };
  const root = xmlPartOf(pkg, main)?.root;
  if (!root || root.ns !== NS.x || root.local !== 'workbook')
    return {
      ok: false,
      error: { reason: 'not-a-workbook', detail: 'the main part is not x:workbook' },
    };
  if (hasLongSheetName(root))
    return {
      ok: false,
      error: {
        reason: 'workbook',
        detail: 'sheet-name-too-long',
        limit: 'workbook.maxSheetNameLength',
      },
    };
  return { ok: true, pkg };
}

/**
 * The longest sheet name that a read takes (`workbook.maxSheetNameLength`): 32,767 UTF-16 code
 * units, as the file stores the name. The schema types `x:sheet/@name` (`CT_Sheet`, ECMA-376
 * Part 1, 18.2.19) as `ST_Xstring` (22.9.2.19), and the standard states no length limit. The
 * length counts the stored name, so each `_xHHHH_` escape counts as its seven stored code
 * units. The value is the most text that one cell holds (`MAX_CELL_TEXT`), so a sheet name
 * never holds more text than a cell. It is far above the 31 code units that a new name may
 * have (`MAX_SHEET_NAME`), so a stored name that is longer than a new name still opens. Each op
 * on a sheet names it, so the limit also bounds the cost of each sheet name lookup.
 */
export const MAX_STORED_SHEET_NAME = 32_767;

/** Whether a sheet name of a workbook part passes `MAX_STORED_SHEET_NAME`. */
function hasLongSheetName(root: Element): boolean {
  for (const sheets of children(root, 'x:sheets'))
    for (const sheet of children(sheets, 'x:sheet'))
      if ((attr(sheet, 'name')?.length ?? 0) > MAX_STORED_SHEET_NAME) return true;
  return false;
}

// ---------------------------------------------------------------- content types

interface ContentTypeIndex {
  readonly defaults: ReadonlyMap<string, string>;
  readonly overrides: ReadonlyMap<string, string>;
}

const ctCache = new WeakMap<XmlPart, ContentTypeIndex>();

function contentTypeIndex(part: XmlPart): ContentTypeIndex {
  const hit = ctCache.get(part);
  if (hit) return hit;
  const defaults = new Map<string, string>();
  const overrides = new Map<string, string>();
  for (const el of elements(part.root)) {
    if (el.ns !== NS.ct) continue;
    const type = attr(el, 'ContentType');
    if (!type) continue;
    if (el.local === 'Default') {
      const ext = attr(el, 'Extension');
      if (ext) defaults.set(asciiFold(ext), type);
    } else if (el.local === 'Override') {
      // `PartName` is a URI, percent-encoded as the ZIP item name is. Decode it the same way,
      // once, so it matches the canonical part name the reader keys parts by.
      const name = attr(el, 'PartName');
      const norm = name ? normalizePartName(name) : undefined;
      if (norm?.ok) overrides.set(asciiFold(norm.partName), type);
    }
  }
  const index = { defaults, overrides };
  ctCache.set(part, index);
  return index;
}

function lookupContentType(index: ContentTypeIndex, partName: string): string | undefined {
  const over = index.overrides.get(asciiFold(partName));
  if (over) return over;
  const dot = partName.lastIndexOf('.');
  return dot < 0 ? undefined : index.defaults.get(asciiFold(partName.slice(dot + 1)));
}

/** The content type of a part, from `[Content_Types].xml`. */
export function contentTypeOf(pkg: XlsxPackage, partName: string): string | undefined {
  const ct = partOf(pkg, CONTENT_TYPES_PART);
  if (!ct || ct.kind !== 'xml') return undefined;
  return lookupContentType(contentTypeIndex(ct.xml), partName);
}

// ---------------------------------------------------------------- relationships

/** One relationship, resolved against its source part. */
export interface Relationship {
  readonly id: string;
  readonly type: string;
  /** Internal: the resolved part name. External: undefined. */
  readonly target?: string;
  /** External: the raw URI, retained verbatim and NEVER fetched. */
  readonly externalTarget?: string;
  readonly external: boolean;
}

/** The `.rels` part name for a source part (`/` for the package). */
export function relsPartName(source: string): string {
  // The writer and the fixtures use lowercase `_rels`; the reader keys parts as named.
  if (source === '/') return '/_rels/.rels';
  const slash = source.lastIndexOf('/');
  return `${source.slice(0, slash)}/_rels/${source.slice(slash + 1)}.rels`;
}

const relCache = new WeakMap<XmlPart, readonly Relationship[]>();

/** All relationships of a source part. Unsafe or unresolvable targets are dropped. */
export function relationshipsOf(pkg: XlsxPackage, source: string): readonly Relationship[] {
  const part = partOf(pkg, relsPartName(source));
  if (!part || part.kind !== 'xml') return [];
  const raw = relCache.get(part.xml) ?? readRelationships(part.xml, source);
  relCache.set(part.xml, raw);
  // The cache holds targets as the relationship spells them; the package's spelling of each
  // part is resolved here, per package, so a later revision never reads a stale spelling.
  return raw.map((r) =>
    r.external || r.target === undefined
      ? r
      : { ...r, target: equivalentPartName(pkg, r.target) ?? r.target }
  );
}

function readRelationships(part: XmlPart, source: string): readonly Relationship[] {
  const out: Relationship[] = [];
  const owner = source === '/' ? '/x' : source;
  for (const el of children(part.root, { ns: NS.rel, local: 'Relationship' })) {
    const id = attr(el, 'Id');
    const type = attr(el, 'Type');
    const target = attr(el, 'Target');
    if (!id || !type || target === undefined) continue;
    if (attr(el, 'TargetMode') === 'External') {
      const ok = validateExternalTarget(target);
      out.push({ id, type, external: true, ...(ok.ok ? { externalTarget: target } : {}) });
      continue;
    }
    const resolved = resolveInternalTarget(owner, target);
    if (resolved.ok) out.push({ id, type, target: resolved.partName, external: false });
  }
  return out;
}

/** The relationship of a source part with id `id`: the first one with that id. */
export function relationshipById(
  pkg: XlsxPackage,
  source: string,
  id: string
): Relationship | undefined {
  return relationshipsOf(pkg, source).find((r) => r.id === id);
}

/** The internal target of a relationship id, or undefined. */
export function relTarget(pkg: XlsxPackage, source: string, id: string): string | undefined {
  const rel = relationshipById(pkg, source, id);
  return rel && !rel.external ? rel.target : undefined;
}

/** The first internal target of a relationship type. */
export function relTargetByType(
  pkg: XlsxPackage,
  source: string,
  type: string
): string | undefined {
  return relationshipsOf(pkg, source).find((r) => r.type === type && !r.external)?.target;
}

/** `/xl/workbook.xml`, or wherever the package relationships point. */
export function mainWorkbookPart(pkg: XlsxPackage): string | undefined {
  return relTargetByType(pkg, '/', REL.officeDocument);
}

const keyMaps = new WeakMap<ReadonlyMap<string, PackagePart>, Map<string, string>>();

/**
 * The existing part whose name is OPC-equivalent to `partName` (ASCII case-insensitive), or
 * undefined. Part names compare without case (OPC), so a relationship to `sheet1.xml` names a
 * zip entry `Sheet1.xml`. A new part name must not match one: the writer folds case. One key
 * map per part map, built on first use.
 */
export function equivalentPartName(pkg: XlsxPackage, partName: string): string | undefined {
  if (pkg.parts.has(partName)) return partName;
  let map = keyMaps.get(pkg.parts);
  if (!map) {
    map = new Map();
    for (const name of pkg.parts.keys()) {
      const key = partNameKey(name);
      if (!map.has(key)) map.set(key, name);
    }
    keyMaps.set(pkg.parts, map);
  }
  return map.get(partNameKey(partName));
}

/** The part with a name OPC-equivalent to `partName`, or undefined. */
export function partOf(pkg: XlsxPackage, partName: string | undefined): PackagePart | undefined {
  if (!partName) return undefined;
  const exact = pkg.parts.get(partName);
  if (exact) return exact;
  const name = equivalentPartName(pkg, partName);
  return name === undefined ? undefined : pkg.parts.get(name);
}

/** The XML tree of a part, or undefined. */
export function xmlPartOf(pkg: XlsxPackage, partName: string | undefined): XmlPart | undefined {
  const part = partOf(pkg, partName);
  return part && part.kind === 'xml' ? part.xml : undefined;
}

/** The bytes of a binary part, or undefined. */
export function binaryPartOf(
  pkg: XlsxPackage,
  partName: string | undefined
): Uint8Array | undefined {
  const part = partOf(pkg, partName);
  return part && part.kind === 'binary' ? part.bytes : undefined;
}

// ---------------------------------------------------------------- round-trip fingerprints

/**
 * Per-part canonical fingerprints. XML parts hash their canonical form; binaries hash bytes.
 * Two packages with the same map are the same document content, however each was written.
 */
export function packageFingerprints(pkg: XlsxPackage): ReadonlyMap<string, string> {
  const out = new Map<string, string>();
  for (const part of [...pkg.parts.values()].sort((a, b) => (a.partName < b.partName ? -1 : 1)))
    out.set(
      part.partName,
      part.kind === 'xml' ? canonicalFingerprint(part.xml) : bytesFingerprint(part.bytes)
    );
  return out;
}

/** One hash over all part fingerprints: the package's canonical fingerprint. */
export function canonicalPackageFingerprint(pkg: XlsxPackage): string {
  const lines = [...packageFingerprints(pkg)].map(([k, v]) => `${k}\t${v}`).join('\n');
  return sha256FontBytes(new TextEncoder().encode(lines));
}
