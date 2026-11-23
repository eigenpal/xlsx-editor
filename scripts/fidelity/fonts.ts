// The reference fonts of the fidelity harness (design D9). They come from the machine, through
// the environment variable `XLSX_REFERENCE_FONTS`: one or more folders, separated by the path
// list separator, that hold `.ttf`, `.otf`, or `.ttc` files. They are never committed. Without
// the variable, the harness measures and paints with the metric-compatible substitutes of
// `@docx-editor.dev/fonts` and says so in the report.
//
// Same role as pptx-editor/scripts/fidelity/local-fonts.ts at 1943d2c9 (EigenPal), reduced to
// what a workbook needs: the family names that a fixture's styles and theme name, and the faces
// of those families. Fonts are file data: the name table is read with bounds checks
// (`font-names.ts`).

import { createHash } from 'node:crypto';
import { existsSync, readdirSync, readFileSync, statSync } from 'node:fs';
import { delimiter, extname, join } from 'node:path';
import { unzipSync, strFromU8 } from 'fflate';
import { familyName, macStyle, postScriptName, u16, u32 } from './font-names.ts';

export const FONTS_ENV = 'XLSX_REFERENCE_FONTS';

/** One face of a font file. */
export interface LocalFace {
  readonly family: string;
  readonly bold: boolean;
  readonly italic: boolean;
  readonly file: string;
  /** The face's own bytes (one face of a collection is cut out as a font file of its own). */
  readonly bytes: Uint8Array;
}

/** The most faces read from one collection file. */
const MAX_COLLECTION_FACES = 64;

/** The offsets of the fonts in a file: one, or each face of a `ttcf` collection. */
function faceOffsets(b: Uint8Array): number[] {
  if (u32(b, 0) !== 0x74746366) return [0];
  const n = Math.min(u32(b, 8), MAX_COLLECTION_FACES);
  return Array.from({ length: n }, (_, i) => u32(b, 12 + i * 4));
}

/** One face of a collection as a font file of its own: its table directory and tables. */
function cutFace(b: Uint8Array, base: number): Uint8Array {
  if (base === 0) return b;
  const n = Math.min(u16(b, base + 4), 512);
  const tables: { r: number; off: number; len: number }[] = [];
  for (let i = 0; i < n; i += 1) {
    const r = base + 12 + i * 16;
    tables.push({ r, off: u32(b, r + 8), len: u32(b, r + 12) });
  }
  const header = 12 + n * 16;
  let size = header;
  for (const t of tables) size += (t.len + 3) & ~3;
  const out = new Uint8Array(size);
  out.set(b.subarray(base, base + header), 0);
  let at = header;
  tables.forEach((t, i) => {
    out.set(b.subarray(t.off, t.off + t.len), at);
    const r = 12 + i * 16 + 8;
    out[r] = at >>> 24;
    out[r + 1] = (at >>> 16) & 255;
    out[r + 2] = (at >>> 8) & 255;
    out[r + 3] = at & 255;
    at += (t.len + 3) & ~3;
  });
  return out;
}

/** The folders of `XLSX_REFERENCE_FONTS`, or an empty list. */
export function referenceFontDirs(env: NodeJS.ProcessEnv = process.env): string[] {
  const raw = env[FONTS_ENV];
  if (!raw) return [];
  return raw.split(delimiter).filter((d) => d.length > 0 && existsSync(d));
}

/** Every face in the reference font folders whose family is in `families` (lower case). */
export function referenceFaces(
  dirs: readonly string[],
  families: ReadonlySet<string>
): LocalFace[] {
  const out: LocalFace[] = [];
  for (const dir of dirs)
    for (const name of readdirSync(dir).sort()) {
      if (!['.ttf', '.otf', '.ttc'].includes(extname(name).toLowerCase())) continue;
      const file = join(dir, name);
      if (!statSync(file).isFile()) continue;
      const bytes = new Uint8Array(readFileSync(file));
      for (const base of faceOffsets(bytes)) {
        const family = familyName(bytes, base);
        if (!family || !families.has(family.toLowerCase())) continue;
        out.push({ family, ...macStyle(bytes, base), file: name, bytes: cutFace(bytes, base) });
      }
    }
  return out;
}

/**
 * The family (name id 1, `familyName`) of each face in the reference font folders, by its
 * PostScript name (name id 6), as PPTX Editor's `postScriptFamilies` reads it: the name that a
 * PDF writes for the face, resolved through the face's own name table.
 */
export function postScriptFamilies(dirs: readonly string[]): Map<string, string> {
  const out = new Map<string, string>();
  for (const dir of dirs)
    for (const name of readdirSync(dir).sort()) {
      if (!['.ttf', '.otf', '.ttc'].includes(extname(name).toLowerCase())) continue;
      const file = join(dir, name);
      if (!statSync(file).isFile()) continue;
      const bytes = new Uint8Array(readFileSync(file));
      for (const base of faceOffsets(bytes)) {
        const family = familyName(bytes, base);
        const postScript = postScriptName(bytes, base);
        if (family && postScript) out.set(postScript, family);
      }
    }
  return out;
}

/** The font families that a workbook's style and theme parts name. */
export function workbookFamilies(xlsx: Uint8Array): Set<string> {
  const parts = unzipSync(xlsx, {
    filter: (f) => f.name === 'xl/styles.xml' || f.name.startsWith('xl/theme/'),
  });
  const families = new Set<string>();
  for (const bytes of Object.values(parts)) {
    const xml = strFromU8(bytes);
    for (const m of xml.matchAll(/<(?:\w+:)?name\s+val="([^"]{1,64})"/g)) families.add(m[1]);
    for (const m of xml.matchAll(/<a:latin\s+typeface="([^"]{1,64})"/g)) families.add(m[1]);
  }
  return new Set([...families].filter((f) => f.length > 0).map((f) => f.toLowerCase()));
}

/** A short identity of a set of faces, for the report: families and a hash of the bytes. */
export function facesIdentity(faces: readonly LocalFace[]): string {
  if (faces.length === 0) return 'none';
  const hash = createHash('sha256');
  const families = new Set<string>();
  for (const f of [...faces].sort((a, b) => (a.file < b.file ? -1 : 1))) {
    hash.update(f.bytes);
    families.add(f.family);
  }
  return `${[...families].sort().join(', ')} (${hash.digest('hex').slice(0, 12)})`;
}
