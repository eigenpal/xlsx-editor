// The names of a font face, for the fidelity harness: the family and style of a font file
// (`fonts.ts`, `page/main.ts`), the PostScript name that a PDF writes for it, and the family of
// a reference PDF font (`foreignReferenceFonts`).
// Fonts and PDFs are file data: tables are read with bounds checks. No Node.js import, so the
// harness page uses it too.
//
// Same role as pptx-editor/scripts/fidelity/local-fonts.ts at 1943d2c9 (EigenPal): it reads the
// family (name id 1) of a font file from its name table.
//
// Compared with 1943d2c9; later commits 67ae7309 and 06534363 checked: 67ae7309 is already
// equal, and 06534363 does not change this file. 67ae7309 removes the fixtures that may embed any
// font (`NON_OPEN_FONT_DECKS`) from `foreignReferenceFonts`. This file never had such a list:
// every reference PDF embeds only fonts of the families that its workbook names.

export const u16 = (b: Uint8Array, o: number): number =>
  o + 2 <= b.length ? (b[o]! << 8) | b[o + 1]! : 0;
export const u32 = (b: Uint8Array, o: number): number =>
  o + 4 <= b.length ? ((b[o]! << 24) | (b[o + 1]! << 16) | (b[o + 2]! << 8) | b[o + 3]!) >>> 0 : 0;

/** The offset and length of a table of the font at `base`. */
export function table(b: Uint8Array, base: number, tag: string): [number, number] | undefined {
  const n = Math.min(u16(b, base + 4), 512);
  for (let i = 0; i < n; i += 1) {
    const r = base + 12 + i * 16;
    if (String.fromCharCode(b[r]!, b[r + 1]!, b[r + 2]!, b[r + 3]!) !== tag) continue;
    const off = u32(b, r + 8);
    const len = u32(b, r + 12);
    return off + len <= b.length ? [off, len] : undefined;
  }
  return undefined;
}

/**
 * The English strings of the name table of the font at `base`, by name id, for the ids in
 * `ids`: Windows (platform 3, US English) or, unless `windowsOnly`, Macintosh (platform 1,
 * Roman) records.
 */
function nameStrings(
  b: Uint8Array,
  base: number,
  ids: readonly number[],
  windowsOnly = false
): Map<number, string> {
  const found = new Map<number, string>();
  const t = table(b, base, 'name');
  if (!t) return found;
  const [off, len] = t;
  const count = Math.min(u16(b, off + 2), 4096);
  const strings = off + u16(b, off + 4);
  for (let i = 0; i < count; i += 1) {
    const r = off + 6 + i * 12;
    if (r + 12 > off + len) break;
    const platform = u16(b, r);
    const lang = u16(b, r + 4);
    const id = u16(b, r + 6);
    const size = Math.min(u16(b, r + 8), 512);
    const at = strings + u16(b, r + 10);
    if (!ids.includes(id) || found.has(id) || at + size > b.length) continue;
    let s = '';
    if (platform === 3 && lang === 0x409)
      for (let k = 0; k + 1 < size; k += 2) s += String.fromCharCode(u16(b, at + k));
    else if (platform === 1 && lang === 0 && !windowsOnly)
      for (let k = 0; k < size; k += 1) s += String.fromCharCode(b[at + k]!);
    else continue;
    found.set(id, s);
  }
  return found;
}

/**
 * The family name (name id 1, platform 3, US English) of the font at `base`, as PPTX Editor's
 * `names()` in `local-fonts.ts` reads it. A cell font names a face by this family: `Calibri
 * Light` is a family of its own, though its typographic family (name id 16) is `Calibri`.
 */
export function familyName(b: Uint8Array, base = 0): string | undefined {
  return nameStrings(b, base, [1], true).get(1);
}

/** The PostScript name (name id 6) of the font at `base`: the name a PDF writes for it. */
export function postScriptName(b: Uint8Array, base = 0): string | undefined {
  return nameStrings(b, base, [6]).get(6);
}

/** Bold and italic from `head.macStyle`. */
export function macStyle(b: Uint8Array, base = 0): { bold: boolean; italic: boolean } {
  const t = table(b, base, 'head');
  const mac = t ? u16(b, t[0] + 44) : 0;
  return { bold: (mac & 1) !== 0, italic: (mac & 2) !== 0 };
}

/** The name of a PDF font without its subset prefix (`AAAAAC+Calibri` is `Calibri`). */
export const baseFontName = (name: string): string => name.replace(/^[A-Z]{6}\+/, '');

/** A face of the harness page: the PostScript name that a PDF writes for it. */
export interface PageFace {
  readonly postScript: string;
}

/** Style words of a PostScript name that are not part of the family. */
const STYLE_WORDS = /^(Regular|Roman|Book|Normal|Bold|Italic|Oblique)$/;

/** A family name compared without case and spaces: `Times New Roman` is `timesnewroman`. */
export const familyId = (family: string): string => family.toLowerCase().replace(/\s+/g, '');

/**
 * The family of a PDF font from its PostScript name alone, for a name that no font file of the
 * harness resolves (`foreignReferenceFonts`). `Calibri-Bold` and `Calibri,Bold` (the form of a
 * font that a PDF names by its family) are Calibri; `Calibri-Light` is Calibri Light. The vendor
 * suffixes `MT`, `PSMT`, and `PS` are not part of the family: `ArialMT` is Arial, and
 * `TimesNewRomanPS-BoldMT` is Times New Roman. Words of the family part that the name joins
 * (`CourierNew`) are split at each capital that follows a lower-case letter. The words of the
 * style part other than those of `STYLE_WORDS` (regular, roman, book, normal, bold, italic, and
 * oblique) stay in the family.
 */
export function referenceFamily(postScript: string): string {
  const name = baseFontName(postScript);
  const comma = name.indexOf(',');
  const at = comma > 0 ? comma : name.lastIndexOf('-');
  const vendor = (t: string): string => t.replace(/(PS)?MT$/, '').replace(/PS$/, '');
  const family = vendor(at > 0 ? name.slice(0, at) : name).replace(/([a-z])([A-Z])/g, '$1 $2');
  const suffix = vendor(at > 0 ? name.slice(at + 1) : '');
  const rest = (suffix.match(/[A-Z][a-z]*/g) ?? []).filter((w) => !STYLE_WORDS.test(w));
  return [family, ...rest].join(' ');
}

/** The remedy for a reference PDF with a foreign font. */
export const FOREIGN_REFERENCE_REMEDY =
  'export the reference PDF again with the fonts of the workbook embedded, and record it with `bun run fidelity:record-reference`';

/** A font of a PDF: its PostScript name, and whether the PDF embeds it. */
export interface PdfFont {
  readonly name: string;
  readonly embedded: boolean;
}

/**
 * The fonts of a reference PDF that are not fonts of its workbook, as PPTX Editor's
 * `foreignReferenceFonts` requires: each must be embedded, and its family must be a family that
 * the workbook names (`named`, by `familyId`). A PDF printed with another font measures that
 * font, not the workbook. The family comes from the name table of a font file of the harness
 * that has the same PostScript name (`known`, from `postScriptFamilies`), as in the source.
 * Difference: the reference fonts are not in the repository and CI has none, so a name that no
 * file resolves takes its family from the name itself (`referenceFamily`), where the source
 * refuses it.
 */
export function foreignReferenceFonts(
  fonts: readonly PdfFont[],
  named: ReadonlySet<string>,
  known: ReadonlyMap<string, string> = new Map()
): string[] {
  return fonts
    .filter((f) => {
      const base = baseFontName(f.name);
      const family = known.get(base) ?? referenceFamily(base);
      return !f.embedded || !named.has(familyId(family));
    })
    .map((f) => (f.embedded ? f.name : `${f.name} (not embedded)`));
}
