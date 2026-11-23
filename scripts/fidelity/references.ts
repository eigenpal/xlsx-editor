// The identity of each fixture's reference PDF, kept in `fixtures/references.json`.
//
// Same role as pptx-editor/scripts/fidelity/reference-identity.ts at 1943d2c9 (EigenPal): a
// reference is current only while the SHA-256 of its workbook and of its PDF match the record.
// A fixture with no record, a stale record, or a PDF that fails the page checks is a gap, never
// a pass. Differences: the page facts are read from the PDF bytes, and each record names the
// number locale of the environment that printed it, because the General number format prints
// with the host's decimal separator.

import { createHash } from 'node:crypto';
import { existsSync, readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';

export const FIXTURE_DIR = join(import.meta.dir, '..', '..', 'fixtures');
export const MANIFEST = join(FIXTURE_DIR, 'references.json');

/**
 * The number locale that the reference PDF printed with: its separators, and the built-in
 * format codes (by id) that the locale replaces (ECMA-376 Part 1, 18.8.30).
 */
export interface NumberLocale {
  readonly decimal: string;
  readonly group: string;
  readonly builtins?: Readonly<Record<string, string>>;
}

export interface ReferenceIdentity {
  readonly workbook: string;
  readonly pdf: string;
  readonly pages: number;
  readonly pageSizePt: readonly [number, number];
  readonly locale: NumberLocale;
}

export interface Manifest {
  readonly version: 1;
  readonly references: Readonly<Record<string, ReferenceIdentity>>;
}

export const sha256 = (bytes: Uint8Array): string =>
  createHash('sha256').update(bytes).digest('hex');

/**
 * The reference manifest at `path`, or an empty one when the file does not exist. Throws when
 * the file is not JSON, or not a manifest of version 1 with a `references` object. Each error
 * names the path of the manifest.
 */
export function readManifest(path: string = MANIFEST): Manifest {
  if (!existsSync(path)) return { version: 1, references: {} };
  let data: unknown;
  try {
    data = JSON.parse(readFileSync(path, 'utf8'));
  } catch (error) {
    throw new Error(`${path}: ${error instanceof Error ? error.message : String(error)}`);
  }
  const manifest = data as Partial<Manifest> | null;
  const references = manifest?.references;
  if (
    manifest?.version !== 1 ||
    typeof references !== 'object' ||
    references === null ||
    Array.isArray(references)
  )
    throw new Error(`${path} is not a reference manifest of version 1`);
  for (const [name, record] of Object.entries(references)) {
    const field = invalidField(record);
    if (field) throw new Error(`${path}: the record of ${name} has no valid ${field}`);
  }
  return manifest as Manifest;
}

/** The first field of a manifest record that does not have its type, or undefined. */
export function invalidField(record: unknown): string | undefined {
  const r = record as Partial<Record<keyof ReferenceIdentity, unknown>> | null;
  if (typeof r !== 'object' || r === null || Array.isArray(r)) return 'record';
  if (typeof r.workbook !== 'string') return 'workbook';
  if (typeof r.pdf !== 'string') return 'pdf';
  if (typeof r.pages !== 'number' || !Number.isInteger(r.pages) || r.pages < 1) return 'pages';
  const size = r.pageSizePt;
  if (!Array.isArray(size) || size.length !== 2 || size.some((n) => typeof n !== 'number'))
    return 'pageSizePt';
  const locale = r.locale as Partial<Record<keyof NumberLocale, unknown>> | null;
  if (typeof locale !== 'object' || locale === null) return 'locale';
  if (typeof locale.decimal !== 'string') return 'locale.decimal';
  if (typeof locale.group !== 'string') return 'locale.group';
  const builtins = locale.builtins;
  if (
    builtins !== undefined &&
    (typeof builtins !== 'object' ||
      builtins === null ||
      Array.isArray(builtins) ||
      Object.values(builtins).some((v) => typeof v !== 'string'))
  )
    return 'locale.builtins';
  return undefined;
}

export function writeManifest(manifest: Manifest): void {
  const sorted = Object.fromEntries(
    Object.entries(manifest.references).sort(([a], [b]) => (a < b ? -1 : 1))
  );
  writeFileSync(MANIFEST, `${JSON.stringify({ version: 1, references: sorted }, null, 2)}\n`);
}

/** The number locale of a fixture's reference, or undefined when it has no record. */
export function referenceLocale(
  name: string,
  manifest: Manifest = readManifest()
): NumberLocale | undefined {
  return Object.hasOwn(manifest.references, name) ? manifest.references[name]!.locale : undefined;
}

/** Page count and the size of page 1, from the PDF's page objects. */
export function pdfPages(bytes: Uint8Array): { pages: number; size: [number, number] } {
  const text = new TextDecoder('latin1').decode(bytes);
  const pages = (text.match(/\/Type\s*\/Page(?![s\w])/g) ?? []).length;
  const box = /\/MediaBox\s*\[\s*([-\d.]+)\s+([-\d.]+)\s+([-\d.]+)\s+([-\d.]+)\s*\]/.exec(text);
  const size: [number, number] = box
    ? [Number(box[3]) - Number(box[1]), Number(box[4]) - Number(box[2])]
    : [0, 0];
  return { pages, size };
}

/** US Letter, the paper of every fixture (scripts/fixtures/builder.ts). */
export const LETTER: readonly [number, number] = [612, 792];

/**
 * One line for each missing workbook or PDF of the fixtures `names` in `dir`, with its next
 * step: build the workbook, or export the PDF (CONTRIBUTING.md, "Reference PDFs").
 */
export function missingFiles(names: readonly string[], dir: string = FIXTURE_DIR): string[] {
  return names.flatMap((name) => [
    ...(existsSync(join(dir, `${name}.xlsx`))
      ? []
      : [`fixtures/${name}.xlsx is missing: build it with \`bun run fixtures:build\`.`]),
    ...(existsSync(join(dir, `${name}.pdf`))
      ? []
      : [
          `fixtures/${name}.pdf is missing: a maintainer makes it (CONTRIBUTING.md, "Reference PDFs").`,
        ]),
  ]);
}

/** Why a fixture's reference is not current, or an empty list. */
export function referenceProblems(name: string, manifest: Manifest = readManifest()): string[] {
  const record = Object.hasOwn(manifest.references, name) ? manifest.references[name] : undefined;
  const workbookPath = join(FIXTURE_DIR, `${name}.xlsx`);
  const pdfPath = join(FIXTURE_DIR, `${name}.pdf`);
  if (!record) return [`${name}: no reference in fixtures/references.json`];
  const missing = missingFiles([name]);
  if (missing.length) return missing.map((line) => `${name}: ${line}`);
  const problems: string[] = [];
  const pdf = new Uint8Array(readFileSync(pdfPath));
  if (sha256(new Uint8Array(readFileSync(workbookPath))) !== record.workbook)
    problems.push(`${name}: the workbook changed after its reference was recorded (stale)`);
  if (sha256(pdf) !== record.pdf) problems.push(`${name}: the PDF changed after it was recorded`);
  const facts = pdfPages(pdf);
  if (facts.pages !== 1) problems.push(`${name}: the reference has ${facts.pages} pages, not 1`);
  if (facts.size[0] !== LETTER[0] || facts.size[1] !== LETTER[1])
    problems.push(`${name}: the reference page is ${facts.size.join('x')} pt, not Letter`);
  return problems;
}
