// Record the reference of each named fixture in `fixtures/references.json`: the SHA-256 of the
// workbook and of its PDF, the page count, the page size, and the number locale of the
// environment that printed it. Run it after a maintainer makes a reference PDF
// (CONTRIBUTING.md, "Reference PDFs"). It refuses a PDF with more than one page
// or another page size, and a PDF that draws with a font that it does not embed or that the
// workbook does not name (`foreignReferenceFonts`). The built-in format codes of the locale stay
// from the previous record of the fixture, or come from another record with the same separators.
//
//   bun run fidelity:record-reference --decimal , --group ' ' basic-values [more names...]
//
// Exit codes: 0 recorded, 1 a refused reference PDF, 2 a usage error (`run-args.ts`), a missing
// file, reference fonts that cannot be read, or another error that stops the recorder. Each
// error and each refusal prints one `fidelity:` line. The recorder checks that the workbook and
// the PDF of each fixture exist before it reads any, and prints one line for each missing file.
// It writes `fixtures/references.json` only after it read every named fixture.
//
// Same role as pptx-editor/scripts/fidelity/record-reference.ts at 1943d2c9 (EigenPal).
// Compared with 1943d2c9; later commits 67ae7309 and 06534363 checked: 67ae7309 is already
// equal, and 06534363 does not change this file. 67ae7309 drops the fixture name from the foreign
// font check of the recorder. `foreignReferenceFonts` here never took one.
// Differences:
// - The source stops at the first refused fixture and writes nothing. This recorder prints a
//   refusal line for each refused fixture, and it writes the fixtures that it accepts.
// - The source exits with code 1 for a missing file. Here a missing file is an input error, and
//   it exits with code 2, before the recorder reads any file.
// - The source prints all problems of a fixture in one line. Here the first problem of a fixture
//   refuses it, in one line: `fidelity: <fixture>: <problem>`.
// - A record holds the number locale that the reference PDF prints (`--decimal`, `--group`).

import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { FIXTURES } from '../fixtures/catalog.ts';
import { firstLine } from './chromium.ts';
import { FOREIGN_REFERENCE_REMEDY, familyId, foreignReferenceFonts } from './font-names.ts';
import { FONTS_ENV, postScriptFamilies, referenceFontDirs, workbookFamilies } from './fonts.ts';
import { pdfItems } from './pdf-lines.ts';
import {
  FIXTURE_DIR,
  LETTER,
  missingFiles,
  pdfPages,
  readManifest,
  sha256,
  writeManifest,
  type ReferenceIdentity,
} from './references.ts';
import { RECORD_REFERENCE_USAGE, recordReferenceArgs } from './run-args.ts';

async function main(): Promise<number> {
  const parsed = recordReferenceArgs(
    process.argv.slice(2),
    FIXTURES.map((f) => f.name)
  );
  if ('help' in parsed) {
    console.log(RECORD_REFERENCE_USAGE);
    return 0;
  }
  if ('error' in parsed) {
    console.error(`fidelity: ${parsed.error}`);
    return 2;
  }
  let knownFamilies: Map<string, string>;
  try {
    knownFamilies = postScriptFamilies(referenceFontDirs());
  } catch (error) {
    console.error(
      `fidelity: the reference fonts in ${FONTS_ENV} cannot be read: ${firstLine(error)}`
    );
    return 2;
  }
  const missing = missingFiles(parsed.names);
  for (const line of missing) console.error(`fidelity: ${line}`);
  if (missing.length) return 2;
  const separators = { decimal: parsed.decimal, group: parsed.group };
  const manifest = readManifest();
  const builtinsOf = (name: string) => {
    const own = Object.hasOwn(manifest.references, name) ? manifest.references[name] : undefined;
    const same = (r: ReferenceIdentity | undefined) =>
      r?.locale.decimal === separators.decimal && r.locale.group === separators.group;
    if (same(own) && own!.locale.builtins) return own!.locale.builtins;
    return Object.values(manifest.references).find((r) => same(r) && r.locale.builtins)?.locale
      .builtins;
  };
  const references: Record<string, ReferenceIdentity> = { ...manifest.references };
  let failed = false;
  for (const name of parsed.names) {
    const pdf = new Uint8Array(readFileSync(join(FIXTURE_DIR, `${name}.pdf`)));
    const facts = pdfPages(pdf);
    if (facts.pages === 0) {
      console.error(
        `fidelity: ${name}: no page objects found; the PDF may use object streams. Not recorded`
      );
      failed = true;
      continue;
    }
    const xlsx = new Uint8Array(readFileSync(join(FIXTURE_DIR, `${name}.xlsx`)));
    const workbook = sha256(xlsx);
    const old = Object.hasOwn(references, name) ? references[name] : undefined;
    if (old && old.workbook !== workbook && old.pdf === sha256(pdf)) {
      console.error(
        `fidelity: ${name}: the workbook changed but the PDF did not; export the PDF again`
      );
      failed = true;
      continue;
    }
    if (facts.pages !== 1 || facts.size[0] !== LETTER[0] || facts.size[1] !== LETTER[1]) {
      console.error(
        `fidelity: ${name}: ${facts.pages} pages of ${facts.size.join('x')} pt; not recorded`
      );
      failed = true;
      continue;
    }
    const named = new Set([...workbookFamilies(xlsx)].map(familyId));
    const foreign = foreignReferenceFonts((await pdfItems(pdf)).fonts, named, knownFamilies);
    if (foreign.length) {
      console.error(
        `fidelity: ${name}: the PDF draws with fonts that it does not embed or that the workbook does not name: ${foreign.join(', ')}; ${FOREIGN_REFERENCE_REMEDY}. Not recorded`
      );
      failed = true;
      continue;
    }
    references[name] = {
      workbook,
      pdf: sha256(pdf),
      pages: facts.pages,
      pageSizePt: facts.size,
      locale: { ...separators, ...(builtinsOf(name) ? { builtins: builtinsOf(name) } : {}) },
    };
    console.log(`recorded ${name}`);
  }
  writeManifest({ version: 1, references });
  return failed ? 1 : 0;
}

/**
 * `bun run fidelity:record-reference` (`cli.ts record-reference`). Any other error, such as a
 * file that cannot be read or a PDF that cannot be parsed, stops the recorder in one line,
 * before it writes the manifest.
 */
export function command(): Promise<number> {
  return main().catch((error: unknown) => {
    console.error(`fidelity: the recorder stopped: ${firstLine(error)}`);
    return 2;
  });
}

if (import.meta.main) process.exit(await command());
