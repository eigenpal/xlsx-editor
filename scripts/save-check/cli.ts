// The committed half of the save check (design D5a). A maintainer opens each copy
// and records the result; see CONTRIBUTING.md, "Save check".
// `record-edit` records the edited copy of a feature edit, which `bun
// scripts/fidelity/feature-edits.ts --out <dir>` writes.

import { join } from 'node:path';
import { existsSync, readFileSync } from 'node:fs';
import { FIXTURES, fixtureFile } from '../fixtures/catalog.ts';
import { FEATURE_EDITS } from '../fidelity/feature-edits.ts';
import { featureCopyName, runFeatureEdit } from '../fidelity/feature-edit-kit.ts';
import { readFeatureRecords, writeFeatureRecords } from './feature-records.ts';
import {
  fixtureCopies,
  readRecords,
  RESULTS,
  sha256,
  writeCopies,
  writeRecords,
  type SaveResult,
} from './records.ts';

const USAGE = `usage:
  bun run save-check copies <dir>
      write the saved and edited copy of every fixture to <dir>
  bun run save-check record <dir> <fixture> <saved-result> <edited-result>
      record the result of opening the copies in <dir>
  bun run save-check record-edit <dir> <row> <result>
      record the result of opening <dir>/<row>.xlsx, the copy of the feature edit of a support
      matrix row that \`bun scripts/fidelity/feature-edits.ts --out <dir>\` writes
  a copy of a macro-enabled fixture has the extension .xlsm in place of .xlsx
  a result is one of: ${RESULTS.join(', ')}`;

function fail(message: string): never {
  console.error(`${message}\n\n${USAGE}`);
  process.exit(1);
}

const isResult = (r: string): r is SaveResult => (RESULTS as readonly string[]).includes(r);

const [command, ...args] = process.argv.slice(2);
if (command === 'copies' && args.length === 1) {
  console.log(writeCopies(args[0]!).join('\n'));
} else if (command === 'record-edit' && args.length === 3) {
  const [dir, row, result] = args as [string, string, string];
  const recipe = Object.hasOwn(FEATURE_EDITS, row) ? FEATURE_EDITS[row] : undefined;
  if (!recipe)
    fail(
      `no feature edit for the row ${row}: the rows are the keys of scripts/fidelity/feature-edits.ts`
    );
  if (!isResult(result)) fail('unknown result');
  // The record names the bytes that were opened, and those must be the bytes the recipe saves now.
  const run = await runFeatureEdit(recipe);
  if (run.problems.length > 0) fail(`${row}: the recipe fails: ${run.problems.join('; ')}`);
  const copy = join(dir, featureCopyName(row, recipe.fixture));
  if (!existsSync(copy))
    fail(
      `${copy} does not exist: write the copies first (bun scripts/fidelity/feature-edits.ts --out ${dir})`
    );
  const opened = new Uint8Array(readFileSync(copy));
  if (sha256(opened) !== sha256(run.bytes))
    fail(`${row}: the copy in ${dir} is not what the recipe saves now; write the copy again`);
  const records = readFeatureRecords();
  records[row] = { sha256: sha256(opened), result };
  writeFeatureRecords(records);
  console.log(`recorded ${row}: ${result}`);
} else if (command === 'record' && args.length === 4) {
  const [dir, name, saved, edited] = args as [string, string, string, string];
  const fixture = FIXTURES.find((f) => f.name === name);
  if (!fixture) fail(`no fixture named ${name}`);
  if (!isResult(saved) || !isResult(edited)) fail('unknown result');
  // The record names the bytes that were opened, and those must be the bytes the writer emits now.
  const current = fixtureCopies(fixture);
  const opened = {
    saved: new Uint8Array(readFileSync(join(dir, fixtureFile(name, '.saved')))),
    edited: new Uint8Array(readFileSync(join(dir, fixtureFile(name, '.edited')))),
  };
  for (const kind of ['saved', 'edited'] as const)
    if (sha256(opened[kind]) !== sha256(current[kind]))
      fail(
        `${name}: the ${kind} copy in ${dir} is not what the writer emits now; write the copies again`
      );
  const records = readRecords();
  records[name] = {
    saved: { sha256: sha256(opened.saved), result: saved },
    edited: { sha256: sha256(opened.edited), result: edited },
  };
  writeRecords(records);
  console.log(`recorded ${name}: saved ${saved}, edited ${edited}`);
} else {
  fail('unknown command');
}
