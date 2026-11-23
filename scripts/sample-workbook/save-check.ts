// The save check of the sample workbook (`examples/vite/public/sample.xlsx`), in the form of
// the fixture save check (`scripts/save-check/`, CONTRIBUTING.md, "Save check"). The demo
// opens the sample at start, and its users edit and save it, so a saved copy and an edited copy
// of the sample must open with no repair prompt too. `save-check.json` beside this file holds
// the SHA-256 of each copy that the local check opened, and the result.
//
//   bun run sample:save-check copies <dir>
//   bun run sample:save-check record <dir> <saved-result> <edited-result>
//
// A difference from `scripts/save-check/`: the sample is not a fixture, so it has its own
// record file and its own edit. The edit is the one that the demo save tests make (bold on
// Sales A5) and one value, so the local check can read the edited cell back.

import { readFileSync, writeFileSync, mkdirSync } from 'node:fs';
import { join } from 'node:path';
import { formatOps } from '../../packages/core/src/automation/index.ts';
import { readXlsx, saveXlsx, WorkbookDocumentStore } from '../../packages/core/src/store/index.ts';
import { RESULTS, sha256, type SaveRecord, type SaveResult } from '../save-check/records.ts';
import { SAMPLE_BOLD_EDIT, SAMPLE_EDIT } from './edits.ts';

// The zip writer stores entry times as local-time fields (see `scripts/save-check/records.ts`).
// Importing this module, as importing `records.ts`, sets the time zone of the whole process to
// UTC. Import the edits from `edits.ts`, which has no side effects.
process.env.TZ = 'UTC';

const SAMPLE = join(import.meta.dir, '../../examples/vite/public/sample.xlsx');
const SAMPLE_RECORD = join(import.meta.dir, 'save-check.json');

const REMEDY =
  'Run `bun run sample:save-check copies <dir>`, open both copies, then `bun run sample:save-check record <dir> <saved-result> <edited-result>` (CONTRIBUTING.md, "Save check").';

/** The saved and the edited copy of the sample, as the editor writes them. */
export function sampleCopies(): { saved: Uint8Array; edited: Uint8Array } {
  const read = readXlsx(new Uint8Array(readFileSync(SAMPLE)));
  if (!read.ok) throw new Error(`sample: does not open (${read.error.reason})`);
  const bold = formatOps(SAMPLE_BOLD_EDIT.sheet, SAMPLE_BOLD_EDIT.ref, { bold: true });
  if (!bold.ok) throw new Error(`sample: the bold edit is refused (${bold.rejection.detail})`);
  const store = new WorkbookDocumentStore(read.pkg);
  const r = store.transact([
    ...bold.ops,
    {
      op: 'setCellValue',
      sheet: SAMPLE_EDIT.sheet,
      ref: SAMPLE_EDIT.ref,
      value: { type: 'number', value: SAMPLE_EDIT.value },
    },
  ]);
  if (!r.ok) throw new Error(`sample: the edit is refused (${r.rejection.detail})`);
  return { saved: saveXlsx(read.pkg), edited: saveXlsx(store.package) };
}

/** The committed save record of the sample, or undefined when it is missing or does not parse. */
function readSampleRecord(): SaveRecord | undefined {
  try {
    return JSON.parse(readFileSync(SAMPLE_RECORD, 'utf8')) as SaveRecord;
  } catch {
    return undefined;
  }
}

/** Why the sample is not save-verified, or an empty list. Each problem names its remedy. */
export function sampleRecordProblems(record = readSampleRecord()): string[] {
  if (!record) return [`sample: no save-check record. ${REMEDY}`];
  const copies = sampleCopies();
  const problems: string[] = [];
  for (const kind of ['saved', 'edited'] as const) {
    if (record[kind].sha256 !== sha256(copies[kind]))
      problems.push(`sample: the ${kind} copy changed since the check opened it. ${REMEDY}`);
    if (record[kind].result !== 'verified')
      problems.push(`sample: ${kind} copy ${record[kind].result}`);
  }
  return problems;
}

function fail(message: string): never {
  console.error(
    `${message}\n\nusage:\n  bun run sample:save-check copies <dir>\n  bun run sample:save-check record <dir> <saved-result> <edited-result>\n  a result is one of: ${RESULTS.join(', ')}`
  );
  process.exit(1);
}

if (import.meta.main) {
  const [command, ...args] = process.argv.slice(2);
  if (command === 'copies' && args.length === 1) {
    const copies = sampleCopies();
    mkdirSync(args[0]!, { recursive: true });
    writeFileSync(join(args[0]!, 'sample.saved.xlsx'), copies.saved);
    writeFileSync(join(args[0]!, 'sample.edited.xlsx'), copies.edited);
    console.log('sample');
  } else if (command === 'record' && args.length === 3) {
    const [dir, saved, edited] = args as [string, string, string];
    const isResult = (r: string): r is SaveResult => (RESULTS as readonly string[]).includes(r);
    if (!isResult(saved) || !isResult(edited)) fail('unknown result');
    // The record names the bytes that were opened, and those must be what the writer emits now.
    const current = sampleCopies();
    const opened = {
      saved: new Uint8Array(readFileSync(join(dir, 'sample.saved.xlsx'))),
      edited: new Uint8Array(readFileSync(join(dir, 'sample.edited.xlsx'))),
    };
    for (const kind of ['saved', 'edited'] as const)
      if (sha256(opened[kind]) !== sha256(current[kind]))
        fail(`sample: the ${kind} copy in ${dir} is not what the writer emits now`);
    const record: SaveRecord = {
      saved: { sha256: sha256(opened.saved), result: saved },
      edited: { sha256: sha256(opened.edited), result: edited },
    };
    writeFileSync(SAMPLE_RECORD, `${JSON.stringify(record, null, 2)}\n`);
    console.log(`recorded sample: saved ${saved}, edited ${edited}`);
  } else {
    fail('unknown command');
  }
}
