// Corpus census: how the engine opens, edits, and saves a set of workbooks, and which
// SpreadsheetML features they use.
//
//   bun run fidelity:census -- --dir <corpus folder>       # or XLSX_CORPUS_DIR=<folder>
//   bun run fidelity:census -- --dir fixtures              # the synthetic fixtures
//   bun run fidelity:census -- --dir <folder> --out <ignored folder> --fresh
//
// `--out` names a folder, outside every git work tree or ignored by git. `--fresh` replaces the
// results of an earlier run there; without it, a rerun stops with a refusal. One workbook may
// take `CENSUS_WORKBOOK_MS` (120 s).
//
// Reports COUNTS, CODES, AND FEATURE NAMES ONLY. It never prints or writes workbook text, file
// names, sheet names, or number format codes. Workbooks are named workbook1 to workbookN, sorted
// by their path in the folder. The output goes to `local/corpus/census/` (git ignores `local/`):
// `results.jsonl` (one line for each workbook), `run.json`, and `census.md`.
//
// After pptx-editor/scripts/fidelity/census.ts at 412d750f (EigenPal). Differences, each for a
// reason:
// - It runs on the memory-safe corpus runner (`scripts/corpus-runner/runner.ts`): one child for
//   each batch, and a watchdog at 3 GiB. PPTX Editor's census reads every deck in one process.
// - Each workbook is read by the engine's bounded reader, not by a separate zip and XML parser,
//   so the census also counts what the engine refuses: open refusals, refused edit probes, and
//   saved copies that do not reopen or lose content.
// - The corpus folder has no default: `--dir` or `XLSX_CORPUS_DIR` names it. The census reads its
//   subfolders too, skips hidden entries, links, and owner files (`~$name.xlsx`), and refuses a
//   missing or empty folder with exit code 1. PPTX Editor reads only the top folder and passes
//   with "Nothing to count" when the folder is missing: a census of nothing is a gap here.
// - `--out` names a folder for the three output files of the runner, not a `.md` file: the
//   runner writes its results there as each workbook ends.
// - The feature counts are every element of the workbook part and the worksheet parts, the
//   content types, and the cell types: the parts that the engine models. PPTX Editor counts
//   every element of a deck. An element name counts only when the ECMA-376 schema of its
//   namespace declares it, because an author can put any name in a namespace. A content type
//   that ECMA-376 does not define for a workbook part, other than `application/xml`,
//   `application/octet-stream`, and the common image types, is `(other)`, not the text of the
//   package.
//
// A census cannot show a repair prompt. The save check (CONTRIBUTING.md, "Save check") covers
// repair prompts; schema validation is `bun run validate:xsd:corpus`.

import { readFileSync, rmSync, writeFileSync } from 'node:fs';
import { join, resolve } from 'node:path';
import {
  contentTypeOf,
  NS,
  readXlsx,
  saveXlsx,
  sheetIndex,
  workbookIndex,
  workbookOracles,
  WorkbookDocumentStore,
  xmlPartOf,
  type Element,
  type XlsxPackage,
} from '../../packages/core/src/store/index.ts';
import { refusalCode } from '../corpus-runner/open-refusal.ts';
import {
  checkFresh,
  checkOut,
  corpusDir,
  defaultOut,
  flag,
  listWorkbooks,
  ROOT,
  runCorpus,
  type CorpusJob,
} from '../corpus-runner/runner.ts';

/** The time that one workbook may take. A stop is recorded, never a pass. */
export const CENSUS_WORKBOOK_MS = 120_000;

type Counts = Record<string, number>;

const bump = (counts: Counts, key: string, n = 1): void => {
  counts[key] = (counts[key] ?? 0) + n;
};

/** The short prefix of a namespace that the engine knows, or `other`. */
const PREFIX = new Map<string, string>(Object.entries(NS).map(([prefix, uri]) => [uri, prefix]));

/** The ECMA-376 schema of each namespace whose element names the census keeps. */
const ELEMENT_SCHEMAS: Readonly<Record<string, string>> = { x: 'sml.xsd' };

/** The elements of markup compatibility (ECMA-376 Part 3), which no schema of Part 4 declares. */
const MC_ELEMENTS = ['AlternateContent', 'Choice', 'Fallback'];

const elementNames = new Map<string, ReadonlySet<string>>();

/** Counts `root` and every element under it, as PPTX Editor counts every element of a slide. */
function countElements(root: Element, counts: Counts): void {
  // A list, not recursion: the depth of a part is the author's choice.
  const stack: Element[] = [root];
  for (let el = stack.pop(); el; el = stack.pop()) {
    bump(counts, featureName(el));
    for (const child of el.children) if (child.type === 'element') stack.push(child);
  }
}

/** The element names that the schema of `prefix` declares, read once. */
function schemaElements(prefix: string): ReadonlySet<string> {
  let names = elementNames.get(prefix);
  if (!names) {
    const schema = ELEMENT_SCHEMAS[prefix];
    const text = schema
      ? readFileSync(join(ROOT, 'reference/ecma-376/part4/transitional-schemas', schema), 'utf8')
      : '';
    names = new Set(
      prefix === 'mc'
        ? MC_ELEMENTS
        : [...text.matchAll(/<xsd:element name="([A-Za-z][A-Za-z0-9_]*)"/g)].map((m) => m[1]!)
    );
    elementNames.set(prefix, names);
  }
  return names;
}

/**
 * A feature name from an element: its known prefix and a local name that the schema of its
 * namespace declares, never free text. The author can put any element name in a namespace, so
 * any other name is `(name)`, and an element of a namespace that the engine does not know is
 * `other:(name)`.
 */
function featureName(el: Element): string {
  const prefix = PREFIX.get(el.ns);
  if (prefix === undefined) return 'other:(name)';
  return `${prefix}:${schemaElements(prefix).has(el.local) ? el.local : '(name)'}`;
}

const SML = 'application/vnd.openxmlformats-officedocument.spreadsheetml.';
const DML = 'application/vnd.openxmlformats-officedocument.drawingml.';
const OFFICE = 'application/vnd.openxmlformats-officedocument.';
const OPC = 'application/vnd.openxmlformats-package.';

/**
 * The content types that the census names: the ones that ECMA-376 defines for the parts of a
 * workbook (Part 1, clauses 11 to 15, and Part 2), the macro-enabled main parts and macro part,
 * `application/xml`, `application/octet-stream`, and common images. A list, not a pattern, because
 * the author of a package can put any words after a known prefix.
 */
const KNOWN_CONTENT_TYPES: ReadonlyMap<string, string> = new Map(
  [
    ...[
      'sheet.main+xml',
      'template.main+xml',
      'worksheet+xml',
      'chartsheet+xml',
      'dialogsheet+xml',
      'styles+xml',
      'sharedStrings+xml',
      'calcChain+xml',
      'comments+xml',
      'connections+xml',
      'externalLink+xml',
      'pivotTable+xml',
      'pivotCacheDefinition+xml',
      'pivotCacheRecords+xml',
      'queryTable+xml',
      'revisionHeaders+xml',
      'revisionLog+xml',
      'sheetMetadata+xml',
      'table+xml',
      'tableSingleCells+xml',
      'userNames+xml',
      'volatileDependencies+xml',
      'customProperty',
      'printerSettings',
    ].map((sub) => SML + sub),
    ...[
      'chart+xml',
      'chartshapes+xml',
      'diagramColors+xml',
      'diagramData+xml',
      'diagramLayout+xml',
      'diagramStyle+xml',
    ].map((sub) => DML + sub),
    ...[
      'theme+xml',
      'themeOverride+xml',
      'drawing+xml',
      'vmlDrawing',
      'oleObject',
      'extended-properties+xml',
      'custom-properties+xml',
      'customXmlProperties+xml',
    ].map((sub) => OFFICE + sub),
    ...[
      'relationships+xml',
      'core-properties+xml',
      'digital-signature-origin',
      'digital-signature-xmlsignature+xml',
      'digital-signature-certificate',
    ].map((sub) => OPC + sub),
    // The macro-enabled main parts and their macro part, which ECMA-376 does not list.
    'application/vnd.ms-excel.sheet.macroEnabled.main+xml',
    'application/vnd.ms-excel.template.macroEnabled.main+xml',
    'application/vnd.ms-office.vbaProject',
    'application/vnd.ms-excel.macrosheet+xml',
    'application/vnd.ms-excel.intlmacrosheet+xml',
    'application/xml',
    'application/octet-stream',
    ...['png', 'jpeg', 'gif', 'bmp', 'tiff', 'x-emf', 'x-wmf', 'svg+xml'].map(
      (sub) => `image/${sub}`
    ),
  ].map((type) => [type.toLowerCase(), type])
);

/**
 * A content type of the list, in the form that the list gives, or `(other)`: the author of a
 * package can choose any other text as a content type. Content types compare without case.
 */
export const contentTypeName = (type: string | undefined): string =>
  type === undefined ? '(none)' : (KNOWN_CONTENT_TYPES.get(type.trim().toLowerCase()) ?? '(other)');

/** The cell types of ECMA-376 (`ST_CellType`), and `d`. Any other value counts as `(other)`. */
const CELL_TYPES = new Set(['b', 'd', 'e', 'inlineStr', 'n', 's', 'str']);

/** One `results.jsonl` line of the census: counts and codes only. */
export interface WorkbookCensus {
  readonly open: 'ok' | 'refused';
  readonly refusal?: string;
  readonly sheets?: number;
  readonly worksheets?: number;
  readonly cells?: number;
  readonly formulas?: number;
  /** `ok`, or the code of the refused value edit in the first worksheet. */
  readonly edit?: string;
  /** `ok`, or the code of the refused row insert in the first worksheet. */
  readonly structure?: string;
  /** `same`: the saved copy reopens to the same content. `changed` or `unreadable` otherwise. */
  readonly save?: 'same' | 'changed' | 'unreadable';
  /** The same check for the copy after the edit probes. */
  readonly editedSave?: 'same' | 'changed' | 'unreadable' | 'none';
  readonly sheetProblems?: Counts;
  readonly contentTypes?: Counts;
  readonly workbookElements?: Counts;
  readonly worksheetElements?: Counts;
  readonly cellTypes?: Counts;
}

/** The writer of the census. Tests pass a faulty one. */
export type CensusSave = (pkg: XlsxPackage) => Uint8Array;

/** Whether a save of `pkg` reopens to the content that it holds. */
function saveOutcome(pkg: XlsxPackage, save: CensusSave): 'same' | 'changed' | 'unreadable' {
  try {
    const oracles = workbookOracles(pkg, save);
    return oracles.fingerprint === oracles.digest ? 'same' : 'changed';
  } catch {
    return 'unreadable';
  }
}

/** The census of one workbook, from its bytes. The task of the corpus runner. */
export function censusWorkbook(bytes: Uint8Array, save: CensusSave = saveXlsx): WorkbookCensus {
  const read = readXlsx(bytes);
  if (!read.ok) return { open: 'refused', refusal: refusalCode(read.error) };
  const pkg = read.pkg;
  const contentTypes: Counts = {};
  for (const part of pkg.parts.values())
    bump(contentTypes, contentTypeName(contentTypeOf(pkg, part.partName)));
  const index = workbookIndex(pkg);
  const workbookElements: Counts = {};
  const root = index ? xmlPartOf(pkg, index.mainPart)?.root : undefined;
  if (root) countElements(root, workbookElements);
  const worksheetElements: Counts = {};
  const cellTypes: Counts = {};
  const sheetProblems: Counts = {};
  let cells = 0;
  let formulas = 0;
  const worksheets = (index?.sheets ?? []).filter((s) => s.kind === 'worksheet');
  for (const sheet of worksheets) {
    const part = xmlPartOf(pkg, sheet.partName);
    if (!part) continue;
    countElements(part.root, worksheetElements);
    const sheetIdx = sheetIndex(part);
    cells += sheetIdx.cells.size;
    formulas += sheetIdx.formulas;
    for (const problem of sheetIdx.problems) bump(sheetProblems, problem.code);
    for (const cell of sheetIdx.cells.values()) {
      const t = cell.el.attrs.find((a) => a.local === 't' && a.ns === '')?.value ?? 'n';
      bump(cellTypes, CELL_TYPES.has(t) ? t : '(other)');
    }
  }
  const first = worksheets[0];
  let edit = 'none';
  let structure = 'none';
  let editedSave: WorkbookCensus['editedSave'] = 'none';
  if (first) {
    const store = new WorkbookDocumentStore(pkg);
    const value = store.transact([
      {
        op: 'setCellValue',
        sheet: first.name,
        ref: 'A1',
        value: { type: 'string', value: 'Census & <edit>' },
      },
    ]);
    edit = value.ok ? 'ok' : value.rejection.code;
    const rows = store.transact([{ op: 'insertRows', sheet: first.name, start: 1, count: 1 }]);
    structure = rows.ok ? 'ok' : rows.rejection.code;
    if (value.ok || rows.ok) editedSave = saveOutcome(store.package, save);
  }
  return {
    open: 'ok',
    sheets: index?.sheets.length ?? 0,
    worksheets: worksheets.length,
    cells,
    formulas,
    edit,
    structure,
    save: saveOutcome(pkg, save),
    editedSave,
    sheetProblems,
    contentTypes,
    workbookElements,
    worksheetElements,
    cellTypes,
  };
}

/** One line of `results.jsonl`: a census, a stop, or an error of the task. */
type ResultLine = WorkbookCensus & {
  readonly id: string;
  readonly stopped?: string;
  readonly error?: string;
};

const COUNT_FIELDS = [
  'sheetProblems',
  'contentTypes',
  'workbookElements',
  'worksheetElements',
  'cellTypes',
] as const;

/** A table of counts. `unit` names what each count counts, for example `Workbooks`. */
function table(title: string, counts: Counts, unit: string, limit = 80): string {
  const rows = Object.entries(counts)
    .sort((a, b) => b[1] - a[1] || (a[0] < b[0] ? -1 : 1))
    .slice(0, limit);
  if (rows.length === 0) return `## ${title}\n\nNone.\n`;
  const body = rows
    .map(([k, v]) => `| \`${k.replace(/\\/g, '\\\\').replace(/\|/g, '\\|')}\` | ${v} |`)
    .join('\n');
  return `## ${title}\n\n| Value | ${unit} |\n| --- | ---: |\n${body}\n`;
}

/**
 * The census as Markdown. `perWorkbook` adds one row for each workbook; it stays off for a
 * private corpus, because the counts of one workbook together can identify it.
 */
export function renderCensus(
  lines: readonly ResultLine[],
  label: string,
  perWorkbook: boolean
): string {
  const outcome: Counts = {};
  const refusals: Counts = {};
  const edits: Counts = {};
  const structures: Counts = {};
  const saves: Counts = {};
  const editedSaves: Counts = {};
  const sums: Record<(typeof COUNT_FIELDS)[number], Counts> = {
    sheetProblems: {},
    contentTypes: {},
    workbookElements: {},
    worksheetElements: {},
    cellTypes: {},
  };
  let sheets = 0;
  let cells = 0;
  let formulas = 0;
  for (const line of lines) {
    if (line.stopped) bump(outcome, `stopped (${line.stopped})`);
    else if (line.error) bump(outcome, 'task error');
    else if (line.open === 'refused') {
      bump(outcome, 'refused at open');
      bump(refusals, line.refusal ?? '(none)');
    } else {
      bump(outcome, 'opened');
      sheets += line.sheets ?? 0;
      cells += line.cells ?? 0;
      formulas += line.formulas ?? 0;
      bump(edits, line.edit ?? '(none)');
      bump(structures, line.structure ?? '(none)');
      bump(saves, line.save ?? '(none)');
      bump(editedSaves, line.editedSave ?? '(none)');
      for (const field of COUNT_FIELDS)
        for (const [k, v] of Object.entries(line[field] ?? {})) bump(sums[field], k, v);
    }
  }
  const perRows = perWorkbook
    ? `## Per workbook\n\n| Workbook | Outcome | Sheets | Cells | Edit | Insert row | Save |\n| --- | --- | ---: | ---: | --- | --- | --- |\n${lines
        .map(
          (l) =>
            `| ${l.id} | ${l.stopped ? `stopped (${l.stopped})` : l.error ? 'task error' : l.open === 'refused' ? `refused (${l.refusal})` : 'opened'} | ${l.sheets ?? ''} | ${l.cells ?? ''} | ${l.edit ?? ''} | ${l.structure ?? ''} | ${l.save ?? ''} |`
        )
        .join('\n')}\n`
    : '';
  return [
    `# ${label} census\n`,
    `Generated by \`scripts/fidelity/census.ts\`. Counts, codes, and feature names only${perWorkbook ? '. Workbooks are named by sorted position' : ', totals across all workbooks'}. A census cannot show a repair prompt: the save check opens saved copies in a desktop program.\n`,
    `- Workbooks: ${lines.length}\n- Sheets in opened workbooks: ${sheets}\n- Cells: ${cells}\n- Cells with a formula: ${formulas}\n`,
    table('Outcome', outcome, 'Workbooks'),
    table('Open refusals', refusals, 'Workbooks'),
    table('Value edit in the first worksheet', edits, 'Workbooks'),
    table('Row insert in the first worksheet', structures, 'Workbooks'),
    table('Save and reopen', saves, 'Workbooks'),
    table('Save and reopen after the edits', editedSaves, 'Workbooks'),
    perRows,
    table('Sheet problems', sums.sheetProblems, 'Problems'),
    table('Content types', sums.contentTypes, 'Parts'),
    table('Workbook part elements', sums.workbookElements, 'Elements'),
    table('Worksheet elements', sums.worksheetElements, 'Elements'),
    table('Cell types', sums.cellTypes, 'Cells'),
  ]
    .filter(Boolean)
    .join('\n');
}

const USAGE = `usage: bun run fidelity:census -- --dir <folder> [--out <folder>] [--fresh]
  --dir    the folder of .xlsx files, with its subfolders; without it,
           XLSX_CORPUS_DIR names the folder, and with neither the command
           refuses to run; hidden entries, links, and owner files
           (~$name.xlsx) are skipped
  --out    the output folder: outside every git work tree or ignored by git
           (default local/corpus/census)
  --fresh  replace the results of an earlier run in the output folder`;

/** What the census runs over, after its checks. */
export interface CensusPlan {
  readonly jobs: readonly CorpusJob[];
  readonly out: string;
  readonly fresh: boolean;
  /** The committed fixtures: the report may have a row for each workbook. */
  readonly fixtureSet: boolean;
}

/**
 * The plan of a census, or a refusal: `usage` when the flags are wrong, so the usage text
 * follows the message. It runs before any child starts.
 */
export function planCensus(
  args: readonly string[],
  env: NodeJS.ProcessEnv = process.env
): CensusPlan | { readonly refused: string; readonly usage: boolean } {
  let dir: string;
  let out: string;
  let jobs: CorpusJob[];
  try {
    dir = corpusDir(args, env);
    out = checkOut(flag(args, 'out') ?? defaultOut('census'));
    jobs = listWorkbooks(dir);
  } catch (error) {
    return { refused: (error as Error).message, usage: true };
  }
  if (jobs.length === 0)
    return {
      refused: 'The corpus folder holds no .xlsx file. A census of nothing is a gap.',
      usage: false,
    };
  const fresh = args.includes('--fresh');
  const refused = checkFresh(out, fresh);
  if (refused) return { refused, usage: false };
  return { jobs, out, fresh, fixtureSet: dir === resolve(ROOT, 'fixtures') };
}

async function main(): Promise<number> {
  const args = process.argv.slice(2);
  if (args.includes('--help')) {
    console.log(USAGE);
    return 0;
  }
  const plan = planCensus(args);
  if ('refused' in plan) {
    console.error(plan.usage ? `${plan.refused}\n\n${USAGE}` : plan.refused);
    return 1;
  }
  const report = join(plan.out, 'census.md');
  // A report of an earlier run would not match the new results while this run goes on.
  rmSync(report, { force: true });
  const run = await runCorpus({
    task: 'census',
    jobs: plan.jobs,
    out: plan.out,
    workbookMs: CENSUS_WORKBOOK_MS,
    fresh: plan.fresh,
    onProgress: (done, total) => console.log(`${done}/${total} workbooks`),
  });
  if ('refused' in run) {
    console.error(run.refused);
    return 1;
  }
  const lines = readFileSync(run.resultsFile, 'utf8')
    .split('\n')
    .filter((l) => l.trim() !== '')
    .map((l) => JSON.parse(l) as ResultLine);
  const label = plan.fixtureSet ? 'Fixture set' : 'Corpus';
  writeFileSync(report, renderCensus(lines, label, plan.fixtureSet));
  console.log(
    `${run.workbooks} workbooks, ${lines.filter((l) => l.open === 'refused').length} refused at open, ${run.stopped.length} stopped. Child processes: ${run.children}, the highest memory reading at ${Math.round(run.peakBytes / 1048576)} MiB. Wrote ${report}.`
  );
  return 0;
}

if (import.meta.main) process.exit(await main());
