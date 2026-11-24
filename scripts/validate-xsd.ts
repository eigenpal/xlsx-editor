// Validate saved SpreadsheetML parts against the ECMA-376 Transitional schemas with xmllint.
//
// Adapted from pptx-editor/scripts/validate-xsd.ts at 1943d2c9 (EigenPal); the corpus mode
// after the same file at 9d7d60b7. For every fixture it validates the ORIGINAL modeled parts,
// the parts after save, and the parts after an edit session through the store ops (the
// fixture's defined edit, a new string, a new number, a boolean, a clear, a format patch with
// every property, a format clear, and a structural session that inserts and deletes rows and
// columns). Differences from PPTX Editor: the fixtures are synthetic, so the gate is zero errors
// in all three, not "no new error". It also validates the demo's sample workbook
// (`examples/vite/public/sample.xlsx`), original, saved, and edited.
//
// The structural session is its own transaction. A fixture whose content refuses it
// (`structure-blocked`, design D4) keeps the edits before it and reports the refusal; any other
// refusal, or a refusal of the `structure` fixture, fails the gate.
//
// Before validation each part goes through Markup Compatibility preprocessing (ECMA-376
// Part 3): ignorable namespaces are dropped and `mc:AlternateContent` keeps its Fallback. The
// content of `x:extLst` is checked only for well-formedness (the requirement "Normalized save
// with fidelity oracles"), and the known schema exceptions below are removed and listed.
//
//   bun run validate:xsd
//   bun run validate:xsd:corpus -- --dir <corpus folder>   # or XLSX_CORPUS_DIR (local only)
//
// `--corpus` also validates every workbook of a private corpus folder, on the memory-safe
// corpus runner (`scripts/corpus-runner/runner.ts`): one child for each batch, and a watchdog
// at 3 GiB. As in PPTX Editor, the corpus gate is "no new error": the saved and edited copies
// have no schema message, as xmllint gives it with its values, that the original did not
// already have. A workbook that is refused at open, stopped, or fails its task also fails the
// run, as an unreadable deck does there. One workbook may take `XSD_WORKBOOK_MS` (300 s).
// Differences from PPTX Editor's `--corpus`, each for a reason:
// - The corpus folder has no default, and the workbooks run in children.
// - The edit session is the one of the fixtures, and each op is its own transaction, so one
//   refused op does not skip the others. It edits the first worksheet that takes at least one
//   op. A workbook with no such worksheet fails: a skipped edit check is a gap, never a pass.
//   As in the fixture gate, a structural session may be refused only with `structure-blocked`.
// - The output (`local/corpus/xsd/` by default) holds ids, counts, and up to five new messages
//   for each workbook, without values, and with `(name)` for each name that the schemas do
//   not declare (`schemaMessage`).
// - The arguments and the output folder are checked before the fixture gate runs.

import { spawnSync } from 'node:child_process';
import { mkdtempSync, readdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { refusalCode } from './corpus-runner/open-refusal.ts';
import {
  NS,
  readXlsx,
  saveXlsx,
  serializeXmlPart,
  workbookIndex,
  WorkbookDocumentStore,
  type Element,
  type WorkbookOp,
  type XlsxPackage,
} from '../packages/core/src/store/index.ts';
import { FIXTURES, fixtureEditOps, fixtureFile } from './fixtures/catalog.ts';
import {
  checkFresh,
  checkOut,
  corpusDir,
  defaultOut,
  flag,
  listWorkbooks,
  runCorpus,
  taskTempRoot,
  type CorpusJob,
} from './corpus-runner/runner.ts';
// The `validate:xsd` script sets `TZ=UTC`, so the zip entry times of the saved copies are the
// same on every machine. Importing `save-check.ts` also sets it, for a direct `bun` run.
import { sampleCopies } from './sample-workbook/save-check.ts';

const SCHEMAS = 'reference/ecma-376/part4/transitional-schemas';
const SAMPLE = 'examples/vite/public/sample.xlsx';

/**
 * Attributes and values that files in common use hold and that the Transitional schema does not
 * declare. Each one is removed before validation and counted in the report.
 */
export const KNOWN_SCHEMA_EXCEPTIONS = [
  // `xml:space="preserve"` on `x:t` keeps leading and trailing spaces. In sml.xsd, `CT_Rst/t`
  // and `CT_RElt/t` are `ST_Xstring` elements with no attributes, so xmllint refuses it.
  'xml:space on x:t',
  // `t="d"` (an ISO 8601 date cell) is not an `ST_CellType` value (quick reference, "Cells and
  // values"). Validated as `t="str"`.
  'x:c/@t="d"',
] as const;

function preprocess(el: Element, ignorable: Set<string>, counts: Map<string, number>): Element {
  const bump = (key: string) => counts.set(key, (counts.get(key) ?? 0) + 1);
  const children: Element['children'][number][] = [];
  if (el.ns === NS.x && el.local === 'extLst') return { ...el, children: [] };
  for (const c of el.children) {
    if (c.type === 'text') {
      children.push(c);
      continue;
    }
    if (c.ns === NS.mc && c.local === 'AlternateContent') {
      const fallback = c.children.find((x) => x.type === 'element' && x.local === 'Fallback');
      if (fallback && fallback.type === 'element')
        for (const f of fallback.children)
          if (f.type === 'element') children.push(preprocess(f, ignorable, counts));
      continue;
    }
    if (ignorable.has(c.ns)) continue;
    children.push(preprocess(c, ignorable, counts));
  }
  let attrs = el.attrs.filter((a) => !ignorable.has(a.ns) && a.ns !== NS.mc);
  if (el.ns === NS.x && el.local === 't' && attrs.some((a) => a.ns === NS.xml)) {
    bump(KNOWN_SCHEMA_EXCEPTIONS[0]);
    attrs = attrs.filter((a) => a.ns !== NS.xml);
  }
  if (el.ns === NS.x && el.local === 'c' && attrs.some((a) => a.local === 't' && a.value === 'd')) {
    bump(KNOWN_SCHEMA_EXCEPTIONS[1]);
    attrs = attrs.map((a) => (a.local === 't' && a.ns === '' ? { ...a, value: 'str' } : a));
  }
  return { ...el, attrs, children };
}

function ignorableNamespaces(root: Element): Set<string> {
  const out = new Set<string>();
  const attr = root.attrs.find((a) => a.ns === NS.mc && a.local === 'Ignorable');
  for (const prefix of attr?.value.split(/\s+/) ?? []) {
    const decl = root.nsDecls.find((d) => d.prefix === prefix);
    if (decl) out.add(decl.uri);
  }
  return out;
}

/** The schema of a modeled part, from its root element. */
function schemaOf(root: Element): string | undefined {
  if (root.ns === NS.x) return 'sml.xsd';
  if (root.ns === NS.a && root.local === 'theme') return 'dml-main.xsd';
  if (root.ns === NS.xdr) return 'dml-spreadsheetDrawing.xsd';
  return undefined;
}

/** The most files that one xmllint call validates. */
const XMLLINT_FILES = 100;

/** The most output that one xmllint call may write: about two million schema messages. */
const XMLLINT_OUTPUT_BYTES = 256 * 1024 * 1024;

/**
 * The xmllint calls for `files`: one schema for each call, and at most `limit` files, which keeps
 * the argument list short.
 */
export function xmllintGroups<T extends { readonly schema: string }>(
  files: readonly T[],
  limit = XMLLINT_FILES
): { schema: string; group: T[] }[] {
  const groups: { schema: string; group: T[] }[] = [];
  for (const schema of new Set(files.map((f) => f.schema))) {
    const all = files.filter((f) => f.schema === schema);
    for (let i = 0; i < all.length; i += limit)
      groups.push({ schema, group: all.slice(i, i + limit) });
  }
  return groups;
}

/**
 * Runs xmllint with `args` and returns its output. A missing xmllint, more output than
 * `maxBuffer` bytes, or a failed start validates nothing: that is an error, never a pass.
 * `command` is for tests.
 */
export function runXmllint(
  args: readonly string[],
  { command = 'xmllint', maxBuffer = XMLLINT_OUTPUT_BYTES } = {}
): string {
  const run = spawnSync(command, args, { encoding: 'utf8', maxBuffer });
  if (run.error) {
    const code = (run.error as NodeJS.ErrnoException).code;
    throw new Error(
      code === 'ENOENT'
        ? 'xmllint did not run: install libxml2.'
        : code === 'ENOBUFS'
          ? `xmllint wrote more than ${maxBuffer} bytes of messages.`
          : `xmllint failed (${code ?? 'error'}).`
    );
  }
  return `${run.stdout}${run.stderr}`;
}

/** Validate a package's modeled parts; returns error lines keyed by part, line numbers stripped. */
function validate(
  pkg: XlsxPackage,
  dir: string,
  label: string,
  counts: Map<string, number>,
  tally: { validated: number; skipped: string[] }
): Map<string, Set<string>> {
  const byPart = new Map<string, Set<string>>();
  const files: { file: string; part: string; schema: string }[] = [];
  for (const part of pkg.parts.values()) {
    if (part.kind !== 'xml') continue;
    const schema = schemaOf(part.xml.root);
    if (!schema) {
      tally.skipped.push(part.partName);
      continue;
    }
    const root = preprocess(part.xml.root, ignorableNamespaces(part.xml.root), counts);
    // The position keeps two part names that map to the same text apart.
    const name = part.partName.replace(/[/[\]]/g, '_');
    const file = join(dir, `${label}-${files.length}-${name}`);
    writeFileSync(file, serializeXmlPart({ ...part.xml, root }));
    files.push({ file, part: part.partName, schema });
  }
  for (const { schema, group } of xmllintGroups(files)) {
    // An unreadable schema or a crash validates nothing: that is a failure, never a pass.
    // Every file must get its own verdict line below.
    const out = runXmllint([
      '--noout',
      '--schema',
      join(SCHEMAS, schema),
      ...group.map((g) => g.file),
    ]);
    for (const g of group) {
      const verdict = out
        .split('\n')
        .some((l) => l === `${g.file} validates` || l === `${g.file} fails to validate`);
      if (!verdict)
        throw new Error(`xmllint gave no verdict for ${g.part}:\n${out.slice(0, 2000)}`);
      tally.validated += 1;
    }
    for (const line of out.split('\n')) {
      const g = group.find((x) => line.startsWith(`${x.file}:`));
      if (!g || line.endsWith('validates')) continue;
      const message = line
        .slice(g.file.length)
        .replace(/^:\d+:/, '')
        .trim();
      if (!byPart.has(g.part)) byPart.set(g.part, new Set());
      byPart.get(g.part)!.add(message);
    }
  }
  return byPart;
}

/** The edit session of every workbook: values of each type, a format patch, and clears. */
export function sessionOps(sheet: string): WorkbookOp[] {
  return [
    {
      op: 'setCellValue',
      sheet,
      ref: 'H40',
      value: { type: 'string', value: 'Edited & <saved>' },
    },
    { op: 'setCellValue', sheet, ref: 'H41', value: { type: 'number', value: 1e21 } },
    { op: 'setCellValue', sheet, ref: 'H42', value: { type: 'boolean', value: true } },
    { op: 'setCellValue', sheet, ref: 'H42', value: { type: 'empty' } },
    {
      op: 'setCellValue',
      sheet,
      ref: 'H43',
      value: { type: 'formula', text: 'SUM(H40:H41)&"<&>"' },
    },
    {
      op: 'applyFormat',
      sheet,
      ref: 'H40:I44',
      patch: {
        bold: true,
        underline: true,
        fontName: 'Mono & <Sans>',
        fontSize: 10.5,
        fontColor: 'C00000',
        fill: { type: 'solid', color: 'DDEBF7' },
        borders: { edges: 'all', style: 'double', color: '7030A0' },
        horizontal: 'centerContinuous',
        vertical: 'distributed',
        wrapText: true,
        numberFormat: { code: '"<&>"#,##0.000' },
      },
    },
    { op: 'applyFormat', sheet, ref: 'H45', patch: { borders: { edges: 'none' } } },
    { op: 'clearFormats', sheet, ref: 'I44' },
  ];
}

/** The structural session: rows and columns inserted and deleted. */
export function structuralOps(sheet: string): WorkbookOp[] {
  return [
    { op: 'insertRows', sheet, start: 2, count: 2 },
    { op: 'insertColumns', sheet, start: 2, count: 1 },
    { op: 'deleteRows', sheet, start: 6, count: 1 },
    { op: 'deleteColumns', sheet, start: 4, count: 2 },
  ];
}

/** The fixture gate: every fixture and the demo sample, original, saved, and edited. */
function validateFixtures(): boolean {
  const dir = mkdtempSync(join(tmpdir(), 'xlsx-xsd-'));
  const counts = new Map<string, number>();
  const tally = { validated: 0, skipped: [] as string[] };
  let failures = 0;
  let structuralSessions = 0;
  try {
    for (const fixture of FIXTURES) {
      const read = readXlsx(new Uint8Array(readFileSync(`fixtures/${fixtureFile(fixture.name)}`)));
      if (!read.ok) {
        console.error(`${fixture.name}: does not open (${read.error.reason})`);
        failures += 1;
        continue;
      }
      const saved = readXlsx(saveXlsx(read.pkg));
      const store = new WorkbookDocumentStore(read.pkg);
      const sheet = fixture.edits[0]!.sheet;
      const edit = store.transact([...fixtureEditOps(fixture, read.pkg), ...sessionOps(sheet)]);
      const structural = store.transact(structuralOps(sheet));
      if (!structural.ok) {
        const allowed =
          structural.rejection.code === 'structure-blocked' && fixture.name !== 'structure';
        console[allowed ? 'log' : 'error'](
          `${fixture.name}: the structural session is refused (${structural.rejection.detail})`
        );
        if (!allowed) failures += 1;
      } else structuralSessions += 1;
      const edited = edit.ok ? readXlsx(saveXlsx(store.package)) : undefined;
      const runs: [string, XlsxPackage | undefined][] = [
        ['original', read.pkg],
        ['saved', saved.ok ? saved.pkg : undefined],
        ['edited', edited?.ok ? edited.pkg : undefined],
      ];
      for (const [label, pkg] of runs) {
        if (!pkg) {
          console.error(`${fixture.name} (${label}): does not save, edit, or reopen`);
          failures += 1;
          continue;
        }
        const errors = validate(pkg, dir, `${fixture.name}-${label}`, counts, tally);
        for (const [part, messages] of errors) {
          failures += messages.size;
          for (const message of messages)
            console.error(`${fixture.name} (${label}) ${part}: ${message}`);
        }
      }
    }
    // The demo's sample workbook (`bun run sample:build`), original, saved, and edited, as PPTX
    // Editor validates its demo decks. The edited copy is the one of its save check.
    const sample = readXlsx(new Uint8Array(readFileSync(SAMPLE)));
    const copies = sample.ok ? sampleCopies() : undefined;
    const sampleSaved = copies ? readXlsx(copies.saved) : undefined;
    const sampleEdited = copies ? readXlsx(copies.edited) : undefined;
    for (const [label, read] of [
      ['original', sample],
      ['saved', sampleSaved],
      ['edited', sampleEdited],
    ] as const) {
      if (!read?.ok) {
        console.error(`sample (${label}): does not open or reopen`);
        failures += 1;
        continue;
      }
      for (const [part, messages] of validate(read.pkg, dir, `sample-${label}`, counts, tally)) {
        failures += messages.size;
        for (const message of messages) console.error(`sample (${label}) ${part}: ${message}`);
      }
    }
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
  console.log(
    `Validated ${tally.validated} parts. Skipped ${tally.skipped.length} modeled parts that no schema covers here (content types and relationships): ${[...new Set(tally.skipped)].join(', ')}.`
  );
  if (tally.validated === 0) failures += 1;
  for (const exception of KNOWN_SCHEMA_EXCEPTIONS)
    console.log(
      `Known schema exception "${exception}": ${counts.get(exception) ?? 0} occurrences.`
    );
  if (failures > 0) {
    console.error(`Schema validation failed: ${failures} errors.`);
    return false;
  }
  console.log(
    `Schema validation passed: ${FIXTURES.length} fixtures, original, saved, and edited, with ${structuralSessions} structural sessions, and the demo sample, original, saved, and edited.`
  );
  return true;
}

/** An element name as xmllint quotes it, `{namespace}local`, and an attribute name. */
const ELEMENT = String.raw`'\{([^{}'\s]*)\}([A-Za-z_][\w.-]*)'`;
const NAME = String.raw`[A-Za-z_][\w.:-]*`;
/** The start of a validity message: the prefix of xmllint, then the element and the attribute. */
const HEAD = new RegExp(
  String.raw`^(?:element ${NAME}: Schemas validity error : )?(Element ${ELEMENT}(?:, attribute '(${NAME})')?: )`
);
/** A type name of the schemas, with its namespace when xmllint gives one. */
const TYPE = String.raw`(?:\{[^{}'\s]*\})?${NAME}`;
const KNOWN_NAMESPACES = new Set<string>(Object.values(NS));
/** The middle of the message of an enumeration facet. */
const SET = ' is not an element of the set {';

let declaredNames: ReadonlySet<string> | undefined;
/**
 * A name that the schemas declare as an element or an attribute, or `(name)`. An author can put
 * any name in a namespace, so only the words of the schemas reach the output, as in the census.
 * A prefixed name is the author's prefix: it becomes `(name)` too.
 */
function declared(name: string): string {
  if (!declaredNames) {
    const names = new Set<string>();
    for (const file of readdirSync(SCHEMAS).filter((f) => f.endsWith('.xsd')))
      for (const m of readFileSync(join(SCHEMAS, file), 'utf8').matchAll(
        /<xsd:(?:element|attribute) name="([A-Za-z_][\w.-]*)"/g
      ))
        names.add(m[1]!);
    declaredNames = names;
  }
  return declaredNames.has(name) ? name : '(name)';
}

/**
 * A schema message with no value from the workbook, for the output. It keeps the element and
 * attribute names that xmllint quotes when the schemas declare them, the facet, and the type,
 * and replaces each value with `<value>` and each other name with `(name)`. A namespace that
 * the engine does not know becomes `{other}`. The prefix of xmllint
 * is left out. A message that does not start with this form becomes
 * `(other schema message)` when it holds a quote, so no value can pass through. The gate
 * compares the messages before this step.
 */
export function schemaMessage(message: string): string {
  const m = HEAD.exec(message);
  const head = !m
    ? ''
    : `Element '{${KNOWN_NAMESPACES.has(m[2]!) ? m[2] : 'other'}}${declared(m[3]!)}'${
        m[4] === undefined ? '' : `, attribute '${declared(m[4])}'`
      }: `;
  let rest = m ? message.slice(m[0].length) : message;
  const facet = /^\[facet '(\w+)'\] /.exec(rest);
  if (facet) rest = rest.slice(facet[0].length);
  const typed = new RegExp(
    String.raw`^'.*' is not a valid value of the (\w+) type '(${TYPE})'\.$`,
    's'
  ).exec(rest);
  const attribute = new RegExp(
    String.raw`^The attribute '(${NAME})' is (not allowed|required but missing)\.$`
  ).exec(rest);
  if (attribute) rest = `The attribute '${declared(attribute[1]!)}' is ${attribute[2]}.`;
  else if (typed) rest = `'<value>' is not a valid value of the ${typed[1]} type '${typed[2]}'.`;
  else if (rest.includes("'")) {
    if (!m) return '(other schema message)';
    if (rest.startsWith("The value '") && rest.endsWith('}.') && rest.includes(SET)) {
      // The value and the allowed set are both file text: nothing of them stays.
      rest = `The value '<value>'${SET}<values>}.`;
    } else {
      // Everything from the first quote to the last is a value. A value that held a line break
      // leaves one quote only: the rest of the line goes too.
      const first = rest.indexOf("'");
      const last = rest.lastIndexOf("'");
      rest = `${rest.slice(0, first)}'<value>'${last > first ? rest.slice(last + 1) : ''}`;
    }
  }
  return `${head}${facet ? `[facet '${facet[1]}'] ` : ''}${rest}`.slice(0, 300);
}

/** The time that one corpus workbook may take: xmllint validates its copies. */
export const XSD_WORKBOOK_MS = 300_000;

/** The schema messages of `pkg`, as xmllint gives them, with no part name and no line. */
function messagesOf(pkg: XlsxPackage, dir: string, label: string, counts: Map<string, number>) {
  const tally = { validated: 0, skipped: [] as string[] };
  const messages = new Set<string>();
  for (const set of validate(pkg, dir, label, counts, tally).values())
    for (const message of set) messages.add(message);
  return { messages, validated: tally.validated };
}

/** The writer of the corpus task. Tests pass a faulty one. */
export type CorpusSave = (pkg: XlsxPackage) => Uint8Array;

/** One `results.jsonl` line of `validate:xsd:corpus`. */
export interface CorpusXsdResult {
  readonly open: 'ok' | 'refused';
  readonly refusal?: string;
  /** The count of parts of the original that a schema covers. */
  readonly parts?: number;
  /** The count of schema messages of the original. */
  readonly original?: number;
  /** The count of schema messages of the saved copy that the original does not have, or `unreadable`. */
  readonly saved?: number | 'unreadable';
  /** The same count for the edited copy, or `skipped` when no worksheet takes the edits. */
  readonly edited?: number | 'unreadable' | 'skipped';
  /** The position of the edited worksheet among the worksheets, from 1. */
  readonly editSheet?: number;
  /** The ops of the edit session that the store refused, each in its own transaction. */
  readonly editRefused?: number;
  readonly structure?: string;
  /** Up to five new messages, without values. */
  readonly newMessages?: readonly string[];
}

/**
 * The corpus task: the original, saved, and edited copies of one workbook, validated. The gate
 * is PPTX Editor's: saving and editing introduce no schema message that the original did not
 * already have. The work files go to the folder that the runner removes when the child ends.
 */
export function validateCorpusWorkbook(
  bytes: Uint8Array,
  save: CorpusSave = saveXlsx
): CorpusXsdResult {
  const read = readXlsx(bytes);
  if (!read.ok) return { open: 'refused', refusal: refusalCode(read.error) };
  const dir = mkdtempSync(join(taskTempRoot(), 'xlsx-xsd-corpus-'));
  try {
    const counts = new Map<string, number>();
    const base = messagesOf(read.pkg, dir, 'o', counts);
    const saved = readXlsx(save(read.pkg));
    const worksheets = (workbookIndex(read.pkg)?.sheets ?? []).filter(
      (s) => s.kind === 'worksheet'
    );
    const store = new WorkbookDocumentStore(read.pkg);
    let editSheet: number | undefined;
    let editRefused = 0;
    let structure = 'none';
    for (const [i, sheet] of worksheets.entries()) {
      const ops = sessionOps(sheet.name);
      const refused = ops.filter((op) => !store.transact([op]).ok).length;
      // A sheet that refuses every op, such as a protected sheet, is left as it was.
      if (refused === ops.length) continue;
      editSheet = i + 1;
      editRefused = refused;
      const structural = store.transact(structuralOps(sheet.name));
      structure = structural.ok ? 'ok' : structural.rejection.code;
      break;
    }
    const fresh: string[] = [];
    const newCount = (copy: ReturnType<typeof readXlsx>, label: string) => {
      if (!copy.ok) return 'unreadable' as const;
      const added = [...messagesOf(copy.pkg, dir, label, counts).messages].filter(
        (m) => !base.messages.has(m)
      );
      fresh.push(...added);
      return added.length;
    };
    const savedNew = newCount(saved, 's');
    const editedNew =
      editSheet === undefined ? ('skipped' as const) : newCount(readXlsx(save(store.package)), 'e');
    return {
      open: 'ok',
      parts: base.validated,
      original: base.messages.size,
      saved: savedNew,
      edited: editedNew,
      ...(editSheet === undefined ? {} : { editSheet }),
      editRefused,
      structure,
      newMessages: [...new Set(fresh.map(schemaMessage))].slice(0, 5),
    };
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
}

/** One line of the results of `validate:xsd:corpus`: a result, a stop, or an error of the task. */
export type CorpusXsdLine = Partial<CorpusXsdResult> & {
  readonly id?: string;
  readonly stopped?: string;
  readonly error?: string;
};

/** The problem of one result line, or undefined when the workbook passes the corpus gate. */
export function corpusProblem(r: CorpusXsdLine): string | undefined {
  if (r.stopped) return `stopped (${r.stopped})`;
  if (r.error) return `task error: ${r.error}`;
  if (r.open === 'refused') return `refused at open (${r.refusal})`;
  if (r.open !== 'ok') return 'the line is not a result of the task';
  if (!r.parts) return 'no part was validated';
  if (r.saved === 'unreadable' || r.edited === 'unreadable') return 'a saved copy does not reopen';
  if (r.edited === 'skipped') return 'the edit check was skipped: no worksheet takes the edits';
  if (r.structure !== 'ok' && r.structure !== 'structure-blocked')
    return `the structural session is refused (${r.structure})`;
  if (typeof r.saved !== 'number' || typeof r.edited !== 'number')
    return 'the line is not a result of the task';
  if (r.saved + r.edited > 0) return `saved: ${r.saved} new, edited: ${r.edited} new`;
  return undefined;
}

/** What `--corpus` runs over, after its checks. */
export interface CorpusPlan {
  readonly jobs: readonly CorpusJob[];
  readonly out: string;
  readonly fresh: boolean;
}

/**
 * The plan of `--corpus`, or a refusal: `usage` when the flags are wrong, so the usage text
 * follows the message. It runs before the fixture gate.
 */
export function planCorpus(
  args: readonly string[],
  env: NodeJS.ProcessEnv = process.env
): CorpusPlan | { readonly refused: string; readonly usage: boolean } {
  let jobs: CorpusJob[];
  let out: string;
  try {
    jobs = listWorkbooks(corpusDir(args, env));
    out = checkOut(flag(args, 'out') ?? defaultOut('xsd'));
  } catch (error) {
    return { refused: (error as Error).message, usage: true };
  }
  if (jobs.length === 0)
    return {
      refused: 'The corpus folder holds no .xlsx file. A check of nothing is a gap.',
      usage: false,
    };
  const fresh = args.includes('--fresh');
  const refused = checkFresh(out, fresh);
  return refused ? { refused, usage: false } : { jobs, out, fresh };
}

/** `--corpus`: the corpus task over every workbook of the plan, on the corpus runner. */
export async function validateCorpus(plan: CorpusPlan): Promise<boolean> {
  const run = await runCorpus({
    task: 'xsd',
    jobs: plan.jobs,
    out: plan.out,
    workbookMs: XSD_WORKBOOK_MS,
    fresh: plan.fresh,
    onProgress: (done, total) => console.log(`${done}/${total} workbooks`),
  });
  if ('refused' in run) {
    console.error(run.refused);
    return false;
  }
  let failed = 0;
  let refused = 0;
  for (const line of readFileSync(run.resultsFile, 'utf8').split('\n')) {
    if (!line.trim()) continue;
    const r = JSON.parse(line) as CorpusXsdLine;
    if (r.open === 'refused') refused += 1;
    const problem = corpusProblem(r);
    if (!problem) continue;
    failed += 1;
    console.error(`${r.id}: ${problem}`);
    for (const m of r.newMessages ?? []) console.error(`  ${m}`);
  }
  console.log(
    `Corpus: ${run.workbooks} workbooks, ${refused} refused at open, ${run.stopped.length} stopped, ${failed} with a problem. Child processes: ${run.children}, the highest memory reading at ${Math.round(run.peakBytes / 1048576)} MiB. Results: ${run.resultsFile}.`
  );
  return failed === 0 && run.results + run.stopped.length === run.workbooks;
}

const USAGE = `usage: bun run validate:xsd
       bun run validate:xsd:corpus -- --dir <folder> [--out <folder>] [--fresh]
  --dir    the folder of .xlsx files, with its subfolders; without it,
           XLSX_CORPUS_DIR names the folder, and with neither the command
           refuses to run; hidden entries, links, and owner files
           (~$name.xlsx) are skipped
  --out    the output folder: outside every git work tree or ignored by git
           (default local/corpus/xsd)
  --fresh  replace the results of an earlier run in the output folder
The corpus mode runs the fixture gate first, then the corpus.`;

async function main(): Promise<number> {
  const args = process.argv.slice(2);
  if (args.includes('--help')) {
    console.log(USAGE);
    return 0;
  }
  let plan: CorpusPlan | undefined;
  if (args.includes('--corpus')) {
    const planned = planCorpus(args);
    if ('refused' in planned) {
      console.error(planned.usage ? `${planned.refused}\n\n${USAGE}` : planned.refused);
      return 1;
    }
    plan = planned;
  }
  const fixtures = validateFixtures();
  const corpus = plan ? await validateCorpus(plan) : true;
  return fixtures && corpus ? 0 : 1;
}

if (import.meta.main) process.exit(await main());
