// The kit of the feature edits (`feature-edits.ts`): the recipe types, the helpers that the
// recipes use, and the runner. A recipe opens a fixture of its row, makes the row's own edit
// through the public editor API, saves, and opens the saved copy again. The edit must be in the
// reopened copy, its check must fail on the unedited workbook, a second save must give the same
// package, and only the parts that the edit owns may change.
//
// Same form as pptx-editor/scripts/fidelity/feature-edit-kit.ts at 30dc48ea (EigenPal).
// Differences: a recipe names a worksheet, not a slide, and the helpers read cells and cell
// formats, not shapes. No recipe collaborates: collaboration starts after core editing
// (AGENTS.md, "Collaboration verification").

import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { Window } from 'happy-dom';
import { createXlsxEditor, runToolbarCommand } from '../../packages/core/src/editor/index.ts';
import type {
  ChromeSlotId,
  EditResult,
  XlsxEditorInstance,
} from '../../packages/core/src/editor/index.ts';
import {
  attr,
  child,
  children,
  intAttr,
  packageFingerprints,
  readXlsx,
  serializeXmlPart,
  styleLists,
  textOf,
  workbookIndex,
  xmlPartOf,
  type Element,
  type XlsxPackage,
} from '../../packages/core/src/store/index.ts';

// The zip writer stores entry times as local-time fields, so the copy bytes depend on the time
// zone. The fixture build and the save check run in UTC for the same reason.
process.env.TZ = 'UTC';

export const FIXTURES_DIR = join(import.meta.dir, '../../fixtures');

/** What a recipe uses to edit the open workbook. */
export interface EditTools {
  readonly editor: XlsxEditorInstance;
  /** Check the result of an editor call; a refusal fails the recipe. */
  run(result: EditResult | { readonly ok: boolean; readonly refusal?: unknown }): void;
  /** Select a cell (`B2`) or a range (`B2:C4`) of the shown sheet. */
  select(ref: string): void;
  /** Run a toolbar slot (`runToolbarCommand`); a refusal fails the recipe. */
  toolbar(id: ChromeSlotId, value?: string): void;
}

/** What a recipe uses to check the reopened copy. */
export interface SavedWorkbook {
  readonly pkg: XlsxPackage;
  /** The XML of a part, as saved, or undefined when the part is absent. */
  partXml(partName: string): string | undefined;
  /** The XML of the worksheet named `sheet`, as saved. */
  sheetXml(sheet: string): string;
  /** The XML of the cell `ref` of `sheet`, as saved, or undefined when it has no `x:c`. */
  cellXml(sheet: string, ref: string): string | undefined;
  /** The text of the cell `ref` of `sheet`: its shared or inline string, or its `x:v`. */
  cellText(sheet: string, ref: string): string | undefined;
  /** The cell format records of the cell `ref` of `sheet`, as saved. */
  cellFormat(sheet: string, ref: string): CellFormatXml;
  /** The sheet names, in tab order. */
  readonly sheetNames: readonly string[];
}

/** The records of a cell's `x:xf`, each as XML, and its number format code. */
export interface CellFormatXml {
  readonly xf: string;
  readonly font: string;
  readonly fill: string;
  readonly border: string;
  /** The custom code of `@numFmtId`, or undefined for a built-in id. */
  readonly numFmtCode: string | undefined;
  readonly numFmtId: number;
}

/** A part the edit may change: any change to matching names, or only a new or removed part. */
export type PartRule = RegExp | { readonly added: RegExp } | { readonly removed: RegExp };

/** Parts that match `re` may appear, but existing parts may not change. */
export const added = (re: RegExp): PartRule => ({ added: re });

/** Parts that match `re` may go, but existing parts may not change. */
export const removed = (re: RegExp): PartRule => ({ removed: re });

/** The edit of one support matrix row. */
export interface FeatureEdit {
  /** A fixture of the row, by catalog name (`scripts/fixtures/catalog.ts`). */
  readonly fixture: string;
  /** The sheet the edit starts on. Absent: the sheet that the workbook opens on. */
  readonly sheet?: string;
  /**
   * Attach the editor to a page first: the cell editor and the formula bar need a mounted
   * surface. The page is a happy-dom window, as in the editor's own tests.
   */
  readonly mount?: boolean;
  readonly edit: (t: EditTools) => void;
  /** A failure message, or undefined when the reopened copy has the edit. */
  readonly check: (d: SavedWorkbook) => string | undefined;
  /**
   * Parts the edit may change besides the worksheet it starts on. Absent: only that worksheet.
   */
  readonly parts?: readonly PartRule[];
}

/** A failure when `xml` lacks one of `parts`. */
export const has = (xml: string | undefined, ...parts: string[]): string | undefined => {
  if (xml === undefined) return 'the XML is missing';
  const missing = parts.filter((p) => !xml.includes(p));
  return missing.length ? `missing ${missing.map((m) => JSON.stringify(m)).join(', ')}` : undefined;
};

/** A failure when `xml` holds one of `parts`. */
export const lacks = (xml: string | undefined, ...parts: string[]): string | undefined => {
  if (xml === undefined) return undefined;
  const present = parts.filter((p) => xml.includes(p));
  return present.length
    ? `still has ${present.map((m) => JSON.stringify(m)).join(', ')}`
    : undefined;
};

/** A failure when the text of a cell is not `expected`. */
export const text = (actual: string | undefined, expected: string): string | undefined =>
  actual === expected
    ? undefined
    : `the text is ${JSON.stringify(actual)}, not ${JSON.stringify(expected)}`;

/** The first failure of `checks`, or undefined. */
export const all = (...checks: (string | undefined)[]): string | undefined =>
  checks.find((c) => c !== undefined);

/** The 0-based row and column of an A1 reference, such as `B2`. */
export function point(ref: string): { row: number; col: number } {
  const m = /^([A-Z]{1,3})([1-9][0-9]{0,6})$/.exec(ref);
  if (!m) throw new Error(`not a cell reference: ${ref}`);
  let col = 0;
  for (const ch of m[1]!) col = col * 26 + (ch.charCodeAt(0) - 64);
  return { row: Number(m[2]) - 1, col: col - 1 };
}

/** The XML of `el`, with the namespace declarations of its part's root. */
function xmlOf(el: Element, root: Element): string {
  return serializeXmlPart({
    partName: '',
    nextId: 0,
    root: { ...el, nsDecls: [...root.nsDecls, ...el.nsDecls] },
  });
}

/** The text of every `x:t` under `el`, in order: a string item or an inline string. */
function textsOf(el: Element | undefined): string | undefined {
  if (!el) return undefined;
  const out: string[] = [];
  const walk = (e: Element) => {
    for (const c of e.children)
      if (c.type === 'element') {
        if (c.local === 't') out.push(textOf(c));
        else if (c.local !== 'rPh') walk(c);
      }
  };
  walk(el);
  return out.join('');
}

/** Open `bytes`, or throw: a saved copy that does not open fails the recipe. */
export function reopen(bytes: Uint8Array): XlsxPackage {
  const read = readXlsx(bytes);
  if (!read.ok) throw new Error(`the saved copy does not open (${read.error.reason})`);
  return read.pkg;
}

/** The saved workbook that a check reads. */
export function savedWorkbook(pkg: XlsxPackage): SavedWorkbook {
  const index = workbookIndex(pkg);
  if (!index) throw new Error('the saved copy has no workbook part');
  const sheetPart = (sheet: string) => {
    const entry = index.sheets.find((s) => s.name === sheet);
    const part = xmlPartOf(pkg, entry?.partName);
    if (!part) throw new Error(`no worksheet "${sheet}"`);
    return part;
  };
  const cellOf = (sheet: string, ref: string): Element | undefined => {
    const part = sheetPart(sheet);
    for (const r of children(child(part.root, 'x:sheetData'), 'x:row'))
      for (const c of children(r, 'x:c')) if (attr(c, 'r') === ref) return c;
    return undefined;
  };
  return {
    pkg,
    sheetNames: index.sheets.map((s) => s.name),
    partXml: (name) => {
      const part = xmlPartOf(pkg, name);
      return part && serializeXmlPart(part);
    },
    sheetXml: (sheet) => serializeXmlPart(sheetPart(sheet)),
    cellXml: (sheet, ref) => {
      const c = cellOf(sheet, ref);
      return c && xmlOf(c, sheetPart(sheet).root);
    },
    cellText: (sheet, ref) => {
      const c = cellOf(sheet, ref);
      if (!c) return undefined;
      const type = attr(c, 't');
      if (type === 'inlineStr') return textsOf(child(c, 'x:is'));
      const value = child(c, 'x:v');
      if (type !== 's') return value && textOf(value);
      const strings = xmlPartOf(pkg, index.sharedStringsPart)?.root;
      return textsOf(children(strings, 'x:si')[Number(textOf(value))]);
    },
    cellFormat: (sheet, ref) => {
      const lists = styleLists(pkg, index);
      const xf = lists.cellXfs[intAttr(cellOf(sheet, ref), 's') ?? 0];
      if (!xf) throw new Error(`no cell format for ${sheet}!${ref}`);
      const styles = xmlPartOf(pkg, index.stylesPart)!.root;
      const record = (list: readonly Element[], at: string) => {
        const el = list[intAttr(xf, at) ?? 0];
        return el ? xmlOf(el, styles) : '';
      };
      const numFmtId = intAttr(xf, 'numFmtId') ?? 0;
      return {
        xf: xmlOf(xf, styles),
        font: record(lists.fonts, 'fontId'),
        fill: record(lists.fills, 'fillId'),
        border: record(lists.borders, 'borderId'),
        numFmtId,
        numFmtCode: lists.numFmts.get(numFmtId),
      };
    },
  };
}

/** The file name of the edited copy of a row, as `feature-edits.ts --out` writes it. */
export const featureCopyName = (row: string): string => `${row}.xlsx`;

/** The fixture as committed. */
export const fixtureBytes = (name: string): Uint8Array =>
  new Uint8Array(readFileSync(join(FIXTURES_DIR, `${name}.xlsx`)));

/** The happy-dom page of each mounted editor, closed by `dispose`. */
const pages = new WeakMap<XlsxEditorInstance, Window>();

/** Destroy `editor`, and close its page when it has one. */
function dispose(editor: XlsxEditorInstance): void {
  try {
    editor.destroy();
  } finally {
    pages.get(editor)?.happyDOM.close();
  }
}

/** Attach `editor` to a new happy-dom page with a grid of 800 by 400 pixels. */
function mountEditor(editor: XlsxEditorInstance): void {
  const win = new Window();
  pages.set(editor, win);
  const host = win.document.createElement('div');
  win.document.body.appendChild(host);
  editor.attach(host as unknown as HTMLElement);
  const scroller = host.querySelector('.xlsx-scroller');
  if (!scroller) throw new Error('the editor mounts no grid');
  Object.defineProperty(scroller, 'clientWidth', { get: () => 800 });
  Object.defineProperty(scroller, 'clientHeight', { get: () => 400 });
}

/** Open `bytes` in a new editor, on `sheet` when it is given, and mounted when `mount` is set. */
async function openEditor(
  bytes: Uint8Array,
  sheet?: string,
  mount = false
): Promise<XlsxEditorInstance> {
  const editor = createXlsxEditor({ translate: (key) => key });
  if (mount) mountEditor(editor);
  const snapshot = await editor.load(bytes);
  if (snapshot.status !== 'ready') {
    dispose(editor);
    throw new Error(`the editor does not open the workbook (${snapshot.status})`);
  }
  if (sheet !== undefined) {
    const at = snapshot.sheets.findIndex((s) => s.name === sheet);
    if (at < 0 || !editor.setActiveSheet(at)) {
      dispose(editor);
      throw new Error(`the editor does not show the sheet "${sheet}"`);
    }
  }
  return editor;
}

/** Save `editor`, or throw. */
function save(editor: XlsxEditorInstance): Uint8Array {
  const bytes = editor.save();
  if (!bytes) throw new Error('the save is refused');
  return bytes;
}

function tools(editor: XlsxEditorInstance): EditTools {
  const run: EditTools['run'] = (result) => {
    if (!result.ok)
      throw new Error(
        `the edit is refused (${JSON.stringify((result as EditResult & { ok: false }).refusal)})`
      );
  };
  return {
    editor,
    run,
    select: (ref) => {
      const [from, to] = ref.split(':') as [string, string | undefined];
      if (!editor.select(point(from), to === undefined ? undefined : point(to)))
        throw new Error(`the editor does not select ${ref}`);
    },
    toolbar: (id, value) => run(runToolbarCommand(editor, id, value)),
  };
}

/** The result of one recipe. */
export interface FeatureEditRun {
  /** The edited copy, as the editor saves it. */
  readonly bytes: Uint8Array;
  /** The sheet count of the edited copy, which the save check compares. */
  readonly sheetCount: number;
  /** Every way the recipe fails, or an empty list. */
  readonly problems: readonly string[];
}

/** Whether `rule` lets the part `name` change in the way `how`. */
function allows(rule: PartRule, name: string, how: 'added' | 'removed' | 'changed'): boolean {
  if (rule instanceof RegExp) return rule.test(name);
  if ('added' in rule) return how === 'added' && rule.added.test(name);
  return how === 'removed' && rule.removed.test(name);
}

/**
 * Run one recipe: open its fixture, edit, save, and reopen. Also save the unedited fixture,
 * and save the edited copy a second time. A recipe that throws fails with its message.
 */
export async function runFeatureEdit(recipe: FeatureEdit): Promise<FeatureEditRun> {
  const source = fixtureBytes(recipe.fixture);
  const base = await openEditor(source, recipe.sheet);
  const startSheet = base.snapshot().sheets[base.snapshot().activeSheet]?.name;
  let unedited: XlsxPackage;
  try {
    unedited = reopen(save(base));
  } finally {
    dispose(base);
  }

  const editor = await openEditor(source, recipe.sheet, recipe.mount);
  let bytes: Uint8Array;
  try {
    recipe.edit(tools(editor));
    bytes = save(editor);
  } finally {
    dispose(editor);
  }
  const pkg = reopen(bytes);

  const problems: string[] = [];
  const fail = (message: string) => problems.push(message);
  const edited = savedWorkbook(pkg);
  const found = recipe.check(edited);
  if (found !== undefined) fail(`the reopened copy does not have the edit: ${found}`);
  if (recipe.check(savedWorkbook(unedited)) === undefined)
    fail('the check passes on the unedited workbook: make it find the edit');

  // A second save of the reopened copy gives the same package.
  const again = await openEditor(bytes);
  let second: XlsxPackage;
  try {
    second = reopen(save(again));
  } finally {
    dispose(again);
  }
  const first = packageFingerprints(pkg);
  const next = packageFingerprints(second);
  const differs = [...new Set([...first.keys(), ...next.keys()])].filter(
    (name) => first.get(name) !== next.get(name)
  );
  if (differs.length > 0) fail(`a second save changes ${differs.join(', ')}`);

  // Only the parts that the edit owns change.
  const index = workbookIndex(unedited);
  const own = index?.sheets.find((s) => s.name === startSheet)?.partName;
  const before = packageFingerprints(unedited);
  for (const name of [...new Set([...before.keys(), ...first.keys()])].sort()) {
    if (before.get(name) === first.get(name) || name === own) continue;
    const how = !before.has(name) ? 'added' : !first.has(name) ? 'removed' : 'changed';
    if (!(recipe.parts ?? []).some((rule) => allows(rule, name, how)))
      fail(
        `the edit ${how === 'changed' ? 'changes' : how === 'added' ? 'adds' : 'removes'} ${name}, which it does not own`
      );
  }

  return { bytes, sheetCount: edited.sheetNames.length, problems };
}
