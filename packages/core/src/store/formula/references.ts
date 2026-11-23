// The references of a typed formula, checked against the workbook: each sheet that a reference
// or a name names, each external workbook index `[n]`, and each table that a structured
// reference names must be one that the workbook holds. A formula that names another one would
// be saved as it was typed, and it reads as a link to a workbook that is not there. A sheet name
// with a space must be quoted (`'My sheet'!A1`): unquoted, its words read as other operands.
//
// Security (AGENTS.md, "Security"): the table parts are file data. They are read with the
// bounded XML reader, once per part, and the scope is cached per package revision. Names are
// compared in a `Set`, never as object keys.

import { XML_HARD_MAX_BYTES } from '@eigenpal/ooxml-core';
import { contentTypeOf, decodeXml, xmlPartOf, type XlsxPackage } from '../package/package.ts';
import { parseXmlPart, type Element } from '../package/tree.ts';
import { attr, child, children } from '../package/query.ts';
import { workbookIndex } from '../workbook/workbook-index.ts';
import type { FormulaToken, SheetPrefix } from './tokenizer.ts';

/** Why a reference of a typed formula names nothing that the workbook holds. */
export type FormulaReferenceProblem =
  'unknown-sheet' | 'unknown-external' | 'unknown-table' | 'unquoted-sheet-name';

/** What a formula of the workbook can name. Names are in upper case. */
export interface FormulaScope {
  readonly sheets: ReadonlySet<string>;
  /**
   * The number of `x:externalReferences/x:externalReference` entries (`CT_ExternalReference`,
   * sml.xsd): `[1]` to `[n]`.
   */
  readonly externals: number;
  readonly tables: ReadonlySet<string>;
  /** The defined names of the workbook (`x:definedName/@name`, `CT_DefinedName`, sml.xsd). */
  readonly names: ReadonlySet<string>;
}

const TABLE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml';
/** The most words before a sheet prefix that can be one unquoted sheet name. */
const MAX_NAME_WORDS = 31;

const fold = (s: string) => s.toUpperCase();
const scopeCache = new WeakMap<XlsxPackage, FormulaScope>();
const tableCache = new WeakMap<object, readonly string[]>();

/**
 * The `@name` and `@displayName` of a table part, read once per part. In `CT_Table` (sml.xsd)
 * `@displayName` is required and `@name` is optional; both are `ST_Xstring`.
 */
function tableNames(pkg: XlsxPackage, partName: string): readonly string[] {
  const part = pkg.parts.get(partName);
  if (!part) return [];
  const key = part.kind === 'xml' ? part.xml : part.bytes;
  const hit = tableCache.get(key);
  if (hit) return hit;
  let root: Element | undefined;
  if (part.kind === 'xml') root = part.xml.root;
  else {
    const read = parseXmlPart(partName, decodeXml(part.bytes), { maxBytes: XML_HARD_MAX_BYTES });
    root = read.ok ? read.part.root : undefined;
  }
  const names = [attr(root, 'name'), attr(root, 'displayName')].filter(
    (n): n is string => n !== undefined
  );
  tableCache.set(key, names);
  return names;
}

/** The sheets, external workbooks, and tables of a package. Cached per package. */
export function formulaScope(pkg: XlsxPackage): FormulaScope {
  const hit = scopeCache.get(pkg);
  if (hit) return hit;
  const index = workbookIndex(pkg);
  const sheets = new Set((index?.sheets ?? []).map((s) => fold(s.name)));
  const root = index ? xmlPartOf(pkg, index.mainPart)?.root : undefined;
  const externals = children(child(root, 'x:externalReferences'), 'x:externalReference').length;
  const tables = new Set<string>();
  for (const name of pkg.parts.keys())
    if (contentTypeOf(pkg, name) === TABLE)
      for (const t of tableNames(pkg, name)) tables.add(fold(t));
  const names = new Set<string>();
  for (const d of children(child(root, 'x:definedNames'), 'x:definedName')) {
    const name = attr(d, 'name');
    if (name) names.add(fold(name));
  }
  const scope = { sheets, externals, tables, names };
  scopeCache.set(pkg, scope);
  return scope;
}

/** Whether `[book]` is an external reference index that the workbook holds. */
function knownBook(book: string, scope: FormulaScope): boolean {
  if (book.length === 0 || book.length > 6) return false;
  for (let i = 0; i < book.length; i += 1) {
    const c = book.charCodeAt(i);
    if (c < 48 || c > 57) return false;
  }
  const n = Number(book);
  return n >= 1 && n <= scope.externals;
}

function prefixProblem(sheet: SheetPrefix, scope: FormulaScope): FormulaReferenceProblem | null {
  if (sheet.book !== undefined) return knownBook(sheet.book, scope) ? null : 'unknown-external';
  return sheet.sheets.every((s) => scope.sheets.has(fold(s))) ? null : 'unknown-sheet';
}

/**
 * Whether the words before token `k` and its unquoted sheet name are together a sheet name that
 * has a space: `My sheet!A1` typed without quotes.
 */
function unquotedSpaced(tokens: readonly FormulaToken[], k: number, scope: FormulaScope): boolean {
  const sheet = (tokens[k] as { sheet?: SheetPrefix }).sheet!;
  if (sheet.book !== undefined || sheet.sheets.length !== 1 || sheet.raw.startsWith("'"))
    return false;
  let text = sheet.sheets[0]!;
  for (let j = k - 1, words = 0; j >= 0 && words < MAX_NAME_WORDS; j -= 1) {
    const t = tokens[j]!;
    const plain = !('sheet' in t) || t.sheet === undefined;
    if (t.kind === 'space') text = t.text + text;
    else if (plain && (t.kind === 'name' || t.kind === 'reference' || t.kind === 'number')) {
      text = t.text + text;
      words += 1;
      if (scope.sheets.has(fold(text))) return true;
    } else break;
  }
  return false;
}

/** The first reference of `tokens` that names nothing the workbook holds, or null. */
export function checkFormulaReferences(
  tokens: readonly FormulaToken[],
  scope: FormulaScope
): { readonly problem: FormulaReferenceProblem; readonly at: number } | null {
  for (let k = 0; k < tokens.length; k += 1) {
    const t = tokens[k]!;
    if (t.kind === 'structured') {
      const name = t.text.slice(0, t.text.indexOf('['));
      if (name.length > 0 && !scope.tables.has(fold(name)))
        return { problem: 'unknown-table', at: t.start };
      continue;
    }
    const sheet = 'sheet' in t ? t.sheet : undefined;
    if (!sheet) continue;
    if (unquotedSpaced(tokens, k, scope)) return { problem: 'unquoted-sheet-name', at: t.start };
    const problem = prefixProblem(sheet, scope);
    if (problem) return { problem, at: t.start };
  }
  return null;
}
