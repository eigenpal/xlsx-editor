// Sheet names and the sheet prefixes of formula text, for the sheet ops of design D1 of the
// change `workbook-structure-editing`.
//
// - Names: a sheet name has 1 to 31 characters, counted in UTF-16 code units, none of
//   `[ ] : * ? / \`, no apostrophe at the start or the end, no control character (U+0000 to
//   U+001F, tab and line breaks included), no other character that XML 1.0 cannot hold, no
//   `_xHHHH_` sequence, and no match with another sheet name without case, after both names
//   take their canonical composition (NFC). These are this editor's rules: the schema types
//   `x:sheet/@name` as `ST_Xstring` (`CT_Sheet`, `sml.xsd`; ECMA-376 Part 1, 22.9.2.19), with no
//   length limit, and in `ST_Xstring` a `_xHHHH_` sequence is an escape of one character. The
//   editor shows and matches a stored name as written, so it refuses to write a name with such
//   a sequence. The rules keep the name readable as a formula prefix (`sheet-ops.test.ts`,
//   "sheet names").
// - Prefixes: the bounded tokenizer reads the text. A rename writes the new name in each prefix
//   that names the sheet, quoted when the name is not a plain word, with each apostrophe
//   doubled, as the tokenizer reads a quoted name (`'It''s'!`). A delete writes `#REF!` in
//   place of each reference that names the sheet, with its address, so `Sheet1!B2` and
//   `Sheet1!A1:B2` become `#REF!`. Measured with the save check: a
//   formula or a defined name that holds `#REF!B2` does not open, and a measured sheet delete
//   writes `#REF!`. A move only
//   checks the text: a sheet range (`Sheet1:Sheet3!A1`) refuses it when the move changes the
//   set of sheets between its ends. A delete refuses a sheet range that starts or ends at the
//   sheet. A sheet between the ends leaves the range, and the text stays the same.
// - A prefix with an external workbook (`[1]Sheet1!`) names another workbook and keeps its
//   text. The prefix `[0]` names this workbook: a move checks its sheet range as a local one,
//   a delete of an end of its sheet range is refused with `sheet-range-blocked`, as for a local
//   one, and a rename or a delete of another sheet that it names is refused.
// - Text that the tokenizer cannot read keeps its text when it cannot name the sheet, and
//   refuses the op otherwise: a move with `sheet-range-blocked`, because only a sheet range
//   (`:` and `!`) can change under it, and a rename or a delete with `structure-blocked`.
//   Every other character keeps its text.

import { tokenizeFormula, type FormulaToken, type SheetPrefix } from '../formula/tokenizer.ts';
import { ambiguousSheetRange, crossSheetRangeAt } from '../formula/format.ts';
import { decodeXstring, xmlInvalidAt } from '../workbook/xstring.ts';
import { sheetNameKey, type WorkbookIndex } from '../workbook/workbook-index.ts';
import type { OpRejection } from '../store/ops.ts';

/** The longest sheet name, in UTF-16 code units (a character outside the BMP counts 2). */
export const MAX_SHEET_NAME = 31;

const FORBIDDEN = new Set(['[', ']', ':', '*', '?', '/', '\\']);

/**
 * Why `name` cannot name a sheet of the workbook, or undefined. `except` is the position of a
 * sheet that the name can match: the sheet that a rename names.
 */
export function sheetNameProblem(
  name: string,
  index: WorkbookIndex,
  except = -1
): OpRejection | undefined {
  const subject = name;
  const invalid = (detail: string): OpRejection => ({
    code: 'sheet-name-invalid',
    subject,
    detail: `${JSON.stringify(name)}: ${detail}`,
  });
  if (typeof name !== 'string' || name.length === 0) return invalid('a sheet name is empty');
  if (name.length > MAX_SHEET_NAME)
    return invalid(`a sheet name has at most ${MAX_SHEET_NAME} characters`);
  for (const c of name)
    if (FORBIDDEN.has(c)) return invalid(`a sheet name cannot hold ${JSON.stringify(c)}`);
  if (name.startsWith("'") || name.endsWith("'"))
    return invalid('a sheet name cannot start or end with an apostrophe');
  // U+0000 to U+001F, tab, line feed, and carriage return too: a sheet name is one line.
  if (/[\u0000-\u001f]/.test(name) || xmlInvalidAt(name) >= 0)
    return invalid('a sheet name holds a control character');
  // `ST_Xstring` reads `_xHHHH_` as an escape of one character (ECMA-376 Part 1, 22.9.2.19).
  if (decodeXstring(name) !== name)
    return invalid('a sheet name cannot hold an escape sequence such as _x0041_');
  // Canonical equivalents (`Café` with U+00E9 and with U+0301, `K` and the Kelvin sign) are
  // one name to a reader, so a new name cannot take a name that is equal after NFC.
  const key = sheetNameKey(name.normalize('NFC'));
  const taken = index.sheets.find(
    (s) => s.index !== except && sheetNameKey(s.name.normalize('NFC')) === key
  );
  // The subject is the name of the sheet that has the name, as the tab shows it.
  if (taken)
    return {
      code: 'sheet-name-taken',
      subject: taken.name,
      detail: `${JSON.stringify(name)}: the sheet ${JSON.stringify(taken.name)} has that name`,
    };
  return undefined;
}

/**
 * Whether a name is a plain word that formula text can hold without quotes: ASCII letters and
 * underscores only. A digit or a dot can make the name read as a cell reference (`A1.B`), an
 * R1C1 reference (`R1C1x`), or a number, so such a name is quoted. So are the R1C1 letters
 * (`R`, `C`, `RC`) and a name that starts with a boolean (`TRUEX`).
 */
function plainWord(name: string): boolean {
  if (!/^[A-Za-z_]+$/.test(name)) return false;
  const upper = name.toUpperCase();
  if (upper === 'R' || upper === 'C' || upper === 'RC') return false;
  return !upper.startsWith('TRUE') && !upper.startsWith('FALSE');
}

/** The prefix text of one sheet or a sheet range, with the `!`: `Data!`, `'Data 1'!`. */
export function sheetPrefixText(sheets: readonly string[]): string {
  if (sheets.every(plainWord)) return `${sheets.join(':')}!`;
  return `'${sheets.join(':').replace(/'/g, "''")}'!`;
}

/** What a sheet op does to the sheet prefixes of formula text. Sheet names are folded. */
export type PrefixEdit =
  | {
      readonly kind: 'rename';
      readonly sheet: string;
      readonly name: string;
      /** The folded names of the sheets of the workbook. A sheet range must name two of them. */
      readonly sheets?: ReadonlySet<string>;
    }
  | { readonly kind: 'delete'; readonly sheet: string }
  | {
      readonly kind: 'move';
      /**
       * Why the move changes the sheets that a sheet range from `first` to `last` (folded
       * names) spans, or undefined when it spans the same sheets after the move.
       */
      readonly spanProblem: (first: string, last: string) => string | undefined;
    };

/**
 * The rewritten text, or why the op cannot keep the text valid. `range` marks a sheet range
 * that a move or a delete would change (`sheet-range-blocked`).
 */
export type RewrittenText =
  | { readonly ok: true; readonly text: string; readonly changed: boolean }
  | { readonly ok: false; readonly reason: string; readonly range?: true };

/**
 * Why a move or a delete changes what a sheet range spans, or undefined. A move changes it when
 * the set of sheets between its ends changes. A delete changes it when the sheet is one of its
 * ends; a sheet between the ends leaves the range, and the text stays the same.
 */
function rangeSpanProblem(prefix: SheetPrefix, edit: PrefixEdit): string | undefined {
  const [first, last] = prefix.sheets.map(sheetNameKey) as [string, string];
  if (edit.kind === 'move') return edit.spanProblem(first, last);
  if (edit.kind === 'delete' && (first === edit.sheet || last === edit.sheet))
    return `the sheet range ${prefix.raw} starts or ends at the sheet`;
  return undefined;
}

/** Whether text that the tokenizer cannot read can name the sheet of the edit. */
function unreadableCanName(text: string, edit: PrefixEdit): boolean {
  const upper = sheetNameKey(text);
  if (edit.kind === 'move') return upper.includes(':') && upper.includes('!');
  // A quoted prefix doubles each apostrophe of the name (`'It''s'!A1`).
  return upper.includes(edit.sheet) || upper.includes(edit.sheet.replace(/'/g, "''"));
}

const isLocal = (prefix: SheetPrefix): boolean => prefix.book === undefined;

/** Why a rename cannot rewrite a sheet range prefix, or undefined. */
function rangeProblem(prefix: SheetPrefix, edit: PrefixEdit): string | undefined {
  if (edit.kind !== 'rename' || prefix.sheets.length !== 2) return undefined;
  if (ambiguousSheetRange(prefix))
    return `the sheet range ${prefix.raw} can be read as a cell reference`;
  if (edit.sheets && prefix.sheets.some((s) => !edit.sheets!.has(sheetNameKey(s))))
    return `the sheet range ${prefix.raw} names a sheet that is not in the workbook`;
  return undefined;
}

/** Rewrite the sheet prefixes of formula text (without the leading `=`) for a sheet op. */
export function rewriteSheetPrefixes(text: string, edit: PrefixEdit): RewrittenText {
  const read = tokenizeFormula(text);
  if (!read.ok) {
    if (!unreadableCanName(text, edit)) return { ok: true, text, changed: false };
    const reason = `the formula text cannot be read (${read.problem} at ${read.at})`;
    // A move reads unreadable text only for a sheet range that it can hold.
    return edit.kind === 'move' ? { ok: false, reason, range: true } : { ok: false, reason };
  }
  const tokens = read.tokens;
  // `A1:S!B2` can also be read as the sheet range `A1:S`: no edit of that sheet keeps its meaning.
  const cross = crossSheetRangeAt(tokens);
  if (cross >= 0) {
    const end = tokens.find((t) => t.start > cross && t.kind !== 'operator')!;
    const names = end.kind === 'reference' || end.kind === 'name' ? end.sheet!.sheets : [];
    const reason = `the range at ${cross} can be read as a sheet range that starts with a cell reference`;
    if (edit.kind === 'move') return { ok: false, reason, range: true };
    if (names.some((n) => sheetNameKey(n) === edit.sheet))
      return edit.kind === 'delete' ? { ok: false, reason, range: true } : { ok: false, reason };
  }
  let out = '';
  let last = 0;
  let changed = false;
  for (let i = 0; i < tokens.length; i += 1) {
    const token = tokens[i]!;
    if (token.kind !== 'reference' && token.kind !== 'name' && token.kind !== 'error') continue;
    const prefix = token.sheet;
    if (!prefix) continue;
    if (!isLocal(prefix)) {
      // `[0]` names this workbook. A move checks its sheet range as it checks a local one. A
      // rename or a delete does not rewrite the prefix, so it cannot name the sheet.
      if (prefix.book === '0' && edit.kind === 'move' && prefix.sheets.length === 2) {
        const reason = rangeSpanProblem(prefix, edit);
        if (reason) return { ok: false, reason, range: true };
      } else if (prefix.book === '0' && edit.kind !== 'move') {
        // A delete of an end of a `[0]` sheet range changes the range, as of a local one.
        const span = prefix.sheets.length === 2 ? rangeSpanProblem(prefix, edit) : undefined;
        if (span) return { ok: false, reason: span, range: true };
        const hit = prefix.sheets.some((s) => sheetNameKey(s) === edit.sheet);
        if (hit) return { ok: false, reason: `the reference ${token.text} cannot be rewritten` };
      }
      continue;
    }
    if (prefix.sheets.length === 2 && edit.kind !== 'rename') {
      const reason = rangeSpanProblem(prefix, edit);
      if (reason) return { ok: false, reason, range: true };
      continue;
    }
    if (edit.kind === 'move') continue;
    const hits = prefix.sheets.map((s) => sheetNameKey(s) === edit.sheet);
    if (!hits.some(Boolean)) continue;
    const problem = rangeProblem(prefix, edit);
    if (problem) return { ok: false, reason: problem };
    let start = token.start;
    let end = token.start + token.text.length;
    let next: string;
    if (edit.kind === 'rename') {
      const names = prefix.sheets.map((s, k) => (hits[k] ? edit.name : s));
      next = sheetPrefixText(names) + token.text.slice(prefix.raw.length);
    } else {
      // A range operand that the `:` operator joins (`Sheet1!A1:Sheet1!B2`) becomes one
      // `#REF!`, never `#REF!:#REF!`.
      const before = rangeEnd(tokens, i - 2, i - 1);
      if (before && before.start >= last) start = before.start;
      while (rangeEnd(tokens, i + 2, i + 1)) {
        i += 2;
        end = tokens[i]!.start + tokens[i]!.text.length;
      }
      next = '#REF!';
    }
    if (next === text.slice(start, end)) continue;
    out += text.slice(last, start) + next;
    last = end;
    changed = true;
  }
  return changed
    ? { ok: true, text: out + text.slice(last), changed }
    : { ok: true, text, changed };
}

/**
 * The operand at `at` when the token at `op` is a `:` operator that joins it, with no space, to
 * the operand on the other side; else undefined.
 */
function rangeEnd(
  tokens: readonly FormulaToken[],
  at: number,
  op: number
): FormulaToken | undefined {
  const colon = tokens[op];
  const operand = tokens[at];
  if (colon?.kind !== 'operator' || colon.text !== ':' || !operand) return undefined;
  const kind = operand.kind;
  if (kind !== 'reference' && kind !== 'name' && kind !== 'error') return undefined;
  const [left, right] = at < op ? [operand, colon] : [colon, operand];
  return left.start + left.text.length === right.start ? operand : undefined;
}
