// The number format code parser (ECMA-376 Part 1, 18.8.31). It reads `x:numFmt/@formatCode`
// (`CT_NumFmt`, sml.xsd). One linear pass over the characters splits the code into sections and
// tokens, then a second pass over each section's tokens assigns the roles that formatting needs:
// integer, fraction, exponent, numerator, and denominator digits, scaling commas, and the date and
// time parts. No regular expression reads the code. The code length, the section count, and the
// bracketed tokens per section are capped; a code over a cap or with an unterminated quote or
// bracket is a problem, and the caller falls back to `General`.
//
// The parser lives in the store lane so that the store can refuse a custom code that it cannot
// read before it writes the code to the styles part (`store/format-patch.ts`). The cascade lane
// formats values with it (`cascade/number-format/format.ts`).

/** Why a format code cannot be read. The cascade reports it as a cascade problem. */
export interface FormatCodeProblem {
  readonly code:
    | 'format-code-too-long'
    | 'format-too-many-sections'
    | 'format-too-many-brackets'
    | 'format-code-unparsable';
  readonly detail?: string;
}

/**
 * The longest format code the parser reads. This is an engine limit: `x:numFmt/@formatCode`
 * (`CT_NumFmt`, sml.xsd) is an `ST_Xstring`, which sets no length limit.
 */
export const MAX_CODE_LENGTH = 255;
/** Positive, negative, zero, and text. */
export const MAX_SECTIONS = 4;
/** Color, condition, locale, and one elapsed time part. */
export const MAX_BRACKETS_PER_SECTION = 4;
/** The fraction denominator digits the approximation searches. */
const MAX_DENOMINATOR_DIGITS = 7;

export type DigitChar = '0' | '#' | '?';
export type DigitRole = 'int' | 'frac' | 'exp' | 'num' | 'den';
export type DatePart = 'y' | 'm' | 'd' | 'h' | 'n' | 's';
export type CondOp = '<' | '<=' | '>' | '>=' | '=' | '<>';

export type Token =
  /** `pad`: the text is one space as wide as this character (`_x`, ECMA-376 Part 1, 18.8.31). */
  | { kind: 'lit'; text: string; pad?: string }
  | { kind: 'digit'; char: DigitChar; role: DigitRole; index: number }
  | { kind: 'point' }
  | { kind: 'comma' }
  | { kind: 'percent' }
  | { kind: 'exp'; source: string; plus: boolean }
  | { kind: 'slash' }
  | { kind: 'text' }
  | { kind: 'general' }
  | { kind: 'fill'; char: string }
  | { kind: 'date'; part: DatePart; length: number }
  | { kind: 'elapsed'; part: 'h' | 'n' | 's'; length: number }
  | { kind: 'ampm'; am: string; pm: string }
  | { kind: 'subsec'; digits: number }
  | { kind: 'den'; value: number };

export interface Section {
  readonly kind: 'number' | 'date' | 'text' | 'general';
  readonly tokens: readonly Token[];
  /** `RRGGBB` of a color token. */
  readonly color?: string;
  readonly cond?: { readonly op: CondOp; readonly value: number };
  /** Digit placeholders by role, left to right. */
  readonly int: readonly DigitChar[];
  readonly frac: readonly DigitChar[];
  readonly exp: readonly DigitChar[];
  readonly num: readonly DigitChar[];
  readonly den: readonly DigitChar[];
  /** A fixed fraction denominator, as in `# ?/8`. */
  readonly fixedDen?: number;
  readonly thousands: boolean;
  /** Powers of ten applied before formatting: +2 for each `%`, -3 for each scaling comma. */
  readonly shift: number;
  readonly scientific: boolean;
  readonly fraction: boolean;
  /** The section shows a 12-hour clock. */
  readonly ampm: boolean;
  /** Digits of fractional seconds. */
  readonly subsec: number;
}

export interface ParsedFormat {
  readonly sections: readonly Section[];
}

export type ParseResult =
  { ok: true; format: ParsedFormat } | { ok: false; problem: FormatCodeProblem };

const NAMED_COLORS: Readonly<Record<string, string>> = {
  black: '000000',
  blue: '0000FF',
  cyan: '00FFFF',
  green: '00FF00',
  magenta: 'FF00FF',
  red: 'FF0000',
  white: 'FFFFFF',
  yellow: 'FFFF00',
};

/** The default indexed palette, entries 8 to 63 (Part 1, 18.8.27); `[Color n]` is entry n + 7. */
const INDEXED_COLORS: readonly string[] = [
  '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
  '800000', '008000', '000080', '808000', '800080', '008080', 'C0C0C0', '808080',
  '9999FF', '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF',
  '000080', 'FF00FF', 'FFFF00', '00FFFF', '800080', '800000', '008080', '0000FF',
  '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99',
  '3366FF', '33CCCC', '99CC00', 'FFCC00', 'FF9900', 'FF6600', '666699', '969696',
  '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333',
]; // prettier-ignore

interface RawSection {
  tokens: Token[];
  color?: string;
  cond?: { op: CondOp; value: number };
}

const lower = (ch: string): string => ch.toLowerCase();

const unparsable = (detail: string): FormatCodeProblem => ({
  code: 'format-code-unparsable',
  detail,
});

/** Reads one format code into sections. Linear in the code length. */
export function parseFormatCode(code: string): ParseResult {
  if (code.length > MAX_CODE_LENGTH)
    return { ok: false, problem: { code: 'format-code-too-long', detail: code.slice(0, 16) } };
  const raw: RawSection[] = [];
  let current: RawSection = { tokens: [] };
  let brackets = 0;
  let i = 0;
  const n = code.length;
  while (i < n) {
    const ch = code[i]!;
    if (ch === ';') {
      raw.push(current);
      if (raw.length >= MAX_SECTIONS)
        return { ok: false, problem: { code: 'format-too-many-sections' } };
      current = { tokens: [] };
      brackets = 0;
      i++;
      continue;
    }
    if (ch === '"') {
      const end = code.indexOf('"', i + 1);
      if (end < 0) return { ok: false, problem: unparsable(code.slice(i, i + 16)) };
      current.tokens.push({ kind: 'lit', text: code.slice(i + 1, end) });
      i = end + 1;
      continue;
    }
    if (ch === '\\') {
      if (i + 1 < n) current.tokens.push({ kind: 'lit', text: code[i + 1]! });
      i += 2;
      continue;
    }
    if (ch === '_') {
      if (i + 1 < n) current.tokens.push({ kind: 'lit', text: ' ', pad: code[i + 1]! });
      i += 2;
      continue;
    }
    if (ch === '*') {
      if (i + 1 < n) current.tokens.push({ kind: 'fill', char: code[i + 1]! });
      i += 2;
      continue;
    }
    if (ch === '[') {
      const end = code.indexOf(']', i + 1);
      if (end < 0) return { ok: false, problem: unparsable(code.slice(i, i + 16)) };
      if (++brackets > MAX_BRACKETS_PER_SECTION)
        return { ok: false, problem: { code: 'format-too-many-brackets' } };
      const problem = readBracket(code.slice(i + 1, end), current);
      if (problem) return { ok: false, problem };
      i = end + 1;
      continue;
    }
    i = readPlain(code, i, current.tokens);
  }
  raw.push(current);
  return { ok: true, format: { sections: raw.map(analyze) } };
}

/** One bracketed token: a color, a condition, a locale or currency, or an elapsed time part. */
function readBracket(body: string, into: RawSection): FormatCodeProblem | undefined {
  const key = body.toLowerCase();
  if (Object.hasOwn(NAMED_COLORS, key)) {
    into.color = NAMED_COLORS[key];
    return undefined;
  }
  if (key.startsWith('color')) {
    const index = readInteger(key, 5);
    if (index === undefined || index < 1 || index > INDEXED_COLORS.length)
      return unparsable(`[${body.slice(0, 16)}]`);
    into.color = INDEXED_COLORS[index - 1];
    return undefined;
  }
  const first = body[0];
  if (first === '<' || first === '>' || first === '=') {
    const cond = readCondition(body);
    if (!cond) return unparsable(`[${body.slice(0, 16)}]`);
    into.cond = cond;
    return undefined;
  }
  if (first === '$') {
    // `[$text-locale]`: the text before `-` is a literal; the locale identifier selects nothing.
    const dash = body.indexOf('-');
    const text = body.slice(1, dash < 0 ? body.length : dash);
    if (text) into.tokens.push({ kind: 'lit', text });
    return undefined;
  }
  if (key.length > 0 && isRun(key, 'h'))
    into.tokens.push({ kind: 'elapsed', part: 'h', length: key.length });
  else if (key.length > 0 && isRun(key, 'm'))
    into.tokens.push({ kind: 'elapsed', part: 'n', length: key.length });
  else if (key.length > 0 && isRun(key, 's'))
    into.tokens.push({ kind: 'elapsed', part: 's', length: key.length });
  // Other bracketed tokens (calendar and numeral system selectors) select nothing here.
  return undefined;
}

function isRun(text: string, ch: string): boolean {
  for (let k = 0; k < text.length; k++) if (text[k] !== ch) return false;
  return true;
}

/** A non-negative decimal integer of at most 9 digits from `at` to the end, or undefined. */
function readInteger(text: string, at: number): number | undefined {
  if (at >= text.length || text.length - at > 9) return undefined;
  let value = 0;
  for (let k = at; k < text.length; k++) {
    const d = text.charCodeAt(k) - 48;
    if (d < 0 || d > 9) return undefined;
    value = value * 10 + d;
  }
  return value;
}

function readCondition(body: string): { op: CondOp; value: number } | undefined {
  let at = 1;
  let op: CondOp;
  const two = body.slice(0, 2);
  if (two === '<=' || two === '>=' || two === '<>') {
    op = two;
    at = 2;
  } else op = body[0] as CondOp;
  const rest = body.slice(at).trim();
  if (rest.length === 0 || rest.length > 32) return undefined;
  // A plain decimal number: sign, digits, one point, an optional exponent. Checked by a loop.
  for (let k = 0; k < rest.length; k++) {
    const c = rest[k]!;
    if (!((c >= '0' && c <= '9') || c === '.' || c === '-' || c === '+' || c === 'e' || c === 'E'))
      return undefined;
  }
  const value = Number(rest);
  return Number.isFinite(value) ? { op, value } : undefined;
}

/** Reads one plain token at `i` and returns the index after it. */
function readPlain(code: string, i: number, out: Token[]): number {
  const ch = code[i]!;
  const lc = lower(ch);
  switch (ch) {
    case '0':
    case '#':
    case '?':
      out.push({ kind: 'digit', char: ch, role: 'int', index: 0 });
      return i + 1;
    case '.':
      out.push({ kind: 'point' });
      return i + 1;
    case ',':
      out.push({ kind: 'comma' });
      return i + 1;
    case '%':
      out.push({ kind: 'percent' });
      return i + 1;
    case '/':
      out.push({ kind: 'slash' });
      return i + 1;
    case '@':
      out.push({ kind: 'text' });
      return i + 1;
  }
  if (lc === 'g' && code.slice(i, i + 7).toLowerCase() === 'general') {
    out.push({ kind: 'general' });
    return i + 7;
  }
  if (lc === 'a') {
    const five = code.slice(i, i + 5);
    if (five.toLowerCase() === 'am/pm') {
      out.push({ kind: 'ampm', am: five.slice(0, 2), pm: five.slice(3, 5) });
      return i + 5;
    }
    const three = code.slice(i, i + 3);
    if (three.toLowerCase() === 'a/p') {
      out.push({ kind: 'ampm', am: three[0]!, pm: three[2]! });
      return i + 3;
    }
  }
  if (lc === 'e' && (code[i + 1] === '+' || code[i + 1] === '-')) {
    out.push({ kind: 'exp', source: code.slice(i, i + 2), plus: code[i + 1] === '+' });
    return i + 2;
  }
  if (lc === 'y' || lc === 'm' || lc === 'd' || lc === 'h' || lc === 's') {
    let end = i + 1;
    while (end < code.length && lower(code[end]!) === lc) end++;
    out.push({ kind: 'date', part: lc, length: end - i });
    return end;
  }
  out.push({ kind: 'lit', text: ch });
  return i + 1;
}

/** Assigns roles to the tokens of one section. */
function analyze(raw: RawSection): Section {
  const tokens = raw.tokens;
  let isDate = false;
  let hasText = false;
  let hasGeneral = false;
  for (const t of tokens) {
    if (t.kind === 'date' || t.kind === 'elapsed' || t.kind === 'ampm') isDate = true;
    else if (t.kind === 'text') hasText = true;
    else if (t.kind === 'general') hasGeneral = true;
  }
  const base = { color: raw.color, cond: raw.cond };
  if (isDate) return { ...base, ...analyzeDate(tokens) };
  if (hasText || hasGeneral) {
    const out = tokens.map((t) => demote(t, true));
    return { ...base, ...emptyNumber(), kind: hasText ? 'text' : 'general', tokens: out };
  }
  return { ...base, ...analyzeNumber(tokens) };
}

function emptyNumber() {
  return {
    int: [] as DigitChar[],
    frac: [] as DigitChar[],
    exp: [] as DigitChar[],
    num: [] as DigitChar[],
    den: [] as DigitChar[],
    thousands: false,
    shift: 0,
    scientific: false,
    fraction: false,
    ampm: false,
    subsec: 0,
  };
}

/** A number token outside a number context is a literal. */
function demote(t: Token, keepText: boolean): Token {
  switch (t.kind) {
    case 'digit':
      return { kind: 'lit', text: t.char === '?' ? ' ' : t.char === '0' ? '0' : '' };
    case 'point':
      return { kind: 'lit', text: '.' };
    case 'comma':
      return { kind: 'lit', text: ',' };
    case 'percent':
      return { kind: 'lit', text: '%' };
    case 'slash':
      return { kind: 'lit', text: '/' };
    case 'exp':
      return { kind: 'lit', text: t.source };
    case 'den':
      return { kind: 'lit', text: String(t.value) };
    case 'text':
    case 'general':
      return keepText ? t : { kind: 'lit', text: '' };
    case 'date':
      return { kind: 'lit', text: t.part.repeat(Math.min(t.length, 8)) };
    default:
      return t;
  }
}

function analyzeDate(input: readonly Token[]): Omit<Section, 'color' | 'cond'> {
  const tokens: Token[] = [];
  let ampm = false;
  let subsec = 0;
  for (let k = 0; k < input.length; k++) {
    const t = input[k]!;
    if (t.kind === 'ampm') ampm = true;
    // `.0`, `.00`, `.000` after seconds: fractional seconds.
    if (t.kind === 'point' && isSeconds(tokens)) {
      let digits = 0;
      while (k + 1 < input.length) {
        const next = input[k + 1]!;
        if (next.kind !== 'digit' || next.char !== '0') break;
        digits++;
        k++;
      }
      if (digits > 0) {
        const kept = Math.min(digits, 3);
        tokens.push({ kind: 'subsec', digits: kept });
        subsec = Math.max(subsec, kept);
        continue;
      }
    }
    if (
      t.kind === 'date' ||
      t.kind === 'elapsed' ||
      t.kind === 'ampm' ||
      t.kind === 'lit' ||
      t.kind === 'fill'
    )
      tokens.push(t);
    else tokens.push(demote(t, false));
  }
  // `m` is minutes after an hour or before a second, ignoring literals between them.
  for (let k = 0; k < tokens.length; k++) {
    const t = tokens[k]!;
    if (t.kind !== 'date' || t.part !== 'm' || t.length > 2) continue;
    const before = neighbor(tokens, k, -1);
    const after = neighbor(tokens, k, 1);
    if (isHour(before) || isSecond(after))
      tokens[k] = { kind: 'date', part: 'n', length: t.length };
  }
  return { ...emptyNumber(), kind: 'date', tokens, ampm, subsec };
}

function isSeconds(tokens: readonly Token[]): boolean {
  const last = neighbor(tokens, tokens.length, -1);
  return isSecond(last);
}

function neighbor(tokens: readonly Token[], at: number, step: 1 | -1): Token | undefined {
  for (let k = at + step; k >= 0 && k < tokens.length; k += step) {
    const t = tokens[k]!;
    if (t.kind === 'date' || t.kind === 'elapsed') return t;
  }
  return undefined;
}

const isHour = (t: Token | undefined): boolean =>
  !!t && (t.kind === 'date' || t.kind === 'elapsed') && t.part === 'h';
const isSecond = (t: Token | undefined): boolean =>
  !!t && (t.kind === 'date' || t.kind === 'elapsed') && t.part === 's';

function analyzeNumber(input: readonly Token[]): Omit<Section, 'color' | 'cond'> {
  const tokens: Token[] = input.slice();
  const isDigit = (k: number) => tokens[k]?.kind === 'digit';
  // The exponent, the decimal point, and the fraction slash, each the first one that counts.
  let expAt = -1;
  for (let k = 0; k < tokens.length; k++) {
    if (tokens[k]!.kind === 'exp') {
      expAt = k;
      break;
    }
  }
  const mantissaEnd = expAt < 0 ? tokens.length : expAt;
  let slashAt = -1;
  if (expAt < 0) {
    for (let k = 1; k < tokens.length; k++) {
      if (tokens[k]!.kind !== 'slash' || !isDigit(k - 1)) continue;
      const next = tokens[k + 1];
      if (isDigit(k + 1) || (next?.kind === 'lit' && isDigitText(next.text))) slashAt = k;
      break;
    }
  }
  let pointAt = -1;
  if (slashAt < 0) {
    for (let k = 0; k < mantissaEnd; k++) {
      if (tokens[k]!.kind === 'point') {
        pointAt = k;
        break;
      }
    }
  }
  const out = emptyNumber();
  let fixedDen: number | undefined;
  if (slashAt >= 0) {
    out.fraction = true;
    // The numerator: the digits right before the slash. The integer: the digits before those.
    let numStart = slashAt;
    while (numStart > 0 && isDigit(numStart - 1)) numStart--;
    for (let k = 0; k < tokens.length; k++) {
      const t = tokens[k]!;
      if (t.kind !== 'digit') continue;
      if (k < numStart) t.role = 'int';
      else if (k < slashAt) t.role = 'num';
      else t.role = 'den';
    }
    // A fixed denominator: literal digits right after the slash, with any `0` among them.
    const first = tokens[slashAt + 1];
    if (first?.kind === 'lit' && isDigitText(first.text)) {
      let value = 0;
      let end = slashAt + 1;
      let digits = 0;
      while (end < tokens.length && digits < 9) {
        const t = tokens[end]!;
        if (t.kind === 'lit' && isDigitText(t.text)) value = value * 10 + Number(t.text);
        else if (t.kind === 'digit' && t.char === '0') value *= 10;
        else break;
        digits++;
        end++;
      }
      fixedDen = value;
      tokens.splice(slashAt + 1, end - slashAt - 1, { kind: 'den', value });
    }
  } else {
    for (let k = 0; k < tokens.length; k++) {
      const t = tokens[k]!;
      if (t.kind !== 'digit') continue;
      if (k > mantissaEnd) t.role = 'exp';
      else if (pointAt >= 0 && k > pointAt) t.role = 'frac';
      else t.role = 'int';
    }
  }
  // Commas: a separator between integer digits sets grouping; commas after the last mantissa
  // digit scale by 1000 each; any other comma is a literal.
  const intEnd = pointAt >= 0 ? pointAt : slashAt >= 0 ? slashAt : mantissaEnd;
  let lastMantissaDigit = -1;
  let firstIntDigit = -1;
  for (let k = 0; k < mantissaEnd; k++) {
    if (!isDigit(k)) continue;
    if (firstIntDigit < 0 && k < intEnd) firstIntDigit = k;
    lastMantissaDigit = k;
  }
  for (let k = 0; k < tokens.length; k++) {
    const t = tokens[k]!;
    if (t.kind === 'percent') out.shift += 2;
    if (t.kind !== 'comma') continue;
    let digitAfterInInt = false;
    for (let j = k + 1; j < intEnd; j++) if (isDigit(j)) digitAfterInInt = true;
    if (k < intEnd && firstIntDigit >= 0 && firstIntDigit < k && digitAfterInInt) {
      out.thousands = true;
      tokens[k] = { kind: 'lit', text: '' };
    } else if (
      lastMantissaDigit >= 0 &&
      k > lastMantissaDigit &&
      k < mantissaEnd &&
      !out.fraction
    ) {
      out.shift -= 3;
      tokens[k] = { kind: 'lit', text: '' };
    } else tokens[k] = { kind: 'lit', text: ',' };
  }
  // Index the digits of each role, left to right; demote stray tokens.
  for (let k = 0; k < tokens.length; k++) {
    const t = tokens[k]!;
    if (t.kind === 'digit') {
      const list = out[t.role];
      t.index = list.length;
      list.push(t.char);
    } else if (t.kind === 'point' && k !== pointAt) tokens[k] = { kind: 'lit', text: '.' };
    else if (t.kind === 'slash' && k !== slashAt) tokens[k] = { kind: 'lit', text: '/' };
    else if (t.kind === 'exp' && k !== expAt) tokens[k] = { kind: 'lit', text: t.source };
  }
  if (out.den.length > MAX_DENOMINATOR_DIGITS) out.den.length = MAX_DENOMINATOR_DIGITS;
  out.scientific = expAt >= 0;
  return {
    ...out,
    kind: 'number',
    tokens,
    ...(fixedDen !== undefined ? { fixedDen } : {}),
  };
}

function isDigitText(text: string): boolean {
  return text.length === 1 && text >= '1' && text <= '9';
}
