// SpreadsheetML string escapes (`ST_Xstring`, shared-commonSimpleTypes.xsd; ECMA-376 Part 1,
// 22.9.2.19): a character that XML 1.0 cannot hold is written as `_xHHHH_`, with four hexadecimal
// digits of its UTF-16 code unit. A literal `_xHHHH_` sequence in the text is itself escaped by
// writing its underscore as `_x005F_`, so it reads back as text. Both functions are one linear pass
// with no regex.

const isHex = (c: number): boolean =>
  (c >= 48 && c <= 57) || (c >= 65 && c <= 70) || (c >= 97 && c <= 102);

/** Whether `s` holds `_xHHHH_` at `i`. */
function escapeAt(s: string, i: number): boolean {
  if (i + 7 > s.length) return false;
  if (s.charCodeAt(i) !== 95 || s.charCodeAt(i + 1) !== 120 || s.charCodeAt(i + 6) !== 95)
    return false;
  for (let k = 2; k < 6; k += 1) if (!isHex(s.charCodeAt(i + k))) return false;
  return true;
}

/** Decode `_xHHHH_` escapes. Text without an escape comes back unchanged. */
export function decodeXstring(s: string): string {
  if (!s.includes('_x')) return s;
  let out = '';
  let i = 0;
  while (i < s.length) {
    if (escapeAt(s, i)) {
      out += String.fromCharCode(parseInt(s.slice(i + 2, i + 6), 16));
      i += 7;
    } else {
      out += s[i];
      i += 1;
    }
  }
  return out;
}

/** Whether XML 1.0 can hold the UTF-16 code unit `c` as character data. */
function xmlCanHold(c: number): boolean {
  if (c === 0x09 || c === 0x0a || c === 0x0d) return true;
  if (c < 0x20) return false;
  return c !== 0xfffe && c !== 0xffff;
}

/**
 * The offset of the first UTF-16 code unit of `s` that XML 1.0 cannot hold as character data
 * (a control character, U+FFFE, U+FFFF, or a lone surrogate), or -1 when every unit fits. A
 * valid surrogate pair fits. One linear pass.
 */
export function xmlInvalidAt(s: string): number {
  for (let i = 0; i < s.length; i += 1) {
    const c = s.charCodeAt(i);
    if (c >= 0xd800 && c <= 0xdbff) {
      const next = s.charCodeAt(i + 1);
      if (next >= 0xdc00 && next <= 0xdfff) {
        i += 1;
        continue;
      }
      return i;
    }
    if ((c >= 0xdc00 && c <= 0xdfff) || !xmlCanHold(c)) return i;
  }
  return -1;
}

/**
 * Encode text for an `ST_Xstring` value: escape each code unit XML cannot hold, and the
 * underscore of each literal `_xHHHH_` sequence. `decodeXstring(encodeXstring(s)) === s`.
 * A lone surrogate is escaped, because XML text cannot hold one.
 * Carriage returns are escaped too, because XML end-of-line handling turns a bare CR into LF.
 */
export function encodeXstring(s: string): string {
  let out = '';
  for (let i = 0; i < s.length; i += 1) {
    const c = s.charCodeAt(i);
    const high = c >= 0xd800 && c <= 0xdbff;
    const low = c >= 0xdc00 && c <= 0xdfff;
    if (high && i + 1 < s.length) {
      const next = s.charCodeAt(i + 1);
      if (next >= 0xdc00 && next <= 0xdfff) {
        out += s[i]! + s[i + 1]!;
        i += 1;
        continue;
      }
    }
    if (escapeAt(s, i)) out += '_x005F_';
    else if (c === 0x0d || high || low || !xmlCanHold(c))
      out += `_x${c.toString(16).toUpperCase().padStart(4, '0')}_`;
    else out += s[i];
  }
  return out;
}
