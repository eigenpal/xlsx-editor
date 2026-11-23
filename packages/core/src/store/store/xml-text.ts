// One check of text that a format patch or a number locale carries into the styles part. The
// store index does not re-export this module, so it is not on a public entry.

/** Whether a string holds only characters that XML 1.0 text can carry, and no control codes. */
export function isPlainXmlText(s: string): boolean {
  for (let i = 0; i < s.length; i += 1) {
    const c = s.charCodeAt(i);
    if (c < 0x20 || c === 0x7f || c === 0xfffe || c === 0xffff) return false;
    // A lone surrogate is not a character.
    if (c >= 0xd800 && c <= 0xdbff) {
      const next = s.charCodeAt(i + 1);
      if (!(next >= 0xdc00 && next <= 0xdfff)) return false;
      i += 1;
    } else if (c >= 0xdc00 && c <= 0xdfff) return false;
  }
  return true;
}
