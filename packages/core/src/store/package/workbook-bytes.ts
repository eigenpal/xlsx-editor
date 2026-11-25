/**
 * The bytes of a workbook source, as DOCX Editor's `DocumentSource` accepts them. The tag check
 * also reads an `ArrayBuffer` or a `Uint8Array` of another window, which `instanceof` misses.
 * A `Uint8Array` of another window gives a view of this window over the same bytes, because
 * the ZIP read tests `instanceof Uint8Array`. A detached buffer gives no bytes, as a detached
 * `Uint8Array` does, so the read fails as a package failure. A `SharedArrayBuffer` gives a
 * copy of its bytes in a buffer that is not shared. Any other value passes through, and the
 * ZIP read refuses it.
 */
export function workbookBytes(source: Uint8Array | ArrayBuffer): Uint8Array {
  if (source instanceof Uint8Array) return source;
  const tag = Object.prototype.toString.call(source);
  try {
    if (tag === '[object SharedArrayBuffer]') {
      return new Uint8Array(source as unknown as SharedArrayBuffer).slice();
    }
    if (tag === '[object ArrayBuffer]') return new Uint8Array(source as ArrayBuffer);
    if (tag === '[object Uint8Array]') {
      const view = source as unknown as Uint8Array;
      return new Uint8Array(view.buffer, view.byteOffset, view.byteLength);
    }
  } catch {
    return new Uint8Array(0);
  }
  return source as unknown as Uint8Array;
}
