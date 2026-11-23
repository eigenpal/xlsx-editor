/**
 * The bytes of a workbook source, as DOCX Editor's `DocumentSource` accepts them. The tag check
 * also reads an `ArrayBuffer` of another window, which `instanceof` misses. A detached buffer gives
 * no bytes, as a detached `Uint8Array` does, so `load()` resolves with the failure `package`.
 * A `SharedArrayBuffer` gives a copy of its bytes in a buffer that is not shared.
 */
export function workbookBytes(source: Uint8Array | ArrayBuffer): Uint8Array {
  const tag = Object.prototype.toString.call(source);
  if (tag === '[object SharedArrayBuffer]') {
    return new Uint8Array(source as unknown as SharedArrayBuffer).slice();
  }
  if (tag !== '[object ArrayBuffer]') {
    return source as Uint8Array;
  }
  try {
    return new Uint8Array(source as ArrayBuffer);
  } catch {
    return new Uint8Array(0);
  }
}
