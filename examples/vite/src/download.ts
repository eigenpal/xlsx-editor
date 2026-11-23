// The download of a saved copy, for a browser with no file save picker.
//
// Taken from pptx-editor/examples/shared/demo-files.ts at e22f3f3b (`download`, EigenPal).
// Differences: it takes only a `Blob` that the page made, so no string URL can reach the link,
// and the caller sets the media type, so the `type` parameter is gone. Renames: `download` is
// `downloadBlob`, `data` is `blob`, `url` is `objectUrl`, and `a` is `link`. `sanitizeHref` drops
// a `blob:` URL, so `bun run check:security` allows this link only while the code of this file equals
// `BLOB_DOWNLOAD_CODE` in `scripts/check-security-grep.mjs`. A change to the code fails the gate.

// Saves `blob` as the file `name` through a temporary link.
export function downloadBlob(blob: Blob, name: string): void {
  const objectUrl = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.setAttribute('href', objectUrl);
  link.download = name;
  link.click();
  setTimeout(() => URL.revokeObjectURL(objectUrl), 1000);
}
