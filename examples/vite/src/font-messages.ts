// The demo's words for fonts that did not load. A stalled or blocked font download ends in one
// of these two messages, so the person knows what happened and what to do.

/** The reason of an open that failed because no font loaded. */
export const FONTS_FAILED = 'the fonts did not load';

/** What to do after an open that failed because no font loaded. */
export const FONTS_NEXT_STEP = 'Check the network connection, and then reload the page.';

/** The notice of a workbook that opened with some of its fonts missing. */
export function fontsMissingNotice(missing: number): string {
  const faces = missing === 1 ? '1 font face' : `${missing} font faces`;
  return `${faces} did not load. Text in ${missing === 1 ? 'that face' : 'those faces'} shows in another font. Reload the page to try again. A reload closes this workbook.`;
}
