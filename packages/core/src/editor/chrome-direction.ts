// The writing direction of a chrome element, for the keys and the popups that move along the
// reading direction. The adapters set `dir` on each chrome part (`useChromeDir`); without it, the
// part inherits the direction of the host page. Adapter support, not a stable API.

/**
 * Whether `el` lays out right to left: its computed `direction` is `rtl`, which covers a CSS
 * `direction`, a `dir` attribute, and `dir="auto"` over right-to-left text. The `dir` attribute
 * sets `direction` through the browser's own style sheet; a DOM without that style sheet, such
 * as a test DOM, reports `ltr`, so the nearest `dir` attribute of `rtl` also counts. The chrome
 * CSS selects mirrored icons with `:dir(rtl)`, which follows the `dir` attribute alone. @internal
 */
export function isRightToLeft(el: Element): boolean {
  const view = el.ownerDocument.defaultView;
  if (view?.getComputedStyle(el).direction === 'rtl') return true;
  return el.closest('[dir]')?.getAttribute('dir')?.toLowerCase() === 'rtl';
}

/**
 * The step of an arrow key along the reading direction of `el`: `1` for the next control, `-1`
 * for the previous one, and `0` for any other key. In a right-to-left element, ArrowLeft moves
 * to the next control, as the WAI-ARIA toolbar and tabs patterns describe. @internal
 */
export function inlineStep(key: string, el: Element): -1 | 0 | 1 {
  if (key !== 'ArrowRight' && key !== 'ArrowLeft') return 0;
  const right = key === 'ArrowRight' ? 1 : -1;
  return isRightToLeft(el) ? (-right as -1 | 1) : right;
}
