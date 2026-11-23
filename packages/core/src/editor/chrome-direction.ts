// The writing direction of a chrome element, for the keys and the popups that move along the
// reading direction. The adapters set `dir` on each chrome part (`useChromeDir`); without it, the
// part inherits the direction of the host page. Adapter support, not a stable API.

/**
 * Whether `el` lays out right to left, by its computed `direction`. That covers a CSS
 * `direction`, a `dir` attribute, and `dir="auto"` over right-to-left text, and a computed `ltr`
 * is trusted, such as a wrapper with `direction: ltr` inside a right-to-left page.
 *
 * A browser maps the `dir` attribute to `direction` through its own style sheet. A DOM without
 * that style sheet, such as a test DOM, reports `ltr` for an element with `dir="rtl"`, or no
 * direction at all. So the computed value decides only when `el` has no `dir="rtl"` ancestor,
 * or when that ancestor itself computes `rtl`. In every other case, the first inline
 * `direction` from `el` up to that ancestor, not counting the ancestor, decides, and without
 * one the attribute decides. A class rule that sets `direction: ltr` on the `dir="rtl"`
 * element itself is not detected: that element then counts as right to left. The chrome CSS
 * selects mirrored icons with `:dir(rtl)`, which follows the `dir` attribute alone. @internal
 */
export function isRightToLeft(el: Element): boolean {
  const view = el.ownerDocument.defaultView;
  const computed = view?.getComputedStyle(el).direction ?? '';
  const marked = el.closest('[dir]');
  const attribute = marked?.getAttribute('dir')?.toLowerCase();
  const mapsDir =
    !marked || attribute !== 'rtl' || view?.getComputedStyle(marked).direction === 'rtl';
  if (computed && mapsDir) return computed === 'rtl';
  for (let node: Element | null = el; node && node !== marked; node = node.parentElement) {
    const inline = (node as HTMLElement).style?.direction;
    if (inline === 'ltr' || inline === 'rtl') return inline === 'rtl';
  }
  return attribute === 'rtl';
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
