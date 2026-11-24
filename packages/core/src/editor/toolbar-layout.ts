// The rules of the toolbar that do not depend on a framework: the groups of the bar from the
// chrome registry, how many fit before the "More" button, the arrow keys of the bar and of a
// popup, where a popup opens, and the typed font size. Moved from the React adapter's
// `toolbar/Toolbar.tsx` and `toolbar/parts.tsx`, so the React and Vue toolbars lay out and move
// the same way. Each key function reads a structural event, which a DOM event and a React
// synthetic event both satisfy. Adapter support, not a stable API.

import { inlineStep, isRightToLeft } from './chrome-direction.ts';
import {
  CHROME_GROUPS,
  chromeSlotId,
  type ChromeGroup,
  type ChromeSlotId,
} from './chrome-controls.ts';
import { visibleArea, windowScale } from './host-frame.ts';
import { TOOLBAR_COLLAPSE_ORDER, collapseOrder } from './toolbar-arrange.ts';
import { scrollIntoNearestScroller } from './scroll-nearest.ts';

/**
 * The groups of the bar, in registry order, without the hidden slots. A group whose slots are
 * all hidden is left out. @internal
 */
export function toolbarGroups(hidden: readonly ChromeSlotId[] = []): readonly ChromeGroup[] {
  const leave = new Set<string>(hidden);
  const groups: ChromeGroup[] = [];
  for (const g of CHROME_GROUPS as readonly ChromeGroup[]) {
    const controls = g.controls.filter((c) => !leave.has(chromeSlotId(g, c)));
    if (controls.length === 0) continue;
    groups.push(controls.length === g.controls.length ? g : { ...g, controls });
  }
  return groups;
}

/** The width of the "More" button with its gap, in CSS pixels. */
const MORE_WIDTH = 40;

/** One group of the bar, as the collapse rule reads it. @internal */
export interface ToolbarFitGroup {
  readonly id: string;
  /** The host's collapse priority, else the group's place in `TOOLBAR_COLLAPSE_ORDER`. */
  readonly priority?: number | undefined;
  /** A pinned group stays in the bar at every width. */
  readonly pinned?: boolean | undefined;
}

/** Nothing collapses. Shared, so an answer with nothing hidden keeps one identity. */
const NONE: readonly string[] = [];

/**
 * The ids of the groups that leave the bar for the "More" panel, in bar order. Nothing leaves
 * while a group has no measured width, so the next render measures it, nor when every group
 * fits in `width` (0 or less: no measure yet). Else groups leave in `collapseOrder` (a lower
 * priority first, a group with no priority before every built-in group), one whole group at a
 * time, until the rest and the More button fit. A pinned group never leaves. Each group after
 * the first is charged its gap, as DOCX Editor's `useToolbarOverflow.ts` charges `groupGaps`,
 * and so is the More button. With the default priorities, the groups that stay are the
 * longest start of the bar that fits, as `groupsThatFit` counts it. @internal
 */
export function toolbarCollapse(
  groups: readonly ToolbarFitGroup[],
  widths: ReadonlyMap<string, number>,
  width: number,
  gap = 0
): readonly string[] {
  if (groups.some((g) => !widths.has(g.id))) return NONE;
  const sizes = groups.map((g) => widths.get(g.id) ?? 0);
  const total = sizes.reduce((a, b) => a + b, 0) + gap * Math.max(0, groups.length - 1);
  if (width <= 0 || total <= width) return NONE;
  let used = MORE_WIDTH + sizes.reduce((a, b) => a + gap + b, 0);
  const priorities = new Map<string, number>();
  for (const g of groups)
    if (g.priority !== undefined && !Number.isNaN(g.priority)) priorities.set(g.id, g.priority);
  const pinned = new Set(groups.filter((g) => g.pinned).map((g) => g.id));
  const hidden = new Set<string>();
  const order = collapseOrder(
    groups.map((g) => g.id),
    TOOLBAR_COLLAPSE_ORDER,
    priorities
  );
  for (const id of order) {
    if (used <= width) break;
    if (pinned.has(id)) continue;
    hidden.add(id);
    used -= gap + (widths.get(id) ?? 0);
  }
  return groups.filter((g) => hidden.has(g.id)).map((g) => g.id);
}

/**
 * How many groups, from the start, fit in a bar whose content box is `width` pixels wide, with
 * `gap` pixels between two items. Each group after the first is charged its gap, as DOCX
 * Editor's `useToolbarOverflow.ts` charges `groupGaps`, and so is the More button. @internal
 */
export function groupsThatFit(widths: readonly number[], width: number, gap = 0): number {
  const total = widths.reduce((a, b) => a + b, 0) + gap * Math.max(0, widths.length - 1);
  if (width <= 0 || total <= width) return widths.length;
  let used = MORE_WIDTH;
  let count = 0;
  for (const w of widths) {
    if (used + gap + w > width) break;
    used += gap + w;
    count += 1;
  }
  return count;
}

/** The content width of the bar (its padding left out) and the gap between its items. */
function barSpace(el: HTMLElement): { width: number; gap: number } {
  const style = el.ownerDocument.defaultView?.getComputedStyle(el);
  const px = (v: string | undefined) => {
    const n = parseFloat(v ?? '');
    return Number.isFinite(n) ? n : 0;
  };
  return {
    width: el.clientWidth - px(style?.paddingLeft) - px(style?.paddingRight),
    gap: px(style?.columnGap),
  };
}

/**
 * The attribute of the block that holds a host's loose toolbar children. The block never
 * collapses, so the bar charges its width, and its gap, as fixed width, as DOCX Editor's
 * `useToolbarOverflow.ts` charges its `FIXED_ATTRIBUTE` block. @internal
 */
export const TOOLBAR_FIXED_ATTRIBUTE = 'data-toolbar-fixed';

/**
 * Measure the groups that the bar `el` shows into `widths`, and return the ids of the groups of
 * `groups` that leave the bar so the rest fit beside the fixed block (`TOOLBAR_FIXED_ATTRIBUTE`),
 * by `toolbarCollapse`. @internal
 */
export function measureToolbar(
  el: HTMLElement,
  groups: readonly ToolbarFitGroup[],
  widths: Map<string, number>
): readonly string[] {
  // A box is in window pixels and the bar's space in its CSS pixels, which differ inside a host
  // that scales the editor.
  const scale = windowScale(el).x;
  // Only the groups of the bar: host content can carry a `data-group` of its own.
  for (const g of el.querySelectorAll<HTMLElement>(':scope > [data-group]'))
    widths.set(g.dataset.group ?? '', g.getBoundingClientRect().width / scale);
  const space = barSpace(el);
  const fixed = el.querySelector(`:scope > [${TOOLBAR_FIXED_ATTRIBUTE}]`);
  // `toolbarCollapse` reads a width of 0 or less as a bar with no measure yet, which shows every
  // group. A measured bar that the fixed block fills keeps a width of 1, so it shows only the
  // pinned groups.
  const width =
    fixed && space.width > 0
      ? Math.max(1, space.width - fixed.getBoundingClientRect().width / scale - space.gap)
      : space.width;
  return toolbarCollapse(groups, widths, width, space.gap);
}

/**
 * Watch the bar `el`, each group that it shows, and its fixed block for a change of size, as
 * when fonts load or a value changes, and call `measure`. Returns the stop function. @internal
 */
export function observeToolbar(el: HTMLElement, measure: () => void): () => void {
  const Observer = el.ownerDocument.defaultView?.ResizeObserver;
  const observer = Observer ? new Observer(measure) : null;
  observer?.observe(el);
  for (const g of el.querySelectorAll(
    `:scope > [data-group], :scope > [${TOOLBAR_FIXED_ATTRIBUTE}]`
  ))
    observer?.observe(g);
  return () => observer?.disconnect();
}

/** The part of a key event that the focus moves read. @internal */
export interface FocusMoveKey {
  readonly key: string;
  readonly target: EventTarget | null;
  readonly currentTarget: EventTarget | null;
  /** True when a control of the host handled the key itself. */
  readonly defaultPrevented?: boolean;
  preventDefault(): void;
}

// The controls of the bar and the keys that a field keeps are after `toolbarItems` and
// `fieldOwnsKey` of pptx-editor/packages/core/src/editor/toolbar-focus.ts at 929ca860
// (EigenPal). Differences: the bar has no roving Tab stop, so a control with a negative
// `tabIndex` is no stop; a popup of the bar (`data-popup`) holds its menus, panels, and dialogs,
// in place of the `.pptx-menu` class, and a host dialog or menu keeps its keys as there, but only
// one inside the bar, so a bar in a host dialog keeps its keys; a control that cannot take the
// focus (a hidden input, a control of a disabled fieldset, a control with no box) is passed over;
// the keys stop at the ends of the bar instead of wrapping, as the bar did before; a key that a
// host control handled (`defaultPrevented`) does not move the focus; and a field whose caret the page
// cannot read (an email or a number field) keeps its keys, a range or a radio keeps the arrows
// that change its value, and an input that acts as a button keeps no key.

const BAR_ITEMS = 'button, input, select, textarea, a[href], [tabindex]';

/**
 * The controls of the "More" panel that take the focus when it opens from the keyboard: the
 * controls of the bar (`BAR_ITEMS`), so a host element of `asChild` (a `<span role="button">`
 * with a Tab stop, a link) counts as a built-in button does. A control with `disabled` takes
 * no focus, so it is left out; a built-in control that is off stays (`aria-disabled`). @internal
 */
export const TOOLBAR_PANEL_ITEMS =
  'button:not([disabled]), input:not([disabled]), select:not([disabled]), ' +
  'textarea:not([disabled]), a[href], [tabindex]:not([tabindex="-1"]):not([disabled]), ' +
  '[role="button"]:not([disabled])';

/**
 * The controls of the bar `bar` in order, at any depth: not those of a popup, a dialog, or a
 * menu, and not hidden ones.
 */
function barItems(bar: HTMLElement): HTMLElement[] {
  return [...bar.querySelectorAll<HTMLElement>(BAR_ITEMS)].filter(
    (el) =>
      // A disabled control of the bar stays a stop (`aria-disabled`), so the keyboard reaches
      // its reason. A host's `disabled` control, also in a disabled fieldset, takes no focus.
      !(el as HTMLButtonElement).disabled &&
      !el.matches(':disabled, input[type="hidden"]') &&
      el.tabIndex >= 0 &&
      !inBar(
        bar,
        el,
        '[hidden], [data-popup], [role="dialog"], [role="menu"], [aria-hidden="true"]'
      )
  );
}

/** Whether `el` is in an element of the bar `bar` that `selector` names: not the bar or above it. */
function inBar(bar: HTMLElement, el: Element, selector: string): boolean {
  const hit = el.closest(selector);
  return hit !== null && hit !== bar && bar.contains(hit);
}

/** Input types that act as a button: the bar takes every key. */
const BUTTON_INPUTS = new Set(['button', 'submit', 'reset', 'image', 'checkbox', 'color', 'file']);

/** The part of a field that `fieldOwnsKey` reads. @internal */
export interface FieldLike {
  readonly localName?: string;
  readonly type?: string;
  readonly isContentEditable?: boolean;
  readonly value?: string;
  readonly selectionStart?: number | null;
  readonly selectionEnd?: number | null;
}

/**
 * Whether a field keeps the key. A text field keeps Home and End for its caret, and keeps
 * ArrowLeft and ArrowRight until the caret is at that end of the text, as the WAI-ARIA toolbar
 * pattern says: so the arrows still leave a field. A text area and an editable element keep
 * every key. A list keeps Home and End. @internal
 */
export function fieldOwnsKey(target: EventTarget | null, key: string, forward: boolean): boolean {
  const el = target as FieldLike | null;
  if (!el) return false;
  if (el.localName === 'textarea' || el.isContentEditable) return true;
  const ends = key === 'Home' || key === 'End';
  if (el.localName === 'select') return ends;
  if (el.localName !== 'input') return false;
  const type = (el.type ?? 'text').toLowerCase();
  if (BUTTON_INPUTS.has(type)) return false;
  // The arrows change the value of a radio and a range. Home and End also set a range.
  if (type === 'radio') return !ends;
  if (type === 'range') return true;
  if (ends) return true;
  const start = el.selectionStart ?? null;
  const end = el.selectionEnd ?? null;
  // A field whose caret the page cannot read, such as an email or a number field, keeps the keys.
  if (start === null || end === null) return true;
  if (start !== end) return true;
  return forward ? end < (el.value ?? '').length : start > 0;
}

/**
 * Move focus between the controls of the bar with the arrow keys, Home, and End. `e.currentTarget`
 * is the bar. A control at any depth is a stop: a built-in control, a host control in a part or
 * a host group, and a loose control. @internal
 */
export function moveBarFocus(e: FocusMoveKey): void {
  if (!['ArrowLeft', 'ArrowRight', 'Home', 'End'].includes(e.key)) return;
  // A control of the host that moves its own focus, such as a radio group, handled the key.
  if (e.defaultPrevented) return;
  const target = e.target as HTMLElement;
  const bar = e.currentTarget as HTMLElement;
  // Only a control of the bar itself: a popup moves its own focus.
  if (inBar(bar, target, '[data-popup]')) return;
  // Along the reading direction: in a right-to-left bar, ArrowLeft moves to the next control.
  const step = inlineStep(e.key, bar);
  // A field keeps an arrow until its caret is at that end of its text. The end follows the
  // field's own direction, not the bar's, and not the direction of its text: at the end of its
  // value, a right-to-left field moves its caret no further on ArrowLeft, also for a value of
  // digits or of letters of a left-to-right script, and a left-to-right field on ArrowRight.
  // A browser text field moves its caret so, as measured in the browser engine of `test:e2e`.
  // A field with `dir="auto"` takes its computed direction from its text.
  const towardEnd = e.key === (isRightToLeft(target) ? 'ArrowLeft' : 'ArrowRight');
  if (fieldOwnsKey(target, e.key, towardEnd)) return;
  const items = barItems(bar);
  const at = items.indexOf(target);
  if (at < 0) return;
  e.preventDefault();
  // The controls to try, nearest first: one that does not take the focus, such as a control
  // that is not rendered, is passed over. At an end of the bar, there is none. A try is good when
  // the focus leaves the control, also when the item hands it on, as a host widget that focuses
  // its own field or a popup outside the bar.
  const root = target.getRootNode() as Document | ShadowRoot;
  const before = items.slice(0, at);
  const after = items.slice(at + 1);
  const tries =
    e.key === 'Home'
      ? before
      : e.key === 'End'
        ? after.reverse()
        : step === 1
          ? after
          : before.reverse();
  for (const item of tries) {
    item.focus();
    if (root.activeElement !== target) return;
  }
}

/**
 * Move focus among the items (`selector`) of a popup, `e.currentTarget`, with the arrow keys,
 * Home, and End. `grid` adds Left and Right, for a palette. Home and End in a text field move its
 * caret, not the focus. @internal
 */
export function movePopupFocus(e: FocusMoveKey, selector: string, grid = false): boolean {
  const target = e.target as HTMLElement;
  const field = target.tagName === 'INPUT';
  const items = [...(e.currentTarget as HTMLElement).querySelectorAll<HTMLElement>(selector)];
  const at = items.indexOf(target);
  const last = items.length - 1;
  // Left and Right follow the reading direction: in a right-to-left palette, ArrowLeft moves to
  // the next swatch, as in the bar (`moveBarFocus`).
  const step = grid ? inlineStep(e.key, e.currentTarget as HTMLElement) : 0;
  const forward = e.key === 'ArrowDown' || step === 1;
  const back = e.key === 'ArrowUp' || step === -1;
  const next = forward
    ? Math.min(last, at + 1)
    : back
      ? Math.max(0, at - 1)
      : e.key === 'Home' && !field
        ? 0
        : e.key === 'End' && !field
          ? last
          : -1;
  if (next < 0) return false;
  e.preventDefault();
  items[next]?.focus();
  return true;
}

/** The space between a popup and its trigger, in CSS pixels. */
const POPUP_GAP = 4;

/** The padding and border of `el` at its top and bottom, in CSS pixels. */
function edgesOf(el: Element): number {
  const style = el.ownerDocument.defaultView?.getComputedStyle(el);
  const px = (v: string | undefined) => parseFloat(v ?? '') || 0;
  return (
    px(style?.paddingTop) +
    px(style?.paddingBottom) +
    px(style?.borderTopWidth) +
    px(style?.borderBottomWidth)
  );
}

/**
 * The height of `el` with its top and bottom margins, in CSS pixels, or 0 with no `el`. `scale`
 * is the host's scale of the editor (`windowScale`), which the measured box includes.
 */
function outerHeightOf(el: Element | null, scale = 1): number {
  if (!el) return 0;
  const style = el.ownerDocument.defaultView?.getComputedStyle(el);
  const px = (v: string | undefined) => parseFloat(v ?? '') || 0;
  return el.getBoundingClientRect().height / scale + px(style?.marginTop) + px(style?.marginBottom);
}

/**
 * The least height that shows one item (`selector`) of a popup: the item, the popup's padding
 * and border, and the entry above the list.
 */
function leastOf(popup: HTMLElement, selector: string, scale: number): number {
  return (
    outerHeightOf(popup.querySelector(selector), scale) +
    edgesOf(popup) +
    outerHeightOf(popup.querySelector('.xlsx-toolbar-entry'), scale)
  );
}

/**
 * Cap a popup to `room` window pixels, so that it scrolls inside itself and never runs past the
 * visible part of the window. No cap when the popup fits, or when `room` cannot show one item
 * (`leastOf`): a popup that shows no item is worse than one that runs past the window. `scale`
 * converts the room to the popup's CSS pixels.
 */
function capPopup(popup: HTMLElement, room: number, selector: string, scale: number): void {
  const css = room / scale;
  if (css <= 0 || css < leastOf(popup, selector, scale) || popup.scrollHeight <= css) return;
  popup.style.maxHeight = `${Math.floor(css)}px`;
  popup.dataset.capped = '';
}

/**
 * Place the open popup (`[data-popup]`) of the toolbar part `root`. A popup that would run past
 * the editor's right edge opens from its trigger's end, and is then held inside the editor's
 * left edge. A popup that would run past the bottom of the visible area opens above its trigger
 * when there is more room there. The visible area is the window, less the host containers that
 * clip their overflow (`visibleArea`). With `inline` (inside the overflow panel), the popup
 * opens in flow, and the panel scrolls it into view. `selector` names the items of the popup.
 * @internal
 */
export function placePopup(root: HTMLElement, inline: boolean, selector: string): void {
  const popup = root.querySelector<HTMLElement>('[data-popup]');
  if (!popup) return;
  // In the overflow panel, the popup opens in flow and can start below the panel's visible
  // part: the panel scrolls it into view. Only the panel scrolls, never the page, and focus
  // stays where it is (a keyboard open then focuses an item). The scroll is PPTX Editor's
  // `scrollIntoNearestScroller` (`scroll-nearest.ts` names its source). Intended difference
  // from docx-editor/packages/react/src/editor/toolbar/ToolbarOverflow.tsx at 03735f11
  // (EigenPal), which does not scroll its panel to an open list, and whose popups have no cap
  // to the room in the window (a few lists have a fixed cap): here a popup taller than its room
  // (`capPopup`) scrolls inside itself.
  // Window pixels to the popup's CSS pixels, inside a host that scales the editor. Read on the
  // editor, whose large box makes the rounding of its layout size negligible.
  const scale = windowScale(root.closest<HTMLElement>('.xlsx-editor') ?? root);
  if (inline) {
    // The room is the visible part of the panel, less its padding.
    const panel = root.closest('[role="dialog"]');
    const box = panel?.getBoundingClientRect();
    if (panel && box) {
      const area = visibleArea(panel);
      const shown = Math.min(area.bottom, box.bottom) - Math.max(area.top, box.top);
      if (area.bottom > area.top)
        capPopup(popup, shown - edgesOf(panel) * scale.y, selector, scale.y);
    }
    scrollIntoNearestScroller(popup, panel);
    return;
  }
  const doc = popup.ownerDocument;
  const bound = (root.closest('.xlsx-editor') ?? doc.documentElement).getBoundingClientRect();
  // The popup opens at the start of its trigger (`inset-inline-start`), so in a right-to-left
  // toolbar it grows to the left: the start and the end edges swap.
  const rtl = isRightToLeft(root);
  const opened = popup.getBoundingClientRect();
  if (rtl ? opened.left < bound.left : opened.right > bound.right) popup.dataset.align = 'end';
  const rect = popup.getBoundingClientRect();
  const trigger = root.getBoundingClientRect();
  if (!rtl && rect.left < bound.left) {
    popup.style.left = `${Math.round((bound.left - trigger.left) / scale.x)}px`;
    popup.style.right = 'auto';
  } else if (rtl && rect.right > bound.right) {
    popup.style.right = `${Math.round((trigger.right - bound.right) / scale.x)}px`;
    popup.style.left = 'auto';
  }
  const area = visibleArea(popup);
  const shows = area.bottom > area.top;
  const gap = POPUP_GAP * scale.y;
  const below = area.bottom - trigger.bottom - gap;
  const above = trigger.top - area.top - gap;
  // A popup opens above only when the room there shows one item: else it would run past the
  // top of the visible area, with no cap.
  const top =
    shows &&
    rect.height > below &&
    above > below &&
    (above - gap) / scale.y >= leastOf(popup, selector, scale.y);
  if (top) popup.dataset.side = 'top';
  // The room on the chosen side, inside the visible area: the page never scrolls to show a
  // popup. Intended difference from docx-editor/packages/react/src/components/ui/FontSizePicker.tsx
  // at 2741998e (EigenPal), which caps its list at a fixed 240 pixels: a fixed cap can still
  // run past a short window, and a long list must not scroll the page under the editor.
  if (shows) capPopup(popup, (top ? above : below) - gap, selector, scale.y);
}

/**
 * Focus the selected item (`selector`) of the open popup of `root`, or else the current item
 * (`aria-current`: an entry that names the value but is not picked), or the first. @internal
 */
export function focusPopupItem(root: HTMLElement | null, selector: string): void {
  const popup = root?.querySelector('[data-popup]');
  const items = popup ? [...popup.querySelectorAll<HTMLElement>(selector)] : [];
  const selected = items.find((el) => el.getAttribute('aria-selected') === 'true');
  const active = items.find((el) => el.getAttribute('aria-current') === 'true');
  (selected ?? active ?? items[0])?.focus();
}

/**
 * What a key on a popup trigger does. Enter, Space, and Down open the popup with the keyboard in
 * it. On an open popup, Down moves the keyboard into it, and Enter or Space close it. Null for
 * another key: the key is not the trigger's. @internal
 */
export function popupTriggerAction(
  key: string,
  open: boolean
): 'open' | 'focusIn' | 'close' | null {
  if (key !== 'ArrowDown' && key !== 'Enter' && key !== ' ') return null;
  if (!open) return 'open';
  return key === 'ArrowDown' ? 'focusIn' : 'close';
}

/** The lowest and highest font size of a cell format, in points (`x:sz`). */
const MIN_FONT_SIZE = 1;
const MAX_FONT_SIZE = 409;

/** A typed font size: a number of points from 1 to 409, in either decimal separator. @internal */
export function parseFontSize(text: string): string | null {
  const trimmed = text.trim().replace(',', '.');
  if (!/^\d{1,3}(\.\d{1,2})?$/.test(trimmed)) return null;
  const size = Number(trimmed);
  return size >= MIN_FONT_SIZE && size <= MAX_FONT_SIZE ? String(size) : null;
}

/**
 * Whether `value` is a six-digit hex color, the only text that a color swatch puts in an inline
 * style. @internal
 */
export const isHexColor = (value: string | null | undefined): value is string =>
  !!value && /^[0-9A-Fa-f]{6}$/.test(value);

/**
 * Run `save` for Control or Command and S in the editor that holds the bar `el`, as DOCX
 * Editor's menu bar (`DocxEditorMenu`) does, key `s` with the platform modifier. Scoped to this
 * editor: only a key in the viewport that holds the bar, so a host page's own field keeps the key.
 * Read in the capture phase, as the cell editor and the grid keep their keys. A key in an open
 * modal dialog does not save: the dialog's edit is not applied yet. A host dialog marks itself
 * with `aria-modal="true"`. Returns the stop function. @internal
 */
export function listenSaveKey(el: HTMLElement, save: () => void): () => void {
  const doc = el.ownerDocument;
  const onKey = (e: KeyboardEvent) => {
    if (!(e.metaKey || e.ctrlKey) || e.altKey || e.shiftKey) return;
    if (e.key.toLowerCase() !== 's') return;
    const scope = el.closest('.xlsx-editor') ?? el;
    const target = e.target as Node | null;
    if (!target || !scope.contains(target)) return;
    e.preventDefault();
    if ((target as Element).closest?.('[data-xlsx-dialog], [aria-modal="true"]')) return;
    save();
  };
  doc.addEventListener('keydown', onKey, true);
  return () => doc.removeEventListener('keydown', onKey, true);
}
