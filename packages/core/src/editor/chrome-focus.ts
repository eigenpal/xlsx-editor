// The focus rules of the editor chrome that do not depend on a framework: chrome `mousedown`
// keeps the keyboard on the grid (AGENTS.md "Conventions", focus stealing), a toolbar press
// commits an open sheet rename first (`rename-commit.ts`), and F6 and Shift+F6 move the keyboard
// among the editor's regions. Moved from the React adapter (`chrome.tsx`, `toolbar/parts.tsx`,
// and `landmarks.ts`), so the React and Vue chrome act the same. Each function reads a
// structural event, which a DOM event and a React synthetic event both satisfy. Adapter
// support, not a stable API.
//
// F6 and Shift+F6 visit, in this order: the menu bar (its tab stop), the toolbar, the formula
// bar (its name box first), the grid, and the footer (the sheet tabs and the zoom control). A region that the host's layout
// does not render is skipped. The grid takes the focus through the editor, so its own focus
// rules apply. F6 in the formula bar with an open cell edit commits the edit first, as F6 in the
// cell editor does; a refused commit keeps the keyboard in the bar.

import { blockNextClick, commitFocusedRename, renameFocused } from './rename-commit.ts';
import type { XlsxEditorInstance } from './xlsx-editor.ts';

/** The part of a mouse event that the press guards read. @internal */
export interface ChromePress {
  readonly target: EventTarget | null;
  readonly currentTarget: EventTarget | null;
  preventDefault(): void;
}

/** Whether `target` is a form field, which keeps its own press. */
function isField(target: EventTarget | null): boolean {
  const tag = (target as { tagName?: string } | null)?.tagName;
  return tag === 'INPUT' || tag === 'SELECT' || tag === 'TEXTAREA';
}

/** Keep the grid's focus when chrome is pressed, except in a form field. @internal */
export function keepGridFocus(e: Pick<ChromePress, 'target' | 'preventDefault'>): void {
  if (isField(e.target)) return;
  e.preventDefault();
}

/** A control that is disabled: native, `aria-disabled`, or `data-disabled`. */
const DISABLED = '[disabled], [aria-disabled="true"], [data-disabled]';

/**
 * Keep the grid's focus when the toolbar is pressed, except in a form field. An open sheet
 * rename commits first, as a blur does (`rename-commit.ts`). The click of the press runs no
 * command when the store refuses the name, so the reason stays in view, and when the pressed
 * control was disabled at the press: the commit can enable it (Undo), and its click must not
 * act. @internal
 */
export function guardToolbarMousedown(e: ChromePress): void {
  if (isField(e.target)) return;
  e.preventDefault();
  const target = e.target as Element;
  const from = e.currentTarget as Node;
  // Read the disabled state before the commit changes it.
  const pressed = target.closest?.(DISABLED);
  const disabled = !!pressed && from.contains(pressed);
  const renaming = renameFocused(from);
  if (!commitFocusedRename(from) || (renaming && disabled)) blockNextClick(from);
}

/** The part of a key event that the region keys read. @internal */
export interface RegionKey {
  readonly key: string;
  readonly altKey: boolean;
  readonly ctrlKey: boolean;
  readonly metaKey: boolean;
  readonly shiftKey: boolean;
  preventDefault(): void;
}

/** The regions that F6 visits, in order. */
const REGIONS = [
  '[role="menubar"]',
  '[role="toolbar"]',
  '.xlsx-formula-bar',
  '.xlsx-scroller',
  '.xlsx-footer',
] as const;

/** The control that takes the keyboard in a region: the selected tab, else the first control. */
function entryOf(region: Element): HTMLElement | null {
  return (
    region.querySelector<HTMLElement>('[role="tab"][aria-selected="true"]') ??
    // The menu bar's one tab stop.
    region.querySelector<HTMLElement>('[role="menuitem"][tabindex="0"]') ??
    region.querySelector<HTMLElement>('button, select, input, textarea, [tabindex="0"]')
  );
}

/** Handle F6 and Shift+F6 inside `root`. Returns whether the key moved the focus. @internal */
export function cycleRegion(
  e: RegionKey,
  root: HTMLElement,
  editor: XlsxEditorInstance | null
): boolean {
  if (e.key !== 'F6' || e.altKey || e.ctrlKey || e.metaKey) return false;
  const regions = REGIONS.map((s) => root.querySelector(s)).filter((r): r is Element => r !== null);
  if (regions.length === 0) return false;
  const active = root.ownerDocument.activeElement;
  const at = regions.findIndex((r) => active !== null && r.contains(active));
  const bar = regions[at]?.matches('.xlsx-formula-bar') === true;
  if (bar && editor?.snapshot().editing && !editor.commitEdit().ok) {
    e.preventDefault();
    return true;
  }
  const step = e.shiftKey ? -1 : 1;
  const from = at < 0 ? (e.shiftKey ? 0 : -1) : at;
  for (let i = 1; i <= regions.length; i += 1) {
    const region = regions[(from + step * i + regions.length * 2) % regions.length]!;
    if (region.matches('.xlsx-scroller')) {
      if (!editor) continue;
      editor.focus();
      e.preventDefault();
      return true;
    }
    const entry = entryOf(region);
    if (!entry) continue;
    entry.focus();
    e.preventDefault();
    return true;
  }
  return false;
}
