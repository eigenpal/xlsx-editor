/** @jsxImportSource react */
// The parts that a host passes to the toolbar: `XlsxEditor.Toolbar.Slot` and
// `XlsxEditor.Toolbar.Group`, after docx-editor/packages/react/src/editor/toolbar/ToolbarGroup.tsx,
// toolbar-children.ts, and toolbar-warnings.ts at 10eeaf67 (EigenPal). Under the preset, both
// parts are descriptions: the toolbar reads their props and renders the content itself
// (`arrangeToolbar` in core). The toolbar matches a part by a static marker, never by its display
// name, which minifies away.
//
// Differences from DOCX Editor: this toolbar has no named parts (`Toolbar.Bold`), so
// `Toolbar.Slot` is the part that replaces or hides a slot; a `Toolbar.Slot` with no children
// renders the built-in control of its slot, as `ToolbarSlot` of
// pptx-editor/packages/react/src/Toolbar.tsx at 929ca860 (EigenPal) does, so a host can place a
// built-in control under `preset={false}` or in a host group; a group has no `priority`,
// `pinned`, or `overflowContent`, and a slot has no `overflowContent` (`toolbar-arrange.ts` names
// why); a Fragment of any number of parts is read as its parts, as the Vue adapter reads one
// (DOCX Editor's `unwrapFragment` reads a Fragment of one part only); and the problems come from
// core as sentences, which `warnToolbarProblems` prints once each in development, with the
// `[xlsx-editor]` prefix. The markers are set with `markPart`, so they are no public member of
// the exported parts.

import {
  Children,
  Fragment,
  cloneElement,
  createContext,
  isValidElement,
  useContext,
  useEffect,
  type ReactElement,
  type ReactNode,
} from 'react';
import {
  CHROME_GROUPS,
  chromeSlotId,
  toolbarGroupHeading,
  type ChromeControl,
  type ChromeGroup,
  type ChromeGroupId,
  type ChromeSlotId,
  type ToolbarGroupRequest,
  type ToolbarSlotRequest,
} from '@xlsx-editor.dev/core/editor';
import { useTranslation } from '../context.ts';
import { isDevelopment } from '../is-development.ts';
import { translateKey } from './slot.tsx';
import {
  ColorPalette,
  FontFamilyList,
  FontSizeList,
  NumberFormatDialogButton,
  NumberFormatList,
  ToolbarButton,
  type PartLayout,
} from './parts.tsx';

/**
 * Mark a part for the toolbar, which finds a part by this static, never by its display name,
 * which minifies away. The marker stays out of the declared type, so it is no public member.
 */
function markPart(part: object, marker: 'xlsxToolbarSlot' | 'xlsxToolbarGroup'): void {
  (part as Record<string, unknown>)[marker] = true;
}

/** Props for `XlsxEditor.Toolbar.Slot`. @public */
export interface XlsxEditorToolbarSlotProps {
  /** The built-in slot that the part names, such as `'font.bold'`. */
  readonly slotId: ChromeSlotId;
  /**
   * Removes the slot where it stands, as the toolbar's `hidden` list does. A group whose slots
   * are all hidden is left out.
   */
  readonly hidden?: boolean;
  /**
   * Content that takes the slot's place: in the bar, and in the group's section of the "More"
   * panel. Without children, the part renders the built-in control of the slot.
   */
  readonly children?: ReactNode;
}

/** Props for `XlsxEditor.Toolbar.Group`. @public */
export interface XlsxEditorToolbarGroupProps {
  /**
   * The group id. A built-in id, such as `'font'` or `'file'`, adds the children to that group,
   * after its own controls. Any other id makes a host group, which the bar measures and moves
   * into the "More" panel as it moves its own groups.
   */
  readonly id: ChromeGroupId | (string & {});
  /**
   * The heading of a host group in the "More" panel, and its accessible name. It wins over
   * `labelKey`. Without either, the heading is the `id`. A built-in group keeps its own label.
   */
  readonly label?: string;
  /** A catalog key for the heading of a host group, translated with the toolbar's locale. */
  readonly labelKey?: string;
  /**
   * The group that a host group follows in the bar. Without it, a host group follows every
   * built-in group, in order of appearance. A built-in group keeps its place.
   */
  readonly after?: ChromeGroupId | (string & {});
  /** Removes the group, built-in or host, from the bar and from the "More" panel. */
  readonly hidden?: boolean;
  /** Appended after the `xlsx-toolbar-group` class. */
  readonly className?: string;
  /**
   * The controls of a host group, or the controls that a built-in group shows after its own.
   * A host group with no children is left out.
   */
  readonly children?: ReactNode;
}

/** Where a part renders: in the bar, or as a row of the "More" panel. */
export const PartLayoutContext = createContext<PartLayout>('bar');

/**
 * One control of a group, as its shape renders it. `replaced` names the slots whose host content
 * takes their place, so the number format list and the dialog opener do not count on a slot that
 * shows host content.
 */
export function Control({
  group,
  control,
  layout,
  replaced,
}: {
  readonly group: ChromeGroup;
  readonly control: ChromeControl;
  readonly layout: PartLayout;
  readonly replaced?: ReadonlySet<string>;
}) {
  const slot = chromeSlotId(group, control);
  if (slot === 'font.family') return <FontFamilyList layout={layout} />;
  if (slot === 'font.size') return <FontSizeList layout={layout} />;
  // The number format list holds the dialog's opener as its last item; alone, it is a button.
  const listed = (id: string) =>
    group.controls.some((c) => c.id === id) && !replaced?.has(`${group.id}.${id}`);
  if (slot === 'number.format') return <NumberFormatList layout={layout} more={listed('custom')} />;
  if (control.shape === 'dialog' && slot === 'number.custom')
    return listed('format') ? null : <NumberFormatDialogButton layout={layout} />;
  if (slot === 'font.color' || slot === 'fill.color')
    return <ColorPalette slot={slot} layout={layout} />;
  return <ToolbarButton slot={slot} layout={layout} />;
}

/** The registry group of a slot with only that slot, or null for an id the registry lacks. */
function groupOfSlot(slot: string): { group: ChromeGroup; control: ChromeControl } | null {
  for (const g of CHROME_GROUPS as readonly ChromeGroup[])
    for (const c of g.controls)
      if (chromeSlotId(g, c) === slot) return { group: { ...g, controls: [c] }, control: c };
  return null;
}

/**
 * One slot of the toolbar. Under the preset toolbar, the part replaces its slot where it
 * stands, or removes it with `hidden`. Anywhere else, such as in a host group or under
 * `preset={false}`, it renders its children, or the built-in control of its slot. The toolbar
 * reads a part only as its direct child or in a Fragment, not in a component of your own.
 *
 * @example
 * ```tsx
 * <XlsxEditor.Toolbar>
 *   <XlsxEditor.Toolbar.Slot slotId="file.save">
 *     <MySaveButton />
 *   </XlsxEditor.Toolbar.Slot>
 *   <XlsxEditor.Toolbar.Slot slotId="font.underline" hidden />
 * </XlsxEditor.Toolbar>
 * ```
 *
 * @public
 */
export function XlsxEditorToolbarSlot({ slotId, hidden, children }: XlsxEditorToolbarSlotProps) {
  const layout = useContext(PartLayoutContext);
  const content = hasContent(children);
  const found = hidden || content ? null : groupOfSlot(slotId);
  const unknown = !hidden && !content && !found;
  useEffect(() => {
    if (unknown) warnToolbarProblems([unknownSlotProblem(slotId, 'XlsxEditor.Toolbar.Slot')]);
  }, [unknown, slotId]);
  if (hidden) return null;
  if (content) return <>{children}</>;
  return found ? <Control group={found.group} control={found.control} layout={layout} /> : null;
}

markPart(XlsxEditorToolbarSlot, 'xlsxToolbarSlot');

/**
 * A group of the toolbar. Under the preset toolbar, a host group is measured with the built-in
 * groups, moves into the "More" panel when it does not fit, and shows there as a section with
 * its heading. A built-in id adds the children to that group, or removes the group with
 * `hidden`. Under `preset={false}`, the part renders its children as one group. The toolbar
 * reads a part only as its direct child or in a Fragment, not in a component of your own.
 *
 * @example
 * ```tsx
 * <XlsxEditor.Toolbar>
 *   <XlsxEditor.Toolbar.Group id="share" label="Share" after="file">
 *     <ShareButton />
 *   </XlsxEditor.Toolbar.Group>
 *   <XlsxEditor.Toolbar.Group id="borders" hidden />
 * </XlsxEditor.Toolbar>
 * ```
 *
 * @public
 */
export function XlsxEditorToolbarGroup(props: XlsxEditorToolbarGroupProps) {
  const { t } = useTranslation();
  // A group with no content is left out, as in the preset bar.
  if (props.hidden || !hasContent(props.children)) return null;
  return (
    <div
      role="group"
      aria-label={toolbarGroupHeading(props, (key) => translateKey(t, key))}
      className={`xlsx-toolbar-group${props.className ? ` ${props.className}` : ''}`}
      data-host-group={props.id}
    >
      {props.children}
    </div>
  );
}

markPart(XlsxEditorToolbarGroup, 'xlsxToolbarGroup');

/**
 * The children as a flat list, with every Fragment read as its children. A child of a Fragment
 * takes the Fragment's key before its own, so two children never share a key, and each keeps
 * its state when a sibling comes or goes.
 */
function flatChildren(children: ReactNode, prefix = ''): ReactNode[] {
  return Children.toArray(children).flatMap((child) => {
    if (!isValidElement(child)) return [child];
    if (child.type === Fragment)
      return flatChildren(
        (child.props as { children?: ReactNode }).children,
        `${prefix}${String(child.key)}/`
      );
    return [prefix ? cloneElement(child, { key: `${prefix}${String(child.key)}` }) : child];
  });
}

/**
 * Whether `children` renders anything: `false`, `null`, `undefined`, an empty string, and an
 * empty Fragment are no content.
 */
export function hasContent(children: ReactNode): boolean {
  return flatChildren(children).some((child) => child !== '');
}

const isPart = (child: ReactNode, marker: 'xlsxToolbarSlot' | 'xlsxToolbarGroup') =>
  isValidElement(child) &&
  (typeof child.type === 'function' || typeof child.type === 'object') &&
  (child.type as unknown as Record<string, unknown>)[marker] === true;

/** A child as a part, or null for any other child. */
function partOf(child: ReactNode): ReactElement | null {
  return isPart(child, 'xlsxToolbarSlot') || isPart(child, 'xlsxToolbarGroup')
    ? (child as ReactElement)
    : null;
}

/** The children of the preset toolbar, by role. */
export interface ToolbarChildren {
  /** The `Toolbar.Slot` parts. The last part for a slot wins. */
  readonly slots: ReadonlyMap<string, XlsxEditorToolbarSlotProps>;
  /** The `Toolbar.Group` parts. The last part for an id wins. */
  readonly groups: ReadonlyMap<string, XlsxEditorToolbarGroupProps>;
  /** Everything else, rendered in one fixed block after the groups. */
  readonly appended: readonly ReactNode[];
  /** The requests of the parts, for `arrangeToolbar`. */
  readonly slotRequests: readonly ToolbarSlotRequest[];
  readonly groupRequests: readonly ToolbarGroupRequest[];
}

/**
 * Read the children of the preset toolbar into its parts. `Children.toArray` does not flatten a
 * Fragment element, so the reader does: a host that maps its parts into a Fragment still gets
 * its overrides, as in DOCX Editor's `unwrapFragment` (merge-arrangement.tsx at 10eeaf67).
 */
export function readToolbarChildren(children: ReactNode): ToolbarChildren {
  const slots = new Map<string, XlsxEditorToolbarSlotProps>();
  const groups = new Map<string, XlsxEditorToolbarGroupProps>();
  const slotRequests: ToolbarSlotRequest[] = [];
  const groupRequests: ToolbarGroupRequest[] = [];
  const appended: ReactNode[] = [];
  for (const child of flatChildren(children)) {
    const part = partOf(child);
    if (part && isPart(part, 'xlsxToolbarGroup')) {
      const props = part.props as XlsxEditorToolbarGroupProps;
      const id = String(props.id);
      groups.set(id, props);
      groupRequests.push({
        id,
        after: props.after,
        hidden: props.hidden === true,
        label: props.label,
        labelKey: props.labelKey,
        content: hasContent(props.children),
      });
    } else if (part) {
      const props = part.props as XlsxEditorToolbarSlotProps;
      const slotId = String(props.slotId);
      slots.set(slotId, props);
      slotRequests.push({
        slotId,
        hidden: props.hidden === true,
        content: hasContent(props.children),
      });
    } else appended.push(child);
  }
  return { slots, groups, appended, slotRequests, groupRequests };
}

/** The problem of a `Toolbar.Slot` outside the preset whose id names no slot. */
export function unknownSlotProblem(slotId: string, part: string): string {
  return (
    `"${slotId}" is not a toolbar slot id, so its ${part} renders nothing. Use a ChromeSlotId, ` +
    `such as "font.bold".`
  );
}

// Each problem prints once per page. The set lives as long as the module, so each test gives
// its problems sentences of their own.
const printed = new Set<string>();

/** Print each problem once per page, in development builds only. */
export function warnToolbarProblems(problems: readonly string[]): void {
  if (problems.length === 0 || !isDevelopment()) return;
  for (const problem of problems) {
    if (printed.has(problem)) continue;
    printed.add(problem);
    console.warn(`[xlsx-editor] ${problem}`);
  }
}
