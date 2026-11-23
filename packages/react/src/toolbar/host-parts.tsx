/** @jsxImportSource react */
// The parts that a host passes to the toolbar: `XlsxEditor.Toolbar.Slot` and
// `XlsxEditor.Toolbar.Group`, after docx-editor/packages/react/src/editor/toolbar/ToolbarGroup.tsx,
// toolbar-children.ts, and toolbar-warnings.ts at 10eeaf67 (EigenPal). Under the preset, both
// parts are descriptions: the toolbar reads their props and renders the content itself
// (`arrangeToolbar` in core). The toolbar matches a part by a static marker, never by its display
// name, which minifies away.
//
// The named parts (`XlsxEditor.Toolbar.Bold`, ...) are `Toolbar.Slot` with the slot pinned, as
// DOCX Editor's `definePart` (parts.tsx at ad686574): each carries its slot as a static
// (`xlsxSlot`), and the children of a button part are icon content (`icon`, else children,
// else the registry icon), not a replacement. `icon`, `asChild`, `className`, and
// `overflowContent` of a slot, and `priority`, `pinned`, and `overflowContent` of a group, are
// DOCX Editor's (ToolbarGroup.tsx at ad686574).
//
// Differences from DOCX Editor: a `Toolbar.Slot` with no children renders the built-in control
// of its slot, as `ToolbarSlot` of pptx-editor/packages/react/src/Toolbar.tsx at 929ca860
// (EigenPal) does, so a host can place a built-in control under `preset={false}` or in a host
// group; `icon`, `asChild`, and `className` also apply on `Toolbar.Slot`, as in PPTX Editor's
// `SlotControl` (toolbar-parts.tsx at 70052cf1), and `asChild` with children that are not one
// element falls back to the built-in button with a warning; a value control (a list, a
// palette, the dialog opener) takes no `icon` or `asChild`, with a warning; a Fragment of any
// number of parts is read as its parts, as the Vue adapter reads one (DOCX Editor's
// `unwrapFragment` reads a Fragment of one part only); and the problems come from core as
// sentences, which `warnToolbarProblems` prints once each in development, with the
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
  isToolbarButtonSlot,
  toolbarAsChildProblem,
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
  SlotOverrideContext,
  ToolbarButton,
  type PartLayout,
  type SlotOverride,
} from './parts.tsx';
import { asChildElement } from './as-child.tsx';

/**
 * Mark a part for the toolbar, which finds a part by this static, never by its display name,
 * which minifies away. The marker stays out of the declared type, so it is no public member.
 */
function markPart(part: object, marker: 'xlsxToolbarSlot' | 'xlsxToolbarGroup'): void {
  (part as Record<string, unknown>)[marker] = true;
}

/** Mark a named part with the slot it stands for. */
export function markNamedPart(part: object, slot: ChromeSlotId): void {
  markPart(part, 'xlsxToolbarSlot');
  (part as Record<string, unknown>).xlsxSlot = slot;
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
   * An icon in place of the registry icon, about 18 pixels, as the packaged icons. Button slots
   * only: a list, a palette, or the dialog opener ignores it, with a development warning.
   */
  readonly icon?: ReactNode;
  /**
   * Render the built-in control as the single child element: the child receives the control's
   * class name, label, state attributes, and press. The child's own value wins for each other
   * prop, so its own `title` or `aria-describedby` hides the disabled reason. An element that is
   * not a `<button>` also takes the button role, a Tab stop, and Enter and Space. Button slots
   * only.
   */
  readonly asChild?: boolean;
  /** Appended after the class of the control, or of the trigger of a value control. */
  readonly className?: string;
  /**
   * Content for the slot's row in the "More" panel, when its group leaves the bar. Without it,
   * the row shows the children, or the built-in control as a row.
   */
  readonly overflowContent?: () => ReactNode;
  /**
   * Content that takes the slot's place: in the bar, and in the group's section of the "More"
   * panel. Without children, the part renders the built-in control of the slot. With `asChild`,
   * the one element that renders the control.
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
  /**
   * Collapse priority: a lower number leaves the bar for the "More" panel first. The built-in
   * groups use 10 (`file`) through 90 (`history`) in steps of 10, from the end of the bar. A
   * host group without a priority leaves before every built-in group. A built-in id takes a
   * new priority.
   */
  readonly priority?: number;
  /** `true` keeps the group in the bar at every width. */
  readonly pinned?: boolean;
  /** Removes the group, built-in or host, from the bar and from the "More" panel. */
  readonly hidden?: boolean;
  /**
   * Content for the group's section of the "More" panel: after the controls of a built-in
   * group, where it shows also when the group has no children, or in place of the children of a
   * host group. Without it, the panel shows the children as rows.
   */
  readonly overflowContent?: () => ReactNode;
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
  override,
}: {
  readonly group: ChromeGroup;
  readonly control: ChromeControl;
  readonly layout: PartLayout;
  readonly replaced?: ReadonlySet<string>;
  /** The changes of the slot's part: `icon`, `asChild`, `className`. */
  readonly override?: SlotOverride | null;
}) {
  const own = (
    <BuiltInControl group={group} control={control} layout={layout} replaced={replaced} />
  );
  return override ? (
    <SlotOverrideContext.Provider value={override}>{own}</SlotOverrideContext.Provider>
  ) : (
    own
  );
}

/** The built-in control of a slot. */
function BuiltInControl({
  group,
  control,
  layout,
  replaced,
}: {
  readonly group: ChromeGroup;
  readonly control: ChromeControl;
  readonly layout: PartLayout;
  readonly replaced?: ReadonlySet<string> | undefined;
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
export function XlsxEditorToolbarSlot(props: XlsxEditorToolbarSlotProps) {
  const { slotId, hidden, children } = props;
  const layout = useContext(PartLayoutContext);
  // `asChild` holds the element only on a button slot: elsewhere the children replace the slot.
  const asChild = !!props.asChild && isToolbarButtonSlot(slotId);
  const content = !asChild && hasContent(children);
  const found = hidden || content ? null : groupOfSlot(slotId);
  const unknown = !hidden && !content && !found;
  const override = found ? slotOverrideOf(props) : null;
  const badChild = asChild && !hidden && !override?.asChild;
  useEffect(() => {
    if (unknown) warnToolbarProblems([unknownSlotProblem(slotId, 'XlsxEditor.Toolbar.Slot')]);
    else if (badChild)
      warnToolbarProblems([toolbarAsChildProblem('XlsxEditor.Toolbar.Slot', slotId)]);
  }, [unknown, badChild, slotId]);
  if (hidden) return null;
  if (content) return <>{children}</>;
  return found ? (
    <Control group={found.group} control={found.control} layout={layout} override={override} />
  ) : null;
}

/**
 * The changes of a slot part to its built-in control, or null for none. `icon` and `asChild`
 * reach a button slot only, as `arrangeToolbar` warns.
 */
export function slotOverrideOf(props: XlsxEditorToolbarSlotProps): SlotOverride | null {
  const button = isToolbarButtonSlot(props.slotId);
  const asChild = button && props.asChild ? asChildElement(props.children) : null;
  const icon = button ? props.icon : undefined;
  if (icon == null && !asChild && !props.className) return null;
  return {
    slot: props.slotId,
    ...(icon != null ? { icon } : {}),
    ...(props.className ? { className: props.className } : {}),
    ...(asChild ? { asChild } : {}),
  };
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

/**
 * The slot props of a `Toolbar.Slot` element, or of a named part, which stands for the slot of
 * its static (`namedPartSlotProps`).
 */
function slotPropsOf(part: ReactElement): XlsxEditorToolbarSlotProps {
  const slot = (part.type as unknown as { xlsxSlot?: ChromeSlotId }).xlsxSlot;
  const props = part.props as XlsxEditorToolbarSlotProps;
  return slot ? namedPartSlotProps(slot, props as unknown as Record<string, unknown>) : props;
}

/**
 * The slot props that a named part stands for, as PPTX Editor's `namedPartSlotProps`
 * (toolbar-parts.tsx at 70052cf1): the children of a button part are icon content, as in DOCX
 * Editor, and the children of a value part are dropped (its props type has none).
 */
export function namedPartSlotProps(
  slot: ChromeSlotId,
  props: Record<string, unknown>
): XlsxEditorToolbarSlotProps {
  const { children, icon, ...rest } = props as { children?: ReactNode; icon?: ReactNode };
  // A value part keeps its other props as given, so a stray `icon` or `asChild` still warns.
  if (!isToolbarButtonSlot(slot)) return { ...rest, icon, slotId: slot };
  if (props.asChild === true) return { ...rest, slotId: slot, icon, children };
  const glyph = icon ?? (hasContent(children) ? children : undefined);
  return { ...rest, slotId: slot, ...(glyph !== undefined ? { icon: glyph } : {}) };
}

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
        priority: typeof props.priority === 'number' ? props.priority : undefined,
        pinned: props.pinned === true,
        content: hasContent(props.children),
        overflowContent: typeof props.overflowContent === 'function',
      });
    } else if (part) {
      const props = slotPropsOf(part);
      const slotId = String(props.slotId);
      slots.set(slotId, props);
      const asChild = props.asChild === true;
      // `asChild` holds the element only on a button slot: elsewhere the children replace it.
      const element =
        asChild && isToolbarButtonSlot(slotId) ? asChildElement(props.children) : null;
      slotRequests.push({
        slotId,
        hidden: props.hidden === true,
        content: !(asChild && isToolbarButtonSlot(slotId)) && hasContent(props.children),
        icon: props.icon != null,
        asChild,
        asChildElement: element !== null,
        className: !!props.className,
        overflowContent: typeof props.overflowContent === 'function',
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
