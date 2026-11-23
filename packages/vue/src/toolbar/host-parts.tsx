/** @jsxImportSource vue */
// The parts that a host passes to the toolbar: `XlsxEditorToolbar.Slot` and
// `XlsxEditorToolbar.Group`. The Vue twin of the React adapter's `toolbar/host-parts.tsx`, which
// is after docx-editor/packages/react/src/editor/toolbar/ToolbarGroup.tsx at 10eeaf67
// (EigenPal); the header there explains the rules and the differences from DOCX Editor. The
// child reader is after docx-editor/packages/vue/src/editor/toolbar/toolbar-children.ts at the
// same commit.
//
// The named parts (`XlsxEditorToolbar.Bold`) are read here as the `Toolbar.Slot` of the slot of
// their static (`xlsxSlot`), with the default slot of a button part as its icon, as in React.
//
// Differences from React: the parts' content is the default slot, which the toolbar calls in its
// render; `icon` is a prop that takes a vnode or a string, as DOCX Editor's Vue parts take it;
// `overflowContent` is a function prop that returns vnodes; the props of a part vnode are read camelized (`slot-id` or `slotId`), and a template's
// bare boolean attribute (`<Slot hidden />`) arrives as an empty string, which is `true`
// (`vnodeFlag`, as in DOCX Editor); a template `class` on a group joins its `className`; and a
// part's content can be a default slot, an array, a vnode, or a string, as a render function
// passes it. Differences from DOCX Editor: the marker is set after `defineComponent`
// (`markPart`), not with `Object.assign`, so the declared type of a part stays a
// `DefineComponent`, which TypeScript 5.7 (the floor) accepts, and the marker is no public
// member; and `visibleNodes` keys each child by its path, as React's `Children.toArray` does,
// where DOCX Editor's `flattenChildren` keeps the child keys, so two lists with the same keys
// collide there, and a child takes the element of a sibling before it that goes.

import {
  Comment,
  Fragment,
  camelize,
  cloneVNode,
  createTextVNode,
  defineComponent,
  h,
  inject,
  isVNode,
  provide,
  type ExtractPublicPropTypes,
  type InjectionKey,
  type PropType,
  type VNode,
  type VNodeChild,
} from 'vue';
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
  ToolbarButton,
  slotOverrideKey,
  type PartLayout,
  type SlotOverride,
} from './parts.tsx';
import { asChildElement } from './as-child.tsx';

/** Where a part renders: in the bar, or as a row of the "More" panel. */
export const partLayoutKey: InjectionKey<PartLayout> = Symbol('xlsx-toolbar-part-layout');

/**
 * One control of a group, as its shape renders it. `replaced` names the slots whose host content
 * takes their place, so the number format list and the dialog opener do not count on a slot that
 * shows host content.
 */
export function control(
  group: ChromeGroup,
  c: ChromeControl,
  layout: PartLayout,
  replaced?: ReadonlySet<string>,
  override?: SlotOverride | null
) {
  const own = builtInControl(group, c, layout, replaced);
  return override && own
    ? h(OverrideScope, { key: c.id, override }, { default: () => [own] })
    : own;
}

/** Gives the controls inside it the changes of a slot part (`slotOverrideKey`). */
const OverrideScope = defineComponent({
  name: 'XlsxToolbarSlotOverride',
  props: { override: { type: Object as PropType<SlotOverride>, required: true } },
  setup(props, { slots }) {
    provide(slotOverrideKey, () => props.override);
    return () => slots.default?.();
  },
});

/** The built-in control of a slot. */
function builtInControl(
  group: ChromeGroup,
  c: ChromeControl,
  layout: PartLayout,
  replaced: ReadonlySet<string> | undefined
) {
  const slot = chromeSlotId(group, c);
  if (slot === 'font.family') return <FontFamilyList key={c.id} layout={layout} />;
  if (slot === 'font.size') return <FontSizeList key={c.id} layout={layout} />;
  // The number format list holds the dialog's opener as its last item; alone, it is a button.
  const listed = (id: string) =>
    group.controls.some((x) => x.id === id) && !replaced?.has(`${group.id}.${id}`);
  if (slot === 'number.format')
    return <NumberFormatList key={c.id} layout={layout} more={listed('custom')} />;
  if (c.shape === 'dialog' && slot === 'number.custom')
    return listed('format') ? null : <NumberFormatDialogButton key={c.id} layout={layout} />;
  if (slot === 'font.color' || slot === 'fill.color')
    return <ColorPalette key={c.id} slot={slot} layout={layout} />;
  return <ToolbarButton key={c.id} slot={slot} layout={layout} />;
}

/** The registry group of a slot with only that slot, or null for an id the registry lacks. */
function groupOfSlot(slot: string): { group: ChromeGroup; control: ChromeControl } | null {
  for (const g of CHROME_GROUPS as readonly ChromeGroup[])
    for (const c of g.controls)
      if (chromeSlotId(g, c) === slot) return { group: { ...g, controls: [c] }, control: c };
  return null;
}

/** The key part of `child` at `index` of its list: its own key, else its place. */
const keyPart = (child: unknown, index: number): string =>
  isVNode(child) && child.key != null ? `k${String(child.key)}` : `#${index}`;

/**
 * The vnodes of `content`, with Fragments and arrays flattened, strings and numbers as text, and
 * comments and empty strings left out. Each vnode gets the key of its path, as React's
 * `Children.toArray` keys a child: the key or the place of each list on the way. So two lists
 * with the same keys stay apart, and a child keeps its element when a sibling before it goes.
 */
export function visibleNodes(content: unknown): VNode[] {
  const out: VNode[] = [];
  const walk = (node: unknown, key: string): void => {
    if (Array.isArray(node)) node.forEach((child, i) => walk(child, `${key}/${keyPart(child, i)}`));
    else if (typeof node === 'string' || typeof node === 'number') {
      if (node !== '') out.push(cloneVNode(createTextVNode(String(node)), { key }));
    } else if (isVNode(node)) {
      if (node.type === Fragment) walk(node.children, key);
      // `v-if` branches that render nothing leave comment nodes, which are no content.
      else if (node.type !== Comment) out.push(cloneVNode(node, { key }));
    }
  };
  walk(content, '');
  return out;
}

/**
 * Mark a part for the toolbar, which finds a part by this static, never by the component name.
 * The marker stays out of the declared type, so it is no public member.
 */
function markPart(part: object, marker: 'xlsxToolbarSlot' | 'xlsxToolbarGroup'): void {
  (part as Record<string, unknown>)[marker] = true;
}

/** Mark a named part with the slot it stands for. */
export function markNamedPart(part: object, slot: ChromeSlotId): void {
  markPart(part, 'xlsxToolbarSlot');
  (part as Record<string, unknown>).xlsxSlot = slot;
}

const xlsxEditorToolbarSlotProps = {
  /** The built-in slot that the part names, such as `'font.bold'`. */
  slotId: { type: String as PropType<ChromeSlotId>, required: true },
  /**
   * Removes the slot where it stands, as the toolbar's `hidden` list does. A group whose slots
   * are all hidden is left out.
   */
  hidden: { type: Boolean, default: undefined },
  /**
   * An icon in place of the registry icon, about 18 pixels, as the packaged icons: a vnode or a
   * string. Button slots only: a list, a palette, or the dialog opener ignores it, with a
   * development warning.
   */
  icon: { type: [Object, String, Array] as PropType<VNodeChild>, default: undefined },
  /**
   * Render the built-in control as the single element of the default slot: the element receives
   * the control's class, label, state attributes, and press. The element's own value wins for
   * each other prop, so its own `title` or `aria-describedby` hides the disabled reason. An
   * element that is not a `<button>` or a link also takes the button role, a Tab stop, and Enter
   * and Space. A link keeps its role. Button slots only.
   */
  asChild: { type: Boolean, default: undefined },
  /** Appended after the class of the control, or of the trigger of a value control. */
  className: { type: String, default: undefined },
  /**
   * Content for the slot's row in the "More" panel, when its group leaves the bar. Without it,
   * the row shows the default slot, or the built-in control as a row.
   */
  overflowContent: { type: Function as PropType<() => VNodeChild>, default: undefined },
} as const;

/** Props of `XlsxEditorToolbar.Slot`, the twin of React's `XlsxEditorToolbarSlotProps`. @public */
export type XlsxEditorToolbarSlotProps = ExtractPublicPropTypes<typeof xlsxEditorToolbarSlotProps>;

/**
 * One slot of the toolbar. Under the preset toolbar, the part replaces its slot where it stands
 * with its default slot, or removes it with `hidden`. Anywhere else, such as in a host group or
 * under `:preset="false"`, it renders its default slot, or the built-in control of its slot.
 * The toolbar reads a part only in its default slot or in a Fragment there, not in a component
 * of your own.
 *
 * @example
 * ```ts
 * h(XlsxEditorToolbar, null, () => [
 *   h(XlsxEditorToolbar.Slot, { slotId: 'file.save' }, () => [h(MySaveButton)]),
 *   h(XlsxEditorToolbar.Slot, { slotId: 'font.underline', hidden: true }),
 * ]);
 * ```
 *
 * @public
 */
export const XlsxEditorToolbarSlot = defineComponent({
  name: 'XlsxEditorToolbarSlot',
  // A template `class` joins `className` on the built-in control: the part has no element.
  inheritAttrs: false,
  props: xlsxEditorToolbarSlotProps,
  setup(props, { attrs, slots }) {
    const layout = inject(partLayoutKey, 'bar');
    return () => {
      if (props.hidden) return null;
      const children = slots.default?.();
      // `asChild` holds the element only on a button slot: elsewhere the default slot replaces it.
      const asChild = !!props.asChild && isToolbarButtonSlot(props.slotId);
      const content = asChild ? [] : visibleNodes(children);
      if (content.length > 0) return content;
      const found = groupOfSlot(props.slotId);
      if (!found) warnToolbarProblems([unknownSlotProblem(props.slotId, 'XlsxEditorToolbar.Slot')]);
      const hostClass = typeof attrs.class === 'string' ? attrs.class : undefined;
      const className = [props.className, hostClass].filter(Boolean).join(' ') || undefined;
      const override = found ? slotOverrideOf({ ...props, className }, children) : null;
      if (found && asChild && !override?.asChild)
        warnToolbarProblems([toolbarAsChildProblem('XlsxEditorToolbar.Slot', props.slotId)]);
      return found ? control(found.group, found.control, layout, undefined, override) : null;
    };
  },
});

markPart(XlsxEditorToolbarSlot, 'xlsxToolbarSlot');

const xlsxEditorToolbarGroupProps = {
  /**
   * The group id. A built-in id, such as `'font'` or `'file'`, adds the default slot to that
   * group, after its own controls. Any other id makes a host group, which the bar measures and
   * moves into the "More" panel as it moves its own groups.
   */
  id: { type: String as PropType<ChromeGroupId | (string & {})>, required: true },
  /**
   * The heading of a host group in the "More" panel, and its accessible name. It wins over
   * `labelKey`. Without either, the heading is the `id`. A built-in group keeps its own label.
   */
  label: { type: String, default: undefined },
  /** A catalog key for the heading of a host group, translated with the toolbar's locale. */
  labelKey: { type: String, default: undefined },
  /**
   * The group that a host group follows in the bar. Without it, a host group follows every
   * built-in group, in order of appearance. A built-in group keeps its place.
   */
  after: { type: String as PropType<ChromeGroupId | (string & {})>, default: undefined },
  /**
   * Collapse priority: a lower number leaves the bar for the "More" panel first. The built-in
   * groups use 10 (`file`) through 90 (`history`) in steps of 10, from the end of the bar. A
   * host group without a priority leaves before every built-in group. A built-in id takes a
   * new priority.
   */
  priority: { type: Number, default: undefined },
  /** `true` keeps the group in the bar at every width. */
  pinned: { type: Boolean, default: undefined },
  /** Removes the group, built-in or host, from the bar and from the "More" panel. */
  hidden: { type: Boolean, default: undefined },
  /**
   * Content for the group's section of the "More" panel: after the controls of a built-in
   * group, where it shows also when the group has no default slot, or in place of the default
   * slot of a host group. Without it, the panel shows the default slot as rows.
   */
  overflowContent: { type: Function as PropType<() => VNodeChild>, default: undefined },
  /** Appended after the `xlsx-toolbar-group` class. A host `class` also works. */
  className: { type: String, default: undefined },
} as const;

/** Props of `XlsxEditorToolbar.Group`, the twin of React's `XlsxEditorToolbarGroupProps`. @public */
export type XlsxEditorToolbarGroupProps = ExtractPublicPropTypes<
  typeof xlsxEditorToolbarGroupProps
>;

/**
 * A group of the toolbar. Under the preset toolbar, a host group is measured with the built-in
 * groups, moves into the "More" panel when it does not fit, and shows there as a section with
 * its heading. A built-in id adds the default slot to that group, or removes the group with
 * `hidden`. Under `:preset="false"`, the part renders its default slot as one group. The
 * toolbar reads a part only in its default slot or in a Fragment there, not in a component of
 * your own.
 *
 * @example
 * ```ts
 * h(XlsxEditorToolbar, null, () => [
 *   h(XlsxEditorToolbar.Group, { id: 'share', label: 'Share', after: 'file' }, () => [
 *     h(ShareButton),
 *   ]),
 *   h(XlsxEditorToolbar.Group, { id: 'borders', hidden: true }),
 * ]);
 * ```
 *
 * @public
 */
export const XlsxEditorToolbarGroup = defineComponent({
  name: 'XlsxEditorToolbarGroup',
  props: xlsxEditorToolbarGroupProps,
  setup(props, { slots }) {
    const { t } = useTranslation();
    return () => {
      const content = props.hidden ? [] : visibleNodes(slots.default?.());
      // A group with no content is left out, as in the preset bar.
      return content.length === 0 ? null : (
        <div
          role="group"
          aria-label={toolbarGroupHeading(props, (key) => translateKey(t, key))}
          class={['xlsx-toolbar-group', props.className]}
          data-host-group={props.id}
        >
          {content}
        </div>
      );
    };
  },
});

markPart(XlsxEditorToolbarGroup, 'xlsxToolbarGroup');

/**
 * A boolean prop as the vnode carries it. A template's bare attribute (`<Slot hidden />`)
 * arrives as an empty string, which is `true`.
 */
export function vnodeFlag(value: unknown): boolean {
  return value === true || value === '';
}

const isPart = (node: VNode, marker: 'xlsxToolbarSlot' | 'xlsxToolbarGroup'): boolean =>
  typeof node.type === 'object' &&
  node.type !== null &&
  (node.type as unknown as Record<string, unknown>)[marker] === true;

/**
 * The props of a part vnode, camelized, and the vnodes of its content: the default slot of a
 * template or a slots object, or the array, vnode, or string that a render function passes.
 * `raw` is the content as the part's default slot returns it.
 */
function partProps(node: VNode): {
  props: Record<string, unknown>;
  content: VNode[];
  raw: unknown;
} {
  const props: Record<string, unknown> = Object.create(null);
  for (const [key, value] of Object.entries(node.props ?? {})) props[camelize(key)] = value;
  const children: unknown = node.children;
  const slot =
    children !== null && typeof children === 'object' && !Array.isArray(children)
      ? (children as { default?: unknown }).default
      : undefined;
  const raw =
    typeof slot === 'function'
      ? slot()
      : Array.isArray(children) || typeof children === 'string'
        ? children
        : undefined;
  return { props, content: visibleNodes(raw), raw };
}

const text = (value: unknown): string | undefined =>
  typeof value === 'string' ? value : undefined;

const fn = (value: unknown): (() => VNodeChild) | undefined =>
  typeof value === 'function' ? (value as () => VNodeChild) : undefined;

/** A priority prop as the vnode carries it: a number, or the text of a template attribute. */
function priorityOf(value: unknown): number | undefined {
  const n = typeof value === 'string' && value.trim() !== '' ? Number(value) : value;
  return typeof n === 'number' && !Number.isNaN(n) ? n : undefined;
}

/**
 * The slot props that a named part stands for, the twin of React's `namedPartSlotProps`: the
 * default slot of a button part is icon content, and that of a value part is dropped. `children`
 * carries the default slot's content for `asChild`.
 */
export function namedPartSlotProps(
  slot: ChromeSlotId,
  props: Record<string, unknown>,
  content: unknown
): Record<string, unknown> {
  // A value part keeps its props as given, so a stray `icon` or `asChild` still warns.
  if (!isToolbarButtonSlot(slot)) return { ...props, slotId: slot };
  if (vnodeFlag(props.asChild)) return { ...props, slotId: slot, children: content };
  const glyph = props.icon ?? (visibleNodes(content).length > 0 ? content : undefined);
  return { ...props, slotId: slot, ...(glyph !== undefined ? { icon: glyph } : {}) };
}

/** A part's `Toolbar.Group` as the toolbar renders it. */
export interface ToolbarGroupPart {
  readonly className: string | undefined;
  readonly content: readonly VNode[];
  readonly overflowContent: (() => VNodeChild) | undefined;
}

/** A part's `Toolbar.Slot`, or a named part, as the toolbar renders it. */
export interface ToolbarSlotPart {
  /** The content that takes the slot's place. */
  readonly content: readonly VNode[];
  /** The changes to the built-in control. */
  readonly override: SlotOverride | null;
  readonly overflowContent: (() => VNodeChild) | undefined;
}

/**
 * The changes of a slot part to its built-in control, or null for none. `icon` and `asChild`
 * reach a button slot only, as `arrangeToolbar` warns. `children` is the default slot's
 * content, which holds the element of `asChild`.
 */
export function slotOverrideOf(
  props: {
    readonly slotId: string;
    readonly icon?: VNodeChild;
    readonly asChild?: boolean | undefined;
    readonly className?: string | undefined;
  },
  children: unknown
): SlotOverride | null {
  const button = isToolbarButtonSlot(props.slotId);
  const asChild = button && props.asChild ? asChildElement(children) : null;
  const icon = button ? props.icon : undefined;
  if (icon == null && !asChild && !props.className) return null;
  return {
    slot: props.slotId,
    ...(icon != null ? { icon } : {}),
    ...(props.className ? { className: props.className } : {}),
    ...(asChild ? { asChild } : {}),
  };
}

/** The children of the preset toolbar, by role. */
export interface ToolbarChildren {
  /** Each `Toolbar.Slot` or named part. The last part for a slot wins. */
  readonly slots: ReadonlyMap<string, ToolbarSlotPart>;
  /** The `Toolbar.Group` parts. The last part for an id wins. */
  readonly groups: ReadonlyMap<string, ToolbarGroupPart>;
  /** Everything else, rendered in one fixed block after the groups. */
  readonly appended: readonly VNode[];
  /** The requests of the parts, for `arrangeToolbar`. */
  readonly slotRequests: readonly ToolbarSlotRequest[];
  readonly groupRequests: readonly ToolbarGroupRequest[];
}

/** Read the children of the preset toolbar into its parts. Call it in a render. */
export function readToolbarChildren(children: unknown): ToolbarChildren {
  const slots = new Map<string, ToolbarSlotPart>();
  const groups = new Map<string, ToolbarGroupPart>();
  const slotRequests: ToolbarSlotRequest[] = [];
  const groupRequests: ToolbarGroupRequest[] = [];
  const appended: VNode[] = [];
  for (const node of visibleNodes(children)) {
    if (isPart(node, 'xlsxToolbarGroup')) {
      const { props, content } = partProps(node);
      const id = String(props.id);
      const className = [text(props.class), text(props.className)].filter(Boolean).join(' ');
      const overflowContent = fn(props.overflowContent);
      groups.set(id, { className: className || undefined, content, overflowContent });
      groupRequests.push({
        id,
        after: text(props.after),
        hidden: vnodeFlag(props.hidden),
        label: text(props.label),
        labelKey: text(props.labelKey),
        priority: priorityOf(props.priority),
        pinned: vnodeFlag(props.pinned),
        content: content.length > 0,
        overflowContent: overflowContent !== undefined,
      });
    } else if (isPart(node, 'xlsxToolbarSlot')) {
      const read = partProps(node);
      const named = (node.type as { xlsxSlot?: ChromeSlotId }).xlsxSlot;
      const props = named ? namedPartSlotProps(named, read.props, read.raw) : read.props;
      const slotId = String(named ?? props.slotId);
      const asChild = vnodeFlag(props.asChild);
      // `asChild` holds the element only on a button slot: elsewhere the default slot replaces it.
      const holds = asChild && isToolbarButtonSlot(slotId);
      const raw = named ? props.children : read.raw;
      const content = holds || named ? [] : read.content;
      const className = [text(props.class), text(props.className)].filter(Boolean).join(' ');
      const override = slotOverrideOf(
        {
          slotId,
          icon: props.icon as VNodeChild,
          asChild,
          className: className || undefined,
        },
        raw
      );
      const overflowContent = fn(props.overflowContent);
      slots.set(slotId, { content, override, overflowContent });
      slotRequests.push({
        slotId,
        hidden: vnodeFlag(props.hidden),
        content: content.length > 0,
        icon: props.icon != null,
        asChild,
        asChildElement: holds ? asChildElement(raw) !== null : undefined,
        className: !!className,
        overflowContent: overflowContent !== undefined,
      });
    } else appended.push(node);
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
