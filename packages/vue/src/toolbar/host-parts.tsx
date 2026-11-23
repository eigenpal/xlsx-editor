/** @jsxImportSource vue */
// The parts that a host passes to the toolbar: `XlsxEditorToolbar.Slot` and
// `XlsxEditorToolbar.Group`. The Vue twin of the React adapter's `toolbar/host-parts.tsx`, which
// is after docx-editor/packages/react/src/editor/toolbar/ToolbarGroup.tsx at 10eeaf67
// (EigenPal); the header there explains the rules and the differences from DOCX Editor. The
// child reader is after docx-editor/packages/vue/src/editor/toolbar/toolbar-children.ts at the
// same commit.
//
// Differences from React: the parts' content is the default slot, which the toolbar calls in its
// render; the props of a part vnode are read camelized (`slot-id` or `slotId`), and a template's
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
  inject,
  isVNode,
  type ExtractPublicPropTypes,
  type InjectionKey,
  type PropType,
  type VNode,
} from 'vue';
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
  replaced?: ReadonlySet<string>
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

const xlsxEditorToolbarSlotProps = {
  /** The built-in slot that the part names, such as `'font.bold'`. */
  slotId: { type: String as PropType<ChromeSlotId>, required: true },
  /**
   * Removes the slot where it stands, as the toolbar's `hidden` list does. A group whose slots
   * are all hidden is left out.
   */
  hidden: { type: Boolean, default: undefined },
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
  props: xlsxEditorToolbarSlotProps,
  setup(props, { slots }) {
    const layout = inject(partLayoutKey, 'bar');
    return () => {
      if (props.hidden) return null;
      const content = visibleNodes(slots.default?.());
      if (content.length > 0) return content;
      const found = groupOfSlot(props.slotId);
      if (!found) warnToolbarProblems([unknownSlotProblem(props.slotId, 'XlsxEditorToolbar.Slot')]);
      return found ? control(found.group, found.control, layout) : null;
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
  /** Removes the group, built-in or host, from the bar and from the "More" panel. */
  hidden: { type: Boolean, default: undefined },
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
 */
function partProps(node: VNode): { props: Record<string, unknown>; content: VNode[] } {
  const props: Record<string, unknown> = Object.create(null);
  for (const [key, value] of Object.entries(node.props ?? {})) props[camelize(key)] = value;
  const children: unknown = node.children;
  const slot =
    children !== null && typeof children === 'object' && !Array.isArray(children)
      ? (children as { default?: unknown }).default
      : undefined;
  const content =
    typeof slot === 'function'
      ? visibleNodes(slot())
      : Array.isArray(children) || typeof children === 'string'
        ? visibleNodes(children)
        : [];
  return { props, content };
}

const text = (value: unknown): string | undefined =>
  typeof value === 'string' ? value : undefined;

/** A part's `Toolbar.Group` as the toolbar renders it. */
export interface ToolbarGroupPart {
  readonly className: string | undefined;
  readonly content: readonly VNode[];
}

/** The children of the preset toolbar, by role. */
export interface ToolbarChildren {
  /** The content of each `Toolbar.Slot`. The last part for a slot wins. */
  readonly slots: ReadonlyMap<string, readonly VNode[]>;
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
  const slots = new Map<string, readonly VNode[]>();
  const groups = new Map<string, ToolbarGroupPart>();
  const slotRequests: ToolbarSlotRequest[] = [];
  const groupRequests: ToolbarGroupRequest[] = [];
  const appended: VNode[] = [];
  for (const node of visibleNodes(children)) {
    if (isPart(node, 'xlsxToolbarGroup')) {
      const { props, content } = partProps(node);
      const id = String(props.id);
      const className = [text(props.class), text(props.className)].filter(Boolean).join(' ');
      groups.set(id, { className: className || undefined, content });
      groupRequests.push({
        id,
        after: text(props.after),
        hidden: vnodeFlag(props.hidden),
        label: text(props.label),
        labelKey: text(props.labelKey),
        content: content.length > 0,
      });
    } else if (isPart(node, 'xlsxToolbarSlot')) {
      const { props, content } = partProps(node);
      const slotId = String(props.slotId);
      slots.set(slotId, content);
      slotRequests.push({ slotId, hidden: vnodeFlag(props.hidden), content: content.length > 0 });
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
