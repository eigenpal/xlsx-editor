// The arrangement of the preset toolbar with the parts that a host passes as its children:
// `Toolbar.Slot` and the named parts (`Toolbar.Bold`), which replace, change, or hide one slot
// where it stands, and `Toolbar.Group`, which adds a host group or changes a built-in group.
// Both adapters read their children into the requests of this file and render the result, so
// the React and Vue toolbars arrange the same way. Adapter support, not a stable API.
//
// `arrangeToolbarGroups` is a copy of docx-editor/packages/core/src/editor/toolbar-fit.ts at
// 10eeaf67 (EigenPal), with no difference. `TOOLBAR_COLLAPSE_ORDER`, `TOOLBAR_PINNED_GROUPS`,
// `toolbarGroupPriority`, and `collapseOrder` are copies of the same file at ad686574, with the
// differences below. `arrangeToolbar` is after `arrangeGroups` of
// docx-editor/packages/react/src/editor/toolbar/DocxEditorToolbar.tsx, and its problems after
// docx-editor/packages/react/src/editor/toolbar/toolbar-warnings.ts, at 10eeaf67.
// Differences from DOCX Editor:
// - The collapse order is the reverse bar order of the workbook groups: the save control leaves
//   first and the history group last, as the bar collapsed from its end before groups had a
//   priority. No built-in group is pinned: the workbook toolbar has no control that is the only
//   way to reach a pane, as DOCX Editor's comments toggle is.
// - The bar collapses whole groups, not one control at a time
//   (docx-editor/packages/core/src/editor/toolbar-item-fit.ts): a "More" panel section shows
//   one whole group, with its heading.
// - A slot with `icon`, `asChild`, or `className` on a group that renders as one menu
//   (`ChromeGroup.menu`) leaves the menu and shows as its own control after the menu button, as
//   the content of a `Toolbar.Slot` does (`customized`).
// - The slots of the `hidden` list are hidden as a `Toolbar.Slot` with `hidden` hides them.
//   The list wins over a `Toolbar.Slot` with content for the same slot.
// - The problems are returned as sentences that name the parts as the adapter names them
//   (`ToolbarPartNames`, after pptx-editor/packages/core/src/editor/dev-warnings.ts at 929ca860
//   (EigenPal)), and each adapter prints them in development, once.

import {
  CHROME_GROUPS,
  CHROME_SLOT_IDS,
  chromeSlotId,
  type ChromeGroup,
  type ChromeSlotId,
} from './chrome-controls.ts';
import { isToolbarButtonSlot } from './toolbar-named-parts.ts';

/**
 * Collapse order: the first id here is the first group to leave the bar. The reverse of the bar
 * order, so the bar keeps its reading order: the save control leaves first, and the history
 * and font groups stay longest. A group that this list does not name, such as a host group,
 * collapses before all of these, in reverse bar order.
 *
 * @internal
 */
export const TOOLBAR_COLLAPSE_ORDER: readonly string[] = [
  'file',
  'clear',
  'cells',
  'number',
  'alignment',
  'borders',
  'fill',
  'font',
  'history',
];

/**
 * Groups that never move into the "More" panel. None of the workbook groups: each control is
 * also in a menu, a key, or the panel. A host group joins with `pinned`.
 *
 * @internal
 */
export const TOOLBAR_PINNED_GROUPS: ReadonlySet<string> = new Set<string>();

/**
 * The priority of a declared group: its place in {@link TOOLBAR_COLLAPSE_ORDER}, in steps of
 * 10, so `file` is 10 and `history` is 90. A host group with a `priority` sorts among these,
 * and a LOWER number leaves the bar first. The steps leave room between two built-in groups.
 *
 * @internal
 */
export function toolbarGroupPriority(
  id: string,
  order: readonly string[] = TOOLBAR_COLLAPSE_ORDER
): number | undefined {
  const index = order.indexOf(id);
  return index === -1 ? undefined : (index + 1) * 10;
}

/**
 * The collapse order actually used: the declared one, with undeclared groups first.
 *
 * `priorities` gives a group an explicit place among the declared ones (see
 * {@link toolbarGroupPriority}), and overrides a declared group's own priority. Groups with
 * neither collapse first, in reverse bar order. Two groups with the same priority collapse
 * in reverse bar order too: the one further along the bar leaves first.
 *
 * @internal
 */
export function collapseOrder(
  groups: readonly string[],
  order: readonly string[] = TOOLBAR_COLLAPSE_ORDER,
  priorities?: ReadonlyMap<string, number>
): readonly string[] {
  const ranked: { id: string; priority: number; position: number }[] = [];
  const undeclared: string[] = [];
  groups.forEach((id, position) => {
    const priority = priorities?.get(id) ?? toolbarGroupPriority(id, order);
    if (priority === undefined || Number.isNaN(priority)) undeclared.push(id);
    else ranked.push({ id, priority, position });
  });
  ranked.sort((a, b) => a.priority - b.priority || b.position - a.position);
  return [...undeclared.reverse(), ...ranked.map((entry) => entry.id)];
}

/** One host group's placement request. @internal */
export interface ToolbarHostGroupPlacement {
  readonly id: string;
  /** The group this one follows. Absent, or unknown, places it after every group. */
  readonly after?: string | undefined;
}

/**
 * Bar order with the host's groups placed.
 *
 * A host group without `after` follows every built-in group, in order of appearance. A group
 * with `after` follows the group it names, after any host groups already placed there, so
 * several groups anchored to one place keep their order of appearance. An anchor may be
 * another host group. An anchor that never resolves (an unknown id, or a cycle) places the
 * group at the end.
 *
 * @internal
 */
export function arrangeToolbarGroups(
  builtIn: readonly string[],
  hosts: readonly ToolbarHostGroupPlacement[]
): readonly string[] {
  const result = [...builtIn];
  // The last id placed after each anchor, so a second group anchored there follows the first.
  const tail = new Map<string, string>();
  let pending = hosts.filter((host) => !builtIn.includes(host.id));
  while (pending.length > 0) {
    const deferred: ToolbarHostGroupPlacement[] = [];
    for (const host of pending) {
      const anchor = host.after;
      if (anchor === undefined || anchor === host.id) {
        result.push(host.id);
        continue;
      }
      if (!result.includes(anchor)) {
        deferred.push(host);
        continue;
      }
      const after = tail.get(anchor) ?? anchor;
      result.splice(result.indexOf(after) + 1, 0, host.id);
      tail.set(anchor, host.id);
    }
    if (deferred.length === pending.length) {
      // Nothing resolved in this pass: the rest name groups that do not exist.
      result.push(...deferred.map((host) => host.id));
      break;
    }
    pending = deferred;
  }
  return result;
}

/** A `Toolbar.Slot` child of the preset toolbar, as the arrangement reads it. @internal */
export interface ToolbarSlotRequest {
  /** The slot that the part names. An id that the registry does not hold is a problem. */
  readonly slotId: string;
  /** The part removes its slot. */
  readonly hidden: boolean;
  /** The part has content of its own, which takes the slot's place. */
  readonly content: boolean;
  /** The part gives an icon in place of the registry icon. Button slots only. */
  readonly icon?: boolean | undefined;
  /** The part renders the control as its single child element. Button slots only. */
  readonly asChild?: boolean | undefined;
  /**
   * With `asChild`: the children are one element. Else the control renders its own button, and
   * the arrangement names the problem.
   */
  readonly asChildElement?: boolean | undefined;
  /** The part adds a class name to the control. */
  readonly className?: boolean | undefined;
  /** The part gives content for the slot's row of the "More" panel. */
  readonly overflowContent?: boolean | undefined;
}

/** A `Toolbar.Group` child of the preset toolbar, as the arrangement reads it. @internal */
export interface ToolbarGroupRequest {
  readonly id: string;
  readonly after?: string | undefined;
  readonly hidden?: boolean | undefined;
  readonly label?: string | undefined;
  readonly labelKey?: string | undefined;
  /**
   * Collapse priority: a lower number leaves the bar first (`collapseOrder`). It also sets the
   * place of a built-in group. NaN counts as unset.
   */
  readonly priority?: number | undefined;
  /** The group stays in the bar at every width. */
  readonly pinned?: boolean | undefined;
  /** The part has content of its own: the controls of a host group, or extra controls. */
  readonly content: boolean;
  /**
   * The part gives content for the group's section of the "More" panel. It shows a built-in
   * group even with no control left, as the content of the part does.
   */
  readonly overflowContent?: boolean | undefined;
}

/** A group of the arranged bar. @internal */
export interface ArrangedToolbarGroup {
  readonly id: string;
  /**
   * The built-in group with its hidden slots left out, or null for a host group. A built-in
   * group can have no control left and still show, when its `Toolbar.Group` adds content.
   */
  readonly chrome: ChromeGroup | null;
  /** The literal heading of a host group. */
  readonly label?: string | undefined;
  /** The catalog key of a host group's heading. */
  readonly labelKey?: string | undefined;
  /** The collapse priority that the host set, or undefined for the default place. */
  readonly priority?: number | undefined;
  /** The group stays in the bar at every width. */
  readonly pinned: boolean;
}

/** The names of the two parts in one adapter, for the problems. @internal */
export interface ToolbarPartNames {
  readonly slot: string;
  readonly group: string;
}

/** The part names of the React adapter, the default. */
const REACT_PART_NAMES: ToolbarPartNames = {
  slot: 'XlsxEditor.Toolbar.Slot',
  group: 'XlsxEditor.Toolbar.Group',
};

/**
 * The host groups whose `after` does not place them after their anchor, by the chain of
 * anchors that `arrangeToolbarGroups` follows. `self`: the group names itself, so it is placed
 * as a group with no `after`. `cycle`: the chain leads into a cycle (`loop`), so the group is
 * placed last. `before`: the chain ends at `root`, whose anchor `missing` names no group, so the
 * groups of the chain are placed last, and this group stands before its anchor in `order`. A
 * group whose own anchor names no group is neither: it has its own problem.
 * @internal
 */
function unresolvedAnchors(
  builtIn: readonly string[],
  hosts: readonly ToolbarHostGroupPlacement[],
  order: readonly string[]
): {
  readonly self: string[];
  readonly cycle: { readonly id: string; readonly after: string; readonly loop: string[] }[];
  readonly before: {
    readonly id: string;
    readonly after: string;
    readonly root: string;
    readonly missing: string;
  }[];
} {
  // A group that names itself is placed with no anchor, so a chain through it resolves.
  const anchorOf = new Map(
    hosts.map((host) => [host.id, host.after === host.id ? undefined : host.after])
  );
  const self = hosts.filter((host) => host.after === host.id).map((host) => host.id);
  const cycle: { id: string; after: string; loop: string[] }[] = [];
  const before: { id: string; after: string; root: string; missing: string }[] = [];
  for (const host of hosts) {
    if (host.after === host.id) continue;
    if (host.after === undefined || !(builtIn.includes(host.after) || anchorOf.has(host.after)))
      continue;
    const seen: string[] = [];
    let at: string | undefined = host.id;
    while (at !== undefined && anchorOf.has(at) && !seen.includes(at)) {
      seen.push(at);
      at = anchorOf.get(at);
    }
    if (at !== undefined && seen.includes(at))
      cycle.push({ id: host.id, after: host.after, loop: seen.slice(seen.indexOf(at)) });
    else if (
      at !== undefined &&
      !builtIn.includes(at) &&
      order.indexOf(host.id) < order.indexOf(host.after)
    )
      before.push({ id: host.id, after: host.after, root: seen.at(-1)!, missing: at });
  }
  return { self, cycle, before };
}

/** The preset bar with the host's parts applied. @internal */
export interface ToolbarArrangement {
  /** The groups in bar order. A group with nothing to show is left out. */
  readonly groups: readonly ArrangedToolbarGroup[];
  /** The slots whose `Toolbar.Slot` content takes the place of the built-in control. */
  readonly replaced: ReadonlySet<ChromeSlotId>;
  /**
   * The shown slots whose part changes the built-in control (`icon`, `asChild`, `className`).
   * In a group that renders as one menu, such a slot shows as its own control after the menu.
   */
  readonly customized: ReadonlySet<ChromeSlotId>;
  /** The `Toolbar.Slot` ids that the registry does not hold. Their content is appended. */
  readonly outside: readonly string[];
  /** Development warnings: compositions that cannot do what they say. */
  readonly problems: readonly string[];
  /**
   * Equal for two arrangements with the same groups, controls, replaced and customized slots,
   * priorities, and pinned groups.
   */
  readonly key: string;
}

/**
 * The heading of a group part: for a built-in id, the group's own translated label; else the
 * literal label, else the translated key, else the id. @internal
 */
export function toolbarGroupHeading(
  group: Pick<ArrangedToolbarGroup, 'id' | 'label' | 'labelKey'>,
  translate: (key: string) => string
): string {
  const builtIn = (CHROME_GROUPS as readonly ChromeGroup[]).find((g) => g.id === group.id);
  if (builtIn) return translate(builtIn.labelKey);
  if (group.label !== undefined) return group.label;
  return group.labelKey !== undefined ? translate(group.labelKey) : group.id;
}

/** A request's priority, with NaN as unset, as `collapseOrder` reads it. */
function priorityOf(request: ToolbarGroupRequest | undefined): number | undefined {
  const priority = request?.priority;
  return priority === undefined || Number.isNaN(priority) ? undefined : priority;
}

const BUILT_IN_IDS: readonly string[] = (CHROME_GROUPS as readonly ChromeGroup[]).map((g) => g.id);
const SLOT_IDS: ReadonlySet<string> = new Set(CHROME_SLOT_IDS);

/**
 * Arrange the preset bar: the groups of `CHROME_GROUPS` in registry order, with the slots of
 * `hidden` and of a hidden `Toolbar.Slot` left out, and the `Toolbar.Group` requests applied.
 * The last request for a slot or a group wins; a group keeps the place of its last request. A
 * built-in group with `hidden` is left out with its separator, and so is a built-in group with
 * no control left, no content, and no overflow content. A host group with `hidden` or no content
 * is left out, but it still anchors the groups that name it in `after`. @internal
 */
export function arrangeToolbar(input: {
  readonly hidden?: readonly ChromeSlotId[];
  readonly slots?: readonly ToolbarSlotRequest[];
  readonly groups?: readonly ToolbarGroupRequest[];
  /** How the problems name the parts. The default is the React adapter's names. */
  readonly names?: ToolbarPartNames;
}): ToolbarArrangement {
  const names = input.names ?? REACT_PART_NAMES;
  const problems: string[] = [];
  const listed = new Set<string>(input.hidden ?? []);
  const leave = new Set<string>(listed);
  const replaced = new Set<ChromeSlotId>();
  const customized = new Set<ChromeSlotId>();
  const outside: string[] = [];
  const slots = new Map<string, ToolbarSlotRequest>();
  for (const slot of input.slots ?? []) slots.set(slot.slotId, slot);
  for (const slot of slots.values()) {
    if (!SLOT_IDS.has(slot.slotId)) {
      if (slot.content && !slot.hidden) {
        outside.push(slot.slotId);
        problems.push(
          `"${slot.slotId}" is not a toolbar slot id, so the content of its ` +
            `${names.slot} is appended after the groups. Use a ChromeSlotId, such as ` +
            `"font.bold".`
        );
      } else
        problems.push(
          `"${slot.slotId}" is not a toolbar slot id, so its ${names.slot} changes nothing. ` +
            `Use a ChromeSlotId, such as "font.bold".`
        );
    } else if (slot.hidden) leave.add(slot.slotId);
    else if (slot.content && listed.has(slot.slotId))
      problems.push(
        `${names.slot} "${slot.slotId}" has content, but the hidden list names ` +
          `the slot, so the slot is left out. Remove it from the hidden list to show the ` +
          `content.`
      );
    else if (slot.content) {
      replaced.add(slot.slotId as ChromeSlotId);
      // On a value slot, `asChild` holds no element: the children replace the control.
      if (slot.asChild && !isToolbarButtonSlot(slot.slotId))
        problems.push(
          `${names.slot} "${slot.slotId}" has asChild, but the slot is not a button (it is a ` +
            `list, a palette, or a dialog opener), so its children replace the control. Remove ` +
            `asChild to replace the control, or name a button slot.`
        );
      // The content takes the control's place, so a setting of the control does nothing.
      for (const setting of ['icon', 'className'] as const)
        if (slot[setting])
          problems.push(
            `${names.slot} "${slot.slotId}" has content, which takes the place of the ` +
              `control, so its ${setting} is ignored. Put the ${setting} on your content, or ` +
              `remove the content to change the built-in control.`
          );
    } else if (!listed.has(slot.slotId)) {
      const button = isToolbarButtonSlot(slot.slotId);
      if (slot.asChild && button && slot.asChildElement === false)
        problems.push(toolbarAsChildProblem(names.slot, slot.slotId));
      for (const setting of ['icon', 'asChild'] as const)
        if (slot[setting] && !button)
          problems.push(
            `${names.slot} "${slot.slotId}" has ${setting}, but the slot is not a button ` +
              `(it is a list, a palette, or a dialog opener), so the ${setting} is ignored. ` +
              `Give the slot content to replace the control.`
          );
      const element = slot.asChild && slot.asChildElement !== false;
      if ((slot.icon && button) || (element && button) || slot.className)
        customized.add(slot.slotId as ChromeSlotId);
    }
  }
  // Re-inserted so the LAST request's position is its order of appearance.
  const requests = new Map<string, ToolbarGroupRequest>();
  for (const group of input.groups ?? []) {
    requests.delete(group.id);
    requests.set(group.id, group);
  }
  const byId = new Map<string, ArrangedToolbarGroup>();
  for (const g of CHROME_GROUPS as readonly ChromeGroup[]) {
    const request = requests.get(g.id);
    if (request) {
      for (const setting of ['label', 'labelKey', 'after'] as const)
        if (request[setting] !== undefined)
          problems.push(
            `${names.group} "${g.id}" is a built-in group, so its ${setting} is ` +
              `ignored. A built-in group keeps its own label and place in the bar.`
          );
    }
    if (request?.hidden) {
      for (const c of g.controls) {
        const slot = chromeSlotId(g, c);
        if (replaced.has(slot as ChromeSlotId))
          problems.push(
            `${names.slot} "${slot}" has content, but its group "${g.id}" is hidden, so ` +
              `the content is left out. Show the group, or place the content in a host group.`
          );
      }
      continue;
    }
    const controls = g.controls.filter((c) => !leave.has(chromeSlotId(g, c)));
    if (controls.length === 0 && !request?.content && !request?.overflowContent) continue;
    byId.set(g.id, {
      id: g.id,
      chrome: controls.length === g.controls.length ? g : { ...g, controls },
      priority: priorityOf(request),
      pinned: request?.pinned ?? TOOLBAR_PINNED_GROUPS.has(g.id),
    });
  }
  const hosts = [...requests.values()].filter((r) => !BUILT_IN_IDS.includes(r.id));
  for (const host of hosts) {
    if (!host.hidden && !host.content && host.overflowContent)
      problems.push(
        `${names.group} "${host.id}" has overflowContent but no children, so the group is ` +
          `left out: a host group shows in the "More" panel only when its controls leave the ` +
          `bar. Give the group children.`
      );
    if (host.hidden || !host.content) continue;
    byId.set(host.id, {
      id: host.id,
      chrome: null,
      label: host.label,
      labelKey: host.labelKey,
      priority: priorityOf(host),
      pinned: host.pinned === true,
    });
  }
  const known = [...BUILT_IN_IDS, ...hosts.map((host) => host.id)];
  for (const host of hosts)
    if (host.after !== undefined && !known.includes(host.after))
      problems.push(
        `${names.group} "${host.id}" has after="${host.after}", but no group has ` +
          `that id. The group is placed after every other group. Known ids: ` +
          `${known.join(', ')}.`
      );
  // Placed over EVERY built-in id, hidden ones included, so `after` still finds its anchor.
  const order = arrangeToolbarGroups(
    BUILT_IN_IDS,
    hosts.map((host) => ({ id: host.id, after: host.after }))
  );
  const unresolved = unresolvedAnchors(BUILT_IN_IDS, hosts, order);
  const quoted = (ids: readonly string[]) => ids.map((id) => `"${id}"`).join(', ');
  for (const { id, after, loop } of unresolved.cycle)
    problems.push(
      `${names.group} "${id}" has after="${after}", which leads into a cycle of anchors: ` +
        `${quoted(loop)}. The group is placed after every other group. Change the after of ` +
        `one group of the cycle.`
    );
  for (const id of unresolved.self)
    problems.push(
      `${names.group} "${id}" has after="${id}", its own id, so the group is placed as a ` +
        `group with no after. Name another group.`
    );
  for (const { id, after, root, missing } of unresolved.before)
    problems.push(
      `${names.group} "${id}" has after="${after}", but its chain of anchors ends at ` +
        `"${root}", whose after="${missing}" names no group. So "${after}" is placed after ` +
        `every other group, and this group stands before it. Fix the after of "${root}".`
    );
  const groups = order.flatMap((id) => {
    const group = byId.get(id);
    return group ? [group] : [];
  });
  // JSON keeps the ids apart, also an id with a space or a comma.
  const key = JSON.stringify([
    groups.map((g) => [
      g.id,
      g.chrome ? g.chrome.controls.map((c) => c.id) : null,
      g.priority ?? null,
      g.pinned,
    ]),
    [...replaced].sort(),
    [...customized].sort(),
  ]);
  return { groups, replaced, customized, outside, problems, key };
}

/**
 * The problem of a part with `asChild` whose children are not one element: `part` names the part
 * as the adapter names it, and `name` its slot id or its label. The part renders its own button.
 * @internal
 */
export function toolbarAsChildProblem(part: string, name: string): string {
  return (
    `${part} "${name}" has asChild, but its children are not one element, so it renders its ` +
    `own button. Pass exactly one element, such as a <button>.`
  );
}
