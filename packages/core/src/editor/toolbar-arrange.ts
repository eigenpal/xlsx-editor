// The arrangement of the preset toolbar with the parts that a host passes as its children:
// `Toolbar.Slot`, which replaces or hides one slot where it stands, and `Toolbar.Group`, which
// adds a host group or changes a built-in group. Both adapters read their children into the
// requests of this file and render the result, so the React and Vue toolbars arrange the same
// way. Adapter support, not a stable API.
//
// `arrangeToolbarGroups` is a copy of docx-editor/packages/core/src/editor/toolbar-fit.ts at
// 10eeaf67 (EigenPal), with no difference. `arrangeToolbar` is after `arrangeGroups` of
// docx-editor/packages/react/src/editor/toolbar/DocxEditorToolbar.tsx, and its problems after
// docx-editor/packages/react/src/editor/toolbar/toolbar-warnings.ts, at the same commit.
// Differences from DOCX Editor:
// - The bar collapses its groups from the end, not in a declared priority order, so a group
//   has no `priority` and no `pinned`. A host group with no `after` stands after every
//   built-in group, so it collapses first, as a DOCX Editor host group with no priority does.
// - A group and a slot have no `overflowContent`: in the "More" panel, a group shows its own
//   content, and a slot shows the content of its `Toolbar.Slot`.
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
}

/** A `Toolbar.Group` child of the preset toolbar, as the arrangement reads it. @internal */
export interface ToolbarGroupRequest {
  readonly id: string;
  readonly after?: string | undefined;
  readonly hidden?: boolean | undefined;
  readonly label?: string | undefined;
  readonly labelKey?: string | undefined;
  /** The part has content of its own: the controls of a host group, or extra controls. */
  readonly content: boolean;
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
  /** The `Toolbar.Slot` ids that the registry does not hold. Their content is appended. */
  readonly outside: readonly string[];
  /** Development warnings: compositions that cannot do what they say. */
  readonly problems: readonly string[];
  /** Equal for two arrangements with the same groups, controls, and replaced slots. */
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

const BUILT_IN_IDS: readonly string[] = (CHROME_GROUPS as readonly ChromeGroup[]).map((g) => g.id);
const SLOT_IDS: ReadonlySet<string> = new Set(CHROME_SLOT_IDS);

/**
 * Arrange the preset bar: the groups of `CHROME_GROUPS` in registry order, with the slots of
 * `hidden` and of a hidden `Toolbar.Slot` left out, and the `Toolbar.Group` requests applied.
 * The last request for a slot or a group wins; a group keeps the place of its last request. A
 * built-in group with `hidden` is left out with its separator, and so is a built-in group with
 * no control left and no content. A host group with `hidden` or no content is left out, but it
 * still anchors the groups that name it in `after`. @internal
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
    else if (slot.content) replaced.add(slot.slotId as ChromeSlotId);
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
    if (controls.length === 0 && !request?.content) continue;
    byId.set(g.id, {
      id: g.id,
      chrome: controls.length === g.controls.length ? g : { ...g, controls },
    });
  }
  const hosts = [...requests.values()].filter((r) => !BUILT_IN_IDS.includes(r.id));
  for (const host of hosts) {
    if (host.hidden || !host.content) continue;
    byId.set(host.id, { id: host.id, chrome: null, label: host.label, labelKey: host.labelKey });
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
    groups.map((g) => (g.chrome ? [g.id, g.chrome.controls.map((c) => c.id)] : g.id)),
    [...replaced].sort(),
  ]);
  return { groups, replaced, outside, problems, key };
}
