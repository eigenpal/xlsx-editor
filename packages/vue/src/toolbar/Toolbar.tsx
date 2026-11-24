/** @jsxImportSource vue */
// The editing toolbar: the chrome registry as one row of controls. The Vue twin of the React
// adapter's `toolbar/Toolbar.tsx`, which is after
// docx-editor/packages/react/src/editor/toolbar/DocxEditorToolbar.tsx at a0825f1f (EigenPal);
// the header there explains the rules and the differences from DOCX Editor. The default
// arrangement is derived from `defaultChromeGroups()` (`toolbarGroups`), the bar measures itself and
// moves the groups that do not fit, in collapse order, into a "More" panel (`measureToolbar`,
// `observeToolbar`), the arrow keys move between its controls (`moveBarFocus`), Control or
// Command and S runs the save (`listenSaveKey`), and the parts in its default slot change the
// bar (`host-parts.tsx`, `named-parts.tsx`, `action.tsx`, and `arrangeToolbar` in core).
//
// Differences from React: the toolbar host is a computed ref that the slots inject
// (`toolbarHostKey`), so a change of `onSave` changes each slot's state at once; the bar takes
// the host's attributes (`inheritAttrs: false`), because the component also renders the
// number format dialog beside the bar; the measure runs in a watcher after the render; and the
// render reads the default slot, and keeps the arrangement in a ref when its groups change, so
// the measure watcher sees them, as docx-editor/packages/vue/src/editor/toolbar/
// DocxEditorToolbar.tsx at 10eeaf67 (EigenPal) keeps its layout.

import {
  Fragment,
  computed,
  defineComponent,
  h,
  provide,
  shallowRef,
  watch,
  type AllowedComponentProps,
  type ExtractPublicPropTypes,
  type PropType,
  type VNode,
} from 'vue';
import {
  TOOLBAR_FIXED_ATTRIBUTE,
  TOOLBAR_PANEL_ITEMS,
  arrangeToolbar,
  chromeSlotId,
  guardToolbarMousedown,
  listenSaveKey,
  markChromeDirection,
  measureToolbar,
  moveBarFocus,
  observeToolbar,
  runToolbarCommand,
  toolbarGroupHeading,
  type ArrangedToolbarGroup,
  type ChromeSlotId,
  type EditorSnapshot,
  type SavedWorkbook,
  type ToolbarArrangement,
  editorClipboardReader,
  type ToolbarHost,
  type ChromePartTranslate,
} from '@xlsx-editor.dev/core/editor';
import { ToolbarAction, ToolbarSeparator } from './action.tsx';
import { TOOLBAR_PARTS, type ToolbarNamedParts } from './named-parts.tsx';
import type { TFunction } from '@xlsx-editor.dev/i18n';
import { useEditorState, useTranslation, useXlsxEditor } from '../context.ts';
import { usePartTranslation } from '../locale.ts';
import { ChromeIcon, toolbarHostKey, translateKey } from './slot.tsx';
import { GroupMenu, type PartLayout } from './parts.tsx';
import { popupTriggerKeys, usePopup } from './popup.ts';
import { ICON_MORE } from './icons.ts';
import { NumberFormatDialogHost, panelCloseKey } from './NumberFormatDialog.tsx';
import {
  XlsxEditorToolbarGroup,
  XlsxEditorToolbarSlot,
  control,
  partLayoutKey,
  readToolbarChildren,
  warnToolbarProblems,
  type ToolbarChildren,
} from './host-parts.tsx';

/** The names of the parts in the problems that the toolbar prints. */
const PART_NAMES = { slot: 'XlsxEditorToolbar.Slot', group: 'XlsxEditorToolbar.Group' } as const;

const xlsxEditorToolbarProps = {
  /**
   * Receives the workbook as bytes of the opened package's kind (`snapshot().workbookKind`: a
   * workbook or a template, macro-enabled when `snapshot().macroEnabled` is true) when the save
   * control is pressed, or when the user presses Control or Command and S in the editor, also with
   * the save control hidden. Without it, the save control is disabled with the engine's reason, and
   * the browser keeps the shortcut. A refused save calls nothing: the snapshot holds the refusal
   * with `command: 'save'`. The bytes arrive after the press, because a save runs in short tasks.
   * `saved.stillOpen` is false when another open, `close()`, or `destroy()` ended the workbook
   * during the save, and `saved.revision` is the revision of the bytes. A `@save` listener sets it
   * too.
   */
  onSave: {
    type: Function as PropType<(bytes: Uint8Array<ArrayBuffer>, saved: SavedWorkbook) => void>,
    default: undefined,
  },
  /**
   * Called once when a press of the save control, or of Control or Command and S in the editor,
   * starts a delivery to `onSave`. It gets the promise of the result, so you can show that the
   * save runs. A press that joins the running save of an earlier press and a refused save do
   * not call it. A press that joins a save from `saveInTasks` calls it, because that save has no
   * delivery yet. The promise resolves to `null` when the save is refused later. A
   * `@save-start` listener sets it too.
   */
  onSaveStart: {
    type: Function as PropType<(saving: Promise<SavedWorkbook | null>) => void>,
    default: undefined,
  },
  /**
   * Whether your page opens another workbook now, while the editor still shows the old one: for
   * example during a download or a file read before `load`. While it is true, the save control
   * is disabled with the reason that a workbook is loading, and Control or Command and S in the
   * editor starts no save work. After such a press, and only then, the snapshot holds the
   * refusal `loading`, with `command: 'save'`. When it changes to false, the toolbar forgets
   * that refusal and keeps any other refusal. The default is `false`.
   */
  opening: { type: Boolean, default: undefined },
  /**
   * `false` wraps the controls onto more rows instead of moving the groups that do not fit
   * into the "More" panel. It has no effect with `:preset="false"`, which always wraps.
   */
  overflow: { type: Boolean, default: true },
  /**
   * `false` renders only the default slot, in its order, with no measure and no "More" panel.
   * An `XlsxEditorToolbar.Slot` there renders the built-in control of its slot, or its default
   * slot. The default is `true`: the groups of `defaultChromeGroups()` (the toolbar groups of
   * `CHROME_GROUPS`, without the menu bar's), changed by the default slot,
   * where `XlsxEditorToolbar.Slot` replaces or hides one slot where it stands and
   * `XlsxEditorToolbar.Group` adds a host group or changes a built-in group. Other content
   * shows in one block after the groups, which never moves into the "More" panel.
   */
  preset: { type: Boolean, default: true },
  /** Appended after the `xlsx-toolbar` class. A host `class` also works. */
  className: { type: String, default: undefined },
  /**
   * The label resolver of the toolbar, its parts, its "More" panel, and its number format
   * dialog: a key and the values of its placeholders, as DOCX Editor's `t` prop. A key that it
   * leaves unresolved reads the root's `locale`. Keep it stable, for example from
   * `useChromeTranslate(overrides)`: a new function measures the bar again.
   */
  t: { type: Function as PropType<ChromePartTranslate>, default: undefined },
  /**
   * Slot ids to leave out, such as `['file.save']`. A group whose slots are all hidden is left
   * out with its separator. The order of the list does not matter.
   */
  hidden: { type: Array as PropType<readonly ChromeSlotId[]>, default: undefined },
} as const;

/** Props of `XlsxEditorToolbar`, the twin of React's `XlsxEditorToolbarProps`. @public */
export type XlsxEditorToolbarProps = ExtractPublicPropTypes<typeof xlsxEditorToolbarProps>;

/** The heading of a group: the registry label of a built-in group, or a host group's own. */
function groupHeading(t: TFunction, group: ArrangedToolbarGroup): string {
  const translate = (key: string) => translateKey(t, key);
  return group.chrome ? translate(group.chrome.labelKey) : toolbarGroupHeading(group, translate);
}

/** The slots that a part replaces, and those whose built-in control a part changes. */
interface SlotSets {
  readonly replaced: ReadonlySet<string>;
  readonly customized: ReadonlySet<string>;
}

/**
 * The controls of a group in `layout`: each built-in control, changed by its part, or the
 * content of the `Toolbar.Slot` that replaces it, then the content of the group's
 * `Toolbar.Group`. A menu group in the bar keeps its menu of the other controls, and the
 * replacing content and the changed controls follow it. In a row of the "More" panel, the
 * `overflowContent` of a slot takes the place of its control, and the `overflowContent` of a
 * group takes the place of its default slot.
 */
function groupContent(
  group: ArrangedToolbarGroup,
  parts: ToolbarChildren,
  sets: SlotSets,
  layout: PartLayout
): (VNode | null)[] {
  const { replaced, customized } = sets;
  const chrome = group.chrome;
  const row = layout === 'row';
  const controls = chrome?.controls ?? [];
  const own = !!chrome?.menu && layout === 'bar';
  const menu =
    chrome && own
      ? controls.filter((c) => {
          const slot = chromeSlotId(chrome, c);
          return !replaced.has(slot) && !customized.has(slot);
        })
      : [];
  // Each item is a Fragment in its own place and with its own key, as in React: so the content
  // of one slot never takes the element of another, as when two slots change at once.
  const keyed = (key: string, nodes: unknown) =>
    nodes == null || nodes === false
      ? null
      : h(Fragment, { key }, (Array.isArray(nodes) ? [...nodes] : [nodes]) as VNode[]);
  const groupPart = parts.groups.get(group.id);
  return [
    chrome && menu.length > 0 ? (
      <GroupMenu key="menu" group={{ ...chrome, controls: menu }} />
    ) : null,
    ...controls.map((c) => {
      const slot = chromeSlotId(chrome!, c);
      const part = parts.slots.get(slot);
      if (row && part?.overflowContent) return keyed(`row:${c.id}`, part.overflowContent());
      if (replaced.has(slot)) return keyed(`slot:${c.id}`, part?.content);
      if (own && !customized.has(slot)) return null;
      const override = customized.has(slot) ? (part?.override ?? null) : null;
      return keyed(`control:${c.id}`, control(chrome!, c, layout, replaced, override));
    }),
    row && groupPart?.overflowContent
      ? keyed('overflow', groupPart.overflowContent())
      : keyed('content', groupPart?.content),
  ];
}

/** A group in the bar: its controls, or its menu, and the host's content. */
function barGroup(
  t: TFunction,
  group: ArrangedToolbarGroup,
  parts: ToolbarChildren,
  sets: SlotSets
) {
  return (
    <div
      key={`group:${group.id}`}
      class={['xlsx-toolbar-group', parts.groups.get(group.id)?.className]}
      role="group"
      aria-label={groupHeading(t, group)}
      data-group={group.id}
    >
      {groupContent(group, parts, sets, 'bar')}
    </div>
  );
}

const isLoading = (s: EditorSnapshot): boolean => s.status === 'loading';

/** The "More" button and its panel: the groups that the bar could not hold, as rows. */
const MorePanel = defineComponent({
  name: 'XlsxToolbarMore',
  props: {
    groups: { type: Array as PropType<readonly ArrangedToolbarGroup[]>, required: true },
    parts: { type: Object as PropType<ToolbarChildren>, required: true },
    sets: { type: Object as PropType<SlotSets>, required: true },
  },
  setup(props) {
    const { t } = useTranslation();
    // The rows of the panel: a part in a section renders as a row.
    provide(partLayoutKey, 'row');
    const popup = usePopup(TOOLBAR_PANEL_ITEMS, { nested: true });
    // An open of another workbook closes the panel: its rows showed the old workbook's state.
    const loading = useEditorState(isLoading);
    watch(loading, (now) => {
      if (now) popup.setOpen(false);
    });
    provide(panelCloseKey, () => popup.setOpen(false));
    return () => {
      const text = t('formattingBar.more');
      return (
        <div ref={popup.root} class="xlsx-toolbar-more">
          <button
            type="button"
            class="xlsx-toolbar-button"
            aria-haspopup="dialog"
            aria-expanded={popup.open.value}
            aria-controls={popup.open.value ? popup.id : undefined}
            aria-label={text}
            title={text}
            onMousedown={guardToolbarMousedown}
            onClick={() => popup.setOpen(!popup.open.value)}
            {...popupTriggerKeys(popup, true)}
          >
            <ChromeIcon paths={ICON_MORE} />
          </button>
          {popup.open.value ? (
            <div
              id={popup.id}
              role="dialog"
              aria-label={text}
              class="xlsx-toolbar-panel"
              data-popup=""
            >
              {props.groups.map((g) => (
                <div
                  key={`group:${g.id}`}
                  role="group"
                  aria-label={groupHeading(t, g)}
                  class="xlsx-toolbar-section"
                >
                  <span class="xlsx-toolbar-heading" aria-hidden="true">
                    {groupHeading(t, g)}
                  </span>
                  {groupContent(g, props.parts, props.sets, 'row')}
                </div>
              ))}
            </div>
          ) : null}
        </div>
      );
    };
  },
});

/** The editing toolbar, with no statics. */
const XlsxEditorToolbarRoot = defineComponent({
  name: 'XlsxEditorToolbar',
  inheritAttrs: false,
  props: xlsxEditorToolbarProps,
  setup(props, { attrs, slots }) {
    const { t, translator, dir } = usePartTranslation('XlsxEditorToolbar', () => props.t);
    const editor = useXlsxEditor();
    // A press reads `opening` at its call, so a press and the state of each control agree. A
    // change of a prop gives a new host, so each control reads its state again.
    // A slot of the menu bar in a toolbar group reads the same host: Paste reads the clipboard
    // of the editor's window. Open needs the menu bar's `onOpen`, so in a toolbar it is disabled.
    const host = computed<ToolbarHost>(() => {
      const opening = props.opening;
      const read = editorClipboardReader(editor.value);
      return {
        ...(props.onSave ? { save: props.onSave } : {}),
        ...(props.onSaveStart ? { saveStart: props.onSaveStart } : {}),
        ...(opening === undefined ? {} : { opening: () => props.opening ?? false }),
        ...(read ? { readClipboard: read } : {}),
      };
    });
    provide(toolbarHostKey, host);
    // The render reads the parts of the default slot and keeps the arrangement here while its
    // groups stay the same (`key`), so the measure watcher runs again only on a change.
    const arranged = shallowRef<ToolbarArrangement>(arrangeToolbar({ hidden: props.hidden }));
    const groups = computed(() => arranged.value.groups);
    // Whether the render shows the fixed block, which takes room from the groups: its arrival
    // or its removal measures again.
    const hasFixed = shallowRef(false);
    const bar = shallowRef<HTMLDivElement | null>(null);
    // When the open ends, the refusal of a save during the open no longer applies: forget it,
    // also when the open ends early. A later refusal of another command stays.
    let wasOpening = false;
    watch(
      [() => props.opening, editor],
      ([opening, ed]) => {
        const ended = wasOpening && !opening;
        wasOpening = !!opening;
        if (!ended || !ed) return;
        const refusal = ed.snapshot().refusal;
        if (refusal?.code === 'loading' && refusal.command === 'save') ed.dismissRefusal();
      },
      { immediate: true }
    );
    // Control or Command and S runs the save of the save control, from `onSave` alone, so a
    // hidden save control does not stop the shortcut. Scoped to this editor (`listenSaveKey`).
    watch(
      [bar, editor, () => !!props.onSave, host],
      ([el, ed, saves, h], _old, onCleanup) => {
        if (!el || !ed || !saves) return;
        onCleanup(listenSaveKey(el, () => runToolbarCommand(ed, 'file.save', undefined, h)));
      },
      { immediate: true, flush: 'post' }
    );
    // Mark the bar's computed direction for the mirrored icons (`markChromeDirection`).
    watch([bar, dir, () => attrs['dir']], ([el]) => markChromeDirection(el), {
      immediate: true,
      flush: 'post',
    });
    const widths = new Map<string, number>();
    // The ids of the groups in the "More" panel, in bar order (`toolbarCollapse`).
    const collapsed = shallowRef<readonly string[]>([]);
    // A new locale or a new set of groups changes the widths: measure every group again.
    let measuredFor: { t: TFunction; groups: readonly ArrangedToolbarGroup[] } | null = null;
    watch(
      [() => props.overflow && props.preset, groups, collapsed, translator, bar, hasFixed],
      ([measuring, all, , tr, el], _old, onCleanup) => {
        if (measuredFor?.t !== tr || measuredFor.groups !== all) widths.clear();
        measuredFor = { t: tr, groups: all };
        if (!el || !measuring) {
          if (collapsed.value.length > 0) collapsed.value = [];
          return;
        }
        const measure = (): void => {
          const next = measureToolbar(el, all, widths);
          const now = collapsed.value;
          if (now.length !== next.length || now.some((id, i) => id !== next[i]))
            collapsed.value = next;
        };
        measure();
        // The bar and each shown group: a group grows when its fonts load or its value changes.
        onCleanup(observeToolbar(el, measure));
      },
      { immediate: true, flush: 'post' }
    );

    return () => {
      const children = slots.default?.();
      const measuring = props.overflow && props.preset;
      let body: unknown = children;
      if (!props.preset) hasFixed.value = false;
      else {
        const parts = readToolbarChildren(children);
        const next = arrangeToolbar({
          hidden: props.hidden,
          slots: parts.slotRequests,
          groups: parts.groupRequests,
          names: PART_NAMES,
        });
        warnToolbarProblems(next.problems);
        if (next.key !== arranged.value.key) arranged.value = next;
        const all = next.groups;
        const sets: SlotSets = { replaced: next.replaced, customized: next.customized };
        const out = new Set(collapsed.value);
        const rest = all.filter((g) => out.has(g.id));
        const items: unknown[] = all
          .filter((g) => !out.has(g.id))
          .map((g) => barGroup(t, g, parts, sets));
        if (rest.length > 0)
          items.push(<MorePanel key="more" groups={rest} parts={parts} sets={sets} />);
        // Loose children, and the content of a `Toolbar.Slot` that names no slot, never
        // collapse: they have no group, and so no heading in the panel. The bar charges them as
        // fixed width. The content of each such slot is keyed by its slot, as in React, so its
        // path keys stay apart from those of the loose children.
        const fixed = [
          ...parts.appended,
          ...next.outside.map((slot) =>
            h(Fragment, { key: `slot:${slot}` }, [...(parts.slots.get(slot)?.content ?? [])])
          ),
        ];
        hasFixed.value = fixed.length > 0;
        if (fixed.length > 0)
          items.push(
            <div key="fixed" class="xlsx-toolbar-group" {...{ [TOOLBAR_FIXED_ATTRIBUTE]: '' }}>
              {fixed}
            </div>
          );
        body = items;
      }
      return (
        <NumberFormatDialogHost>
          {{
            default: () => (
              <div
                {...attrs}
                ref={bar}
                role="toolbar"
                // The chrome direction, else a host `dir` attribute: never `undefined` over it.
                dir={dir.value ?? (attrs['dir'] as string | undefined)}
                aria-label={t('formattingBar.label')}
                class={['xlsx-toolbar', props.className, attrs.class]}
                {...(measuring ? { 'data-overflow': '' } : {})}
                // The bar guards every press on it, also between its controls.
                onMousedown={guardToolbarMousedown}
                onKeydown={moveBarFocus}
              >
                {body}
              </div>
            ),
          }}
        </NumberFormatDialogHost>
      );
    };
  },
});

/**
 * The toolbar and its parts, the twin of React's `XlsxEditorToolbarNamespace`, as DOCX Editor's
 * Vue adapter declares `DocxEditorToolbarNamespace`. @public
 */
export interface XlsxEditorToolbarNamespace extends ToolbarNamedParts {
  /** The toolbar. A host `class` and `style` land on the bar, as on a component. */
  (props: XlsxEditorToolbarProps & AllowedComponentProps): VNode;
  /** Replaces, changes, or hides one slot where it stands, or renders its built-in control. */
  readonly Slot: typeof XlsxEditorToolbarSlot;
  /** A host group that the bar measures and collapses, or a change of a built-in group. */
  readonly Group: typeof XlsxEditorToolbarGroup;
  /** A host-owned action, styled and behaved like the packaged controls. */
  readonly Action: typeof ToolbarAction;
  /** A rule between controls. */
  readonly Separator: typeof ToolbarSeparator;
}

/**
 * The editing toolbar: every group of the chrome registry in bar order. Place it inside
 * `XlsxEditorViewport`, which carries the styling scope. Its parts change the bar, in its
 * default slot: the named parts (`XlsxEditorToolbar.Bold`, ...), `XlsxEditorToolbar.Slot`,
 * `XlsxEditorToolbar.Group`, `XlsxEditorToolbar.Action`, and `XlsxEditorToolbar.Separator`.
 *
 * @example
 * ```ts
 * h(XlsxEditorToolbar, { preset: false }, () => [
 *   h(XlsxEditorToolbar.Group, { id: 'text', label: 'Text' }, () => [
 *     h(XlsxEditorToolbar.Slot, { slotId: 'font.bold' }),
 *     h(XlsxEditorToolbar.Slot, { slotId: 'font.italic' }),
 *   ]),
 * ]);
 * ```
 *
 * @public
 */
export const XlsxEditorToolbar = Object.assign(XlsxEditorToolbarRoot, TOOLBAR_PARTS, {
  Slot: XlsxEditorToolbarSlot,
  Group: XlsxEditorToolbarGroup,
  Action: ToolbarAction,
  Separator: ToolbarSeparator,
}) as unknown as XlsxEditorToolbarNamespace;
