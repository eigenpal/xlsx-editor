/** @jsxImportSource react */
// The editing toolbar: the chrome registry as one row of controls. After
// docx-editor/packages/react/src/editor/toolbar/DocxEditorToolbar.tsx at a0825f1f (EigenPal):
// the default arrangement is derived from `defaultChromeGroups()`, never listed by hand, so a slot
// added to the registry appears here without an edit; the bar measures itself and moves the
// groups that do not fit, from the end, into a "More" panel of full-width rows; and the bar's
// own `mousedown` keeps the keyboard on the grid; and a group whose slots are all hidden is
// left out. The host's parts change the preset bar, after the same file at 10eeaf67:
// `XlsxEditor.Toolbar.Slot` replaces or hides a slot where it stands, `XlsxEditor.Toolbar.Group`
// adds a host group that the bar measures and moves into the panel with its own groups, or
// changes a built-in group, other children show in one fixed block after the groups, and
// `preset={false}` renders only the children (`host-parts.tsx`, and `arrangeToolbar` in core).
// The named parts (`Toolbar.Bold`), `Toolbar.Action`, `Toolbar.Separator`, and the collapse by
// `priority` and `pinned` are after the same files at ad686574 (`named-parts.tsx`,
// `action.tsx`, and `toolbarCollapse` in core).
// Differences: slots can also be hidden with one `hidden` list of slot ids; the bar collapses
// whole groups, not one control at a time; the `t` prop is DOCX Editor's root resolver shape,
// `(key, params) => string`, not the toolbar's `(key) => string`, because the catalog here
// fills values into its messages, and a key that `t` leaves unresolved reads the root's
// `locale` (`withChromeTranslate`); a group can render as one menu
// (`ChromeGroup.menu`), and the content of a `Toolbar.Slot` for one of its slots shows after the
// menu button; the arrow keys move between the controls of the bar
// as the toolbar pattern of WAI-ARIA describes, while Tab still reaches each control, as in
// DOCX Editor (no roving tab index); and the group widths are measured again after a change of
// locale or of a group's own size. `onSave` receives the saved bytes (`onSave(bytes, saved)`), where
// DocxEditorToolbar's `onSave()` takes none: the toolbar saves through the engine and hands the
// result over, as docx-editor's PaginatedDocxEditorShell.tsx `onSave(bytes)` does, so a host
// never calls `save()` a second time.

import {
  Fragment,
  useCallback,
  useEffect,
  useLayoutEffect,
  useMemo,
  useRef,
  useState,
  type ReactElement,
  type ReactNode,
} from 'react';
import {
  TOOLBAR_FIXED_ATTRIBUTE,
  TOOLBAR_PANEL_ITEMS,
  arrangeToolbar,
  chromeSlotId,
  guardToolbarMousedown,
  listenSaveKey,
  measureToolbar,
  moveBarFocus,
  markChromeDirection,
  observeToolbar,
  toolbarGroupHeading,
  type ArrangedToolbarGroup,
  runToolbarCommand,
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
import { useChromeDir, withChromeTranslate } from '../locale.tsx';
import { ChromeIcon, ToolbarHostContext, translateKey } from './slot.tsx';
import { GroupMenu, type PartLayout } from './parts.tsx';
import { popupTriggerKeys, usePopup } from './popup.ts';
import { ICON_MORE } from './icons.ts';
import { NumberFormatDialogHost, PanelCloseContext } from './NumberFormatDialog.tsx';
import {
  Control,
  PartLayoutContext,
  XlsxEditorToolbarGroup,
  XlsxEditorToolbarSlot,
  readToolbarChildren,
  slotOverrideOf,
  warnToolbarProblems,
  type ToolbarChildren,
} from './host-parts.tsx';

/** Props for `XlsxEditor.Toolbar`. @public */
export interface XlsxEditorToolbarProps {
  /**
   * Receives the workbook as bytes of the opened package's kind (`snapshot().workbookKind`: a
   * workbook or a template) when the save control is pressed, or when the user presses Control
   * or Command and S in the editor, also with the save control hidden. Without it, the save
   * control is disabled with the engine's reason, and the browser keeps the shortcut. A refused
   * save calls nothing: the snapshot holds the refusal with `command: 'save'`. The bytes arrive
   * after the press, because a save runs in short tasks. `saved.stillOpen` is false when
   * another open, `close()`, or `destroy()` ended the workbook during the save, and
   * `saved.revision` is the revision of the bytes.
   */
  readonly onSave?: (bytes: Uint8Array<ArrayBuffer>, saved: SavedWorkbook) => void;
  /**
   * Called once when a press of the save control, or of Control or Command and S in the editor,
   * starts a delivery to `onSave`. It gets the promise of the result, so you can show that the
   * save runs. A press that joins the running save of an earlier press and a refused save do
   * not call it. A press that joins a save from `saveInTasks` calls it, because that save has no
   * delivery yet. The promise resolves to `null` when the save is refused later.
   */
  readonly onSaveStart?: (saving: Promise<SavedWorkbook | null>) => void;
  /**
   * Whether your page opens another workbook now, while the editor still shows the old one: for
   * example during a download or a file read before `load`. While it is true, the save control
   * is disabled with the reason that a workbook is loading, and Control or Command and S in the
   * editor starts no save work. After such a press, and only then, the snapshot holds the
   * refusal `loading`, with `command: 'save'`. A press reads the value of the latest render. When
   * it changes to false, the toolbar forgets that refusal and keeps any other refusal. The default
   * is `false`. Pass a boolean from your state, as the other props of the toolbar.
   */
  readonly opening?: boolean;
  /**
   * `false` wraps the controls onto more rows instead of moving the groups that do not fit
   * into the "More" panel. It has no effect with `preset={false}`, which always wraps.
   */
  readonly overflow?: boolean;
  /**
   * `false` renders only the children, in their order, with no measure and no "More" panel. An
   * `XlsxEditor.Toolbar.Slot` there renders the built-in control of its slot, or its children.
   * The default is `true`: the groups of `defaultChromeGroups()` (the toolbar groups of
   * `CHROME_GROUPS`, without the menu bar's), changed by the children.
   */
  readonly preset?: boolean;
  /**
   * Parts that change the preset bar. `XlsxEditor.Toolbar.Slot` replaces or hides one slot
   * where it stands, and `XlsxEditor.Toolbar.Group` adds a host group or changes a built-in
   * group. The last part for a slot or a group wins. Other children show in one block after
   * the groups, which never moves into the "More" panel.
   */
  readonly children?: ReactNode;
  /** Appended after the `xlsx-toolbar` class. */
  readonly className?: string;
  /**
   * The label resolver of the toolbar, its parts, its "More" panel, and its number format
   * dialog: a key and the values of its placeholders, as DOCX Editor's `t` prop. A key that it
   * leaves unresolved reads the root's `locale`. Keep it stable, for example from
   * `useChromeTranslate(overrides)`: a new function measures the bar again.
   */
  readonly t?: ChromePartTranslate;
  /**
   * Slot ids to leave out, such as `['file.save']`. A group whose slots are all hidden is left
   * out with its separator. The order of the list does not matter.
   */
  readonly hidden?: readonly ChromeSlotId[];
}

/** How the toolbar renders the parts of one group of the arrangement. */
interface GroupParts {
  readonly group: ArrangedToolbarGroup;
  readonly parts: ToolbarChildren;
  readonly replaced: ReadonlySet<string>;
  readonly customized: ReadonlySet<string>;
}

/** The heading of a group: the registry label of a built-in group, or a host group's own. */
function groupHeading(t: TFunction, group: ArrangedToolbarGroup): string {
  const translate = (key: string) => translateKey(t, key);
  return group.chrome ? translate(group.chrome.labelKey) : toolbarGroupHeading(group, translate);
}

/**
 * The controls of a group in `layout`: each built-in control, changed by its part, or the
 * content of the `Toolbar.Slot` that replaces it, then the content of the group's
 * `Toolbar.Group`. A menu group in the bar keeps its menu of the other controls, and the
 * replacing content and the changed controls follow it. In a row of the "More" panel, the
 * `overflowContent` of a slot takes the place of its control, and the `overflowContent` of a
 * group takes the place of its children.
 */
function GroupContent({
  group,
  parts,
  replaced,
  customized,
  layout,
}: GroupParts & { layout: PartLayout }) {
  const chrome = group.chrome;
  const row = layout === 'row';
  const controls = chrome?.controls ?? [];
  const own = chrome?.menu && layout === 'bar';
  const menu = own
    ? controls.filter((c) => {
        const slot = chromeSlotId(chrome, c);
        return !replaced.has(slot) && !customized.has(slot);
      })
    : [];
  const groupPart = parts.groups.get(group.id);
  return (
    <>
      {chrome && menu.length > 0 ? <GroupMenu group={{ ...chrome, controls: menu }} /> : null}
      {chrome
        ? controls.map((c) => {
            const slot = chromeSlotId(chrome, c);
            const part = parts.slots.get(slot);
            if (row && part?.overflowContent)
              return <Fragment key={c.id}>{part.overflowContent()}</Fragment>;
            if (replaced.has(slot)) return <Fragment key={c.id}>{part?.children}</Fragment>;
            if (own && !customized.has(slot)) return null;
            return (
              <Control
                key={c.id}
                group={chrome}
                control={c}
                layout={layout}
                replaced={replaced}
                override={customized.has(slot) && part ? slotOverrideOf(part) : null}
              />
            );
          })
        : null}
      {row && groupPart?.overflowContent ? groupPart.overflowContent() : groupPart?.children}
    </>
  );
}

/** A group in the bar: its controls, or its menu, and the host's content. */
function BarGroup(props: GroupParts) {
  const { t } = useTranslation();
  const own = props.parts.groups.get(props.group.id)?.className;
  return (
    <div
      className={`xlsx-toolbar-group${own ? ` ${own}` : ''}`}
      role="group"
      aria-label={groupHeading(t, props.group)}
      data-group={props.group.id}
    >
      <GroupContent {...props} layout="bar" />
    </div>
  );
}

const isLoading = (s: EditorSnapshot): boolean => s.status === 'loading';

/** The "More" button and its panel: the groups that the bar could not hold, as rows. */
function MorePanel({
  groups,
  parts,
  replaced,
  customized,
}: {
  readonly groups: readonly ArrangedToolbarGroup[];
  readonly parts: ToolbarChildren;
  readonly replaced: ReadonlySet<string>;
  readonly customized: ReadonlySet<string>;
}) {
  const { t } = useTranslation();
  const popup = usePopup(TOOLBAR_PANEL_ITEMS, { nested: true });
  const text = t('formattingBar.more');
  // An open of another workbook closes the panel: its rows showed the old workbook's state.
  const loading = useEditorState(isLoading);
  const { setOpen } = popup;
  useEffect(() => {
    if (loading) setOpen(false);
  }, [loading, setOpen]);
  const closePanel = useCallback(() => setOpen(false), [setOpen]);
  return (
    <div ref={popup.root} className="xlsx-toolbar-more">
      <button
        type="button"
        className="xlsx-toolbar-button"
        aria-haspopup="dialog"
        aria-expanded={popup.open}
        aria-controls={popup.open ? popup.id : undefined}
        aria-label={text}
        title={text}
        onMouseDown={guardToolbarMousedown}
        onClick={() => popup.setOpen(!popup.open)}
        {...popupTriggerKeys(popup, true)}
      >
        <ChromeIcon paths={ICON_MORE} />
      </button>
      {popup.open ? (
        <div
          id={popup.id}
          role="dialog"
          aria-label={text}
          className="xlsx-toolbar-panel"
          data-popup=""
        >
          <PanelCloseContext.Provider value={closePanel}>
            <PartLayoutContext.Provider value="row">
              {groups.map((g) => (
                <div
                  key={`group:${g.id}`}
                  role="group"
                  aria-label={groupHeading(t, g)}
                  className="xlsx-toolbar-section"
                >
                  <span className="xlsx-toolbar-heading" aria-hidden="true">
                    {groupHeading(t, g)}
                  </span>
                  <GroupContent
                    group={g}
                    parts={parts}
                    replaced={replaced}
                    customized={customized}
                    layout="row"
                  />
                </div>
              ))}
            </PartLayoutContext.Provider>
          </PanelCloseContext.Provider>
        </div>
      ) : null}
    </div>
  );
}

/** The children of a toolbar with no parts. */
const NO_PARTS: ToolbarChildren = readToolbarChildren(null);

/** No group in the "More" panel. */
const NONE: readonly string[] = [];

/** Whether two lists of ids are equal. */
const sameIds = (a: readonly string[], b: readonly string[]): boolean =>
  a.length === b.length && a.every((id, i) => id === b[i]);

/**
 * The editing toolbar: every group of the chrome registry in bar order, changed by its parts.
 * Place it inside `XlsxEditor.Viewport`, which carries the styling scope.
 */
function XlsxEditorToolbarRoot({
  onSave,
  onSaveStart,
  opening,
  overflow = true,
  className,
  hidden,
  preset = true,
  children,
}: Omit<XlsxEditorToolbarProps, 't'>) {
  const { t } = useTranslation();
  const dir = useChromeDir();
  const editor = useXlsxEditor();
  // The value of `opening` at the latest render. A press reads it here, so a press and the state
  // of each control agree. A change of the prop gives a new host, so each control reads its
  // state again.
  const openingNow = useRef(opening ?? false);
  openingNow.current = opening ?? false;
  // A slot of the menu bar in a toolbar group reads the same host: Paste reads the clipboard of
  // the editor's window. Open needs the menu bar's `onOpen`, so in a toolbar it is disabled.
  const host = useMemo<ToolbarHost>(() => {
    const read = editorClipboardReader(editor);
    return {
      ...(onSave ? { save: onSave } : {}),
      ...(onSaveStart ? { saveStart: onSaveStart } : {}),
      ...(opening === undefined ? {} : { opening: () => openingNow.current }),
      ...(read ? { readClipboard: read } : {}),
    };
  }, [onSave, onSaveStart, opening, editor]);
  // The children are new elements on each render, and a list written inline is a new array:
  // key the arrangement by the content of the requests, so the groups keep their identity.
  const parts = preset ? readToolbarChildren(children) : NO_PARTS;
  const requestKey = JSON.stringify([
    [...(hidden ?? [])].sort(),
    parts.slotRequests,
    parts.groupRequests,
  ]);
  const arrangement = useMemo<ToolbarArrangement>(() => {
    const [h, slots, groups] = JSON.parse(requestKey) as [
      ChromeSlotId[],
      ToolbarChildren['slotRequests'],
      ToolbarChildren['groupRequests'],
    ];
    return arrangeToolbar({ hidden: h, slots, groups });
  }, [requestKey]);
  useEffect(() => warnToolbarProblems(arrangement.problems), [arrangement]);
  const groups = arrangement.groups;
  // Loose children, and the content of a `Toolbar.Slot` that names no slot, never collapse:
  // they have no group, and so no heading in the panel. The bar charges them as fixed width.
  const fixed: ReactNode[] = preset
    ? [
        ...parts.appended.filter((child) => child !== ''),
        ...arrangement.outside.map((slot) => (
          <Fragment key={`slot:${slot}`}>{parts.slots.get(slot)?.children}</Fragment>
        )),
      ]
    : [];
  // The fixed block takes room from the groups: its arrival or its removal measures again.
  const hasFixed = fixed.length > 0;
  // Only the preset bar measures: `preset={false}` is the host's own markup.
  const measuring = preset && overflow;
  const bar = useRef<HTMLDivElement>(null);
  // When the open ends, the refusal of a save during the open no longer applies: forget it, also
  // when the open ends early. A later refusal of another command stays.
  const wasOpening = useRef(false);
  useEffect(() => {
    const ended = wasOpening.current && !opening;
    wasOpening.current = !!opening;
    if (!ended || !editor) return;
    const refusal = editor.snapshot().refusal;
    if (refusal?.code === 'loading' && refusal.command === 'save') editor.dismissRefusal();
  }, [opening, editor]);
  // From `onSave` alone, as DOCX Editor claims the key from its save handler: a hidden save
  // control does not stop the shortcut.
  const saves = !!onSave;

  // Control or Command and S runs the save of the save control, as DOCX Editor's menu bar does
  // (packages/react/src/editor/menu/DocxEditorMenu.tsx, key `s` with the platform modifier).
  // Scoped to this editor: only a key in the viewport that holds the bar, so a host page's own
  // field keeps the key. Read in the capture phase, as the cell editor and the grid keep their
  // keys. A save that the engine refuses (an input method composes, a workbook loads) takes
  // the same refusal path as a press of the control.
  useEffect(() => {
    const el = bar.current;
    if (!el || !editor || !saves) return undefined;
    return listenSaveKey(el, () => runToolbarCommand(editor, 'file.save', undefined, host));
  }, [editor, saves, host]);
  const widths = useRef(new Map<string, number>());
  // The ids of the groups in the "More" panel, in bar order (`toolbarCollapse`).
  const [collapsed, setCollapsed] = useState<readonly string[]>(NONE);

  // Mark the bar's computed direction for the mirrored icons (`markChromeDirection`).
  useLayoutEffect(() => markChromeDirection(bar.current), [dir]);

  // A new locale or a new set of groups changes the widths: measure every group again.
  useLayoutEffect(() => {
    widths.current.clear();
  }, [t, groups]);

  useLayoutEffect(() => {
    const el = bar.current;
    if (!el || !measuring) {
      setCollapsed(NONE);
      return undefined;
    }
    const measure = () => {
      const next = measureToolbar(el, groups, widths.current);
      setCollapsed((current) => (sameIds(current, next) ? current : next));
    };
    measure();
    // The bar and each shown group: a group grows when its fonts load or its value changes.
    return observeToolbar(el, measure);
  }, [measuring, groups, collapsed, t, hasFixed]);

  const { replaced, customized } = arrangement;
  let body: ReactNode;
  if (!preset) body = children;
  else {
    const out = new Set(collapsed);
    const rest = groups.filter((g) => out.has(g.id));
    const items: ReactNode[] = groups
      .filter((g) => !out.has(g.id))
      .map((g) => (
        <BarGroup
          key={`group:${g.id}`}
          group={g}
          parts={parts}
          replaced={replaced}
          customized={customized}
        />
      ));
    if (rest.length > 0)
      items.push(
        <MorePanel
          key="more"
          groups={rest}
          parts={parts}
          replaced={replaced}
          customized={customized}
        />
      );
    if (hasFixed)
      items.push(
        <div key="fixed" className="xlsx-toolbar-group" {...{ [TOOLBAR_FIXED_ATTRIBUTE]: '' }}>
          {fixed}
        </div>
      );
    body = items;
  }
  return (
    <ToolbarHostContext.Provider value={host}>
      <NumberFormatDialogHost>
        <div
          ref={bar}
          role="toolbar"
          dir={dir}
          aria-label={t('formattingBar.label')}
          className={`xlsx-toolbar${className ? ` ${className}` : ''}`}
          {...(measuring ? { 'data-overflow': '' } : {})}
          // The bar guards every press on it, also between its controls.
          onMouseDown={guardToolbarMousedown}
          onKeyDown={moveBarFocus}
        >
          {body}
        </div>
      </NumberFormatDialogHost>
    </ToolbarHostContext.Provider>
  );
}

/**
 * The toolbar with its parts as statics: `Slot`, `Group`, `Action`, `Separator`, and one named
 * part for each slot (`ToolbarNamedParts`). @public
 */
export interface XlsxEditorToolbarNamespace extends ToolbarNamedParts {
  /** The toolbar. */
  (props: XlsxEditorToolbarProps): ReactElement;
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
 * `XlsxEditor.Viewport`, which carries the styling scope. Its parts change the bar: the named
 * parts (`XlsxEditor.Toolbar.Bold`, ...), `XlsxEditor.Toolbar.Slot`, `XlsxEditor.Toolbar.Group`,
 * `XlsxEditor.Toolbar.Action`, and `XlsxEditor.Toolbar.Separator`.
 *
 * @example
 * ```tsx
 * <XlsxEditor.Toolbar onSave={upload}>
 *   <XlsxEditor.Toolbar.Underline hidden />
 *   <XlsxEditor.Toolbar.Bold className="my-bold" />
 *   <XlsxEditor.Toolbar.Group id="share" label="Share" priority={85}>
 *     <XlsxEditor.Toolbar.Action label="Share" icon={<ShareIcon />} onSelect={share} />
 *   </XlsxEditor.Toolbar.Group>
 * </XlsxEditor.Toolbar>
 * ```
 *
 * @public
 */
export const XlsxEditorToolbar: XlsxEditorToolbarNamespace = Object.assign(
  withChromeTranslate<XlsxEditorToolbarProps>(XlsxEditorToolbarRoot, 'XlsxEditorToolbar'),
  TOOLBAR_PARTS,
  {
    Slot: XlsxEditorToolbarSlot,
    Group: XlsxEditorToolbarGroup,
    Action: ToolbarAction,
    Separator: ToolbarSeparator,
  }
);
