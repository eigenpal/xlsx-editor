// The chrome registry: the toolbar as data, DOM-free. After
// docx-editor/packages/core/src/editor/chrome-controls.ts at a0825f1f (EigenPal): the
// `ChromeControl` and `ChromeGroup` shapes, the `${groupId}.${controlId}` slot ids, the
// `CHROME_GROUPS` taxonomy that both adapters walk, and the rule that a control never says
// whether it is enabled. Differences: the groups and controls of the editing commands of a
// workbook; the shapes are `icon`, `dropdown` (a list of values), `color` (a palette), and
// `dialog` (a button that opens a dialog, which produces the value); a group can render as one
// `menu` of its slots; and the label keys name the shared catalog directly. The number group is
// `number.format`, `number.custom`, then `number.increaseDecimals` before
// `number.decreaseDecimals`, as the command that adds comes before the command that removes
// (insert before delete in the cells group). DOCX Editor puts its `list.outdent` slot before
// `list.indent`; this group keeps the add-first order of its own cells group. The
// order of a group is its order on the bar and in the "More" panel.
//
// SLOT IDS ARE PUBLIC API. A host places, replaces, or hides chrome by slot id, and
// `toolbarCommandState` and `runToolbarCommand` resolve a slot id to the engine. Renaming a
// group or a control is a breaking change. Control ids are short lowerCamel words that never
// repeat their group's name (`font.bold`, not `font.fontBold`).
//
// Enabled state has exactly one source: `toolbarCommandState` (`toolbar-commands.ts`), which
// reads the editor snapshot. A registry entry carries no enabled or disabled flag.

import { ICON_PATHS } from './icon-paths.ts';

/**
 * How a control reaches the engine, never whether it is enabled. `command`: a press runs one
 * fixed command. `value`: a press needs a picked value (a font, a size, a color, a number
 * format). `save`: the host receives the workbook bytes. `open`: the host opens a workbook
 * (`ToolbarHost.open`).
 *
 * @public
 */
export type ChromeControlKind = 'command' | 'value' | 'save' | 'open';

/**
 * How a control renders: an icon button, a dropdown list of values, a color palette, or a
 * button that opens a dialog. A `dialog` control is a `value` control: the dialog produces the
 * value, and its apply runs the slot with it. DOCX Editor's `paragraph.dialog` is a `command`
 * slot with no fixed command; a `value` slot here says that a bare press has nothing to send.
 * DOCX Editor draws that slot with the `icon` shape. Here `dialog` tells a host that walks
 * the registry to render a dialog opener (`aria-haspopup="dialog"`), not a command button.
 *
 * @public
 */
export type ChromeControlShape = 'icon' | 'dropdown' | 'color' | 'dialog';

/**
 * One toolbar control: what it renders as and how it reaches the engine.
 *
 * @public
 */
export interface ChromeControl<Id extends string = string> {
  /** Unique within its group. Public API: renames are breaking. */
  readonly id: Id;
  /** `icon` when absent. */
  readonly shape?: ChromeControlShape;
  /** The catalog key of the accessible name and the tooltip. */
  readonly labelKey: string;
  /** Material Symbols path data (view box `0 -960 960 960`), or null for a picker or a menu item. */
  readonly paths: readonly string[] | null;
  readonly kind: ChromeControlKind;
  /** The control shows a pressed state (`aria-pressed`). */
  readonly toggle?: true;
  /**
   * `false` keeps this public slot composable but leaves it out of the preset toolbar, as DOCX
   * Editor's flag of the same name. The menu bar places it.
   */
  readonly defaultToolbar?: false;
  /**
   * A wide field. In a bar that is too short, it leaves for the "More" panel before its group
   * does, and the rest of its group stays in the bar.
   */
  readonly wide?: true;
}

/**
 * One toolbar group, in bar order. The adapters build the default toolbar by walking
 * `CHROME_GROUPS`, never by listing controls by hand.
 *
 * @public
 */
export interface ChromeGroup<Id extends string = string, ControlId extends string = string> {
  /** Public API: renames are breaking. */
  readonly id: Id;
  /** The catalog key of the group's name. */
  readonly labelKey: string;
  /**
   * The default toolbar renders the group as one menu button with this icon, and its controls
   * as the menu's items. The slots stay public: a host can place each one alone.
   */
  readonly menu?: { readonly paths: readonly string[] };
  /**
   * The group is not part of the preset toolbar, as DOCX Editor's flag of the same name. Its
   * slots stay public: the menu bar places them, and a host can place each one alone.
   */
  readonly contextual?: true;
  readonly controls: readonly ChromeControl<ControlId>[];
}

const icon = (name: keyof typeof ICON_PATHS): readonly string[] => ICON_PATHS[name];

/**
 * The editing chrome, in bar order. Literal-typed, so `ChromeSlotId` is derived from the data
 * and cannot drift from it.
 *
 * @public
 */
export const CHROME_GROUPS = [
  {
    id: 'history',
    labelKey: 'formattingBar.groups.history',
    controls: [
      { id: 'undo', labelKey: 'formattingBar.undo', paths: icon('undo'), kind: 'command' },
      { id: 'redo', labelKey: 'formattingBar.redo', paths: icon('redo'), kind: 'command' },
    ],
  },
  {
    id: 'font',
    labelKey: 'formattingBar.groups.font',
    controls: [
      {
        id: 'family',
        shape: 'dropdown',
        labelKey: 'font.selectAriaLabel',
        paths: null,
        kind: 'value',
        wide: true,
      },
      { id: 'size', shape: 'dropdown', labelKey: 'fontSize.label', paths: null, kind: 'value' },
      {
        id: 'bold',
        labelKey: 'formattingBar.bold',
        paths: icon('format_bold'),
        kind: 'command',
        toggle: true,
      },
      {
        id: 'italic',
        labelKey: 'formattingBar.italic',
        paths: icon('format_italic'),
        kind: 'command',
        toggle: true,
      },
      {
        id: 'underline',
        labelKey: 'formattingBar.underline',
        paths: icon('format_underlined'),
        kind: 'command',
        toggle: true,
      },
      {
        id: 'color',
        shape: 'color',
        labelKey: 'formattingBar.fontColor',
        paths: icon('format_color_text'),
        kind: 'value',
      },
    ],
  },
  {
    id: 'fill',
    labelKey: 'formattingBar.groups.fill',
    controls: [
      {
        id: 'color',
        shape: 'color',
        labelKey: 'formattingBar.fillColor',
        paths: icon('format_color_fill'),
        kind: 'value',
      },
    ],
  },
  {
    id: 'borders',
    labelKey: 'formattingBar.groups.borders',
    controls: [
      {
        id: 'outline',
        labelKey: 'borders.outline',
        paths: icon('border_outer'),
        kind: 'command',
      },
      { id: 'all', labelKey: 'borders.all', paths: icon('border_all'), kind: 'command' },
      { id: 'none', labelKey: 'borders.none', paths: icon('border_clear'), kind: 'command' },
    ],
  },
  {
    id: 'alignment',
    labelKey: 'formattingBar.groups.alignment',
    controls: [
      {
        id: 'left',
        labelKey: 'alignment.alignLeft',
        paths: icon('format_align_left'),
        kind: 'command',
        toggle: true,
      },
      {
        id: 'center',
        labelKey: 'alignment.center',
        paths: icon('format_align_center'),
        kind: 'command',
        toggle: true,
      },
      {
        id: 'right',
        labelKey: 'alignment.alignRight',
        paths: icon('format_align_right'),
        kind: 'command',
        toggle: true,
      },
      {
        id: 'top',
        labelKey: 'alignment.alignTop',
        paths: icon('vertical_align_top'),
        kind: 'command',
        toggle: true,
      },
      {
        id: 'middle',
        labelKey: 'alignment.alignMiddle',
        paths: icon('vertical_align_center'),
        kind: 'command',
        toggle: true,
      },
      {
        id: 'bottom',
        labelKey: 'alignment.alignBottom',
        paths: icon('vertical_align_bottom'),
        kind: 'command',
        toggle: true,
      },
      {
        id: 'wrap',
        labelKey: 'alignment.wrapText',
        paths: icon('wrap_text'),
        kind: 'command',
        toggle: true,
      },
    ],
  },
  {
    id: 'number',
    labelKey: 'formattingBar.groups.number',
    controls: [
      {
        id: 'format',
        shape: 'dropdown',
        labelKey: 'numberFormat.label',
        paths: null,
        kind: 'value',
        wide: true,
      },
      {
        id: 'custom',
        shape: 'dialog',
        labelKey: 'numberFormat.more',
        paths: icon('123'),
        kind: 'value',
      },
      {
        id: 'increaseDecimals',
        labelKey: 'numberFormat.increaseDecimals',
        paths: icon('decimal_increase'),
        kind: 'command',
      },
      {
        id: 'decreaseDecimals',
        labelKey: 'numberFormat.decreaseDecimals',
        paths: icon('decimal_decrease'),
        kind: 'command',
      },
    ],
  },
  {
    id: 'cells',
    labelKey: 'formattingBar.groups.cells',
    menu: { paths: icon('add_row_above') },
    controls: [
      {
        id: 'insertRows',
        labelKey: 'cells.insertRows',
        paths: icon('add_row_above'),
        kind: 'command',
      },
      {
        id: 'insertColumns',
        labelKey: 'cells.insertColumns',
        paths: icon('add_column_left'),
        kind: 'command',
      },
      { id: 'deleteRows', labelKey: 'cells.deleteRows', paths: icon('delete'), kind: 'command' },
      {
        id: 'deleteColumns',
        labelKey: 'cells.deleteColumns',
        paths: icon('delete'),
        kind: 'command',
      },
    ],
  },
  {
    id: 'clear',
    labelKey: 'formattingBar.groups.clear',
    menu: { paths: icon('delete_sweep') },
    controls: [
      { id: 'contents', labelKey: 'clear.contents', paths: icon('backspace'), kind: 'command' },
      { id: 'formats', labelKey: 'clear.formats', paths: icon('format_clear'), kind: 'command' },
      { id: 'all', labelKey: 'clear.all', paths: icon('ink_eraser'), kind: 'command' },
    ],
  },
  {
    id: 'file',
    labelKey: 'formattingBar.groups.file',
    controls: [
      {
        id: 'open',
        labelKey: 'menu.open',
        paths: icon('folder_open'),
        kind: 'open',
        defaultToolbar: false,
      },
      { id: 'save', labelKey: 'toolbar.save', paths: icon('file_download'), kind: 'save' },
    ],
  },
  {
    id: 'edit',
    labelKey: 'formattingBar.groups.edit',
    contextual: true,
    controls: [
      { id: 'cut', labelKey: 'menu.cut', paths: icon('content_cut'), kind: 'command' },
      { id: 'copy', labelKey: 'menu.copy', paths: icon('content_copy'), kind: 'command' },
      { id: 'paste', labelKey: 'menu.paste', paths: icon('content_paste'), kind: 'command' },
      { id: 'selectAll', labelKey: 'menu.selectAll', paths: null, kind: 'command' },
    ],
  },
  {
    id: 'view',
    labelKey: 'formattingBar.groups.view',
    contextual: true,
    controls: [
      { id: 'zoomIn', labelKey: 'zoom.zoomIn', paths: null, kind: 'command' },
      { id: 'zoomOut', labelKey: 'zoom.zoomOut', paths: null, kind: 'command' },
      { id: 'zoomReset', labelKey: 'menu.zoomReset', paths: null, kind: 'command' },
      // `x:sheetView/@showRowColHeaders` (`CT_SheetView`, ECMA-376 Part 1, 18.3.1.87).
      { id: 'headings', labelKey: 'menu.headings', paths: null, kind: 'command', toggle: true },
    ],
  },
  {
    id: 'sheet',
    labelKey: 'formattingBar.groups.sheet',
    contextual: true,
    controls: [{ id: 'add', labelKey: 'menu.addSheet', paths: icon('add'), kind: 'command' }],
  },
  {
    id: 'merge',
    labelKey: 'formattingBar.groups.merge',
    contextual: true,
    controls: [
      { id: 'merge', labelKey: 'merge.merge', paths: null, kind: 'command' },
      { id: 'mergeCenter', labelKey: 'merge.mergeCenter', paths: null, kind: 'command' },
      { id: 'unmerge', labelKey: 'merge.unmerge', paths: null, kind: 'command' },
    ],
  },
] as const satisfies readonly ChromeGroup[];

/**
 * A group id of `CHROME_GROUPS`.
 *
 * @public
 */
export type ChromeGroupId = (typeof CHROME_GROUPS)[number]['id'];

/**
 * The slot ids of one group of `CHROME_GROUPS`: `${groupId}.${controlId}` for each of its
 * controls. {@link ChromeSlotId} applies it to every group.
 *
 * @public
 */
export type SlotOf<G> = G extends {
  readonly id: infer GroupId extends string;
  readonly controls: readonly (infer C)[];
}
  ? C extends { readonly id: infer ControlId extends string }
    ? `${GroupId}.${ControlId}`
    : never
  : never;

/**
 * The public slot vocabulary: `${groupId}.${controlId}` for every control of
 * `CHROME_GROUPS`, such as `font.bold` or `cells.deleteRows`. Renaming a slot is a breaking
 * change.
 *
 * @public
 */
export type ChromeSlotId = SlotOf<(typeof CHROME_GROUPS)[number]>;

/**
 * The slot id of a control of a group.
 *
 * @public
 */
export function chromeSlotId(
  group: { readonly id: string },
  control: { readonly id: string }
): ChromeSlotId {
  return `${group.id}.${control.id}` as ChromeSlotId;
}

/** Every slot of the registry, in bar order. @public */
export const CHROME_SLOT_IDS: readonly ChromeSlotId[] = (
  CHROME_GROUPS as readonly ChromeGroup[]
).flatMap((g) => g.controls.map((c) => chromeSlotId(g, c)));

/**
 * The groups of the preset toolbar, in bar order: the groups of `CHROME_GROUPS` without the
 * `contextual` ones, each without its `defaultToolbar: false` controls. After DOCX Editor's
 * function of the same name.
 *
 * @public
 */
export function defaultChromeGroups(): readonly ChromeGroup[] {
  return (CHROME_GROUPS as readonly ChromeGroup[])
    .filter((g) => !g.contextual)
    .map((g) =>
      g.controls.some((c) => c.defaultToolbar === false)
        ? { ...g, controls: g.controls.filter((c) => c.defaultToolbar !== false) }
        : g
    );
}

/** The registry control of a slot, or null for an id that the registry does not hold. @public */
export function chromeControlForSlot(slot: string): ChromeControl | null {
  for (const group of CHROME_GROUPS as readonly ChromeGroup[])
    for (const control of group.controls) if (`${group.id}.${control.id}` === slot) return control;
  return null;
}
