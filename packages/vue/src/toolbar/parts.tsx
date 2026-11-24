/** @jsxImportSource vue */
// The toolbar parts: an icon button, a dropdown list of values, a color palette, and a group
// menu. The Vue twin of the React adapter's `toolbar/parts.tsx`; the header there explains the
// rules. Each part renders one slot of the chrome registry through `useToolbarSlot`, and each
// popup closes through the one close rule, `useDropdownClose`. Chrome `mousedown` keeps the
// keyboard on the grid, except in a form field (AGENTS.md "Conventions", focus stealing). A
// pick returns the keyboard to the grid and leaves the selection where it was. A disabled
// control keeps its focus and carries `aria-disabled` and its reason as the description.
//
// Differences from React: `usePopup` returns refs; the trigger and the reason text are render
// functions of the part that owns them, not components; and a style value carries its unit.

import {
  computed,
  defineComponent,
  inject,
  shallowRef,
  useId,
  watch,
  type InjectionKey,
  type PropType,
  type VNode,
  type VNodeChild,
} from 'vue';
import {
  AUTO_COLOR,
  CLEAR_CHIP_HEX,
  FONT_FAMILY_CHOICES,
  FONT_SIZE_CHOICES,
  NO_FILL,
  guardToolbarMousedown,
  isHexColor,
  colorBarTone,
  colorPickLabel,
  colorPickerGrid,
  customHexDraft,
  focusColorSwatch,
  hexFromText,
  moveColorSwatchFocus,
  movePopupFocus,
  workbookThemeHexes,
  type ColorPick,
  numberFormatListView,
  parseFontSize,
  popupTriggerAction,
  type ChromeGroup,
  type ChromeSlotId,
  type EditorSnapshot,
} from '@xlsx-editor.dev/core/editor';
import type { TFunction } from '@xlsx-editor.dev/i18n';
import { useEditorState, useTranslation } from '../context.ts';
import {
  ChromeIcon,
  translateKey,
  useRowsState,
  useToolbarSlot,
  type ToolbarSlot,
} from './slot.tsx';
import { popupTriggerKeys, usePopup, type Popup } from './popup.ts';
import { useNumberFormatDialogOpener } from './NumberFormatDialog.tsx';
import { ICON_CHECK, ICON_FIELD_CARET, ICON_MENU_CARET } from './icons.ts';
import { renderAsChild } from './as-child.tsx';

/** How a part renders: in the bar, or as a full-width row of the overflow panel. */
export type PartLayout = 'bar' | 'row';

const layoutProp = { type: String as PropType<PartLayout>, default: 'bar' } as const;

/**
 * The changes that a `Toolbar.Slot` or a named part makes to the built-in control of its slot,
 * the twin of React's `SlotOverride`: an icon in place of the registry icon and the host element
 * that renders the control (button slots only), and a class name for the control or the
 * trigger of a value control.
 */
export interface SlotOverride {
  readonly slot: string;
  readonly icon?: VNodeChild;
  readonly className?: string;
  readonly asChild?: VNode | null;
}

/**
 * The override of the control that renders inside it, as a getter, so a control reads the
 * override of the latest render.
 */
export const slotOverrideKey: InjectionKey<() => SlotOverride | null> = Symbol(
  'xlsx-toolbar-slot-override'
);

/**
 * The override of `slot` at the call, or null: an override of another slot never reaches this
 * control. Call it in `setup`, and the getter in the render.
 */
function useSlotOverride(slot: () => string): () => SlotOverride | null {
  const read = inject(slotOverrideKey, () => null);
  return () => {
    const override = read();
    return override?.slot === slot() ? override : null;
  };
}

/** A base class with the override's class name after it. */
function withClass(base: string, override: SlotOverride | null): string {
  return override?.className ? `${base} ${override.className}` : base;
}

/**
 * The attributes that every slot control carries. A disabled control stays focusable
 * (`aria-disabled`), so the keyboard reaches it and its reason, and arrow navigation never
 * loses a control that becomes disabled under focus. Its press does nothing. The reason is the
 * tooltip and the accessible description (`aria-describedby` on a hidden reason text).
 */
function slotAttributes(s: ToolbarSlot, reasonId: string) {
  const words = s.words.value;
  return {
    'data-slot': s.slot,
    'aria-disabled': s.isEnabled.value ? undefined : ('true' as const),
    ...(words.description ? { 'aria-describedby': reasonId } : {}),
    title: words.title,
    onMousedown: guardToolbarMousedown,
    ...(s.isActive.value ? { 'data-active': '' } : {}),
    ...(s.isEnabled.value ? {} : { 'data-disabled': '' }),
  };
}

/** The disabled reason or the hint of a control, as its hidden accessible description. */
function reasonText(id: string, reason: string | null) {
  return reason ? (
    <span id={id} hidden>
      {reason}
    </span>
  ) : null;
}

/** One command slot as a button. In a row, the label shows beside the icon. */
export const ToolbarButton = defineComponent({
  name: 'XlsxToolbarButton',
  props: {
    slot: { type: String as PropType<ChromeSlotId>, required: true },
    layout: layoutProp,
  },
  setup(props) {
    const s = useToolbarSlot(props.slot);
    const reasonId = useId();
    const overrideOf = useSlotOverride(() => props.slot);
    return () => {
      const words = s.words.value;
      const override = overrideOf();
      const shared = {
        class: withClass(
          props.layout === 'row' ? 'xlsx-toolbar-row' : 'xlsx-toolbar-button',
          override
        ),
        'aria-label': words.label,
        ...(s.control.toggle ? { 'aria-pressed': s.isActive.value } : {}),
        ...slotAttributes(s, reasonId),
        // A held Enter repeats the click of a button; a command such as a decimal step would
        // run on each repeat, so only the first press runs it, as in React.
        onKeydown: (e: KeyboardEvent) => {
          if (e.key === 'Enter' && e.repeat) e.preventDefault();
        },
        // A disabled control also stops the default of its element, so an `asChild` link does
        // not navigate.
        onClick: (e: MouseEvent) => {
          if (s.isEnabled.value) s.execute();
          else e.preventDefault();
        },
      };
      const reason = reasonText(reasonId, words.description);
      // `asChild`: the host element renders the control, with its state and press merged in.
      if (override?.asChild) return [renderAsChild(override.asChild, shared), reason];
      const icon = override?.icon ?? <ChromeIcon paths={s.control.paths} />;
      // A slot with no icon (a menu item of the registry) shows its label in the bar too.
      const text = props.layout === 'row' || (override?.icon == null && !s.control.paths);
      return (
        <button type="button" {...shared}>
          {icon}
          {text ? <span>{words.label}</span> : null}
          {reason}
        </button>
      );
    };
  },
});

/** The trigger of a popup: a button that opens it with a press, Enter, Space, or Down. */
function trigger(
  t: TFunction,
  options: {
    readonly s: ToolbarSlot;
    readonly popup: Popup;
    readonly kind: 'listbox' | 'menu' | 'dialog';
    readonly className: string;
    /** The caret: a field's (`expand_more`) or a menu button's (`arrow_drop_down`). */
    readonly caret: 'field' | 'menu';
    readonly name: string;
    readonly reasonId: string;
    /**
     * The text of the title in place of the slot label, such as the label and the value of a
     * list. The disabled reason or the hint follows on its own line, so the title never has two
     * colons, as in "Number format: General: <reason>".
     */
    readonly label?: string;
    /** The changes of the slot's part, for the class name of the trigger. */
    readonly override?: SlotOverride | null;
  },
  children: VNodeChild
) {
  const { s, popup, kind, caret, name, reasonId, label } = options;
  const className = withClass(options.className, options.override ?? null);
  const attributes = slotAttributes(s, reasonId);
  const after = s.words.value.description;
  const title =
    label === undefined
      ? attributes.title
      : after
        ? t('toolbar.hintTitle', { label, hint: after })
        : label;
  return (
    <button
      type="button"
      class={className}
      aria-haspopup={kind}
      aria-expanded={popup.open.value}
      aria-controls={popup.open.value ? popup.id : undefined}
      aria-label={name}
      {...attributes}
      title={title}
      onClick={() => {
        if (s.isEnabled.value) popup.setOpen(!popup.open.value);
      }}
      {...popupTriggerKeys(popup, s.isEnabled.value)}
    >
      {children}
      <ChromeIcon
        paths={caret === 'field' ? ICON_FIELD_CARET : ICON_MENU_CARET}
        className="xlsx-toolbar-caret"
      />
      {reasonText(reasonId, after)}
    </button>
  );
}

/** The class of a popup, and the attributes that keep one in flow inside the overflow panel. */
function popupAttributes(layout: PartLayout, className: string) {
  return {
    class: `xlsx-toolbar-popup ${className}`,
    'data-popup': '',
    ...(layout === 'row' ? { 'data-inline': '' } : {}),
  };
}

/** A typed value of a list: a font size from 1 to 409 points. */
export interface ListEntry {
  readonly label: string;
  /** The message that shows under the entry when the text is not a value. */
  readonly invalid: string;
  /** The value to apply, or null when the text is not one. */
  parse(text: string): string | null;
}

/** A list of values for a value slot: font family, font size, or number format. */
const ToolbarList = defineComponent({
  name: 'XlsxToolbarList',
  props: {
    slot: { type: String as PropType<ChromeSlotId>, required: true },
    options: {
      type: Array as PropType<readonly { readonly value: string; readonly label: string }[]>,
      required: true,
    },
    text: { type: String, required: true },
    /** What the trigger's accessible name and title say after the slot label, when not `text`. */
    name: { type: String, default: undefined },
    /** The current entry (`aria-current`), which the keyboard focuses on open when none is picked. */
    active: { type: String, default: undefined },
    layout: { type: String as PropType<PartLayout>, required: true },
    width: { type: String as PropType<'wide' | 'narrow'>, required: true },
    entry: { type: Object as PropType<ListEntry>, default: undefined },
    /** A last item after the list, such as the number format dialog's opener. */
    more: { type: Function as PropType<(close: () => void) => VNodeChild>, default: undefined },
  },
  setup(props) {
    const { t } = useTranslation();
    const s = useToolbarSlot(props.slot);
    const items = `${props.entry ? 'input, ' : ''}[role="option"]${props.more ? ', [data-list-more]' : ''}`;
    const popup = usePopup(items, {
      inline: props.layout === 'row',
      closed: () => !s.isEnabled.value,
    });
    const invalid = shallowRef(false);
    const reasonId = useId();
    const overrideOf = useSlotOverride(() => props.slot);
    const pick = (value: string): void => {
      popup.setOpen(false);
      s.execute(value);
      s.focusGrid();
    };
    return () => {
      const { layout, entry, text, name } = props;
      const label = s.words.value.label;
      return (
        <div
          ref={popup.root}
          class={`xlsx-toolbar-dropdown${layout === 'row' ? ' xlsx-toolbar-labelled' : ''}`}
        >
          {layout === 'row' ? <span class="xlsx-toolbar-row-label">{label}</span> : null}
          {trigger(
            t,
            {
              s,
              popup,
              kind: 'listbox',
              className: `xlsx-toolbar-select xlsx-toolbar-select-${props.width}`,
              override: overrideOf(),
              caret: 'field',
              name: `${label}: ${name ?? text}`,
              reasonId,
              ...(name !== undefined ? { label: `${label}: ${name}` } : {}),
            },
            <span class="xlsx-toolbar-select-text">{text}</span>
          )}
          {popup.open.value ? (
            <div
              {...popupAttributes(layout, 'xlsx-toolbar-list-popup')}
              onKeydown={(e: KeyboardEvent) => movePopupFocus(e, items)}
            >
              {entry ? (
                <input
                  type="text"
                  inputmode="decimal"
                  autocomplete="off"
                  class="xlsx-toolbar-entry"
                  aria-label={entry.label}
                  aria-invalid={invalid.value}
                  {...(invalid.value ? { 'aria-describedby': `${popup.id}-entry-error` } : {})}
                  placeholder={text}
                  onInput={() => (invalid.value = false)}
                  onKeydown={(e: KeyboardEvent) => {
                    if (e.key !== 'Enter') return;
                    e.preventDefault();
                    const value = entry.parse((e.currentTarget as HTMLInputElement).value);
                    if (value === null) invalid.value = true;
                    else pick(value);
                  }}
                />
              ) : null}
              {entry && invalid.value ? (
                <div id={`${popup.id}-entry-error`} role="alert" class="xlsx-toolbar-entry-error">
                  {entry.invalid}
                </div>
              ) : null}
              <ul id={popup.id} role="listbox" aria-label={label} class="xlsx-toolbar-list">
                {props.options.map((o) => (
                  <li
                    key={o.value}
                    role="option"
                    tabindex={-1}
                    aria-selected={o.value === s.value.value}
                    {...(o.value === props.active ? { 'aria-current': 'true' } : {})}
                    onMousedown={guardToolbarMousedown}
                    onClick={() => pick(o.value)}
                    onKeydown={(e: KeyboardEvent) => {
                      if (e.key !== 'Enter' && e.key !== ' ') return;
                      e.preventDefault();
                      pick(o.value);
                    }}
                  >
                    {o.label}
                    {o.value === s.value.value ? (
                      <ChromeIcon paths={ICON_CHECK} className="xlsx-toolbar-check" />
                    ) : null}
                  </li>
                ))}
              </ul>
              {props.more?.(() => popup.setOpen(false))}
            </div>
          ) : null}
        </div>
      );
    };
  },
});

/** The font family list: the choices, and the active cell's own family when it is not one. */
export const FontFamilyList = defineComponent({
  name: 'XlsxFontFamilyList',
  props: { layout: layoutProp },
  setup(props) {
    const s = useToolbarSlot('font.family');
    return () => {
      const current = s.value.value ?? '';
      const names = current && !FONT_FAMILY_CHOICES.includes(current) ? [current] : [];
      const options = [...names, ...FONT_FAMILY_CHOICES].map((n) => ({ value: n, label: n }));
      return (
        <ToolbarList
          slot="font.family"
          options={options}
          text={current}
          layout={props.layout}
          width="wide"
        />
      );
    };
  },
});

/** The font size list, in points, with the active cell's own size when it is not a choice. */
export const FontSizeList = defineComponent({
  name: 'XlsxFontSizeList',
  props: { layout: layoutProp },
  setup(props) {
    const { t } = useTranslation();
    const s = useToolbarSlot('font.size');
    return () => {
      const current = Number(s.value.value);
      const sizes = [...FONT_SIZE_CHOICES];
      if (Number.isFinite(current) && current > 0 && !sizes.includes(current)) sizes.push(current);
      sizes.sort((a, b) => a - b);
      const options = sizes.map((n) => ({ value: String(n), label: String(n) }));
      return (
        <ToolbarList
          slot="font.size"
          options={options}
          text={s.value.value ?? ''}
          layout={props.layout}
          width="narrow"
          entry={{
            label: t('fontSize.typeSize'),
            invalid: t('fontSize.invalid'),
            parse: parseFontSize,
          }}
        />
      );
    };
  },
});

const numberLocaleOf = (snapshot: EditorSnapshot) => snapshot.numberLocale;
const themeColorsOf = (snapshot: EditorSnapshot) => snapshot.themeColors;

/**
 * The number format list: the built-in choices, and the code of any other format. A format
 * that is no choice has the value `CUSTOM_NUMBER_FORMAT_PREFIX` and its code, so its code never
 * shows as a choice.
 */
export const NumberFormatList = defineComponent({
  name: 'XlsxNumberFormatList',
  props: {
    layout: layoutProp,
    /** Show the number format dialog's opener (`number.custom`) after the list. */
    more: { type: Boolean, default: false },
  },
  setup(props) {
    const { t } = useTranslation();
    const s = useToolbarSlot('number.format');
    const locale = useEditorState(numberLocaleOf);
    // The engine names the entries and the closed list (`numberFormatListView`). An entry has
    // no title, so a screen reader does not read a format code on each arrow key; the number
    // format dialog shows the codes.
    const view = computed(() =>
      numberFormatListView(
        s.value.value ?? undefined,
        (key, params) => translateKey(t, key, params),
        locale.value
      )
    );
    const moreItem = (close: () => void) => <MoreFormatsItem close={close} />;
    return () => {
      const v = view.value;
      return (
        <ToolbarList
          slot="number.format"
          options={v.entries}
          text={v.text}
          {...(v.name ? { name: v.name } : {})}
          {...(v.current !== undefined ? { active: v.current } : {})}
          layout={props.layout}
          width="wide"
          {...(props.more ? { more: moreItem } : {})}
        />
      );
    };
  },
});

/**
 * The last item of the number format list: it opens the number format dialog. Focus goes back
 * to the list's trigger on a cancel when the keyboard opened the dialog, and else stays where
 * it was (the grid), as a press of a toolbar control keeps it there.
 */
const MoreFormatsItem = defineComponent({
  name: 'XlsxMoreFormatsItem',
  props: { close: { type: Function as PropType<() => void>, required: true } },
  setup(props) {
    const s = useToolbarSlot('number.custom');
    const dialog = useNumberFormatDialogOpener();
    const reasonId = useId();
    return () => {
      // With no dialog host, the opener renders nothing, after DOCX Editor's dialog hosts.
      if (!dialog) return null;
      const words = s.words.value;
      return (
        <button
          type="button"
          class="xlsx-toolbar-list-more"
          tabindex={-1}
          data-list-more=""
          aria-haspopup="dialog"
          {...slotAttributes(s, reasonId)}
          onClick={(e: MouseEvent) => {
            if (!s.isEnabled.value) return;
            const item = e.currentTarget as HTMLElement;
            const root = item.closest('.xlsx-toolbar-dropdown');
            const fromKeyboard = item.ownerDocument.activeElement === item;
            const back = root?.querySelector<HTMLElement>('[aria-haspopup="listbox"]') ?? null;
            props.close();
            // A refused start opens no dialog, and the list with this item closes: the keyboard
            // goes to the grid, as a refused choice of a list option does.
            if (!dialog.open(fromKeyboard ? back : null) && fromKeyboard) s.focusGrid();
          }}
        >
          {words.label}
          {reasonText(reasonId, words.description)}
        </button>
      );
    };
  },
});

/**
 * The `number.custom` slot alone, when the bar has no number format list to hold it: an icon
 * button that opens the number format dialog.
 */
export const NumberFormatDialogButton = defineComponent({
  name: 'XlsxNumberFormatDialogButton',
  props: { layout: layoutProp },
  setup(props) {
    const s = useToolbarSlot('number.custom');
    const dialog = useNumberFormatDialogOpener();
    const reasonId = useId();
    const overrideOf = useSlotOverride(() => 'number.custom');
    return () => {
      if (!dialog) return null;
      const words = s.words.value;
      return (
        <button
          type="button"
          class={withClass(
            props.layout === 'row' ? 'xlsx-toolbar-row' : 'xlsx-toolbar-button',
            overrideOf()
          )}
          aria-label={words.label}
          aria-haspopup="dialog"
          {...slotAttributes(s, reasonId)}
          onClick={(e: MouseEvent) => {
            if (!s.isEnabled.value) return;
            const self = e.currentTarget as HTMLElement;
            dialog.open(self.ownerDocument.activeElement === self ? self : null);
          }}
        >
          <ChromeIcon paths={s.control.paths} />
          {props.layout === 'row' ? <span>{words.label}</span> : null}
          {reasonText(reasonId, words.description)}
        </button>
      );
    };
  },
});

/**
 * The font color or the fill color as a split button, the twin of React's `ColorPalette` and of
 * docx-editor/packages/vue/src/editor/toolbar/ColorSplit.tsx at 60a86b48 (EigenPal). The main
 * half applies the last color that it applied, which its bar shows. The caret opens the picker:
 * Automatic (or No fill), the workbook theme colors with five rows of tints and shades, the
 * standard colors, and a custom hex field, from core (`color-picker.ts`). A swatch paints six
 * hex digits that core checked, so no inline style carries other text.
 */
export const ColorPalette = defineComponent({
  name: 'XlsxColorPalette',
  props: {
    slot: { type: String as PropType<'font.color' | 'fill.color'>, required: true },
    layout: layoutProp,
  },
  setup(props) {
    const { t } = useTranslation();
    const themeColors = useEditorState(themeColorsOf);
    const s = useToolbarSlot(props.slot);
    const popup = usePopup('[data-color]', {
      inline: props.layout === 'row',
      closed: () => !s.isEnabled.value,
    });
    const reasonId = useId();
    const headingId = useId();
    const overrideOf = useSlotOverride(() => props.slot);
    const fill = props.slot === 'fill.color';
    const clear = fill ? NO_FILL : AUTO_COLOR;
    const last = shallowRef(fill ? 'FFFF00' : 'FF0000');
    const draft = shallowRef('');
    let keyboardOpen = false;
    const popupEl = (): HTMLElement | null =>
      popup.root.value?.querySelector<HTMLElement>('[data-popup]') ?? null;
    watch(
      popup.open,
      (isOpen) => {
        if (!isOpen) {
          draft.value = '';
          return;
        }
        if (!keyboardOpen) return;
        keyboardOpen = false;
        focusColorSwatch(popupEl());
      },
      { flush: 'post' }
    );
    const apply = (value: string): void => {
      popup.setOpen(false);
      if (!s.isEnabled.value) return;
      s.execute(value);
      if (value !== clear) last.value = value;
      s.focusGrid();
    };
    const tr = (key: string, params?: Readonly<Record<string, string | number>>): string =>
      translateKey(t, key, params);
    return () => {
      const { layout } = props;
      const label = s.words.value.label;
      const current = s.value.value?.toUpperCase() ?? null;
      const override = overrideOf();
      const attributes = slotAttributes(s, reasonId);
      // The caret is part of the slot's control: the main half alone carries `data-slot`.
      const { 'data-slot': _slot, ...caretAttributes } = attributes;
      const grid = popup.open.value ? colorPickerGrid(workbookThemeHexes(themeColors.value)) : null;
      const custom = hexFromText(draft.value);
      const swatch = (pick: ColorPick, hex: string) => {
        const value = hex.slice(1);
        const name = colorPickLabel(pick, tr);
        return (
          <button
            key={pick}
            type="button"
            tabindex={-1}
            class="xlsx-toolbar-swatch"
            data-color-swatch=""
            data-color={value}
            data-tone={colorBarTone(value)}
            aria-label={name}
            title={name}
            aria-pressed={current === value}
            style={{ backgroundColor: hex }}
            onMousedown={guardToolbarMousedown}
            onClick={() => apply(value)}
          />
        );
      };
      return (
        <div
          ref={popup.root}
          class={`xlsx-toolbar-dropdown xlsx-toolbar-colorsplit${layout === 'row' ? ' xlsx-toolbar-labelled' : ''}`}
        >
          {layout === 'row' ? <span class="xlsx-toolbar-row-label">{label}</span> : null}
          <button
            type="button"
            class={withClass('xlsx-toolbar-button xlsx-toolbar-colorsplit-main', override)}
            aria-label={label}
            {...attributes}
            onClick={() => apply(last.value)}
          >
            <span class="xlsx-toolbar-color-icon">
              {override?.icon ?? <ChromeIcon paths={s.control.paths} />}
              <span
                class="xlsx-toolbar-color-bar"
                data-tone={colorBarTone(last.value)}
                style={isHexColor(last.value) ? { backgroundColor: `#${last.value}` } : undefined}
              />
            </span>
          </button>
          <button
            type="button"
            class="xlsx-toolbar-button xlsx-toolbar-colorsplit-caret"
            aria-haspopup="dialog"
            aria-expanded={popup.open.value}
            aria-controls={popup.open.value ? popup.id : undefined}
            aria-label={t('colorPicker.colorOptions', { label })}
            {...caretAttributes}
            onClick={() => {
              if (s.isEnabled.value) popup.setOpen(!popup.open.value);
            }}
            onKeydown={(e: KeyboardEvent) => {
              const action = popupTriggerAction(e.key, popup.open.value);
              if (!action) return;
              e.preventDefault();
              if (!s.isEnabled.value) return;
              if (action === 'open') {
                keyboardOpen = true;
                popup.setOpen(true);
              } else if (action === 'focusIn') focusColorSwatch(popupEl());
              else popup.setOpen(false);
            }}
            onKeyup={(e: KeyboardEvent) => {
              if (e.key === ' ') e.preventDefault();
            }}
          >
            <ChromeIcon paths={ICON_MENU_CARET} className="xlsx-toolbar-caret" />
          </button>
          {reasonText(reasonId, s.words.value.description)}
          {grid ? (
            <div
              id={popup.id}
              role="dialog"
              aria-label={label}
              {...popupAttributes(layout, 'xlsx-toolbar-color-popup')}
              onKeydown={(e: KeyboardEvent) => {
                if ((e.target as HTMLElement).tagName === 'INPUT') return;
                if (moveColorSwatchFocus(e.currentTarget as HTMLElement, e.key)) e.preventDefault();
              }}
            >
              <button
                type="button"
                class="xlsx-toolbar-color-clear"
                data-color-none=""
                data-color={clear}
                aria-pressed={fill ? current === NO_FILL.toUpperCase() : undefined}
                onMousedown={guardToolbarMousedown}
                onClick={() => apply(clear)}
              >
                <span
                  class={`xlsx-toolbar-color-chip${fill ? ' xlsx-toolbar-color-chip-none' : ''}`}
                  style={{ backgroundColor: `#${CLEAR_CHIP_HEX[fill ? 'none' : 'auto']}` }}
                  aria-hidden="true"
                />
                {t(fill ? 'colorPicker.noFill' : 'colorPicker.automatic')}
              </button>
              <div class="xlsx-toolbar-color-section">
                <div class="xlsx-toolbar-color-heading" id={`${headingId}-theme`}>
                  {t('colorPicker.themeColors')}
                </div>
                <div
                  role="group"
                  aria-labelledby={`${headingId}-theme`}
                  class="xlsx-toolbar-color-grid xlsx-toolbar-color-grid-theme"
                >
                  {grid.theme.flat().map((c) => swatch(c.pick, c.hex))}
                </div>
              </div>
              <div class="xlsx-toolbar-color-section">
                <div class="xlsx-toolbar-color-heading" id={`${headingId}-standard`}>
                  {t('colorPicker.standardColors')}
                </div>
                <div
                  role="group"
                  aria-labelledby={`${headingId}-standard`}
                  class="xlsx-toolbar-color-grid"
                >
                  {grid.standard.map((c) => swatch(c.pick, c.hex))}
                </div>
              </div>
              <div class="xlsx-toolbar-color-section">
                <label class="xlsx-toolbar-color-heading" for={`${headingId}-hex`}>
                  {t('colorPicker.customColor')}
                </label>
                <div class="xlsx-toolbar-color-custom">
                  <span class="xlsx-toolbar-color-hash" aria-hidden="true">
                    #
                  </span>
                  <input
                    id={`${headingId}-hex`}
                    type="text"
                    class="xlsx-toolbar-color-hex"
                    dir="ltr"
                    value={draft.value}
                    spellcheck={false}
                    autocomplete="off"
                    aria-label={t('colorPicker.hexColor')}
                    aria-invalid={draft.value !== '' && !custom}
                    onInput={(e: Event) => {
                      const input = e.target as HTMLInputElement;
                      draft.value = customHexDraft(input.value);
                      input.value = draft.value;
                    }}
                    onKeydown={(e: KeyboardEvent) => {
                      if (e.key === 'Enter' && custom) {
                        e.preventDefault();
                        apply(custom);
                      }
                    }}
                  />
                  <button
                    type="button"
                    class="xlsx-toolbar-color-apply"
                    disabled={!custom}
                    onMousedown={guardToolbarMousedown}
                    onClick={() => {
                      if (custom) apply(custom);
                    }}
                  >
                    {t('colorPicker.apply')}
                  </button>
                </div>
              </div>
            </div>
          ) : null}
        </div>
      );
    };
  },
});

/**
 * One slot as a menu item: its label, and its reason as visible text and the tooltip when it
 * is disabled. Menus hold no format slot, so a menu item has no hint.
 */
const MenuItem = defineComponent({
  name: 'XlsxToolbarMenuItem',
  props: {
    slot: { type: String as PropType<ChromeSlotId>, required: true },
    close: { type: Function as PropType<() => void>, required: true },
  },
  setup(props) {
    const s = useToolbarSlot(props.slot);
    const reasonId = useId();
    return () => {
      const words = s.words.value;
      const reason = s.disabledReason.value;
      // A disabled menu item stays focusable (`aria-disabled`), so the keyboard reaches its
      // reason.
      return (
        <button
          type="button"
          role="menuitem"
          class="xlsx-toolbar-row"
          tabindex={-1}
          data-slot={s.slot}
          aria-disabled={!s.isEnabled.value}
          {...(reason ? { 'aria-describedby': reasonId } : {})}
          title={words.title}
          onMousedown={guardToolbarMousedown}
          onClick={() => {
            if (!s.isEnabled.value) return;
            props.close();
            s.execute();
            s.focusGrid();
          }}
        >
          <ChromeIcon paths={s.control.paths} />
          <span>{words.label}</span>
          {reason ? (
            <span id={reasonId} class="xlsx-toolbar-reason">
              {reason}
            </span>
          ) : null}
        </button>
      );
    };
  },
});

/** A group that renders as one menu button (insert and delete, clear). */
export const GroupMenu = defineComponent({
  name: 'XlsxGroupMenu',
  props: { group: { type: Object as PropType<ChromeGroup>, required: true } },
  setup(props) {
    const { t } = useTranslation();
    const slots = computed(() =>
      props.group.controls.map((c) => `${props.group.id}.${c.id}` as ChromeSlotId)
    );
    // To open the menu is no command: it is disabled when every row is (`toolbarRowsState`).
    const rows = useRowsState(slots);
    const popup = usePopup('[role="menuitem"]', { closed: () => !rows.isEnabled.value });
    const reasonId = useId();
    const close = (): void => popup.setOpen(false);
    return () => {
      const { group } = props;
      const name = translateKey(t, group.labelKey);
      const reason = rows.disabledReason.value;
      return (
        <div ref={popup.root} class="xlsx-toolbar-dropdown">
          <button
            type="button"
            class="xlsx-toolbar-button xlsx-toolbar-menu-button"
            data-group-menu={group.id}
            aria-haspopup="menu"
            aria-expanded={popup.open.value}
            aria-controls={popup.open.value ? popup.id : undefined}
            aria-label={name}
            title={reason ? t('toolbar.disabledTitle', { label: name, reason }) : name}
            {...(reason ? { 'aria-disabled': 'true', 'data-disabled': '' } : {})}
            {...(reason ? { 'aria-describedby': reasonId } : {})}
            onMousedown={guardToolbarMousedown}
            onClick={() => rows.isEnabled.value && popup.setOpen(!popup.open.value)}
            {...popupTriggerKeys(popup, rows.isEnabled.value)}
          >
            <ChromeIcon paths={group.menu?.paths ?? null} />
            <ChromeIcon paths={ICON_MENU_CARET} className="xlsx-toolbar-caret" />
            {reasonText(reasonId, reason)}
          </button>
          {popup.open.value ? (
            <div
              id={popup.id}
              role="menu"
              aria-label={name}
              class="xlsx-toolbar-popup xlsx-toolbar-menu"
              data-popup=""
              onKeydown={(e: KeyboardEvent) => movePopupFocus(e, '[role="menuitem"]')}
            >
              {slots.value.map((slot) => (
                <MenuItem key={slot} slot={slot} close={close} />
              ))}
            </div>
          ) : null}
        </div>
      );
    };
  },
});
