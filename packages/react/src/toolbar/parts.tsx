/** @jsxImportSource react */
// The toolbar parts: an icon button, a dropdown list of values, a color palette, and a group
// menu. Each part renders one slot of the chrome registry through `useToolbarSlot`, and each
// popup closes through the one close rule, `useDropdownClose`. Chrome `mousedown` keeps the
// keyboard on the grid, except in a form field (AGENTS.md "Conventions", focus stealing). A
// pick returns the keyboard to the grid and leaves the selection where it was.
//
// A disabled control keeps its focus and carries `aria-disabled` and its reason as the
// description. An enabled control with a hint keeps its label as the tooltip, with the hint on a
// second line, and the hint as the description. DOCX Editor's toolbar buttons use `disabled`
// with the reason as the tooltip; this toolbar differs because a `disabled` button takes no
// focus, so the keyboard could not reach the reason, and a control that becomes disabled under
// focus would drop the focus to the page. Inside the overflow panel, a popup renders in flow
// (`data-inline`), so the panel's scroll box never clips it.

import {
  createContext,
  useContext,
  useEffect,
  useId,
  useRef,
  useMemo,
  useState,
  type KeyboardEvent,
  type ReactElement,
  type ReactNode,
} from 'react';
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

/**
 * The changes that a `Toolbar.Slot` or a named part makes to the built-in control of its slot,
 * after `SlotOverrideContext` of pptx-editor/packages/react/src/slot-override.ts at 70052cf1
 * (EigenPal): an icon in place of the registry icon and the host element that renders the
 * control (button slots only), and a class name for the control or the trigger of a value
 * control.
 */
export interface SlotOverride {
  readonly slot: string;
  readonly icon?: ReactNode;
  readonly className?: string;
  readonly asChild?: ReactElement | null;
}

/** The override of the control that renders inside it. */
export const SlotOverrideContext = createContext<SlotOverride | null>(null);

/** The override of `slot`, or null: an override of another slot never reaches this control. */
function useSlotOverride(slot: string): SlotOverride | null {
  const override = useContext(SlotOverrideContext);
  return override?.slot === slot ? override : null;
}

/** A base class with the override's class name after it. */
function withClass(base: string, override: SlotOverride | null): string {
  return override?.className ? `${base} ${override.className}` : base;
}

/**
 * The attributes that every slot control carries. A disabled control stays focusable
 * (`aria-disabled`), so the keyboard reaches it and its reason, and arrow navigation never loses
 * a control that becomes disabled under focus. Its press does nothing. The reason is the
 * tooltip and the accessible description (`aria-describedby` on a hidden `ReasonText`).
 */
function slotAttributes(s: ToolbarSlot, reasonId: string) {
  return {
    'data-slot': s.slot,
    'aria-disabled': s.isEnabled ? undefined : ('true' as const),
    ...(s.description ? { 'aria-describedby': reasonId } : {}),
    title: s.title,
    onMouseDown: guardToolbarMousedown,
    ...(s.isActive ? { 'data-active': '' } : {}),
    ...(s.isEnabled ? {} : { 'data-disabled': '' }),
  };
}

/** The disabled reason or the hint of a control, as its hidden accessible description. */
function ReasonText({ id, reason }: { readonly id: string; readonly reason: string | null }) {
  return reason ? (
    <span id={id} hidden>
      {reason}
    </span>
  ) : null;
}

/** One command slot as a button. In a row, the label shows beside the icon. */
export function ToolbarButton({
  slot,
  layout = 'bar',
}: {
  readonly slot: ChromeSlotId;
  readonly layout?: PartLayout;
}) {
  const s = useToolbarSlot(slot);
  const reasonId = useId();
  const override = useSlotOverride(slot);
  const shared = {
    className: withClass(layout === 'row' ? 'xlsx-toolbar-row' : 'xlsx-toolbar-button', override),
    'aria-label': s.label,
    ...(s.control.toggle ? { 'aria-pressed': s.isActive } : {}),
    ...slotAttributes(s, reasonId),
    // A held Enter repeats the click of a button; a command such as a decimal step would
    // run on each repeat, so only the first press runs it. The guard applies to every toolbar
    // button, undo and redo too. Intended difference from
    // docx-editor/packages/react/src/components/Toolbar.tsx at 2741998e (EigenPal), whose
    // buttons have no repeat guard.
    onKeyDown: (e: KeyboardEvent<HTMLElement>) => {
      if (e.key === 'Enter' && e.repeat) e.preventDefault();
    },
    // A disabled control also stops the default of its element, so an `asChild` link does not
    // navigate.
    onClick: (e: { preventDefault(): void }) => {
      if (s.isEnabled) s.execute();
      else e.preventDefault();
    },
  };
  const reason = <ReasonText id={reasonId} reason={s.description} />;
  // `asChild`: the host element renders the control, with its state and press merged in.
  if (override?.asChild)
    return (
      <>
        {renderAsChild(override.asChild, shared)}
        {reason}
      </>
    );
  const icon = override?.icon ?? <ChromeIcon paths={s.control.paths} />;
  // A slot with no icon (a menu item of the registry) shows its label in the bar too.
  const text = layout === 'row' || (override?.icon == null && !s.control.paths);
  return (
    <button type="button" {...shared}>
      {icon}
      {text ? <span>{s.label}</span> : null}
      {reason}
    </button>
  );
}

/** The trigger of a popup: a button that opens it with a press, Enter, Space, or Down. */
function Trigger({
  s,
  popup,
  kind,
  className,
  caret,
  name,
  children,
  label,
}: {
  readonly s: ToolbarSlot;
  readonly popup: Popup;
  readonly kind: 'listbox' | 'menu' | 'dialog';
  readonly className: string;
  /** The caret: a field's (`expand_more`) or a menu button's (`arrow_drop_down`). */
  readonly caret: 'field' | 'menu';
  readonly name: string;
  readonly children: ReactNode;
  /**
   * The text of the title in place of the slot label, such as the label and the value of a
   * list. The disabled reason or the hint follows on its own line, so the title never has two
   * colons, as in "Number format: General: <reason>".
   */
  readonly label?: string;
}) {
  const { t } = useTranslation();
  const reasonId = useId();
  const override = useSlotOverride(s.slot);
  const attributes = slotAttributes(s, reasonId);
  const after = s.description;
  const title =
    label === undefined
      ? attributes.title
      : after
        ? t('toolbar.hintTitle', { label, hint: after })
        : label;
  return (
    <button
      type="button"
      className={withClass(className, override)}
      aria-haspopup={kind}
      aria-expanded={popup.open}
      aria-controls={popup.open ? popup.id : undefined}
      aria-label={name}
      {...attributes}
      title={title}
      onClick={() => {
        if (s.isEnabled) popup.setOpen(!popup.open);
      }}
      {...popupTriggerKeys(popup, s.isEnabled)}
    >
      {children}
      <ChromeIcon
        paths={caret === 'field' ? ICON_FIELD_CARET : ICON_MENU_CARET}
        className="xlsx-toolbar-caret"
      />
      <ReasonText id={reasonId} reason={s.description} />
    </button>
  );
}

/** The class of a popup, and the attributes that keep one in flow inside the overflow panel. */
function popupAttributes(layout: PartLayout, className: string) {
  return {
    className: `xlsx-toolbar-popup ${className}`,
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
function ToolbarList({
  slot,
  options,
  text,
  layout,
  width,
  entry,
  more,
  name,
  active,
}: {
  readonly slot: ChromeSlotId;
  readonly options: readonly { readonly value: string; readonly label: string }[];
  readonly text: string;
  /** What the trigger's accessible name and title say after the slot label, when not `text`. */
  readonly name?: string;
  /** The current entry (`aria-current`), which the keyboard focuses on open when none is picked. */
  readonly active?: string;
  readonly layout: PartLayout;
  readonly width: 'wide' | 'narrow';
  readonly entry?: ListEntry;
  /** A last item after the list, such as the number format dialog's opener. */
  readonly more?: (close: () => void) => ReactNode;
}) {
  const s = useToolbarSlot(slot);
  const items = `${entry ? 'input, ' : ''}[role="option"]${more ? ', [data-list-more]' : ''}`;
  const popup = usePopup(items, {
    inline: layout === 'row',
    closed: !s.isEnabled,
  });
  const [invalid, setInvalid] = useState(false);
  const pick = (value: string) => {
    popup.setOpen(false);
    s.execute(value);
    s.focusGrid();
  };
  return (
    <div
      ref={popup.root}
      className={`xlsx-toolbar-dropdown${layout === 'row' ? ' xlsx-toolbar-labelled' : ''}`}
    >
      {layout === 'row' ? <span className="xlsx-toolbar-row-label">{s.label}</span> : null}
      <Trigger
        s={s}
        popup={popup}
        kind="listbox"
        className={`xlsx-toolbar-select xlsx-toolbar-select-${width}`}
        caret="field"
        name={`${s.label}: ${name ?? text}`}
        {...(name !== undefined ? { label: `${s.label}: ${name}` } : {})}
      >
        <span className="xlsx-toolbar-select-text">{text}</span>
      </Trigger>
      {popup.open ? (
        <div
          {...popupAttributes(layout, 'xlsx-toolbar-list-popup')}
          onKeyDown={(e) => movePopupFocus(e, items)}
        >
          {entry ? (
            <input
              type="text"
              inputMode="decimal"
              autoComplete="off"
              className="xlsx-toolbar-entry"
              aria-label={entry.label}
              aria-invalid={invalid}
              {...(invalid ? { 'aria-describedby': `${popup.id}-entry-error` } : {})}
              placeholder={text}
              onChange={() => setInvalid(false)}
              onKeyDown={(e) => {
                if (e.key !== 'Enter') return;
                e.preventDefault();
                const value = entry.parse(e.currentTarget.value);
                if (value === null) setInvalid(true);
                else pick(value);
              }}
            />
          ) : null}
          {entry && invalid ? (
            <div id={`${popup.id}-entry-error`} role="alert" className="xlsx-toolbar-entry-error">
              {entry.invalid}
            </div>
          ) : null}
          <ul id={popup.id} role="listbox" aria-label={s.label} className="xlsx-toolbar-list">
            {options.map((o) => (
              <li
                key={o.value}
                role="option"
                tabIndex={-1}
                aria-selected={o.value === s.value}
                {...(o.value === active ? { 'aria-current': 'true' as const } : {})}
                onMouseDown={guardToolbarMousedown}
                onClick={() => pick(o.value)}
                onKeyDown={(e) => {
                  if (e.key !== 'Enter' && e.key !== ' ') return;
                  e.preventDefault();
                  pick(o.value);
                }}
              >
                {o.label}
                {o.value === s.value ? (
                  <ChromeIcon paths={ICON_CHECK} className="xlsx-toolbar-check" />
                ) : null}
              </li>
            ))}
          </ul>
          {more?.(() => popup.setOpen(false))}
        </div>
      ) : null}
    </div>
  );
}

/** The font family list: the choices, and the active cell's own family when it is not one. */
export function FontFamilyList({ layout = 'bar' }: { readonly layout?: PartLayout }) {
  const s = useToolbarSlot('font.family');
  const current = s.value ?? '';
  const names = current && !FONT_FAMILY_CHOICES.includes(current) ? [current] : [];
  const options = [...names, ...FONT_FAMILY_CHOICES].map((n) => ({ value: n, label: n }));
  return (
    <ToolbarList slot="font.family" options={options} text={current} layout={layout} width="wide" />
  );
}

/** The font size list, in points, with the active cell's own size when it is not a choice. */
export function FontSizeList({ layout = 'bar' }: { readonly layout?: PartLayout }) {
  const { t } = useTranslation();
  const s = useToolbarSlot('font.size');
  const current = Number(s.value);
  const sizes = [...FONT_SIZE_CHOICES];
  if (Number.isFinite(current) && current > 0 && !sizes.includes(current)) sizes.push(current);
  sizes.sort((a, b) => a - b);
  const options = sizes.map((n) => ({ value: String(n), label: String(n) }));
  return (
    <ToolbarList
      slot="font.size"
      options={options}
      text={s.value ?? ''}
      layout={layout}
      width="narrow"
      entry={{
        label: t('fontSize.typeSize'),
        invalid: t('fontSize.invalid'),
        parse: parseFontSize,
      }}
    />
  );
}

const numberLocaleOf = (snapshot: EditorSnapshot) => snapshot.numberLocale;
const themeColorsOf = (snapshot: EditorSnapshot) => snapshot.themeColors;

/**
 * The number format list: the built-in choices, and the code of any other format. A format
 * that is no choice has the value `CUSTOM_NUMBER_FORMAT_PREFIX` and its code, so its code never
 * shows as a choice.
 */
export function NumberFormatList({
  layout = 'bar',
  more = false,
}: {
  readonly layout?: PartLayout;
  /** Show the number format dialog's opener (`number.custom`) after the list. */
  readonly more?: boolean;
}) {
  const { t } = useTranslation();
  const s = useToolbarSlot('number.format');
  const locale = useEditorState(numberLocaleOf);
  // The engine names the entries and the closed list (`numberFormatListView`). An entry has no
  // title: as in DOCX Editor's lists
  // (docx-editor/packages/react/src/components/ui/StylePicker.tsx at 1d092ab8, EigenPal), an
  // option says only its name, so a screen reader does not read a format code on each arrow key.
  // The number format dialog shows the codes.
  const view = useMemo(
    () =>
      numberFormatListView(
        s.value ?? undefined,
        (key, params) => translateKey(t, key, params),
        locale
      ),
    [s.value, t, locale]
  );
  return (
    <ToolbarList
      slot="number.format"
      options={view.entries}
      text={view.text}
      {...(view.name ? { name: view.name } : {})}
      {...(view.current !== undefined ? { active: view.current } : {})}
      layout={layout}
      width="wide"
      {...(more ? { more: (close: () => void) => <MoreFormatsItem close={close} /> } : {})}
    />
  );
}

/**
 * The last item of the number format list: it opens the number format dialog. Focus goes back
 * to the list's trigger on a cancel when the keyboard opened the dialog, and else stays where
 * it was (the grid), as a press of a toolbar control keeps it there.
 */
function MoreFormatsItem({ close }: { readonly close: () => void }) {
  const s = useToolbarSlot('number.custom');
  const dialog = useNumberFormatDialogOpener();
  const reasonId = useId();
  // With no dialog host, the opener renders nothing, after DOCX Editor's dialog hosts.
  if (!dialog) return null;
  return (
    <button
      type="button"
      className="xlsx-toolbar-list-more"
      tabIndex={-1}
      data-list-more=""
      aria-haspopup="dialog"
      {...slotAttributes(s, reasonId)}
      onClick={(e) => {
        if (!s.isEnabled) return;
        const item = e.currentTarget;
        const root = item.closest('.xlsx-toolbar-dropdown');
        const fromKeyboard = item.ownerDocument.activeElement === item;
        const trigger = root?.querySelector<HTMLElement>('[aria-haspopup="listbox"]') ?? null;
        close();
        // A refused start opens no dialog, and the list with this item closes: the keyboard goes
        // to the grid, as a refused choice of a list option does.
        if (!dialog.open(fromKeyboard ? trigger : null) && fromKeyboard) s.focusGrid();
      }}
    >
      {s.label}
      <ReasonText id={reasonId} reason={s.description} />
    </button>
  );
}

/**
 * The `number.custom` slot alone, when the bar has no number format list to hold it: an icon
 * button that opens the number format dialog.
 */
export function NumberFormatDialogButton({ layout = 'bar' }: { readonly layout?: PartLayout }) {
  const s = useToolbarSlot('number.custom');
  const dialog = useNumberFormatDialogOpener();
  const reasonId = useId();
  const override = useSlotOverride('number.custom');
  if (!dialog) return null;
  return (
    <button
      type="button"
      className={withClass(layout === 'row' ? 'xlsx-toolbar-row' : 'xlsx-toolbar-button', override)}
      aria-label={s.label}
      aria-haspopup="dialog"
      {...slotAttributes(s, reasonId)}
      onClick={(e) => {
        if (!s.isEnabled) return;
        const self = e.currentTarget;
        dialog.open(self.ownerDocument.activeElement === self ? self : null);
      }}
    >
      <ChromeIcon paths={s.control.paths} />
      {layout === 'row' ? <span>{s.label}</span> : null}
      <ReasonText id={reasonId} reason={s.description} />
    </button>
  );
}

/**
 * The font color or the fill color as a split button, after `ColorSplit` of
 * docx-editor/packages/react/src/editor/toolbar/ColorSplit.tsx at 60a86b48 (EigenPal). The main
 * half applies the last color that it applied, which its bar shows: red for the font color and
 * yellow for the fill color until the first pick. The caret opens the picker: Automatic (or No
 * fill), the workbook theme colors with five rows of tints and shades, the standard colors, and
 * a custom hex field. The grid and the keys come from core (`color-picker.ts`), so the Vue twin
 * shows the same picker. A swatch paints six hex digits that core checked, so no inline style
 * carries other text.
 *
 * Intended differences: the caret has its own accessible name ("Font color options"), so the
 * two halves differ for assistive technology, and a disabled control stays focusable as every
 * other toolbar control here does.
 */
export function ColorPalette({
  slot,
  layout = 'bar',
}: {
  readonly slot: 'font.color' | 'fill.color';
  readonly layout?: PartLayout;
}) {
  const { t } = useTranslation();
  const themeColors = useEditorState(themeColorsOf);
  const s = useToolbarSlot(slot);
  const override = useSlotOverride(slot);
  const reasonId = useId();
  const headingId = useId();
  const popup = usePopup('[data-color]', { inline: layout === 'row', closed: !s.isEnabled });
  const fill = slot === 'fill.color';
  const clear = fill ? NO_FILL : AUTO_COLOR;
  const [last, setLast] = useState(fill ? 'FFFF00' : 'FF0000');
  const [draft, setDraft] = useState('');
  const keyboardOpen = useRef(false);
  const current = s.value?.toUpperCase() ?? null;
  useEffect(() => {
    if (!popup.open) {
      setDraft('');
      return;
    }
    if (!keyboardOpen.current) return;
    keyboardOpen.current = false;
    focusColorSwatch(popup.root.current?.querySelector<HTMLElement>('[data-popup]') ?? null);
  }, [popup.open, popup.root]);
  const apply = (value: string) => {
    popup.setOpen(false);
    if (!s.isEnabled) return;
    s.execute(value);
    if (value !== clear) setLast(value);
    s.focusGrid();
  };
  const themeHexes = popup.open ? workbookThemeHexes(themeColors) : [];
  const grid = popup.open ? colorPickerGrid(themeHexes) : null;
  const custom = hexFromText(draft);
  const tr = (key: string, params?: Readonly<Record<string, string | number>>) =>
    translateKey(t, key, params);
  const swatch = (pick: ColorPick, hex: string) => {
    const value = hex.slice(1);
    return (
      <button
        key={pick}
        type="button"
        tabIndex={-1}
        className="xlsx-toolbar-swatch"
        data-color-swatch=""
        data-color={value}
        data-tone={colorBarTone(value)}
        aria-label={colorPickLabel(pick, tr)}
        title={colorPickLabel(pick, tr)}
        aria-pressed={current === value}
        style={{ backgroundColor: hex }}
        onMouseDown={guardToolbarMousedown}
        onClick={() => apply(value)}
      />
    );
  };
  const attributes = slotAttributes(s, reasonId);
  // The caret is part of the slot's control: the main half alone carries `data-slot`.
  const { 'data-slot': _slot, ...caretAttributes } = attributes;
  return (
    <div
      ref={popup.root}
      className={`xlsx-toolbar-dropdown xlsx-toolbar-colorsplit${layout === 'row' ? ' xlsx-toolbar-labelled' : ''}`}
    >
      {layout === 'row' ? <span className="xlsx-toolbar-row-label">{s.label}</span> : null}
      <button
        type="button"
        className={withClass('xlsx-toolbar-button xlsx-toolbar-colorsplit-main', override)}
        aria-label={s.label}
        {...attributes}
        onClick={() => apply(last)}
      >
        <span className="xlsx-toolbar-color-icon">
          {override?.icon ?? <ChromeIcon paths={s.control.paths} />}
          <span
            className="xlsx-toolbar-color-bar"
            data-tone={colorBarTone(last)}
            style={isHexColor(last) ? { backgroundColor: `#${last}` } : undefined}
          />
        </span>
      </button>
      <button
        type="button"
        className="xlsx-toolbar-button xlsx-toolbar-colorsplit-caret"
        aria-haspopup="dialog"
        aria-expanded={popup.open}
        aria-controls={popup.open ? popup.id : undefined}
        aria-label={t('colorPicker.colorOptions', { label: s.label })}
        {...caretAttributes}
        onClick={() => {
          if (s.isEnabled) popup.setOpen(!popup.open);
        }}
        onKeyDown={(e) => {
          const action = popupTriggerAction(e.key, popup.open);
          if (!action) return;
          e.preventDefault();
          if (!s.isEnabled) return;
          if (action === 'open') {
            keyboardOpen.current = true;
            popup.setOpen(true);
          } else if (action === 'focusIn')
            focusColorSwatch(
              popup.root.current?.querySelector<HTMLElement>('[data-popup]') ?? null
            );
          else popup.setOpen(false);
        }}
        onKeyUp={(e) => {
          if (e.key === ' ') e.preventDefault();
        }}
      >
        <ChromeIcon paths={ICON_MENU_CARET} className="xlsx-toolbar-caret" />
      </button>
      <ReasonText id={reasonId} reason={s.description} />
      {grid ? (
        <div
          id={popup.id}
          role="dialog"
          aria-label={s.label}
          {...popupAttributes(layout, 'xlsx-toolbar-color-popup')}
          onKeyDown={(e) => {
            if ((e.target as HTMLElement).tagName === 'INPUT') return;
            if (moveColorSwatchFocus(e.currentTarget, e.key)) e.preventDefault();
          }}
        >
          <button
            type="button"
            className="xlsx-toolbar-color-clear"
            data-color-none=""
            data-color={clear}
            aria-pressed={fill ? current === NO_FILL.toUpperCase() : undefined}
            onMouseDown={guardToolbarMousedown}
            onClick={() => apply(clear)}
          >
            <span
              className={`xlsx-toolbar-color-chip${fill ? ' xlsx-toolbar-color-chip-none' : ''}`}
              style={{ backgroundColor: `#${CLEAR_CHIP_HEX[fill ? 'none' : 'auto']}` }}
              aria-hidden="true"
            />
            {t(fill ? 'colorPicker.noFill' : 'colorPicker.automatic')}
          </button>
          <div className="xlsx-toolbar-color-section">
            <div className="xlsx-toolbar-color-heading" id={`${headingId}-theme`}>
              {t('colorPicker.themeColors')}
            </div>
            <div
              role="group"
              aria-labelledby={`${headingId}-theme`}
              className="xlsx-toolbar-color-grid xlsx-toolbar-color-grid-theme"
            >
              {grid.theme.flat().map((c) => swatch(c.pick, c.hex))}
            </div>
          </div>
          <div className="xlsx-toolbar-color-section">
            <div className="xlsx-toolbar-color-heading" id={`${headingId}-standard`}>
              {t('colorPicker.standardColors')}
            </div>
            <div
              role="group"
              aria-labelledby={`${headingId}-standard`}
              className="xlsx-toolbar-color-grid"
            >
              {grid.standard.map((c) => swatch(c.pick, c.hex))}
            </div>
          </div>
          <div className="xlsx-toolbar-color-section">
            <label className="xlsx-toolbar-color-heading" htmlFor={`${headingId}-hex`}>
              {t('colorPicker.customColor')}
            </label>
            <div className="xlsx-toolbar-color-custom">
              <span className="xlsx-toolbar-color-hash" aria-hidden="true">
                #
              </span>
              <input
                id={`${headingId}-hex`}
                type="text"
                className="xlsx-toolbar-color-hex"
                dir="ltr"
                value={draft}
                spellCheck={false}
                autoComplete="off"
                aria-label={t('colorPicker.hexColor')}
                aria-invalid={draft !== '' && !custom}
                onChange={(e) => setDraft(customHexDraft(e.target.value))}
                onKeyDown={(e) => {
                  if (e.key === 'Enter' && custom) {
                    e.preventDefault();
                    apply(custom);
                  }
                }}
              />
              <button
                type="button"
                className="xlsx-toolbar-color-apply"
                disabled={!custom}
                onMouseDown={guardToolbarMousedown}
                onClick={() => custom && apply(custom)}
              >
                {t('colorPicker.apply')}
              </button>
            </div>
          </div>
        </div>
      ) : null}
    </div>
  );
}

/**
 * One slot as a menu item: its label, and its reason as visible text and the tooltip when it
 * is disabled. Menus hold no format slot, so a menu item has no hint.
 */
function MenuItem({ slot, close }: { readonly slot: ChromeSlotId; readonly close: () => void }) {
  const s = useToolbarSlot(slot);
  const reasonId = useId();
  return (
    // A disabled menu item stays focusable (`aria-disabled`), so the keyboard reaches its reason.
    <button
      type="button"
      role="menuitem"
      className="xlsx-toolbar-row"
      tabIndex={-1}
      data-slot={s.slot}
      aria-disabled={!s.isEnabled}
      {...(s.disabledReason ? { 'aria-describedby': reasonId } : {})}
      title={s.title}
      onMouseDown={guardToolbarMousedown}
      onClick={() => {
        if (!s.isEnabled) return;
        close();
        s.execute();
        s.focusGrid();
      }}
    >
      <ChromeIcon paths={s.control.paths} />
      <span>{s.label}</span>
      {s.disabledReason ? (
        <span id={reasonId} className="xlsx-toolbar-reason">
          {s.disabledReason}
        </span>
      ) : null}
    </button>
  );
}

/** A group that renders as one menu button (insert and delete, clear). */
export function GroupMenu({ group }: { readonly group: ChromeGroup }) {
  const { t } = useTranslation();
  const slots = group.controls.map((c) => `${group.id}.${c.id}` as ChromeSlotId);
  // To open the menu is no command: it is disabled when every row is (`toolbarRowsState`).
  const rows = useRowsState(slots);
  const popup = usePopup('[role="menuitem"]', { closed: !rows.isEnabled });
  const reasonId = useId();
  const name = translateKey(t, group.labelKey);
  const reason = rows.disabledReason;
  return (
    <div ref={popup.root} className="xlsx-toolbar-dropdown">
      <button
        type="button"
        className="xlsx-toolbar-button xlsx-toolbar-menu-button"
        data-group-menu={group.id}
        aria-haspopup="menu"
        aria-expanded={popup.open}
        aria-controls={popup.open ? popup.id : undefined}
        aria-label={name}
        title={reason ? t('toolbar.disabledTitle', { label: name, reason }) : name}
        {...(reason ? { 'aria-disabled': 'true' as const, 'data-disabled': '' } : {})}
        {...(reason ? { 'aria-describedby': reasonId } : {})}
        onMouseDown={guardToolbarMousedown}
        onClick={() => {
          if (rows.isEnabled) popup.setOpen(!popup.open);
        }}
        {...popupTriggerKeys(popup, rows.isEnabled)}
      >
        <ChromeIcon paths={group.menu?.paths ?? null} />
        <ChromeIcon paths={ICON_MENU_CARET} className="xlsx-toolbar-caret" />
        <ReasonText id={reasonId} reason={reason} />
      </button>
      {popup.open ? (
        <div
          id={popup.id}
          role="menu"
          aria-label={name}
          className="xlsx-toolbar-popup xlsx-toolbar-menu"
          data-popup=""
          onKeyDown={(e) => movePopupFocus(e, '[role="menuitem"]')}
        >
          {slots.map((slot) => (
            <MenuItem key={slot} slot={slot} close={() => popup.setOpen(false)} />
          ))}
        </div>
      ) : null}
    </div>
  );
}
