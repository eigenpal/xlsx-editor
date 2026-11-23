// What the number format list shows for the value of its slot, DOM-free. Every adapter shows
// the list from this one function, so the naming rule of the closed list (an entry, an entry
// with other decimal places, or a custom code) exists once, in the engine.

import { EN_US, type NumberLocale } from '../cascade/index.ts';
import type { EditorTranslate } from './cell-input.ts';
import {
  CUSTOM_NUMBER_FORMAT_PREFIX,
  numberFormatChoiceCode,
  numberFormatChoices,
  numberFormatFamilyOf,
  numberFormatPlaces,
  type NumberFormatChoiceValue,
} from './number-format-commands.ts';

/**
 * One entry of the number format list, with its name translated.
 *
 * @public
 */
export interface NumberFormatListEntry {
  readonly value: NumberFormatChoiceValue;
  readonly label: string;
}

/**
 * What the number format list shows for the value of the `number.format` slot.
 *
 * @public
 */
export interface NumberFormatListView {
  /** The entries that have a code in the number locale, in list order. */
  readonly entries: readonly NumberFormatListEntry[];
  /**
   * The text of the closed list: the name of the entry, or `Custom:` and the code of a format
   * that names no entry. Empty with no value.
   */
  readonly text: string;
  /**
   * What the accessible name and the title of the closed list say after the slot label: the
   * text, and the decimal places of a format that has them, such as `Currency, 3 decimal
   * places`. Never the code of an entry. Empty with no value.
   */
  readonly name: string;
  /**
   * The entry whose code differs from the format only in its decimal places, as after a
   * decimal step, or undefined. The closed list shows its name. The entry is current, not
   * picked, because a pick writes the entry's own code: the keyboard opens the list on it.
   */
  readonly current?: NumberFormatChoiceValue;
}

/**
 * What the number format list shows for a value of the `number.format` slot
 * (`ToolbarCommandState.value`), in the number locale. `t` translates the entry names and the
 * `numberFormat.custom` and `numberFormat.family` texts.
 *
 * - An entry shows its name. An entry with decimal places names them, so the name says what a
 *   decimal step changes.
 * - A code that differs from an entry only in its decimal places shows the entry's name, names
 *   its places, and sets `current` to that entry.
 * - Any other code shows `Custom:` and the code, and names its places when a step can change
 *   them.
 *
 * @public
 */
export function numberFormatListView(
  value: string | undefined,
  t: EditorTranslate,
  locale: NumberLocale = EN_US
): NumberFormatListView {
  const entries = numberFormatChoices(locale).map((c) => ({
    value: c.value,
    label: t(c.labelKey),
  }));
  const named = (text: string, places: number | undefined) =>
    places === undefined ? text : t('numberFormat.family', { name: text, places });
  const known = entries.find((e) => e.value === value);
  if (known) {
    const code = numberFormatChoiceCode(known.value, locale);
    const places = code === undefined ? undefined : numberFormatPlaces(code);
    return { entries, text: known.label, name: named(known.label, places) };
  }
  if (value === undefined || !value.startsWith(CUSTOM_NUMBER_FORMAT_PREFIX))
    return { entries, text: '', name: '' };
  const code = value.slice(CUSTOM_NUMBER_FORMAT_PREFIX.length);
  const family = numberFormatFamilyOf(code, locale);
  const entry = family && entries.find((e) => e.value === family.value);
  if (family && entry)
    return {
      entries,
      text: entry.label,
      name: named(entry.label, family.places),
      current: entry.value,
    };
  const text = t('numberFormat.custom', { code });
  return { entries, text, name: named(text, numberFormatPlaces(code)) };
}
