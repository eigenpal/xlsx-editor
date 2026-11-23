/** @jsxImportSource vue */
// The number format dialog of the `number.custom` slot (ECMA-376 Part 1, 18.8.31): the Vue
// twin of the React adapter's `toolbar/NumberFormatDialog.tsx`. The header there explains the
// rules: a format code field, a live preview of the active cell's value through the engine's
// own renderer (`previewNumberFormat`), the custom codes of the workbook
// (`workbookNumberFormatCodes`), and Apply through the slot, in one transaction and one undo
// step. One host per toolbar renders the dialog, after DOCX Editor's `ParagraphDialogHost`.
// The workbook's codes are file text: they render as text only.
//
// Differences from React: the opener and the panel close are injection keys, not contexts; the
// preview and the codes are computed refs that read the snapshot, so each editor change reads
// them again, as React's render does.

import {
  computed,
  defineComponent,
  inject,
  onMounted,
  provide,
  shallowRef,
  useId,
  watch,
  type InjectionKey,
  type PropType,
} from 'vue';
import {
  numberFormatDialogApplyStep,
  numberFormatDialogCurrent,
  numberFormatDialogStart,
  previewNumberFormat,
  refusalText,
  workbookNumberFormatCodes,
  type NumberFormatDialogStart,
} from '@xlsx-editor.dev/core/editor';
import { DialogFrame } from '../dialog-frame.tsx';
import { useEditorSnapshot, useEditorState, useTranslation, useXlsxEditor } from '../context.ts';
import { translateKey, useToolbarSlot } from './slot.tsx';

/** Opens the number format dialog. `returnTo` takes focus on a cancel, when it is in the page. */
export interface NumberFormatDialogOpener {
  /** True when the dialog opened; false when the start was refused (`numberFormatDialogStart`). */
  open(returnTo: HTMLElement | null): boolean;
}

const openerKey: InjectionKey<NumberFormatDialogOpener> = Symbol('xlsxNumberFormatOpener');

/** Closes the toolbar panel that holds an opener, such as the "More" panel. */
export const panelCloseKey: InjectionKey<() => void> = Symbol('xlsxPanelClose');

/**
 * The opener of the toolbar's number format dialog, or null outside a toolbar. An open from the
 * mouse also closes the panel that holds the opener, because focus does not go back to it.
 */
export function useNumberFormatDialogOpener(): NumberFormatDialogOpener | null {
  const opener = inject(openerKey, null);
  const closePanel = inject(panelCloseKey, null);
  if (!opener || !closePanel) return opener;
  return {
    open(returnTo: HTMLElement | null) {
      const opened = opener.open(returnTo);
      if (opened && !returnTo) closePanel();
      return opened;
    },
  };
}

/** Renders the number format dialog of a toolbar, once, beside its controls. */
export const NumberFormatDialogHost = defineComponent({
  name: 'XlsxNumberFormatDialogHost',
  setup(_props, { slots }) {
    const editor = useXlsxEditor();
    const open = shallowRef<{
      readonly returnTo: HTMLElement | null;
      readonly start: NumberFormatDialogStart;
    } | null>(null);
    provide(openerKey, {
      open: (returnTo) => {
        // The start commits an open cell edit; a refused commit keeps the dialog closed, and
        // the status message gives the reason.
        const start = editor.value ? numberFormatDialogStart(editor.value) : null;
        if (start) open.value = { returnTo, start };
        return start !== null;
      },
    });
    const close = (): void => {
      open.value = null;
    };
    return () => [
      slots.default?.(),
      open.value ? (
        <NumberFormatDialog
          returnTo={open.value.returnTo}
          start={open.value.start}
          onClose={close}
        />
      ) : null,
    ];
  },
});

const HEX6 = /^[0-9A-Fa-f]{6}$/;

const NumberFormatDialog = defineComponent({
  name: 'XlsxNumberFormatDialog',
  props: {
    returnTo: { type: Object as PropType<HTMLElement | null>, default: null },
    start: { type: Object as PropType<NumberFormatDialogStart>, required: true },
    onClose: { type: Function as PropType<() => void>, required: true },
  },
  setup(props) {
    const editor = useXlsxEditor();
    const { t } = useTranslation();
    const s = useToolbarSlot('number.custom');
    // The code that the dialog opened with: `numberFormatDialogApplyStep` decides when an Apply
    // of it unchanged writes nothing.
    const code = shallowRef(props.start.initial);
    const error = shallowRef<string | null>(null);
    const errorCount = shallowRef(0);
    let applied = false;
    // After an Apply that the code check refused, the field stays invalid until its code reads.
    const attempted = shallowRef(false);
    const field = shallowRef<HTMLInputElement | null>(null);
    const list = shallowRef<HTMLSelectElement | null>(null);
    const id = useId();
    const ids = {
      field: `${id}-code`,
      hint: `${id}-hint`,
      preview: `${id}-preview`,
      error: `${id}-error`,
      list: `${id}-list`,
      sample: `${id}-sample`,
      blank: `${id}-blank`,
      previewLabel: `${id}-preview-label`,
    };
    // The frame focuses the field; its text is selected, so a typed code replaces it.
    onMounted(() => field.value?.select());
    // Each change of the editor state reads the dialog again: the preview reads the cell anew,
    // and the dialog closes when its start no longer holds (another workbook, a moved
    // selection, or a disabled slot such as view mode or protection that locks formats).
    const snapshot = useEditorSnapshot();
    // The workbook's codes, read again on each editor change: a host call can add a code while
    // the dialog is open. A keystroke in the field does not read them again.
    const codes = computed(() => {
      void snapshot.value;
      return editor.value ? workbookNumberFormatCodes(editor.value) : [];
    });
    // The close check of the host procedure in `packages/react/docs/toolbar-and-keyboard.md`.
    const current = useEditorState(
      () => !!editor.value && numberFormatDialogCurrent(editor.value, props.start)
    );
    watch(
      current,
      (now) => {
        if (!now) props.onClose();
      },
      { immediate: true, flush: 'post' }
    );
    const preview = computed(() => {
      void snapshot.value;
      return editor.value ? previewNumberFormat(editor.value, code.value) : null;
    });
    const problem = computed(() => {
      const p = preview.value;
      return p && !p.ok ? translateKey(t, p.problem.key, p.problem.params) : null;
    });
    // A code that reads ends the invalid state.
    watch(problem, (p) => {
      if (p === null) attempted.value = false;
    });
    // A list box with no matching value selects its first option in some browsers, and a click
    // on a selected option changes nothing: select none when the field holds no listed code.
    watch(
      [code, codes, list],
      ([now, listed, el]) => {
        if (!el) return;
        if (listed.includes(now)) el.value = now;
        else el.selectedIndex = -1;
      },
      { flush: 'post', immediate: true }
    );
    const apply = (): void => {
      const ed = editor.value;
      if (!ed) return;
      const step = numberFormatDialogApplyStep(ed, props.start, code.value);
      if (step.kind === 'closed' || step.kind === 'unchanged') {
        // An unchanged Apply returns the keyboard to the grid, as a writing Apply does.
        if (step.kind === 'unchanged') applied = true;
        props.onClose();
        return;
      }
      if (step.kind === 'problem') {
        attempted.value = true;
        error.value = translateKey(t, step.problem.key, step.problem.params);
        errorCount.value += 1;
        field.value?.focus();
        return;
      }
      if (s.execute(code.value)) {
        applied = true;
        props.onClose();
        return;
      }
      const after = ed.snapshot();
      // A refusal with no reason in the snapshot still tells the user that nothing changed.
      error.value = after.refusal
        ? refusalText(after.refusal, t, after.numberLocale)
        : t('dialogs.numberFormat.notApplied');
      errorCount.value += 1;
    };
    const pickCode = (value: string): void => {
      code.value = value;
      error.value = null;
    };
    return () => {
      const p = preview.value;
      const listed = codes.value;
      const invalid = problem.value !== null && (attempted.value || code.value !== '');
      const sample = !!p?.ok && p.sample;
      // A code can hide the value (`;;;`): then a note says so, as the preview box is empty.
      const blank = !!p?.ok && p.text === '';
      const shownColor = p?.ok && p.color && HEX6.test(p.color) ? p.color : null;
      return (
        <DialogFrame
          kind="numberFormat"
          label={t('dialogs.numberFormat.title')}
          onClose={props.onClose}
          onRestoreFocus={() => {
            if (applied && editor.value) {
              editor.value.focus();
              return true;
            }
            const back = props.returnTo;
            if (!back?.isConnected) return false;
            back.focus({ preventScroll: true });
            return true;
          }}
          onKeyDown={(event: KeyboardEvent) => {
            // Enter applies from the field and from the list; a button keeps its own Enter.
            const tag = (event.target as Element).tagName;
            if (event.key !== 'Enter' || (tag !== 'INPUT' && tag !== 'SELECT')) return;
            event.preventDefault();
            apply();
          }}
        >
          {{
            default: () => (
              <>
                <div class="xlsx-dialog__header">
                  <h2 class="xlsx-dialog__title">{t('dialogs.numberFormat.title')}</h2>
                </div>
                <div class="xlsx-dialog__body">
                  <label class="xlsx-dialog__label" for={ids.field}>
                    {t('dialogs.numberFormat.code')}
                  </label>
                  <input
                    ref={field}
                    id={ids.field}
                    name="xlsx-number-format-code"
                    type="text"
                    dir="ltr"
                    autocomplete="off"
                    spellcheck={false}
                    class="xlsx-dialog__input"
                    value={code.value}
                    aria-invalid={invalid}
                    aria-describedby={[
                      ids.hint,
                      ...(error.value ? [ids.error] : [ids.previewLabel, ids.preview]),
                      sample ? ids.sample : '',
                      blank && !error.value ? ids.blank : '',
                    ]
                      .filter(Boolean)
                      .join(' ')}
                    onInput={(e: Event) => pickCode((e.currentTarget as HTMLInputElement).value)}
                  />
                  {/* The reason of a refused Apply shows under the field, in place of the preview. */}
                  {/* A new key on each refused Apply mounts the alert again, so it is announced again. */}
                  <p key={errorCount.value} id={ids.error} role="alert" class="xlsx-dialog__error">
                    {error.value}
                  </p>
                  <p id={ids.hint} class="xlsx-dialog__hint">
                    {t(
                      listed.length > 0
                        ? 'dialogs.numberFormat.codeHint'
                        : 'dialogs.numberFormat.codeHintNoList'
                    )}
                  </p>
                  <div class="xlsx-dialog__preview" hidden={error.value !== null}>
                    <span id={ids.previewLabel} class="xlsx-dialog__preview-label">
                      {t('dialogs.numberFormat.preview')}
                    </span>
                    <output
                      id={ids.preview}
                      for={ids.field}
                      aria-labelledby={ids.previewLabel}
                      // The field's description carries the preview: no announcement for each key.
                      aria-live="off"
                      class="xlsx-dialog__preview-value"
                      data-problem={problem.value ? '' : undefined}
                      style={shownColor ? { color: `#${shownColor}` } : undefined}
                    >
                      {p?.ok
                        ? p.stale
                          ? t('grid.staleResult', { text: p.text })
                          : p.text
                        : (problem.value ?? '')}
                    </output>
                  </div>
                  {blank && error.value === null ? (
                    <p id={ids.blank} class="xlsx-dialog__hint">
                      {t('dialogs.numberFormat.previewBlank')}
                    </p>
                  ) : null}
                  {sample ? (
                    <p id={ids.sample} class="xlsx-dialog__hint">
                      {t('dialogs.numberFormat.previewSample')}
                    </p>
                  ) : null}
                  {listed.length > 0 ? (
                    <label class="xlsx-dialog__label" for={ids.list}>
                      {t('dialogs.numberFormat.existing')}
                    </label>
                  ) : (
                    <span class="xlsx-dialog__label">{t('dialogs.numberFormat.existing')}</span>
                  )}
                  {listed.length > 0 ? (
                    <select
                      id={ids.list}
                      name="xlsx-number-format-existing"
                      size={Math.min(6, Math.max(2, listed.length))}
                      dir="ltr"
                      class="xlsx-dialog__list"
                      ref={list}
                      onChange={(e: Event) =>
                        pickCode((e.currentTarget as HTMLSelectElement).value)
                      }
                      // A double-click on a code applies it; the first click already set the field.
                      // With no listed code selected, the first arrow key starts at an end of the
                      // list.
                      onKeydown={(e: KeyboardEvent) => {
                        const el = e.currentTarget as HTMLSelectElement;
                        const down = e.key === 'ArrowDown';
                        if (el.selectedIndex !== -1 || (!down && e.key !== 'ArrowUp')) return;
                        e.preventDefault();
                        const option = el.options[down ? 0 : el.options.length - 1];
                        if (!option) return;
                        pickCode(option.value);
                      }}
                      onDblclick={(e: MouseEvent) => {
                        if ((e.target as Element).tagName === 'OPTION') apply();
                      }}
                    >
                      {listed.map((c) => (
                        <option key={c} value={c} title={c}>
                          {c}
                        </option>
                      ))}
                    </select>
                  ) : (
                    <p class="xlsx-dialog__hint">{t('dialogs.numberFormat.noExisting')}</p>
                  )}
                </div>
                <div class="xlsx-dialog__footer">
                  <button type="button" class="xlsx-dialog__button" onClick={props.onClose}>
                    {t('common.cancel')}
                  </button>
                  <button
                    type="button"
                    class="xlsx-dialog__button xlsx-dialog__button--primary"
                    data-slot="number.custom"
                    onClick={apply}
                  >
                    {t('common.apply')}
                  </button>
                </div>
              </>
            ),
          }}
        </DialogFrame>
      );
    };
  },
});
