/** @jsxImportSource react */
// The number format dialog of the `number.custom` slot (ECMA-376 Part 1, 18.8.31): a format
// code field, a live preview of the active cell's value through the engine's own renderer
// (`previewNumberFormat`), and the custom codes of the workbook (`workbookNumberFormatCodes`).
// Apply runs the slot with the code, so the store writes or reuses the `x:numFmt` in one
// transaction and one undo step. A code that the store would refuse shows its reason, and the
// field is marked invalid while it shows it. Apply then keeps the dialog open. A refusal of the
// engine (protection, a refused commit) shows in the dialog too, but leaves the field valid:
// the code is valid, and the selection or the sheet is what the engine refused.
//
// One host per toolbar renders the dialog, after DOCX Editor's `ParagraphDialogHost`
// (docx-editor/packages/react/src/editor/paragraph-dialog-host.tsx at 03735f11): the bar moves
// its controls between the bar and the "More" panel when it measures, and a dialog mounted in
// a control would be lost. The workbook's codes are file text: they render as text only.
//
// Difference: DOCX Editor exports its dialogs (`DocxEditorParagraphDialog` and its hook) from
// its React entry. This dialog is internal to `XlsxEditor.Toolbar`, as every toolbar part is,
// until the toolbar takes part children and host groups (task 8.4 of the active change); an
// export of this dialog and its opener is a follow-up of that task. A host with its own
// chrome uses `useEditorCommand('number.custom')` and the core functions that this dialog
// reads. A second editor dialog moves both to the shape of DOCX Editor's `DialogHost`
// registry, so a second host does not grow beside this one.

import {
  createContext,
  useCallback,
  useContext,
  useEffect,
  useId,
  useMemo,
  useLayoutEffect,
  useRef,
  useState,
  type ReactNode,
} from 'react';
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

const OpenerContext = createContext<NumberFormatDialogOpener | null>(null);

/** Closes the toolbar panel that holds an opener, such as the "More" panel. */
export const PanelCloseContext = createContext<(() => void) | null>(null);

/**
 * The opener of the toolbar's number format dialog, or null outside a toolbar. An open from the
 * mouse also closes the panel that holds the opener, because focus does not go back to it.
 */
export function useNumberFormatDialogOpener(): NumberFormatDialogOpener | null {
  const opener = useContext(OpenerContext);
  const closePanel = useContext(PanelCloseContext);
  return useMemo(
    () =>
      opener && closePanel
        ? {
            open(returnTo: HTMLElement | null) {
              const opened = opener.open(returnTo);
              if (opened && !returnTo) closePanel();
              return opened;
            },
          }
        : opener,
    [opener, closePanel]
  );
}

/** Renders the number format dialog of a toolbar, once, beside its controls. */
export function NumberFormatDialogHost({ children }: { readonly children: ReactNode }) {
  const editor = useXlsxEditor();
  const [open, setOpen] = useState<{
    readonly returnTo: HTMLElement | null;
    readonly start: NumberFormatDialogStart;
  } | null>(null);
  const opener = useMemo<NumberFormatDialogOpener>(
    () => ({
      open: (returnTo) => {
        // The start commits an open cell edit; a refused commit keeps the dialog closed, and
        // the status message gives the reason.
        const start = editor ? numberFormatDialogStart(editor) : null;
        if (start) setOpen({ returnTo, start });
        return start !== null;
      },
    }),
    [editor]
  );
  const close = useCallback(() => setOpen(null), []);
  return (
    <OpenerContext.Provider value={opener}>
      {children}
      {open ? (
        <NumberFormatDialog returnTo={open.returnTo} start={open.start} onClose={close} />
      ) : null}
    </OpenerContext.Provider>
  );
}

const HEX6 = /^[0-9A-Fa-f]{6}$/;

function NumberFormatDialog({
  returnTo,
  start,
  onClose,
}: {
  readonly returnTo: HTMLElement | null;
  readonly start: NumberFormatDialogStart;
  readonly onClose: () => void;
}) {
  const editor = useXlsxEditor();
  const { t } = useTranslation();
  const s = useToolbarSlot('number.custom');
  // The code that the dialog opened with: `numberFormatDialogApplyStep` decides when an Apply of
  // it unchanged writes nothing.
  const initial = start.initial;
  const [code, setCode] = useState(initial);
  const [error, setError] = useState<string | null>(null);
  const [errorCount, setErrorCount] = useState(0);
  const applied = useRef(false);
  // After an Apply that the code check refused, the field stays invalid until its code reads.
  const [attempted, setAttempted] = useState(false);
  const field = useRef<HTMLInputElement>(null);
  const list = useRef<HTMLSelectElement>(null);
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
  useEffect(() => field.current?.select(), []);
  // Each change of the editor state renders the dialog again, as DOCX Editor's
  // `useDialogDocument` listens to each change: the preview reads the cell anew, and the dialog
  // closes when its start no longer holds (another workbook, a moved selection, or a disabled
  // slot such as view mode or protection that locks formats).
  const snapshot = useEditorSnapshot();
  // The workbook's codes, read again on each editor change: a host call can add a code while
  // the dialog is open. A keystroke in the field does not read them again: `snapshot` changes
  // on each editor change only.
  const codes = useMemo(
    () => (editor && snapshot ? workbookNumberFormatCodes(editor) : []),
    [editor, snapshot]
  );
  // The close check of the host procedure in `packages/react/docs/toolbar-and-keyboard.md`.
  const current = useEditorState(
    useCallback(() => !!editor && numberFormatDialogCurrent(editor, start), [editor, start])
  );
  useEffect(() => {
    if (!current) onClose();
  }, [current, onClose]);
  const preview = editor ? previewNumberFormat(editor, code) : null;
  const problem =
    preview && !preview.ok ? translateKey(t, preview.problem.key, preview.problem.params) : null;
  // A code that reads ends the invalid state. A field with a code that the preview refuses is
  // invalid at once; an empty field is invalid only after an Apply.
  if (attempted && problem === null) setAttempted(false);
  const invalid = problem !== null && (attempted || code !== '');
  // A list box with no matching value selects its first option in some browsers, and a click
  // on a selected option changes nothing: select none when the field holds no listed code.
  useLayoutEffect(() => {
    const el = list.current;
    if (!el) return;
    if (codes.includes(code)) el.value = code;
    else el.selectedIndex = -1;
  }, [code, codes]);
  const apply = () => {
    if (!editor) return;
    const step = numberFormatDialogApplyStep(editor, start, code);
    if (step.kind === 'closed' || step.kind === 'unchanged') {
      // An unchanged Apply returns the keyboard to the grid, as a writing Apply does.
      if (step.kind === 'unchanged') applied.current = true;
      onClose();
      return;
    }
    if (step.kind === 'problem') {
      setAttempted(true);
      setError(translateKey(t, step.problem.key, step.problem.params));
      setErrorCount((n) => n + 1);
      field.current?.focus();
      return;
    }
    if (s.execute(code)) {
      applied.current = true;
      onClose();
      return;
    }
    const after = editor.snapshot();
    // A refusal with no reason in the snapshot still tells the user that nothing changed.
    setError(
      after.refusal
        ? refusalText(after.refusal, t, after.numberLocale)
        : t('dialogs.numberFormat.notApplied')
    );
    setErrorCount((n) => n + 1);
  };
  const sample = !!preview?.ok && preview.sample;
  // A code can hide the value (`;;;`): then a note says so, as the preview box is empty.
  const blank = !!preview?.ok && preview.text === '';
  const shownColor =
    preview?.ok && preview.color && HEX6.test(preview.color) ? preview.color : null;
  return (
    <DialogFrame
      kind="numberFormat"
      label={t('dialogs.numberFormat.title')}
      onClose={onClose}
      onRestoreFocus={() => {
        if (applied.current && editor) {
          editor.focus();
          return true;
        }
        if (!returnTo?.isConnected) return false;
        returnTo.focus({ preventScroll: true });
        return true;
      }}
      onKeyDown={(event) => {
        // Enter applies from the field and from the list; a button keeps its own Enter.
        const tag = (event.target as Element).tagName;
        if (event.key !== 'Enter' || (tag !== 'INPUT' && tag !== 'SELECT')) return;
        event.preventDefault();
        apply();
      }}
    >
      <div className="xlsx-dialog__header">
        <h2 className="xlsx-dialog__title">{t('dialogs.numberFormat.title')}</h2>
      </div>
      <div className="xlsx-dialog__body">
        <label className="xlsx-dialog__label" htmlFor={ids.field}>
          {t('dialogs.numberFormat.code')}
        </label>
        <input
          ref={field}
          id={ids.field}
          name="xlsx-number-format-code"
          type="text"
          dir="ltr"
          autoComplete="off"
          spellCheck={false}
          className="xlsx-dialog__input"
          value={code}
          aria-invalid={invalid}
          aria-describedby={[
            ids.hint,
            ...(error ? [ids.error] : [ids.previewLabel, ids.preview]),
            sample ? ids.sample : '',
            blank && !error ? ids.blank : '',
          ]
            .filter(Boolean)
            .join(' ')}
          onChange={(e) => {
            setCode(e.currentTarget.value);
            setError(null);
          }}
        />
        {/* The reason of a refused Apply shows under the field, in place of the preview. */}
        {/* A new key on each refused Apply mounts the alert again, so it is announced again. */}
        <p key={errorCount} id={ids.error} role="alert" className="xlsx-dialog__error">
          {error}
        </p>
        <p id={ids.hint} className="xlsx-dialog__hint">
          {t(
            codes.length > 0
              ? 'dialogs.numberFormat.codeHint'
              : 'dialogs.numberFormat.codeHintNoList'
          )}
        </p>
        <div className="xlsx-dialog__preview" hidden={error !== null}>
          <span id={ids.previewLabel} className="xlsx-dialog__preview-label">
            {t('dialogs.numberFormat.preview')}
          </span>
          <output
            id={ids.preview}
            htmlFor={ids.field}
            aria-labelledby={ids.previewLabel}
            // The field's description carries the preview: no announcement for each key.
            aria-live="off"
            className="xlsx-dialog__preview-value"
            data-problem={problem ? '' : undefined}
            style={shownColor ? { color: `#${shownColor}` } : undefined}
          >
            {preview?.ok
              ? preview.stale
                ? t('grid.staleResult', { text: preview.text })
                : preview.text
              : (problem ?? '')}
          </output>
        </div>
        {blank && error === null ? (
          <p id={ids.blank} className="xlsx-dialog__hint">
            {t('dialogs.numberFormat.previewBlank')}
          </p>
        ) : null}
        {sample ? (
          <p id={ids.sample} className="xlsx-dialog__hint">
            {t('dialogs.numberFormat.previewSample')}
          </p>
        ) : null}
        {codes.length > 0 ? (
          <label className="xlsx-dialog__label" htmlFor={ids.list}>
            {t('dialogs.numberFormat.existing')}
          </label>
        ) : (
          <span className="xlsx-dialog__label">{t('dialogs.numberFormat.existing')}</span>
        )}
        {codes.length > 0 ? (
          <select
            id={ids.list}
            name="xlsx-number-format-existing"
            size={Math.min(6, Math.max(2, codes.length))}
            dir="ltr"
            className="xlsx-dialog__list"
            ref={list}
            onChange={(e) => {
              setCode(e.currentTarget.value);
              setError(null);
            }}
            // A double-click on a code applies it; the first click already set the field.
            // With no listed code selected, the first arrow key starts at an end of the list.
            onKeyDown={(e) => {
              const el = e.currentTarget;
              const down = e.key === 'ArrowDown';
              if (el.selectedIndex !== -1 || (!down && e.key !== 'ArrowUp')) return;
              e.preventDefault();
              const option = el.options[down ? 0 : el.options.length - 1];
              if (!option) return;
              setCode(option.value);
              setError(null);
            }}
            onDoubleClick={(e) => {
              if ((e.target as Element).tagName === 'OPTION') apply();
            }}
          >
            {codes.map((c) => (
              <option key={c} value={c} title={c}>
                {c}
              </option>
            ))}
          </select>
        ) : (
          <p className="xlsx-dialog__hint">{t('dialogs.numberFormat.noExisting')}</p>
        )}
      </div>
      <div className="xlsx-dialog__footer">
        <button type="button" className="xlsx-dialog__button" onClick={onClose}>
          {t('common.cancel')}
        </button>
        <button
          type="button"
          className="xlsx-dialog__button xlsx-dialog__button--primary"
          data-slot="number.custom"
          onClick={apply}
        >
          {t('common.apply')}
        </button>
      </div>
    </DialogFrame>
  );
}
