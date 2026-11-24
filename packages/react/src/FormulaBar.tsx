/** @jsxImportSource react */
// The formula bar: the name box and the content field of the active cell.
//
// The content field edits the open cell edit of the engine, never a copy of it. Each change
// calls `setEditText`, which opens the edit over the active cell when none is open; the grid
// shows the same text in the cell, and the field shows `snapshot().editText`. Enter and Tab
// commit through `commitEdit(move)`, the commit of the cell editor, and Escape cancels. A
// refused commit keeps the edit and the keyboard in the field, and the status strip shows the
// engine's reason, as for the grid. This differs from the chrome fields of DOCX Editor and
// PPTX Editor, which keep a draft and commit it on Enter and on blur: a cell has one edit,
// and the grid shows it while the user types in the bar. For the same reason, blur does not
// commit: the edit stays open, and the next command or a press on the grid commits it.
//
// The content field is a one-row text area, not an input: an input drops the line breaks of a
// cell's text, so an edit in it would change the cell without a word. Alt+Enter adds a line
// break, as in the cell editor, and a line count beside the text area shows the lines of a
// text with more than one. A press on the line count, or on the gap beside the text area,
// puts the caret at the end of the line that the field shows.
//
// The name box keeps a draft of the typed reference, because it edits no cell. Enter selects
// the reference through `selectReference`; a refused reference keeps the text and marks the
// field invalid. Blur drops the draft and its reason: the name box shows the selection while
// it does not have the keyboard, so a typed reference cannot stay in it. DOCX Editor's Go To
// box keeps its query in a pane that stays open, so it has no blur rule.
//
// An invalid field names its reason (`aria-describedby`), as the fields of DOCX Editor and PPTX
// Editor do; the status strip shows the same reason, and the grid's live region announces it
// once. The content field is invalid while it holds the text whose commit the engine refused,
// by Enter, Tab, F6, or another command.
//
// The key rules are the engine's (`editMoveOf`, `composesKey`, `fieldHistoryKey`), so each
// adapter's bar acts the same. While an input method composes, no key commits, cancels, or
// selects. A composition that ends without its end event (the field loses the keyboard or
// unmounts) is ended for the engine on blur and on unmount.
//
// A text past `MAX_CONTENT_FIELD_CHARS` shows shortened and read-only, and a visible note says so.
// The open edit keeps the complete text, so the field never gives a shortened text to the engine.

import { useCallback, useId, useLayoutEffect, useRef, useState } from 'react';
import type { KeyboardEvent, ReactElement } from 'react';
import {
  MAX_GO_TO_LENGTH,
  composesKey,
  contentFieldText,
  editMoveOf,
  formulaSlice,
  insertFieldText,
  keepGridFocus,
  pressFieldBox,
  refusalText,
  refusesEditText,
  refusesReference,
  sameSlice,
  showCaretLine,
  stepFieldHistory,
} from '@xlsx-editor.dev/core/editor';
import type {
  ChromePartTranslate,
  EditorSnapshot,
  XlsxEditorInstance,
} from '@xlsx-editor.dev/core/editor';
import { useTranslation, useEditorState, useXlsxEditor } from './context.ts';
import { useChromeDir, withChromeTranslate } from './locale.tsx';

/** Props for `XlsxEditor.FormulaBar`. @public */
export interface FormulaBarProps {
  /**
   * The label resolver of the bar: a key and the values of its placeholders, as DOCX Editor's
   * `t` prop. A key that it leaves unresolved reads the root's `locale`. Keep it stable, for
   * example from `useChromeTranslate(overrides)`.
   */
  readonly t?: ChromePartTranslate;
}

/**
 * The active cell's reference and content, both editable. The content is formula text, or the
 * value in a form that reads back, such as `2026-11-12` for a date. @public
 */
export const FormulaBar: (props: FormulaBarProps) => ReactElement = withChromeTranslate(
  FormulaBarPart,
  'FormulaBar'
);

function FormulaBarPart() {
  const { t } = useTranslation();
  const dir = useChromeDir();
  const editor = useXlsxEditor();
  const snap = useEditorState(formulaSlice, sameSlice);
  const composing = useRef(false);
  const field = useRef<HTMLTextAreaElement | null>(null);
  // The caret after a line break that Alt+Enter added, placed after the next render.
  const caret = useRef<number | null>(null);
  // The text of the open edit when the engine refused its commit: invalid while it stays.
  const [refusedText, setRefusedText] = useState<string | null>(null);
  const hintId = useId();
  const reasonId = useId();
  const shortenedId = useId();
  const { editing, refusal, text } = snap;
  // The refusal at the last run: a refusal that came before the edit opened, such as a
  // refused sheet rename, does not refuse the text of the edit.
  // It starts empty, so a field that mounts during a refused edit is marked invalid.
  const seen = useRef<EditorSnapshot['refusal']>(null);
  // A layout effect, so the field is marked invalid in the same frame as the refusal.
  useLayoutEffect(() => {
    if (!editing) setRefusedText(null);
    else if (refusal !== seen.current)
      setRefusedText(refusal && refusesEditText(refusal.code) ? text : null);
    seen.current = refusal;
    // Only a new refusal, or the end of the edit, sets the refused text. A run for typed text
    // finds the same refusal and changes nothing.
  }, [refusal, editing, text]);
  useLayoutEffect(() => {
    const at = caret.current;
    caret.current = null;
    if (at !== null) field.current?.setSelectionRange(at, at);
    showCaretLine(field.current);
  });
  // The field unmounts with no end of its composition (another workbook loads, or the host
  // removes the bar): the composition ends for the engine and for the field.
  const bindField = useCallback(
    (el: HTMLTextAreaElement | null) => {
      field.current = el;
      if (el || !composing.current) return;
      composing.current = false;
      const s = editor?.snapshot();
      if (s?.editing) editor?.setEditText(s.editText ?? '', { composing: false });
    },
    [editor]
  );
  if (snap.status !== 'ready') return null;
  const type = (value: string, composes: boolean) =>
    editor?.setEditText(value, { composing: composes }).ok === true;
  const shown = contentFieldText(text);
  const shortened = shown.length < text.length;
  const readOnly = snap.view || shortened;
  const shortenedNote = `${t('formulaBar.shortened')}${snap.view ? '' : ` ${t('formulaBar.shortenedEdit')}`}`;
  const endComposition = (value: string) => {
    if (!composing.current) return;
    composing.current = false;
    const open = editor?.snapshot();
    if (open?.editing) type(shortened ? (open.editText ?? '') : value, false);
  };
  const invalid = editing && refusal !== null && refusedText === text;
  const onKeyDown = (e: KeyboardEvent<HTMLTextAreaElement>) => {
    if (!editor || composesKey(e.nativeEvent, composing.current)) return;
    if (stepFieldHistory(e, editor, editing)) return;
    const move = editMoveOf(e);
    if (move) {
      e.preventDefault();
      editor.commitEdit(move);
    } else if (e.key === 'Enter' && e.altKey && !readOnly) {
      // A line break at the caret, as Alt+Enter adds in the cell editor. The field inserts it
      // as typed text, so the field's own undo can remove it, and the change event gives the
      // text to the engine; a refused edit restores the field. Without the insert command,
      // the engine gets the text first, and the field shows it only when the engine takes it.
      e.preventDefault();
      if (insertFieldText(e.currentTarget, '\n')) return;
      const { value, selectionStart, selectionEnd } = e.currentTarget;
      const next = `${value.slice(0, selectionStart)}\n${value.slice(selectionEnd)}`;
      if (type(next, false)) caret.current = selectionStart + 1;
    } else if (e.key === 'Escape') {
      e.preventDefault();
      e.stopPropagation();
      if (editor.snapshot().editing) editor.cancelEdit();
      else editor.focus();
    }
  };
  const lines = shortened ? 1 : text.split('\n').length;
  const reason = invalid && refusal ? refusalText(refusal, t, snap.numberLocale) : '';
  return (
    <div className="xlsx-formula-bar" dir={dir} onMouseDown={keepGridFocus}>
      <NameBox reference={snap.ref ?? ''} contentField={field} />
      <div
        className="xlsx-cell-content-box"
        data-lines={lines > 1 ? '' : undefined}
        onMouseDown={(e) => pressFieldBox(field.current, e)}
      >
        <textarea
          // The text sets its own direction, as in the cell: a formula reads left to right in a
          // right-to-left chrome, and a text keeps the direction of its script.
          dir="auto"
          ref={bindField}
          className="xlsx-cell-content"
          name="xlsx-cell-content"
          rows={1}
          wrap="off"
          spellCheck={false}
          autoComplete="off"
          autoCapitalize="off"
          aria-label={t('formulaBar.content')}
          aria-describedby={
            shortened
              ? shortenedId
              : snap.view
                ? undefined
                : invalid
                  ? `${reasonId} ${hintId}`
                  : hintId
          }
          readOnly={readOnly}
          aria-invalid={invalid || undefined}
          value={shown}
          onChange={(e) => {
            if (shortened) return;
            const event = e.nativeEvent as { readonly isComposing?: boolean };
            type(e.currentTarget.value, composing.current || event.isComposing === true);
          }}
          onCompositionStart={() => {
            composing.current = true;
          }}
          // The last input event can come before the end of the composition: the edit learns
          // here that its text is no longer composed.
          onCompositionEnd={(e) => endComposition(e.currentTarget.value)}
          onBlur={(e) => endComposition(e.currentTarget.value)}
          onKeyDown={onKeyDown}
          onSelect={(e) => showCaretLine(e.currentTarget)}
        />
        {shortened ? (
          <span
            id={shortenedId}
            className="xlsx-cell-content-lines xlsx-cell-content-note"
            title={shortenedNote}
          >
            {shortenedNote}
          </span>
        ) : lines > 1 ? (
          <span className="xlsx-cell-content-lines" aria-hidden="true">
            {t('formulaBar.lines', { count: lines })}
          </span>
        ) : null}
      </div>
      <span id={hintId} hidden>
        {t('formulaBar.contentHint')}
      </span>
      <span id={reasonId} hidden>
        {reason}
      </span>
    </div>
  );
}

/** The name box: the selected range, and a field to type a reference to select. */
function NameBox(props: {
  reference: string;
  contentField: { readonly current: HTMLTextAreaElement | null };
}) {
  const { t } = useTranslation();
  const editor = useXlsxEditor();
  const reason = useEditorState(
    (s) => (s.refusal ? refusalText(s.refusal, t, s.numberLocale) : ''),
    Object.is
  );
  const [draft, setDraft] = useState<string | null>(null);
  const [refused, setRefused] = useState(false);
  const composing = useRef(false);
  const hintId = useId();
  const reasonId = useId();
  // Whether the name box raised the refusal that shows: typing after it does not change that.
  const raised = useRef(false);
  const reset = () => {
    setDraft(null);
    setRefused(false);
    raised.current = false;
  };
  // Drop the typed reference, and the reason of a refused one with it.
  const drop = (editor: XlsxEditorInstance) => {
    if (raised.current && refusesReference(editor.snapshot().refusal?.code ?? 'loading'))
      editor.dismissRefusal();
    reset();
  };
  const go = (editor: XlsxEditorInstance, text: string) => {
    if (text.trim() === '') {
      // An empty name box selects nothing and is no error: it shows the selection again.
      drop(editor);
      editor.focus();
      return;
    }
    const r = editor.selectReference(text);
    if (r.ok) {
      reset();
      editor.focus();
    } else if (refusesReference(r.refusal.code)) {
      setDraft(text);
      setRefused(true);
      raised.current = true;
    } else {
      // The open cell edit refused its commit, not the reference: the content field holds
      // that edit and its reason, so it takes the keyboard.
      reset();
      props.contentField.current?.focus();
    }
  };
  return (
    <>
      <input
        className="xlsx-name-box"
        // The text sets its own direction: a reference reads left to right in a right-to-left
        // chrome, and a defined name keeps the direction of its script.
        dir="auto"
        name="xlsx-name-box"
        type="text"
        spellCheck={false}
        autoComplete="off"
        autoCapitalize="off"
        maxLength={MAX_GO_TO_LENGTH}
        aria-label={t('formulaBar.nameBox')}
        aria-describedby={refused ? `${reasonId} ${hintId}` : hintId}
        aria-invalid={refused || undefined}
        value={draft ?? props.reference}
        onFocus={(e) => e.currentTarget.select()}
        onChange={(e) => {
          setDraft(e.currentTarget.value);
          setRefused(false);
        }}
        onBlur={() => {
          if (editor) drop(editor);
          else reset();
        }}
        onCompositionStart={() => {
          composing.current = true;
        }}
        onCompositionEnd={() => {
          composing.current = false;
        }}
        onKeyDown={(e) => {
          if (!editor || composesKey(e.nativeEvent, composing.current)) return;
          // An open cell edit, or a typed reference, keeps the keys for the typing in a field.
          if (stepFieldHistory(e, editor, draft !== null || editor.snapshot().editing)) return;
          if (e.key === 'Enter') {
            e.preventDefault();
            go(editor, e.currentTarget.value);
          } else if (e.key === 'Escape') {
            e.preventDefault();
            e.stopPropagation();
            drop(editor);
            editor.focus();
          }
        }}
      />
      <span id={hintId} hidden>
        {t('formulaBar.nameBoxHint')}
      </span>
      <span id={reasonId} hidden>
        {refused ? reason : ''}
      </span>
    </>
  );
}
