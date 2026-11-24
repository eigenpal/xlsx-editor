/** @jsxImportSource vue */
// The formula bar: the name box and the content field of the active cell. The Vue twin of the
// React adapter's `FormulaBar.tsx`; the rules are the same, and the header there explains them.
// In short: the content field edits the open cell edit of the engine (`setEditText`), Enter and
// Tab commit (`commitEdit(move)`), Escape cancels, blur does not commit, Alt+Enter adds a line
// break, and the name box keeps a draft of the typed reference until Enter or blur. The key
// rules are the engine's (`editMoveOf`, `composesKey`, `fieldHistoryKey`), so each adapter's bar
// acts the same.
//
// Differences from React:
// - Vue does not restore a controlled value that the engine refused: nothing renders again. A
//   refused `setEditText` puts the shown text back into the field by hand.
// - The field shows its caret line (`showCaretLine`) after each render (`onMounted` and
//   `onUpdated`), and on the native events that move the caret: `select`, `keyup`, `mouseup`,
//   `focus`, and the document's `selectionchange`, as React's synthetic `onSelect` listens.
// - The name box reads the refusal from the snapshot and translates it in its render, so a
//   change of the root's `locale` translates it again.

import {
  defineComponent,
  onBeforeUnmount,
  onMounted,
  onUpdated,
  shallowRef,
  useId,
  watch,
  type ExtractPublicPropTypes,
  type PropType,
} from 'vue';
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
  type EditorSnapshot,
  type XlsxEditorInstance,
  type ChromePartTranslate,
} from '@xlsx-editor.dev/core/editor';
import { useEditorState, useTranslation, useXlsxEditor } from './context.ts';
import { usePartTranslation } from './locale.ts';

const formulaBarProps = {
  /**
   * The label resolver of the bar: a key and the values of its placeholders, as DOCX Editor's
   * `t` prop. A key that it leaves unresolved reads the root's `locale`. Keep it stable, for
   * example from `useChromeTranslate(overrides)`.
   */
  t: { type: Function as PropType<ChromePartTranslate>, default: undefined },
} as const;

/** Props for `FormulaBar`, the twin of React's `FormulaBarProps`. @public */
export type FormulaBarProps = ExtractPublicPropTypes<typeof formulaBarProps>;

const refusalOf = (s: EditorSnapshot) => ({ refusal: s.refusal, numberLocale: s.numberLocale });

/**
 * The active cell's reference and content, both editable. The content is formula text, or the
 * value in a form that reads back, such as `2026-11-12` for a date. @public
 */
export const FormulaBar = defineComponent({
  name: 'FormulaBar',
  props: formulaBarProps,
  setup(props) {
    const { t, dir } = usePartTranslation('FormulaBar', () => props.t);
    const editor = useXlsxEditor();
    const snap = useEditorState(formulaSlice, sameSlice);
    let composing = false;
    let field: HTMLTextAreaElement | null = null;
    // The caret after a line break that Alt+Enter added, placed after the next render.
    let caret: number | null = null;
    // The text of the open edit when the engine refused its commit: invalid while it stays.
    const refusedText = shallowRef<string | null>(null);
    const hintId = useId();
    const reasonId = useId();
    const shortenedId = useId();
    // The refusal at the last run: a refusal that came before the edit opened, such as a
    // refused sheet rename, does not refuse the text of the edit. It starts empty, so a field
    // that mounts during a refused edit is marked invalid.
    let seen: EditorSnapshot['refusal'] = null;
    // Before the render, so the field is marked invalid in the same frame as the refusal.
    watch(
      () => [snap.value.refusal, snap.value.editing, snap.value.text] as const,
      ([refusal, editing, text]) => {
        if (!editing) refusedText.value = null;
        else if (refusal !== seen)
          refusedText.value = refusal && refusesEditText(refusal.code) ? text : null;
        seen = refusal;
        // Only a new refusal, or the end of the edit, sets the refused text. A run for typed
        // text finds the same refusal and changes nothing.
      },
      { immediate: true }
    );
    const afterRender = (): void => {
      const at = caret;
      caret = null;
      if (at !== null) field?.setSelectionRange(at, at);
      showCaretLine(field);
    };
    onMounted(afterRender);
    onUpdated(afterRender);
    const onSelectionChange = (): void => showCaretLine(field);
    // The document of the field's own element, not the script's global (AGENTS.md, host
    // window). `bindField` adds the listener when the field mounts.
    let doc: Document | null = null;
    // The field unmounts with no end of its composition (another workbook loads, or the host
    // removes the bar): the composition ends for the engine and for the field.
    const endOpenComposition = (): void => {
      if (!composing) return;
      composing = false;
      const s = editor.value?.snapshot();
      if (s?.editing) editor.value?.setEditText(s.editText ?? '', { composing: false });
    };
    const bindField = (el: unknown): void => {
      const next = el as HTMLTextAreaElement | null;
      if (next === field) return;
      field = next;
      if (!next) endOpenComposition();
      else if (doc !== next.ownerDocument) {
        doc?.removeEventListener('selectionchange', onSelectionChange);
        doc = next.ownerDocument;
        doc.addEventListener('selectionchange', onSelectionChange);
      }
    };
    onBeforeUnmount(() => {
      doc?.removeEventListener('selectionchange', onSelectionChange);
      doc = null;
      endOpenComposition();
    });
    const type = (value: string, composes: boolean): boolean =>
      editor.value?.setEditText(value, { composing: composes }).ok === true;
    const nameBoxFocus = (): void => field?.focus();

    return () => {
      const s = snap.value;
      if (s.status !== 'ready') return null;
      const { editing, refusal, text } = s;
      const shown = contentFieldText(text);
      const shortened = shown.length < text.length;
      const readOnly = s.view || shortened;
      const shortenedNote = `${t('formulaBar.shortened')}${s.view ? '' : ` ${t('formulaBar.shortenedEdit')}`}`;
      const endComposition = (value: string): void => {
        if (!composing) return;
        composing = false;
        const open = editor.value?.snapshot();
        if (open?.editing) type(shortened ? (open.editText ?? '') : value, false);
      };
      const invalid = editing && refusal !== null && refusedText.value === text;
      const onKeydown = (e: KeyboardEvent): void => {
        const ed = editor.value;
        if (!ed || composesKey(e, composing)) return;
        if (stepFieldHistory(e, ed, editing)) return;
        const el = e.currentTarget as HTMLTextAreaElement;
        const move = editMoveOf(e);
        if (move) {
          e.preventDefault();
          ed.commitEdit(move);
        } else if (e.key === 'Enter' && e.altKey && !readOnly) {
          // A line break at the caret, as Alt+Enter adds in the cell editor. The field inserts
          // it as typed text, so the field's own undo can remove it, and the input event gives
          // the text to the engine; a refused edit restores the field. Without the insert
          // command, the engine gets the text first, and the field shows it only when the
          // engine takes it.
          e.preventDefault();
          if (insertFieldText(el, '\n')) return;
          const { value, selectionStart, selectionEnd } = el;
          const next = `${value.slice(0, selectionStart)}\n${value.slice(selectionEnd)}`;
          if (type(next, false)) caret = selectionStart + 1;
        } else if (e.key === 'Escape') {
          e.preventDefault();
          e.stopPropagation();
          if (ed.snapshot().editing) ed.cancelEdit();
          else ed.focus();
        }
      };
      const lines = shortened ? 1 : text.split('\n').length;
      const reason = invalid && refusal ? refusalText(refusal, t, s.numberLocale) : '';
      const showLine = (e: Event): void => showCaretLine(e.currentTarget as HTMLTextAreaElement);
      return (
        <div class="xlsx-formula-bar" dir={dir.value} onMousedown={keepGridFocus}>
          <NameBox reference={s.ref ?? ''} focusContent={nameBoxFocus} />
          <div
            class="xlsx-cell-content-box"
            data-lines={lines > 1 ? '' : undefined}
            onMousedown={(e: MouseEvent) => pressFieldBox(field, e)}
          >
            <textarea
              // The text sets its own direction, as in the cell: a formula reads left to right in a
              // right-to-left chrome, and a text keeps the direction of its script.
              dir="auto"
              ref={bindField}
              class="xlsx-cell-content"
              name="xlsx-cell-content"
              rows={1}
              wrap="off"
              spellcheck={false}
              autocomplete="off"
              autocapitalize="off"
              aria-label={t('formulaBar.content')}
              aria-describedby={
                shortened
                  ? shortenedId
                  : s.view
                    ? undefined
                    : invalid
                      ? `${reasonId} ${hintId}`
                      : hintId
              }
              readonly={readOnly}
              aria-invalid={invalid || undefined}
              value={shown}
              onInput={(e: Event) => {
                const el = e.currentTarget as HTMLTextAreaElement;
                const composes = composing || (e as InputEvent).isComposing === true;
                // Vue renders nothing for a refused text: the field shows the edit again.
                if (shortened || !type(el.value, composes)) {
                  const open = editor.value?.snapshot();
                  const now = contentFieldText(open?.editing ? (open.editText ?? '') : text);
                  if (el.value !== now) el.value = now;
                }
              }}
              onCompositionstart={() => {
                composing = true;
              }}
              // The last input event can come before the end of the composition: the edit
              // learns here that its text is no longer composed.
              onCompositionend={(e: CompositionEvent) =>
                endComposition((e.currentTarget as HTMLTextAreaElement).value)
              }
              onBlur={(e: FocusEvent) =>
                endComposition((e.currentTarget as HTMLTextAreaElement).value)
              }
              onKeydown={onKeydown}
              onKeyup={showLine}
              onMouseup={showLine}
              onFocus={showLine}
              onSelect={showLine}
            />
            {shortened ? (
              <span
                id={shortenedId}
                class="xlsx-cell-content-lines xlsx-cell-content-note"
                title={shortenedNote}
              >
                {shortenedNote}
              </span>
            ) : lines > 1 ? (
              <span class="xlsx-cell-content-lines" aria-hidden="true">
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
    };
  },
});

/** The name box: the selected range, and a field to type a reference to select. */
const NameBox = defineComponent({
  name: 'XlsxNameBox',
  props: {
    reference: { type: String, required: true },
    /** Give the keyboard to the content field. */
    focusContent: { type: Function as PropType<() => void>, required: true },
  },
  setup(props) {
    const { t } = useTranslation();
    const editor = useXlsxEditor();
    const shown = useEditorState(refusalOf, sameSlice);
    const draft = shallowRef<string | null>(null);
    const refused = shallowRef(false);
    let composing = false;
    const hintId = useId();
    const reasonId = useId();
    // Whether the name box raised the refusal that shows: typing after it does not change that.
    let raised = false;
    const reset = (): void => {
      draft.value = null;
      refused.value = false;
      raised = false;
    };
    // Drop the typed reference, and the reason of a refused one with it.
    const drop = (ed: XlsxEditorInstance): void => {
      if (raised && refusesReference(ed.snapshot().refusal?.code ?? 'loading')) ed.dismissRefusal();
      reset();
    };
    const go = (ed: XlsxEditorInstance, text: string): void => {
      if (text.trim() === '') {
        // An empty name box selects nothing and is no error: it shows the selection again.
        drop(ed);
        ed.focus();
        return;
      }
      const r = ed.selectReference(text);
      if (r.ok) {
        reset();
        ed.focus();
      } else if (refusesReference(r.refusal.code)) {
        draft.value = text;
        refused.value = true;
        raised = true;
      } else {
        // The open cell edit refused its commit, not the reference: the content field holds
        // that edit and its reason, so it takes the keyboard.
        reset();
        props.focusContent();
      }
    };
    return () => {
      const { refusal, numberLocale } = shown.value;
      const reason = refused.value && refusal ? refusalText(refusal, t, numberLocale) : '';
      return [
        <input
          class="xlsx-name-box"
          // The text sets its own direction: a reference reads left to right in a right-to-left
          // chrome, and a defined name keeps the direction of its script.
          dir="auto"
          name="xlsx-name-box"
          type="text"
          spellcheck={false}
          autocomplete="off"
          autocapitalize="off"
          maxlength={MAX_GO_TO_LENGTH}
          aria-label={t('formulaBar.nameBox')}
          aria-describedby={refused.value ? `${reasonId} ${hintId}` : hintId}
          aria-invalid={refused.value || undefined}
          value={draft.value ?? props.reference}
          onFocus={(e: FocusEvent) => (e.currentTarget as HTMLInputElement).select()}
          onInput={(e: Event) => {
            draft.value = (e.currentTarget as HTMLInputElement).value;
            refused.value = false;
          }}
          onBlur={() => {
            if (editor.value) drop(editor.value);
            else reset();
          }}
          onCompositionstart={() => {
            composing = true;
          }}
          onCompositionend={() => {
            composing = false;
          }}
          onKeydown={(e: KeyboardEvent) => {
            const ed = editor.value;
            if (!ed || composesKey(e, composing)) return;
            // An open cell edit, or a typed reference, keeps the keys for the typing in a field.
            if (stepFieldHistory(e, ed, draft.value !== null || ed.snapshot().editing)) return;
            if (e.key === 'Enter') {
              e.preventDefault();
              go(ed, (e.currentTarget as HTMLInputElement).value);
            } else if (e.key === 'Escape') {
              e.preventDefault();
              e.stopPropagation();
              drop(ed);
              ed.focus();
            }
          }}
        />,
        <span id={hintId} hidden>
          {t('formulaBar.nameBoxHint')}
        </span>,
        <span id={reasonId} hidden>
          {reason}
        </span>,
      ];
    };
  },
});
