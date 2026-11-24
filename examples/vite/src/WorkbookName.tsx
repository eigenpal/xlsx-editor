/** @jsxImportSource react */
// The name of the open workbook in the demo top bar, as a text field, after the deck name of the
// PPTX Editor demo (pptx-editor/examples/vite/src/App.tsx). The Vue demo has the same field in
// examples/vue/src/App.vue, with the same rules from `workbook-names.ts`. Its text is the demo's
// own English and stays out of `packages/i18n` (AGENTS.md, "i18n").
//
// Differences from the source, each on purpose:
// - Escape puts back the name before the edit, and an empty name keeps the name before it
//   (`renamedTo`). The source keeps what the field holds.
// - Enter and Escape give the focus back to the grid, so the keys go on to the cells.
// - The name never leaves the bar, so File > Rename workbook focuses the field in place.

import { useRef, useState, type KeyboardEvent, type RefObject } from 'react';
import { useEditorState, type EditorSnapshot } from '@xlsx-editor.dev/react';
import { NAME_FIELD_LABEL, renamedTo, shownExtension, shownName } from './workbook-names.ts';

const kindOf = (snapshot: EditorSnapshot) => snapshot.workbookKind;

/**
 * The name of the open workbook, with its extension muted, as the family demos show a file. With
 * no workbook ready, the place stays empty. The name is a text field, as the deck name of the
 * PPTX Editor demo: a click or a double click edits it. Enter and a move of the focus away keep
 * the typed name (`renamedTo`). Escape puts back the name before the edit. Enter and Escape
 * give the focus back to the grid. The muted extension is the one that Save writes for the
 * package kind (`shownExtension`).
 */
export function WorkbookName(props: {
  label: string | null;
  chosen: string | null;
  field: RefObject<HTMLInputElement | null>;
  onRename: (name: string) => void;
  onDone: () => void;
}) {
  const { label, chosen, field, onRename, onDone } = props;
  const kind = useEditorState(kindOf);
  const shown = label === null ? null : shownName(label);
  const name = chosen ?? shown?.name ?? '';
  const [draft, setDraft] = useState<string | null>(null);
  // Set by Escape, so the blur that follows keeps the name before the edit.
  const cancel = useRef(false);
  if (shown === null) return <div className="demo-file" />;
  const ext = shownExtension(shown, kind);
  const commit = () => {
    const typed = draft;
    setDraft(null);
    if (cancel.current || typed === null) {
      cancel.current = false;
      return;
    }
    const next = renamedTo(typed, name);
    if (next !== name) onRename(next);
  };
  const onKeyDown = (e: KeyboardEvent<HTMLInputElement>) => {
    if (e.nativeEvent.isComposing) return;
    if (e.key !== 'Enter' && e.key !== 'Escape') return;
    e.preventDefault();
    cancel.current = e.key === 'Escape';
    e.currentTarget.blur();
    onDone();
  };
  return (
    <div className="demo-file">
      <input
        ref={field}
        className="demo-file__name"
        aria-label={NAME_FIELD_LABEL}
        title={name + ext}
        spellCheck={false}
        autoComplete="off"
        maxLength={64}
        value={draft ?? name}
        onFocus={() => setDraft(name)}
        onChange={(e) => setDraft(e.target.value)}
        onBlur={commit}
        onKeyDown={onKeyDown}
      />
      {ext ? <span className="demo-file__ext">{ext}</span> : null}
    </div>
  );
}
