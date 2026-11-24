// The scripted assistant. A real host calls its model here; the shape of the answer stays the
// same: text to stream, an optional workbook to attach, and an optional action on the open
// editor. The editor refuses an action that it cannot do, such as an edit in view mode, and
// the chat then shows the editor's reason in place of this text. The actions use only the public editor API (`setActiveSheet`, `selectReference`,
// `enterValue`, `toggleFormat`, `addSheet`, `undo`, `redo`, `zoomStep`, the `mode` prop, and
// `saveInTasks`). Adapted from pptx-editor/examples/chat/src/assistant.ts at e0ac8351
// (EigenPal), with the commands of a workbook in place of the commands of a deck.

export type AssistantAction =
  | { readonly kind: 'open' }
  | { readonly kind: 'sheet'; readonly index: number }
  | { readonly kind: 'select'; readonly ref: string }
  | { readonly kind: 'write'; readonly ref: string; readonly text: string }
  | { readonly kind: 'format'; readonly toggle: 'bold' | 'italic' | 'underline' }
  | { readonly kind: 'addSheet' }
  | { readonly kind: 'undo' }
  | { readonly kind: 'redo' }
  | { readonly kind: 'zoom'; readonly direction: 1 | -1 }
  | { readonly kind: 'mode'; readonly mode: 'edit' | 'view' }
  | { readonly kind: 'download' };

export interface AssistantReply {
  readonly text: string;
  /** Attach a new workbook card under the text. */
  readonly workbook?: boolean;
  readonly action?: AssistantAction;
}

/** What the assistant knows about the editor when it answers. */
export interface WorkbookContext {
  readonly hasWorkbook: boolean;
  /** The names of the sheets that have a tab, in tab order, with their positions. */
  readonly sheets: readonly { readonly index: number; readonly name: string }[];
  readonly activeSheet: string | null;
  readonly selection: string | null;
}

export const FIRST_WORKBOOK_TEXT =
  'Here is a first draft. It opens in the editor next to this chat, so you can change any cell in place. Ask me to go to a sheet, select a range, fill a cell, or download the file.';

/** A cell or a range in A1 notation, such as `B5` or `B5:E8`. */
const RANGE = /\b(\$?[a-z]{1,3}\$?\d{1,7}(?::\$?[a-z]{1,3}\$?\d{1,7})?)\b/i;

/** `set B5 to 1200`, `put 1200 in B5`, `write Done into C3`. */
const WRITE_TO = /\b(?:set|change)\s+(\$?[a-z]{1,3}\$?\d{1,7})\s+to\s+(.+)$/i;
const WRITE_IN =
  /\b(?:put|write|enter|type)\s+(.+?)\s+(?:in|into)\s+(\$?[a-z]{1,3}\$?\d{1,7})\s*$/i;

/** The sheet that `text` names, by a whole word, or null. */
function namedSheet(text: string, context: WorkbookContext) {
  const lower = text.toLowerCase();
  // The longest name first, so "Sales 2026" wins over "Sales".
  const byLength = [...context.sheets].sort((a, b) => b.name.length - a.name.length);
  return (
    byLength.find((sheet) => {
      const name = sheet.name.toLowerCase();
      const at = lower.indexOf(name);
      if (at === -1) return false;
      const before = lower[at - 1];
      const after = lower[at + name.length];
      return !/\w/.test(before ?? '') && !/\w/.test(after ?? '');
    }) ?? null
  );
}

const sheetList = (context: WorkbookContext): string =>
  context.sheets.map((sheet) => sheet.name).join(', ');

export function reply(input: string, context: WorkbookContext): AssistantReply {
  const text = input.trim();
  const lower = text.toLowerCase();
  if (!context.hasWorkbook) return { text: FIRST_WORKBOOK_TEXT, workbook: true };

  const write = WRITE_TO.exec(text) ?? null;
  const writeIn = write ? null : WRITE_IN.exec(text);
  if (write || writeIn) {
    const ref = (write ? write[1] : writeIn![2])!.toUpperCase();
    const value = (write ? write[2] : writeIn![1])!.trim().replace(/^["“](.*)["”]$/, '$1');
    return {
      text: `I wrote ${value} in ${ref}${context.activeSheet ? ` on ${context.activeSheet}` : ''}. Press Control+Z in the grid, or say “undo”, to take it back.`,
      action: { kind: 'write', ref, text: value },
    };
  }

  if (/\b(view only|read only|lock)\b/.test(lower))
    return {
      text: 'The workbook is now view only. You can still select, copy, switch sheets, and download.',
      action: { kind: 'mode', mode: 'view' },
    };
  if (/\b(allow edits|unlock|edit mode|let me edit)\b/.test(lower))
    return { text: 'You can edit the cells again.', action: { kind: 'mode', mode: 'edit' } };

  if (/\b(undo)\b/.test(lower))
    return { text: 'I took back the last edit.', action: { kind: 'undo' } };
  if (/\b(redo)\b/.test(lower)) return { text: 'I did the edit again.', action: { kind: 'redo' } };

  const toggle = /\bbold\b/.test(lower)
    ? 'bold'
    : /\bitalic\b/.test(lower)
      ? 'italic'
      : /\bunderline\b/.test(lower)
        ? 'underline'
        : null;
  if (toggle) {
    return {
      text: `I switched ${toggle} on ${context.selection ?? 'the active cell'}, from the format of its first cell.`,
      action: { kind: 'format', toggle },
    };
  }

  if (/\b(add|new|insert)\b.*\bsheet\b/.test(lower)) {
    return {
      text: 'I added a sheet at the end. Its tab is under the grid.',
      action: { kind: 'addSheet' },
    };
  }

  if (/\bzoom in\b|\bbigger\b|\blarger\b/.test(lower))
    return { text: 'Zoomed in one step.', action: { kind: 'zoom', direction: 1 } };
  if (/\bzoom out\b|\bsmaller\b/.test(lower))
    return { text: 'Zoomed out one step.', action: { kind: 'zoom', direction: -1 } };

  if (/\b(download|export|save|xlsx)\b/.test(lower))
    return {
      text: 'Here is the workbook as an .xlsx file, with every edit that you made.',
      action: { kind: 'download' },
    };

  const sheet = namedSheet(text, context);
  if (sheet)
    return {
      text: `${sheet.name} is open in the editor. Select a cell and type to change it.`,
      action: { kind: 'sheet', index: sheet.index },
    };
  if (/\bsheet\b/.test(lower))
    return { text: `The workbook has these sheets: ${sheetList(context)}. Which one do you want?` };

  const range = RANGE.exec(text);
  if (range) {
    const ref = range[1]!.toUpperCase();
    return { text: `I selected ${ref}.`, action: { kind: 'select', ref } };
  }

  if (/\b(open|edit|show)\b/.test(lower))
    return { text: 'The workbook is open next to the chat.', action: { kind: 'open' } };

  return {
    text: `You are on ${context.activeSheet ?? 'the workbook'}${context.selection ? `, at ${context.selection}` : ''}. This mock does not write tables for you, but the editor lets you: select a cell and type. You can also say “go to Sales”, “select B5:E8”, “set B5 to 1200”, “bold”, or “download”.`,
  };
}
