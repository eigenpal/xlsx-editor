// The save of the demo: the name of the saved copy, and the write to a file that the user picks
// with the File System Access save picker, or a download in a browser with no picker. Each save
// reports one outcome, which sets both the error and the notice of the demo, so a good save
// clears the error of an earlier failed one.

import { displayName } from '@xlsx-editor.dev/core/editor';
import { downloadBlob } from './download.ts';

/** The browser's file save picker, where it exists. */
export type SavePicker = (options: {
  suggestedName: string;
  types: { description: string; accept: Record<string, string[]> }[];
}) => Promise<{
  readonly name?: string;
  createWritable(): Promise<{ write(data: Blob): Promise<void>; close(): Promise<void> }>;
}>;

/** The file save picker of this browser, or `undefined` when it has none. */
export const savePicker = (): SavePicker | undefined =>
  (window as unknown as { showSaveFilePicker?: SavePicker }).showSaveFilePicker;

/** Whether the page has the activation of a recent click, which the save picker needs. */
export const hasActivation = (): boolean =>
  (navigator as { userActivation?: { isActive: boolean } }).userActivation?.isActive ?? false;

/** Whether the editor refused the last save (`refusal.command` is `save`): it wrote nothing. */
export const saveRefused = (s: { readonly refusal: { readonly command?: string } | null }) =>
  s.refusal?.command === 'save';

/** The error of a save that the browser could not write. */
export const SAVE_FAILED = 'Nothing was saved: the browser could not write the file. Try again.';

/** The notice of a save whose bytes took longer to prepare than the browser allows after a click. */
export const SAVE_READY =
  'The workbook is ready to save. Select Save, or press Control or Command and S, to choose the file.';

/** What the demo shows after a save: an error, a notice, or neither. */
export interface SaveOutcome {
  /** Whether the file is known to be written. Only such a save clears the unsaved state. */
  readonly written: boolean;
  readonly error: string | null;
  readonly notice: string | null;
  /** The browser refused the picker because the click was too long ago. */
  readonly late?: true;
}

/**
 * The name of the saved copy: the open name without its extension and without an `-edited`
 * end that an earlier save added, then `-edited` and `ext`.
 */
export const savedName = (name: string, ext: string): string =>
  `${name.replace(/\.(xlsx|xltx)$/i, '').replace(/-edited$/i, '') || 'workbook'}-edited${ext}`;

/**
 * The notice of a started download. The page cannot confirm that the file arrives, so it names
 * the question before the page closes only when `unsaved` edits make the page ask.
 */
export const downloadNotice = (name: string, unsaved: boolean): string => {
  const started = `Started the download of ${displayName(name)}.`;
  return unsaved ? `${started} This page still asks before you leave.` : started;
};

/** How long after a started download a new save of the same revision is a repeat. */
export const SAVE_REPEAT_MS = 1000;

/**
 * A guard for a save with no picker, where each save starts a download at once. It drops a save
 * while another save runs, and a save of the same revision within `SAVE_REPEAT_MS` of a started
 * download, as a double click gives. `save` returns whether a download started. Only a started
 * download starts the time: a failed save never blocks a retry, and a dropped save never extends
 * the time. A held key does not come here, because the page drops its repeated key events.
 */
export function saveRepeatGuard(now: () => number = () => performance.now()) {
  let started: { readonly at: number; readonly revision: number } | null = null;
  let running = false;
  return async (revision: number, save: () => Promise<boolean>): Promise<void> => {
    if (running) return;
    if (started?.revision === revision && now() - started.at < SAVE_REPEAT_MS) return;
    running = true;
    try {
      if (await save()) started = { at: now(), revision };
    } finally {
      running = false;
    }
  };
}

/**
 * Write `bytes` to a file that the user picks, or download them when the browser has no picker.
 * A cancel of the picker shows nothing. A browser does not report the end of a download, and the
 * user can cancel it, so a started download leaves `unsaved` edits unsaved and the page still
 * asks before it closes. Difference from pptx-editor `examples/vite/src/App.tsx` at e22f3f3b
 * (`exportDeck`), which marks the deck saved after it starts the download.
 */
export async function saveFile(
  picker: SavePicker | undefined,
  bytes: Uint8Array,
  options: Parameters<SavePicker>[0],
  unsaved: boolean,
  download: (blob: Blob, name: string) => void = downloadBlob,
  activation: () => boolean = hasActivation
): Promise<SaveOutcome> {
  // Read before the call, which uses up the activation of the click.
  const late = picker !== undefined && !activation();
  let picking = false;
  try {
    if (!picker) {
      const type = Object.keys(options.types[0]?.accept ?? {})[0] ?? '';
      download(new Blob([new Uint8Array(bytes)], { type }), options.suggestedName);
      return {
        written: false,
        error: null,
        notice: downloadNotice(options.suggestedName, unsaved),
      };
    }
    picking = true;
    const handle = await picker(options);
    picking = false;
    const out = await handle.createWritable();
    await out.write(new Blob([new Uint8Array(bytes)]));
    await out.close();
    return {
      written: true,
      error: null,
      notice: `Saved ${bytes.byteLength} bytes to ${handle.name ? displayName(handle.name) : 'the picked file'}.`,
    };
  } catch (cause) {
    // The user sees a short text, not the browser's error text, which can be long and names no
    // next step.
    const name = (cause as { name?: string }).name;
    if (name === 'AbortError') return { written: false, error: null, notice: null };
    // The browser opens the picker only soon after a click, and no file was touched. A refusal
    // with a recent click, such as in a sandboxed frame, is a failure.
    if (name === 'SecurityError' && picking && late)
      return { written: false, error: null, notice: SAVE_READY, late: true };
    // The browser's error text goes to the console, for the developer who reads it.
    console.error('The save of the demo failed:', cause);
    return { written: false, error: SAVE_FAILED, notice: null };
  }
}

/**
 * Whether the notice of a save still describes the workbook. `savedAt` is the
 * content revision of the save, or null when no notice of a save shows; `now` is the
 * revision now. An edit after the save moves the revision, and the notice must go.
 */
export const savedNoticeStays = (savedAt: number | null, now: number): boolean =>
  savedAt !== null && savedAt === now;

/**
 * Whether kept bytes still hold the open workbook. `kept` gives the load and the revision of the
 * bytes; `now` gives the load and the revision of the editor, and whether an open replaces the
 * workbook. An edit, a headings switch, or another open makes the bytes stale.
 */
export const keptSaveHolds = (
  kept: { readonly loadAt: number; readonly revision: number },
  now: { readonly load: number; readonly revision: number; readonly replacing: boolean }
): boolean => kept.loadAt === now.load && !now.replacing && kept.revision === now.revision;
