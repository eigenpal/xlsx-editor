// The open and save logic of the Vue demo (`App.vue`), in plain TypeScript, so that
// `session.test.ts` runs it under Bun. It follows the React demo's page (examples/vite/src/
// App.tsx) in a smaller form:
// - An open reads the file and checks it (`whyNotOpen`) before it asks about unsaved edits, so a
//   file that does not open leaves the workbook as it is. A file over 64 MiB is refused.
// - A newer pick stops an older open that has not loaded, and a no to the question stops it too.
// - The name and the clean revision of a workbook are kept only when its load ends ready, and
//   only when no newer load replaced it.
// - The editor shows the row and column headings of each sheet as the file sets them
//   (`x:sheetView/@showRowColHeaders`), and Save writes the bytes that the editor saves.
// - A save is refused while an open runs.
// - One save runs at a time over the top bar's Save, the toolbar's Save, and the keys, and a
//   second download of the same revision soon after the first is dropped, as in the React demo.
// - A name that the user gives the workbook in the top bar (`rename`) is the name of its
//   download, as in the React demo. Each open starts with no such name.
// The state is one object that `App.vue` makes reactive. The demo's text is English and stays
// out of `packages/i18n` (AGENTS.md, "i18n").

import type {
  EditorSnapshot,
  SavedWorkbook,
  XlsxEditorInstance,
} from '@xlsx-editor.dev/core/editor';
import { FILE_KINDS, fileKindOf } from '../../vite/src/file-kinds.ts';
import { readForOpen, whyNotOpen } from '../../vite/src/open-check.ts';
import { SAMPLE } from '../../vite/src/sample.ts';
import { saveFileName, saveRepeatGuard } from '../../vite/src/save.ts';
import { renamedTo, shownName, titleOf } from '../../vite/src/workbook-names.ts';

/** The largest file that the demo reads into memory: 64 MiB, as in the React demo. */
export const MAX_FILE_BYTES = 64 * 1024 * 1024;

const NO_WORKBOOK = 'Could not save: no workbook is open.';
const SAVE_REPLACED = 'Nothing was saved: another workbook opened during the save.';

/** The part of the editor handle (`XlsxEditorRef`) that the session uses. */
export interface SessionEditor {
  load(bytes: Uint8Array): Promise<EditorSnapshot | null>;
  saveInTasks(): Promise<SavedWorkbook | null>;
  snapshot(): EditorSnapshot;
  getEditor(): Pick<XlsxEditorInstance, 'commitEdit'> | null;
}

/** The page services of the session. */
export interface SessionHost {
  /** The editor, or null before it mounts. */
  editor(): SessionEditor | null;
  /** Settles when the editor is attached. */
  readonly ready: Promise<void>;
  /** Asks the question before unsaved edits are lost. */
  confirm(text: string): boolean;
  /** Downloads a saved copy. */
  download(blob: Blob, name: string): void;
  /** Gives the keys to the grid after an open, when no control that the user chose has them. */
  focusGrid(): void;
}

/** The state that the page shows. */
export interface SessionState {
  /** The open workbook (the sample's key or a picked file's name), or null before one is ready. */
  label: string | null;
  /** The name that the user gave the open workbook in the top bar, with no extension, or null. */
  chosen: string | null;
  error: string | undefined;
  /** Whether an open runs, from its read or download to the end of its load. */
  opening: boolean;
}

/** The state at start: no workbook. */
export const initialSessionState = (): SessionState => ({
  label: null,
  chosen: null,
  error: undefined,
  opening: false,
});

/** A picked file, as the file input gives it. */
export interface PickedFile {
  readonly name: string;
  readonly size: number;
  arrayBuffer(): Promise<ArrayBuffer>;
}

/** The reason of a load that did not end ready, for a message. */
const loadFailure = (result: EditorSnapshot | null | undefined) =>
  result?.failure?.detail ?? result?.failure?.code ?? 'the editor is not ready';

export class DemoSession {
  // Each pick takes the next number. Only the newest pick goes on after each step.
  private picks = 0;
  // Each load takes the next number. Only the newest load may record its workbook.
  private loads = 0;
  // The opens that run, from the read or download to the end of the load.
  private openRuns = 0;
  // The opens that the user accepted: the workbook on screen is about to go.
  private replacing = 0;
  // The revision of the last open or save.
  private cleanAt: number | null = null;
  private saving = false;
  // Drops a second download of the same revision soon after the first.
  private readonly downloadOnce = saveRepeatGuard();

  constructor(
    private readonly host: SessionHost,
    readonly state: SessionState = initialSessionState()
  ) {}

  private get editor(): SessionEditor | null {
    return this.host.editor();
  }

  private say(error: string): void {
    this.state.error = error;
  }

  private showOpening(): void {
    this.state.opening = this.openRuns > 0;
  }

  /** True when the open workbook has an edit after the last open or save. */
  hasUnsavedEdits(): boolean {
    const s = this.editor?.snapshot();
    if (!s || s.status !== 'ready' || this.cleanAt === null) return false;
    return s.revision !== this.cleanAt || (s.editing && s.editChanged);
  }

  /**
   * Keeps the name that the user typed in the top bar for the open workbook (`renamedTo`): an
   * empty name keeps the name before it. Does nothing with no workbook open.
   */
  rename(typed: string): void {
    const { label, chosen } = this.state;
    if (label === null) return;
    const before = chosen ?? shownName(label).name;
    const next = renamedTo(typed, before);
    if (next !== before) this.state.chosen = next;
  }

  /** Clears the message. */
  clearMessage(): void {
    this.state.error = undefined;
  }

  /**
   * Asks before a new workbook replaces unsaved edits. True when the open may go on. A value
   * typed with no Enter commits first. A no, and a commit that the editor refuses, also stop an
   * earlier open that has not loaded.
   */
  private mayReplace(label: string): boolean {
    const editor = this.editor;
    // The load of an accepted open already replaces the workbook, so no answer can keep it.
    if (this.replacing > 0 && editor?.snapshot().status === 'loading') return true;
    const keepEdits = () => {
      this.picks += 1;
      return false;
    };
    if (editor?.snapshot().editing) {
      const done = editor.getEditor()?.commitEdit();
      if (done && !done.ok) {
        this.say(`Could not open ${titleOf(label)}: finish or cancel the cell edit first.`);
        return keepEdits();
      }
    }
    const shown = titleOf(this.state.chosen ?? this.state.label ?? SAMPLE);
    return (
      !this.hasUnsavedEdits() ||
      this.host.confirm(
        `Open ${titleOf(label)}? The edits to ${shown} are lost unless you saved them to a file.`
      ) ||
      keepEdits()
    );
  }

  /** The end of a message about a file that did not open. */
  private staysOpen(): string {
    if (this.replacing > 0) return '';
    return this.state.label === null
      ? ' Open a file to go on.'
      : ' The workbook before it stays open.';
  }

  /** Opens a picked file. */
  async openFile(file: PickedFile): Promise<void> {
    if (file.size > MAX_FILE_BYTES) {
      this.say(`Could not open ${titleOf(file.name)}: the file is larger than 64 MiB.`);
      return;
    }
    await this.open(() => file.arrayBuffer(), file.name, false);
  }

  /** Downloads the sample workbook with `download` and opens it: the open at start. */
  async openSample(download: () => Promise<ArrayBuffer>): Promise<void> {
    await this.open(download, SAMPLE, true);
  }

  private async open(
    read: () => Promise<ArrayBuffer>,
    label: string,
    downloads: boolean
  ): Promise<void> {
    const pick = ++this.picks;
    this.openRuns += 1;
    this.showOpening();
    let replaces = false;
    try {
      let bytes: Uint8Array;
      try {
        bytes = new Uint8Array(await read());
      } catch (cause) {
        if (pick !== this.picks) return;
        this.say(
          downloads
            ? `Could not open ${titleOf(label)}: the file did not download. Open a file to go on.`
            : `Could not open ${titleOf(label)}: ${String(cause)}`
        );
        return;
      }
      if (pick !== this.picks) return;
      const read1 = await readForOpen(bytes, () => pick === this.picks);
      if (!read1 || pick !== this.picks) return;
      // A load that fails closes the open workbook, so check first what the editor checks.
      const problem = whyNotOpen(read1);
      if (problem || !read1.ok) {
        this.say(`Could not open ${titleOf(label)}: ${problem}.${this.staysOpen()}`);
        return;
      }
      if (!this.mayReplace(label)) return;
      replaces = true;
      this.replacing += 1;
      await this.host.ready;
      if (pick !== this.picks) return;
      const at = ++this.loads;
      const result = await this.editor?.load(bytes);
      // Only a newer load replaces this workbook. A newer pick that has not loaded leaves this
      // workbook on screen, so its name must still be recorded.
      if (at !== this.loads) return;
      if (result?.status !== 'ready') {
        if (pick === this.picks)
          this.say(
            `Could not open ${titleOf(label)}: ${loadFailure(result)}. Open a file to go on.`
          );
        return;
      }
      this.state.label = label;
      this.state.chosen = null;
      this.cleanAt = result.revision;
      // A newer pick can have its own message.
      if (pick !== this.picks) return;
      this.clearMessage();
      this.host.focusGrid();
    } finally {
      this.openRuns -= 1;
      if (replaces) this.replacing -= 1;
      this.showOpening();
    }
  }

  /**
   * The toolbar's save: Control+S, Command+S, or its Save. It runs as one save with the top bar's
   * Save (`saveOnce`), so a press during a save of either control downloads nothing more.
   */
  async save(bytes: Uint8Array<ArrayBuffer>, result: SavedWorkbook): Promise<void> {
    await this.saveOnce(() => this.write(bytes, result));
  }

  /** The Save of the top bar. Control+S and Command+S go to the toolbar, which calls `save`. */
  async saveNow(): Promise<void> {
    await this.saveOnce(async () => {
      const saved = await this.editor?.saveInTasks();
      if (saved) await this.write(saved.bytes, saved);
      // With no open workbook, the save gives nothing and says nothing, so the demo says why.
      else if (this.editor?.snapshot().status !== 'ready') this.say(NO_WORKBOOK);
    });
  }

  /**
   * Runs one save at a time, as the React demo's `saveOnce` does: a press during a save, of either
   * Save or of the keys, does nothing. A save is refused during an open.
   */
  private async saveOnce(work: () => Promise<void>): Promise<void> {
    if (this.openRuns > 0) {
      this.say('Could not save: a workbook is opening. Try again when it shows.');
      return;
    }
    if (this.saving) return;
    this.saving = true;
    try {
      await work();
    } finally {
      this.saving = false;
    }
  }

  /** Writes the saved `bytes` to a download. */
  private async write(bytes: Uint8Array<ArrayBuffer>, result: SavedWorkbook): Promise<void> {
    // Another open ended this workbook during the save: these bytes are not the open workbook.
    const label = this.state.label;
    const chosen = this.state.chosen;
    if (!result.stillOpen || this.replacing > 0) {
      this.say(SAVE_REPLACED);
      return;
    }
    if (label === null) {
      this.say(NO_WORKBOOK);
      return;
    }
    const at = this.loads;
    const snap = this.editor?.snapshot();
    const kind = (snap ? fileKindOf(snap) : null) ?? 'workbook';
    const file = FILE_KINDS[kind];
    const blob = new Blob([bytes], { type: file.type });
    // Each save starts a download at once, so a repeat of the same revision, as a double click
    // gives, is dropped (`saveRepeatGuard`, as the React demo's `downloadOnce`). A reload starts
    // the revisions again, so the key holds the number of the load too.
    await this.downloadOnce(at * 1_000_000_000 + result.revision, async () => {
      this.host.download(blob, saveFileName(label, chosen, file.ext));
      return true;
    });
    if (at === this.loads) this.cleanAt = result.revision;
  }
}
