/** @jsxImportSource react */
// The demo opens the sample workbook at start, as the PPTX Editor demo opens its sample deck. The
// user can then open a file that they pick, and edit it with the toolbar. Save writes the edited
// workbook to a file that the user picks, with the extension of the package kind (`.xlsx` or
// `.xltx`), or downloads it where the browser has no file save picker. The headings option, on by
// default, is a view option of the demo: Save writes the heading flags of the opened file, not the
// option. A pill switches the editor to view mode. The page shell, its top bar, and its light and
// dark theme follow the PPTX Editor and DOCX Editor demos (`demo.css`): the brand lockup, the React
// | Vue switcher, the name of the open workbook with its muted extension, the menu bar, then the
// actions. The name is a text field, as the deck name of the PPTX Editor demo: a click edits it,
// and File > Rename workbook focuses it. Save suggests the name that the user gave. The demo's
// own text is English and stays out of `packages/i18n` (AGENTS.md, "i18n").
//
// The demo reads each file before the editor gets it, so a file that does not read leaves the
// open workbook, its name, and its flags as they were. It opens a file with
// `XlsxEditorRef.load` and keeps its name and heading flags only when the editor reports it
// ready. Save and the headings switch refuse while the editor is not ready. The content
// revision of the editor snapshot tells an edit from a change of selection: a message clears
// on the next edit. Opening another workbook or leaving the page over unsaved edits asks first
// (`unsaved-edits.ts`). Save writes through the browser's file save picker, or a download
// (`save.ts`).

import { useCallback, useEffect, useRef, useState, type ChangeEvent } from 'react';
import {
  XlsxEditor,
  useEditorState,
  type EditorSnapshot,
  type SavedWorkbook,
  type WorkbookKind,
  type XlsxEditorRef,
} from '@xlsx-editor.dev/react';
import type { readXlsx } from '@xlsx-editor.dev/core/store';
import { AdapterSwitcher, SHOW_ADAPTER_SWITCHER } from './AdapterSwitcher.tsx';
import { BrandLogo } from './BrandLogo.tsx';
import { chromeLocale } from './chrome-locale.ts';
import { useBusyStatus } from './busy-status.ts';
import { FONTS_FAILED, FONTS_NEXT_STEP, fontsMissingNotice } from './font-messages.ts';
import { Icon } from './icons.tsx';
import { initialTheme, rememberTheme, ThemeToggle, type Theme } from './ThemeToggle.tsx';
import { readForOpen, whyNotOpen } from './open-check.ts';
import { runInTasks } from './steps-in-tasks.ts';
import { SAMPLE } from './sample.ts';
import { SAMPLE_URL } from './workbooks.ts';
import { RENAME_ROW_LABEL, titleOf } from './workbook-names.ts';
import { WorkbookName } from './WorkbookName.tsx';
import {
  keptSaveHolds,
  saveFile,
  saveFileName,
  savedNoticeStays as messageStays,
  savePicker,
  saveRepeatGuard,
  saveRefused,
  type SaveOutcome,
} from './save.ts';
import { useUnsavedEdits } from './unsaved-edits.ts';
import { useSheetZooms } from './sheet-zooms.ts';
import { usePageKeys } from './page-keys.ts';
import {
  HEADINGS_FROM_FILE,
  HEADINGS_SHOWN,
  headingFlags,
  packageWithHeadingsSteps,
  withHeadingsSteps,
  type HeadingFlags,
} from './headings.ts';

/** The largest file that the demo reads into memory: 64 MiB. */
const MAX_FILE_BYTES = 64 * 1024 * 1024;

const NO_WORKBOOK = 'Could not save: no workbook is open.';

/** What to do when the sample does not download. */
const OPEN_A_FILE = 'Open a file to go on.';

/** The error of a press of Save while a workbook opens. The open does no save work for it. */
const SAVE_OPENING = 'Could not save: a workbook is opening. Try again when it shows.';

/** The extension and the media type of each package kind. */
const KIND_FILE: Readonly<Record<WorkbookKind, { ext: string; type: string; label: string }>> = {
  workbook: {
    ext: '.xlsx',
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    label: 'Workbook',
  },
  template: {
    ext: '.xltx',
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.template',
    label: 'Workbook template',
  },
};

/** The reason of a load that did not end ready, for a message. */
const loadFailure = (result: EditorSnapshot | null | undefined) =>
  result?.failure?.code === 'fonts'
    ? FONTS_FAILED
    : (result?.failure?.detail ?? result?.failure?.code ?? 'the editor is not ready');

/** What to do after a load that did not end ready. */
const loadNextStep = (result: EditorSnapshot | null | undefined) =>
  result?.failure?.code === 'fonts' ? FONTS_NEXT_STEP : 'Open a workbook to go on.';

/** The kind of the message for a load that did not end ready. The editor itself alerts for fonts. */
const loadMessageKind = (result: EditorSnapshot | null | undefined) =>
  result?.failure?.code === 'fonts' ? ('notice' as const) : ('error' as const);

/** Saved bytes that are ready for the file dialog. */
interface ReadySave {
  readonly bytes: Uint8Array<ArrayBuffer>;
  /** The revision that the bytes hold. */
  readonly revision: number;
  /** The number of the load that the bytes belong to. */
  readonly loadAt: number;
  readonly label: string;
  /** The name that the user gave the workbook in the top bar, or null. */
  readonly chosen: string | null;
  readonly kind: WorkbookKind | null | undefined;
}

/** The error of a save whose workbook another workbook replaced. */
const SAVE_REPLACED = 'Nothing was saved: another workbook opened during the save.';

/** The notice after a save that started before the newest edits. */
const SAVE_AGAIN = 'The save holds the workbook before your later edits. Select Save to save them.';

/** The error of a headings switch whose saved bytes the editor cannot read again. */
const HEADINGS_UNREADABLE =
  'Could not switch the headings: the editor cannot read the saved workbook again. The workbook and your edits stay open.';

/** The error of a save whose bytes the editor cannot read to put back the heading flags. */
const SAVE_UNREADABLE =
  'Nothing was saved: the editor cannot read the saved workbook to put back its own heading flags.';

/** The error of a key that the hidden grid of a load cannot take. */
const KEY_LOADING = 'The key did nothing: a workbook is opening. Try again when it shows.';

/** The error of a key during the reload of a headings switch. */
const KEY_SWITCHING =
  'The key did nothing: the headings are switching. Try again when the switch ends.';

/** The bytes of the committed workbook at `url`. */
async function download(url: string): Promise<Uint8Array> {
  const response = await fetch(url);
  if (!response.ok) throw new Error(`HTTP ${response.status}`);
  return new Uint8Array(await response.arrayBuffer());
}

export function App() {
  // The workbook that the top bar names: none until a workbook is ready, so a sample that does
  // not download is not named.
  const [name, setName] = useState<string | null>(null);
  // The name that the user gave the open workbook in the top bar, with no extension, or null.
  // Each open starts with none.
  const [chosen, setChosen] = useState<string | null>(null);
  // One message at a time: a new message of either kind replaces the one before it.
  const [message, setMessage] = useState<{ text: string; kind: 'notice' | 'error' } | null>(null);
  // Show the row and column headings of every sheet, also in a file that hides them, so whole
  // rows and columns can be selected with the pointer. On by default. Switching reopens the
  // edited workbook, so the edits stay and the undo history starts again.
  const [headings, setHeadings] = useState(true);
  const [viewOnly, setViewOnly] = useState(false);
  // The shell's light or dark theme, as in the family demos: `?theme=`, else the theme picked
  // before, else light. A pick is kept for the next visit.
  const [theme, setThemeState] = useState<Theme>(() => initialTheme(window.location.search));
  // The chrome catalog from `?lang=`, such as `?lang=ar` for a right-to-left chrome. Read once.
  const [locale] = useState(() => chromeLocale(window.location.search));
  const setTheme = (next: Theme) => {
    setThemeState(next);
    rememberTheme(next);
  };
  // The option as the newest switch set it. A load that ends later reads it here, not from the
  // render that started the load.
  const headingsNow = useRef(true);
  // The name and heading flags of the workbook that the editor shows.
  // None until a workbook is ready.
  const nameNow = useRef<string | null>(null);
  const chosenNow = useRef<string | null>(null);
  // The workbook name field in the top bar, which File > Rename workbook focuses.
  const nameField = useRef<HTMLInputElement>(null);
  const flags = useRef<HeadingFlags>(new Map());
  // The opens from the end of their download or file read to the end of their load, and the
  // headings switches during their save and rewrite.
  const opening = useRef(0);
  // The opens that run, from the download or the file read to the end of the load.
  const openRuns = useRef(0);
  /**
   * Whether an open runs, from its download or file read to the end of its load, or the save and
   * rewrite steps of a headings switch run: the toolbar Save then starts no save.
   */
  const isOpening = useCallback(() => opening.current > 0 || openRuns.current > 0, []);
  // The answer of `isOpening`, for the `opening` prop of the toolbar: its Save shows that it
  // waits, and Control+S in the editor starts no save.
  const [openingShown, setOpeningShown] = useState(false);
  const showOpening = useCallback(() => setOpeningShown(isOpening()), [isOpening]);
  // The opens that the user accepted: the workbook on screen is about to go.
  const replacing = useRef(0);
  const editor = useRef<XlsxEditorRef>(null);
  // Each load takes the next number. Only the newest load may show its result.
  const request = useRef(0);
  // The number of the load that the editor shows. A request takes its number before its load
  // starts, so only this number tells which workbook a save holds.
  const loaded = useRef(0);
  // The newest workbook that the user asked for. An older fetch or file read stops at its end.
  const picks = useRef(0);
  // The revision at which the current message was shown. The next edit clears it.
  const shownAt = useRef<number | null>(null);
  // The message on screen. The same message again renders nothing, so repeated presses during a
  // load take no time from the load.
  const shownMessage = useRef<{ text: string; kind: 'notice' | 'error' } | null>(null);

  // True while a save runs, so a second press does not start a second save.
  const saving = useRef(false);
  // The revision that the running save holds, when it is known.
  const savingRevision = useRef<number | null>(null);
  // True when the user pressed Save for edits that the running save does not hold.
  const laterEdits = useRef(false);
  // True when the running save wrote its file.
  const wroteFile = useRef(false);
  // The toolbar saves that have not delivered their bytes.
  const toolbarSaves = useRef(0);
  // True while a headings switch runs: its reload would refuse a save.
  const switching = useRef(false);
  // The bytes that wait for a press of Save, after the browser refused its file dialog.
  const readySave = useRef<ReadySave | null>(null);
  const [busyText, busy] = useBusyStatus();

  const notice = message?.kind === 'notice' ? message : null;

  const { revision, hasUnsavedEdits, holdEdits, markClean } = useUnsavedEdits(editor);
  const [downloadOnce] = useState(() => saveRepeatGuard());

  /** Replaces the workbook of the editor with `bytes`. */
  const loadInEditor = useCallback((bytes: Uint8Array) => {
    loaded.current += 1;
    return editor.current?.load(bytes);
  }, []);

  /** Shows a message until the next edit or its close button. It replaces the one before. */
  const say = useCallback(
    (text: string, kind: 'notice' | 'error') => {
      shownAt.current = revision();
      const shown = shownMessage.current;
      if (shown?.text === text && shown.kind === kind) return;
      shownMessage.current = { text, kind };
      setMessage(shownMessage.current);
    },
    [revision]
  );

  /** Clears the message. */
  const clearMessages = useCallback(() => {
    shownAt.current = null;
    if (shownMessage.current === null) return;
    shownMessage.current = null;
    setMessage(null);
  }, []);

  /** Closes the message from its button and gives the focus back to the grid. */
  const closeMessage = useCallback(() => {
    clearMessages();
    editor.current?.focus();
  }, [clearMessages]);

  const zooms = useSheetZooms(editor);

  /**
   * The editor's change callback. When a sheet first shows after a load, it gets its zoom from
   * before the load, or on a phone the phone zoom. The first edit after a message clears it. A
   * save that the editor refused wrote nothing, and the editor shows why: an old message would
   * mislead, so it clears too.
   */
  const onChange = useCallback(
    (snapshot: EditorSnapshot) => {
      zooms.onSnapshot(snapshot);
      // Kept bytes that no longer hold the workbook go, so stale bytes do not stay in memory.
      const kept = readySave.current;
      const now = {
        load: loaded.current,
        revision: snapshot.revision,
        replacing: replacing.current > 0,
      };
      if (kept && !keptSaveHolds(kept, now)) readySave.current = null;
      if (saveRefused(snapshot)) {
        clearMessages();
        return;
      }
      // A message stays until the revision moves past the one at which it showed.
      if (messageStays(shownAt.current, snapshot.revision)) return;
      if (shownAt.current === null) return;
      clearMessages();
    },
    [zooms, clearMessages]
  );

  // Leaving the page or reloading it over unsaved edits asks first.
  useEffect(() => {
    const warn = (e: BeforeUnloadEvent) => {
      if (hasUnsavedEdits()) e.preventDefault();
    };
    window.addEventListener('beforeunload', warn);
    return () => window.removeEventListener('beforeunload', warn);
  }, [hasUnsavedEdits]);

  /**
   * Asks before a new workbook replaces unsaved edits. True when the open may go on. A value
   * typed with no Enter is still in the cell editor, with no revision, so it commits first and
   * then counts as an edit. A commit that the editor refuses stops the open.
   */
  const mayReplace = useCallback(
    (label: string) => {
      const instance = editor.current?.getEditor();
      // The load of an accepted open already replaces the workbook, so no answer can keep it.
      if (replacing.current > 0 && instance?.snapshot().status === 'loading') return true;
      // A no also stops an earlier open that has not started its load.
      const keepEdits = () => {
        picks.current += 1;
        return false;
      };
      if (instance?.snapshot().editing) {
        const done = instance.commitEdit();
        if (!done.ok) {
          say(`Could not open ${titleOf(label)}: finish or cancel the cell edit first.`, 'error');
          return keepEdits();
        }
      }
      return (
        !hasUnsavedEdits() ||
        window.confirm(
          `Open ${titleOf(label)}? The edits to ${titleOf(chosenNow.current ?? nameNow.current ?? SAMPLE)} are lost unless you saved them to a file.`
        ) ||
        keepEdits()
      );
    },
    [hasUnsavedEdits, say]
  );

  /**
   * The end of a message about a file that did not open: the open workbook stays, or with no
   * workbook ready, what to do.
   */
  const staysOpen = useCallback(() => {
    if (replacing.current > 0) return '';
    return nameNow.current === null ? ` ${OPEN_A_FILE}` : ' The workbook before it stays open.';
  }, []);

  /**
   * Opens `data` as the workbook `label`, with headings if the option is on. `pick` is the
   * number of the user's request for it. `askedAt` is the revision at which the user already
   * answered about unsaved edits. With none, it asks only after the file passes the checks. The
   * load takes the next request number only then, so a file that does not open drops no load
   * that runs.
   */
  const openNow = useCallback(
    async (data: Uint8Array, label: string, pick: number, askedAt?: number) => {
      // The steps before the load yield to the page, so the option must wait for this open.
      opening.current += 1;
      showOpening();
      let replaces = false;
      let result: EditorSnapshot | null | undefined;
      let opened: HeadingFlags;
      let id: number;
      let at: number;
      try {
        // Only the newest workbook that the user asked for goes on after its read.
        let read: ReturnType<typeof readXlsx> | null = await readForOpen(
          data,
          () => pick === picks.current
        );
        if (!read || pick !== picks.current) return;
        // A load that fails closes the open workbook, so check first what the editor checks.
        const problem = whyNotOpen(read);
        if (problem || !read.ok) {
          say(`Could not open ${titleOf(label)}: ${problem}.${staysOpen()}`, 'error');
          return;
        }
        if (askedAt === undefined && !mayReplace(label)) return;
        const asked = askedAt ?? revision();
        replaces = true;
        replacing.current += 1;
        id = ++request.current;
        opened = headingFlags(read);
        const shown = headingsNow.current
          ? await runInTasks(packageWithHeadingsSteps(read.pkg, true, opened))
          : data;
        // The editor reads its own copy, so this one must not double the memory of the load.
        read = null;
        if (id !== request.current || pick !== picks.current) return;
        // The grid stays editable during the steps above, and the answer covers no later edit.
        const now = editor.current?.snapshot();
        const touched = revision() !== asked || (now?.editing && now.editChanged);
        if (touched && !mayReplace(label)) return;
        zooms.forget();
        const loading = loadInEditor(shown);
        at = loaded.current;
        result = await loading;
      } finally {
        opening.current -= 1;
        showOpening();
        if (replaces) replacing.current -= 1;
      }
      // Only a newer load replaces this workbook. A newer request that has not loaded leaves
      // this workbook on screen, so its name and flags must still be recorded.
      if (at !== loaded.current) return;
      const newest = id === request.current;
      // The file passed the checks above, which copy the editor's checks, and the editor loads
      // its fonts once. So this path runs only when the fonts fail at the first load, or for a
      // defect.
      if (result?.status !== 'ready') {
        const why = loadFailure(result);
        const text = `Could not open ${titleOf(label)}: ${why}. ${loadNextStep(result)}`;
        say(text, loadMessageKind(result));
        return;
      }
      flags.current = opened;
      nameNow.current = label;
      setName(label);
      chosenNow.current = null;
      setChosen(null);
      markClean(result.revision);
      // A newer pick can have its own message, such as a download that failed.
      if (!newest || pick !== picks.current) return;
      clearMessages();
      // After its dialog, a file input keeps the focus or leaves none. So the grid takes the
      // keys of the user.
      const active = document.activeElement;
      const fileInput = active instanceof HTMLInputElement && active.type === 'file';
      if (active === document.body || fileInput) editor.current?.focus();
      // A workbook with some fonts missing still opens, so the person learns why text looks
      // different.
      const missing = editor.current?.getEditor()?.fontReport().substituteFailures.length ?? 0;
      if (missing > 0) say(fontsMissingNotice(missing), 'notice');
    },
    [
      say,
      clearMessages,
      mayReplace,
      markClean,
      zooms,
      revision,
      loadInEditor,
      staysOpen,
      showOpening,
    ]
  );

  /** {@link openNow}, with a status text while it runs. */
  const open = useCallback(
    (data: Uint8Array, label: string, pick: number, askedAt?: number) =>
      busy.during('Opening the workbook.', openNow(data, label, pick, askedAt)),
    [busy, openNow]
  );

  /** Downloads the sample workbook and opens it: the open at start. */
  const openSample = useCallback(async () => {
    if (!mayReplace(SAMPLE)) return;
    const asked = revision();
    // A file that the user picks during the download is newer, and this open gives way to it.
    const pick = ++picks.current;
    // This number drops a fetch that a newer open replaced. `open` takes its own number for the
    // load.
    const id = ++request.current;
    const fetchAndOpen = async () => {
      let data: Uint8Array;
      try {
        data = await download(SAMPLE_URL);
      } catch {
        if (id === request.current && pick === picks.current)
          say(
            `Could not open ${titleOf(SAMPLE)}: the file did not download. ${OPEN_A_FILE}`,
            'error'
          );
        return;
      }
      try {
        if (id === request.current && pick === picks.current) await open(data, SAMPLE, pick, asked);
      } catch (cause) {
        if (id === request.current && pick === picks.current)
          say(`Could not open ${titleOf(SAMPLE)}: ${String(cause)}`, 'error');
      }
    };
    // One status text covers the download and the open.
    openRuns.current += 1;
    showOpening();
    try {
      await busy.during('Opening the workbook.', fetchAndOpen());
    } finally {
      openRuns.current -= 1;
      showOpening();
    }
  }, [open, mayReplace, say, revision, busy, showOpening]);

  // Open the sample once, at start, also when the effect runs twice in development.
  const started = useRef(false);
  useEffect(() => {
    if (started.current) return;
    started.current = true;
    // The grid gets the focus, so a visitor can type at once.
    void openSample().then(() => editor.current?.focus());
  }, [openSample]);

  /** Switches the headings option and reopens the edited workbook on the same sheet and cells. */
  const switchHeadings = async (show: boolean) => {
    const instance = editor.current?.getEditor();
    if (!instance) return;
    // The save and the rewrite yield to the page. A workbook that opens during them wins: the
    // check after them ends this switch.
    const loadAt = request.current;
    opening.current += 1;
    showOpening();
    let shown: Uint8Array;
    let revisionAt: number;
    try {
      // The save commits an open cell edit first. When it cannot, it shows the reason; keep the
      // option.
      const saved = await editor.current?.saveInTasks();
      if (!saved) return;
      revisionAt = saved.revision;
      const rewritten = await runInTasks(withHeadingsSteps(saved.bytes, show, flags.current));
      // A reload of bytes that do not read would close the workbook and lose the edits.
      if (!rewritten) {
        say(HEADINGS_UNREADABLE, 'error');
        return;
      }
      shown = rewritten;
    } finally {
      opening.current -= 1;
      showOpening();
    }
    if (loadAt !== request.current) {
      say('Could not switch the headings: another workbook opened.', 'error');
      return;
    }
    // The reopen shows the saved bytes, so an edit made during the save must not be lost. That
    // includes text that is still in the cell editor.
    const typing = instance.snapshot();
    if (revision() !== revisionAt || (typing.editing && typing.editChanged)) {
      say('Could not switch the headings: the workbook changed. Try again.', 'error');
      return;
    }
    const { activeSheet: sheet, selection } = instance.snapshot();
    const edited = hasUnsavedEdits();
    // The switch is the newest load: a workbook that is still loading does not show.
    const id = ++request.current;
    headingsNow.current = show;
    zooms.reopen();
    // During the reopen, only the saved bytes hold the edits, so they stay protected.
    holdEdits(edited);
    let result: EditorSnapshot | null | undefined;
    let at: number;
    try {
      const loading = loadInEditor(shown);
      at = loaded.current;
      result = await loading;
    } finally {
      holdEdits(false);
    }
    // Only a newer load replaces this workbook. A newer request that has not loaded leaves it.
    if (at !== loaded.current) return;
    if (result?.status !== 'ready') {
      headingsNow.current = !show;
      const why = loadFailure(result);
      say(
        `Could not switch the headings: ${why}. ${loadNextStep(result)}`,
        loadMessageKind(result)
      );
      return;
    }
    if (sheet >= 0) instance.setActiveSheet(sheet);
    if (selection) instance.select(selection.anchor, selection.active);
    markClean(instance.snapshot().revision, edited);
    // A newer request that failed has its own message, which stays.
    if (id !== request.current) return;
    // The hidden grid drops the focus to the body. A control that the user chose keeps it.
    if (document.activeElement === document.body) editor.current?.focus();
    const text = show ? HEADINGS_SHOWN : HEADINGS_FROM_FILE;
    say(edited ? `${text} Your edits are kept. Undo starts again from here.` : text, 'notice');
  };

  /** Why the headings option cannot switch now, or null. */
  const headingsRefusal = () => {
    if (saving.current || toolbarSaves.current > 0)
      return 'a save is running. Try again when it ends';
    if (switching.current) return 'a switch is running. Try again when it ends';
    // A switch takes a new request, so the sample that is still in its download would not open.
    if (opening.current > 0 || openRuns.current > 0)
      return 'a workbook is opening. Try again when it shows';
    if (editor.current?.snapshot().status !== 'ready') return 'no workbook is open';
    return null;
  };

  const onHeadings = async (show: boolean) => {
    const refusal = headingsRefusal();
    if (refusal) {
      say(`Could not switch the headings: ${refusal}.`, 'error');
      return;
    }
    switching.current = true;
    // The box answers the press at once. The end of the switch sets the option that is in force.
    setHeadings(show);
    try {
      await busy.during('Switching the headings.', switchHeadings(show));
    } finally {
      switching.current = false;
      setHeadings(headingsNow.current);
    }
  };

  const onFile = async (e: ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    e.target.value = '';
    if (!file) return;
    if (file.size > MAX_FILE_BYTES) {
      say(`Could not open ${titleOf(file.name)}: the file is larger than 64 MiB.`, 'error');
      return;
    }
    // This pick is newer than the sample that is still in its download.
    const pick = ++picks.current;
    openRuns.current += 1;
    showOpening();
    try {
      const data = new Uint8Array(await file.arrayBuffer());
      if (pick === picks.current) await open(data, file.name, pick);
    } catch (cause) {
      if (pick === picks.current)
        say(`Could not open ${titleOf(file.name)}: ${String(cause)}`, 'error');
    } finally {
      openRuns.current -= 1;
      showOpening();
    }
  };

  /** The options of the file dialog for the workbook `label`, or the name that the user gave. */
  const saveOptions = (
    label: string,
    chosenName: string | null,
    kind: WorkbookKind | null | undefined
  ) => {
    const file = KIND_FILE[kind ?? 'workbook'];
    return {
      suggestedName: saveFileName(label, chosenName, file.ext),
      types: [{ description: file.label, accept: { [file.type]: [file.ext] } }],
    };
  };

  /** Shows how a save ended. */
  const showSaveOutcome = (outcome: SaveOutcome) => {
    if (outcome.error) say(outcome.error, 'error');
    else if (outcome.notice) say(outcome.notice, 'notice');
  };

  /**
   * Asks for the file and writes `save`, or downloads it where the browser has no picker. The
   * picker opens in the task of the call. With no picker, each save starts a download at once, so
   * the repeat guard drops a repeat of the same revision, as a double click gives.
   */
  const writeSave = async (save: ReadySave) => {
    readySave.current = null;
    savingRevision.current = save.revision;
    const picker = savePicker();
    const write = async () => {
      // Inside the guard: a dropped repeat keeps the message of the save before it.
      clearMessages();
      const options = saveOptions(save.label, save.chosen, save.kind);
      const outcome = await saveFile(picker, save.bytes, options, hasUnsavedEdits());
      if (outcome.late) readySave.current = save;
      wroteFile.current = outcome.written;
      if (outcome.written && save.loadAt === loaded.current) markClean(save.revision);
      // The file is good, but an edit made after the save started is not in it.
      const later =
        outcome.written && save.loadAt === loaded.current && save.revision !== revision();
      showSaveOutcome(later ? { ...outcome, notice: `${outcome.notice} ${SAVE_AGAIN}` } : outcome);
      return outcome.error === null;
    };
    if (picker) await write();
    else await downloadOnce(save.revision, write);
  };

  /** Prepares the saved bytes for the file, and writes them. */
  const onSave = async (edited: Uint8Array<ArrayBuffer>, result: SavedWorkbook) => {
    // The heading flags below belong to the open workbook, not to the one that was saved.
    if (!result.stillOpen) {
      say(SAVE_REPLACED, 'error');
      return;
    }
    const snapshot = editor.current?.snapshot();
    if (snapshot?.status !== 'ready') {
      say(NO_WORKBOOK, 'error');
      return;
    }
    savingRevision.current = result.revision;
    // The user accepted another workbook, so these bytes are not the user's next file.
    if (replacing.current > 0) {
      say(SAVE_REPLACED, 'error');
      return;
    }
    // The load of the save: a reopen during the write starts a new store, whose revisions start
    // again, so the clean mark applies only to the same load. An edit made while the save ran
    // has a later revision, so it stays unsaved.
    const loadAt = loaded.current;
    // The editor is ready, so a workbook opened and has a name.
    const label = nameNow.current ?? SAMPLE;
    // The file keeps its own heading flags: the headings option changes only the view.
    const restored = headingsNow.current
      ? await runInTasks(withHeadingsSteps(edited, false, flags.current))
      : edited;
    // Without the flags of the file, the save would write the flags of the headings option.
    if (!restored) {
      say(SAVE_UNREADABLE, 'error');
      return;
    }
    const bytes: Uint8Array<ArrayBuffer> = restored === edited ? edited : new Uint8Array(restored);
    // The rewrite yields to the page, so the name above can belong to a replaced workbook.
    if (loadAt !== loaded.current || replacing.current > 0) {
      say(SAVE_REPLACED, 'error');
      return;
    }
    const kind = snapshot.workbookKind;
    const chosenName = chosenNow.current;
    await writeSave({ bytes, revision: result.revision, loadAt, label, chosen: chosenName, kind });
  };

  /**
   * Runs one save at a time: a press during a save does nothing. During an open or a headings
   * switch, a press of the top bar Save, of the keys, or of the toolbar Save comes here
   * (`page-keys.ts`) and is refused before any save work, so presses cannot delay the open. The
   * toolbar also starts no save then (its `opening` prop). Its bytes come here during an open only
   * from a save that started before the open.
   */
  const saveOnce = async (work: () => Promise<void>) => {
    if (switching.current) {
      say('Could not save: the headings are switching. Try again when the switch ends.', 'error');
      return;
    }
    if (opening.current > 0 || openRuns.current > 0) {
      say(SAVE_OPENING, 'error');
      return;
    }
    if (saving.current) {
      const held = savingRevision.current;
      if (held !== null && held !== revision()) laterEdits.current = true;
      return;
    }
    saving.current = true;
    wroteFile.current = false;
    try {
      await busy.during('Saving the workbook.', work());
    } finally {
      saving.current = false;
      savingRevision.current = null;
      // A save that wrote no file keeps its own message: the error or the cancel.
      if (laterEdits.current && wroteFile.current) say(SAVE_AGAIN, 'notice');
      laterEdits.current = false;
    }
  };

  /** Shows the status of a toolbar save from the press, before its bytes arrive. */
  const onToolbarSaveStart = (started: Promise<SavedWorkbook | null>) => {
    toolbarSaves.current += 1;
    const ended = () => (toolbarSaves.current -= 1);
    // The status keeps the open: a text for each press would render the page for each press.
    if (opening.current > 0) void started.finally(ended);
    else void busy.during('Saving the workbook.', started).finally(ended);
  };

  /** Whether the bytes of `save` still hold the open workbook as it is now. */
  const stillReady = (save: ReadySave) => {
    const now = editor.current?.snapshot();
    const typing = now?.editing && now.editChanged;
    const load = { load: loaded.current, revision: revision(), replacing: replacing.current > 0 };
    return keptSaveHolds(save, load) && !typing;
  };

  // The Save of the top bar. Control+S and Command+S go to the toolbar, which calls `onSave`,
  // except during an open and while kept bytes wait: then they come here (`page-keys.ts`).
  const saveNow = useRef<() => void>(() => {});
  saveNow.current = () => {
    void saveOnce(async () => {
      // Bytes that wait for this press go to the picker at once, while the click is recent.
      const ready = readySave.current;
      if (ready && stillReady(ready)) {
        clearMessages();
        await writeSave(ready);
        return;
      }
      const saved = await editor.current?.saveInTasks();
      if (saved) await onSave(saved.bytes, saved);
      // With no open workbook, the save gives nothing and says nothing, so the demo says why.
      else if (editor.current?.snapshot().status !== 'ready') say(NO_WORKBOOK, 'error');
    });
  };
  const openInput = useRef<HTMLInputElement>(null);
  usePageKeys({
    topBarStop: openInput,
    messageOpen: message !== null,
    editing: () => editor.current?.snapshot().editing ?? false,
    closeMessage,
    // A headings switch counts too, from the press to its end: its keys and its toolbar Save go
    // to Save, which gives one text for the switch.
    loading: () =>
      switching.current || isOpening() || editor.current?.snapshot().status === 'loading',
    readyToSave: () => readySave.current !== null,
    onSaveKey: () => saveNow.current(),
    onKeyWhileLoading: () => say(switching.current ? KEY_SWITCHING : KEY_LOADING, 'error'),
  });

  return (
    <XlsxEditor.Root
      ref={editor}
      mode={viewOnly ? 'view' : 'edit'}
      colorMode={theme}
      onChange={onChange}
      {...(locale ? { locale } : {})}
    >
      {/* The page has its own landmarks, such as the top bar's banner, so the viewport renders
          no region of its own. */}
      <XlsxEditor.Viewport className="demo-root" label={false}>
        <header className="demo-topbar">
          <BrandLogo />
          {SHOW_ADAPTER_SWITCHER ? <AdapterSwitcher current="react" theme={theme} /> : null}
          <span className="demo-topbar__divider" />
          <WorkbookName
            label={name}
            chosen={chosen}
            field={nameField}
            onRename={(next) => {
              chosenNow.current = next;
              setChosen(next);
            }}
            onDone={() => editor.current?.focus()}
          />
          {/* The menu bar after the name, as the family demos place theirs. File > Open runs the
              demo's own Open, which asks first over unsaved edits, and File > Save its save.
              File > Rename workbook focuses the name field, as in the PPTX Editor demo. */}
          <XlsxEditor.Menu
            className="demo-menubar"
            opening={openingShown}
            onOpen={() => openInput.current?.click()}
            onSaveStart={onToolbarSaveStart}
            onSave={(bytes, saved) => void saveOnce(() => onSave(bytes, saved))}
          >
            <XlsxEditor.Menu.File>
              <XlsxEditor.Menu.Separator />
              <XlsxEditor.Menu.Row
                label={RENAME_ROW_LABEL}
                icon={<Icon name="text_fields" />}
                disabled={name === null}
                disabledReason="No workbook is open."
                onSelect={() => {
                  nameField.current?.focus();
                  nameField.current?.select();
                }}
              />
            </XlsxEditor.Menu.File>
          </XlsxEditor.Menu>
          <div className="demo-topbar__actions">
            <span className="demo-busy" role="status">
              {busyText}
            </span>
            <label className="demo-pill" title="Show the row and column headings of every sheet">
              <input
                type="checkbox"
                className="demo-sr-only"
                aria-label="Show all headings"
                checked={headings}
                onChange={(e) => void onHeadings(e.target.checked)}
              />
              <Icon name="table_chart" />
              <span className="demo-pill__label">Show all headings</span>
            </label>
            <label className="demo-pill" title="Refuse every edit">
              <input
                type="checkbox"
                className="demo-sr-only"
                aria-label="View only"
                checked={viewOnly}
                onChange={(e) => setViewOnly(e.target.checked)}
              />
              <Icon name="visibility" />
              <span className="demo-pill__label">View only</span>
            </label>
            <ThemeToggle value={theme} onChange={setTheme} />
            <span className="demo-topbar__divider" />
            {/* A styled button for the file input. The input keeps its name when a narrow window
                hides the label text. */}
            <label className="demo-btn demo-btn--quiet demo-btn--sm" title="Open a file">
              <Icon name="folder_open" />
              <span className="demo-btn__label">Open a file</span>
              <input
                ref={openInput}
                className="demo-sr-only"
                type="file"
                aria-label="Open a file"
                aria-keyshortcuts="F6"
                accept=".xlsx,.xltx"
                onChange={(e) => void onFile(e)}
              />
            </label>
            <button
              type="button"
              className="demo-btn demo-btn--primary demo-btn--sm"
              title="Save"
              onMouseDown={(e) => e.preventDefault()}
              // A held Enter repeats the click, and each click would build the file again.
              onKeyDown={(e) => {
                if (e.key === 'Enter' && e.repeat) e.preventDefault();
              }}
              onClick={() => saveNow.current()}
            >
              <Icon name="file_download" />
              <span className="demo-btn__label">Save</span>
            </button>
          </div>
        </header>
        <main className="demo-main">
          <h1 className="demo-sr-only">XLSX Editor</h1>
          <div className="demo-dock">
            <XlsxEditor.Toolbar
              opening={openingShown}
              onSaveStart={onToolbarSaveStart}
              onSave={(bytes, saved) => void saveOnce(() => onSave(bytes, saved))}
            />
          </div>
          <div className="demo-workspace">
            <Workbench />
            {/* The message bar sits under the editor: the grid gets shorter, and it does not
                move. */}
            {/* The status region stays mounted and empty until a notice sets its text, so a
                screen reader announces the notice. An alert is announced when it mounts. */}
            <div
              className={message ? `demo-message demo-message--${message.kind}` : 'demo-sr-only'}
            >
              <span
                className={
                  notice ? 'demo-message__status demo-message__text' : 'demo-message__status'
                }
                role="status"
              >
                {notice?.text}
              </span>
              {message?.kind === 'error' ? (
                <span className="demo-message__text" role="alert">
                  {message.text}
                </span>
              ) : null}
              {message ? (
                <button
                  type="button"
                  className="demo-icon-btn demo-close"
                  aria-label="Close the message"
                  onMouseDown={(e) => e.preventDefault()}
                  onClick={closeMessage}
                >
                  <Icon name="close" />
                </button>
              ) : null}
            </div>
          </div>
        </main>
      </XlsxEditor.Viewport>
    </XlsxEditor.Root>
  );
}

const statusOf = (snapshot: EditorSnapshot) => snapshot.status;

/**
 * The formula bar, the grid, the footer with the sheet tabs and the zoom, and the grid context
 * menu, in one card.
 */
function Workbench() {
  // The status alone: this renders again only when it changes, not on each key of an edit.
  const status = useEditorState(statusOf);
  return (
    <div className="demo-canvas">
      <XlsxEditor.FormulaBar />
      <XlsxEditor.StatusMessage />
      <XlsxEditor.Content className={status === 'ready' ? undefined : 'xlsx-canvas-idle'} />
      <div className="xlsx-footer">
        <XlsxEditor.SheetTabs />
        <XlsxEditor.ZoomControl />
      </div>
      <XlsxEditor.ContextMenu />
    </div>
  );
}
