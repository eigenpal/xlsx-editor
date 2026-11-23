// The workbook editor facade (tasks 6.4 and 7.2). It opens a workbook into a
// `WorkbookDocumentStore`, holds the view state, and mounts the grid surface: a sticky stage that
// paints the visible window into a canvas, with its cell texts as accessible rows. The adapters
// render chrome from `snapshot()` and hold no state. Every edit is one store transaction with its
// undo step; a refused edit returns its typed reason and shows it. DOM code resolves its document
// and window from the host element at use time (`host-realm.ts`), also in a pop-out or an iframe.

import type { GridPoint, GridSelection } from '../contracts/index.ts';
import {
  saveXlsx,
  sheetIndex,
  workbookIndex,
  WorkbookDocumentStore,
  xmlPartOf,
} from '../store/index.ts';
import { EN_US, sameNumberLocale, type NumberLocale } from '../cascade/index.ts';
import { mergedRanges, zoomProfile, type SheetLayout } from '../layout/index.ts';
import { forgetLongTexts } from '../layout/cell-text.ts';
import { createPictureStore } from '../output/index.ts';
import { EditorFonts, type FontMeasurer } from './editor-fonts.ts';
import type { FontReport } from './substitute-fonts.ts';
import { hostTimers, hostWindowOf, registerEditorHost, type HostTimers } from './host-realm.ts';
import { buildSurface, clearSurface, type StageBox, type SurfaceElements } from './grid-surface.ts';
import { scrollToShow, viewStateOf } from './view-map.ts';
import { bindSurface, layoutWindow, openingView, paintWindow } from './surface-render.ts';
import type { SheetView } from './surface-render.ts';
import { shownGeneralOf } from './decimal-selection.ts';
import { GridPointer } from './grid-pointer.ts';
import { englishText } from './english.ts';
import { followSheets, SheetCommands, sheetStepAnnouncement } from './sheet-commands.ts';
import { WorkbookClipboard, type CopyResult, type PasteData } from './clipboard.ts';
import { SurfaceScroll, extentReaching, surfaceExtent } from './scroll-extent.ts';
import { paintedGrid, runResize, type ResizeAxis } from './grid-resize.ts';
import { runFit } from './grid-fit.ts';
import {
  clampPoint,
  extendTo,
  historyGridSelection,
  selectBand,
  selectionWords,
  stepAnnouncement,
  selectCell,
  type Merges,
} from './selection.ts';
import type { CellEditMode, EditMove, EditTextOptions } from './cell-editor.ts';
import { goTo } from './go-to.ts';
import { saveInSlices } from './editor-save.ts';
import { readForLoad } from './load-steps.ts';
import { CellEditing, registerEditing, type EditTarget } from './cell-editing.ts';
import { LiveAnnouncer } from './live-announcer.ts';
import { commitMove } from './grid-keys.ts';
import type { EditResult } from './cell-input.ts';
import type { FormatPatch, HistorySelection, StructuralOpName } from '../store/index.ts';
import type { FormatToggle, MergeOptions } from '../automation/index.ts';
import type { SavedWorkbook, XlsxEditorInstance } from './editor-instance.ts';
import { dataEdge, hiddenLines, NONE_HIDDEN, sheetFacts, type HiddenLines } from './sheet-facts.ts';
import { ConditionalCut } from './conditional-cut.ts';
import { WheelZoom } from './wheel-zoom.ts';
import { cellAnnouncement, selectionAnnouncement } from './cell-announcement.ts';
import {
  EMPTY_SNAPSHOT,
  NO_SHEET_STATE,
  cellFont,
  clampZoom,
  zoomStepFrom,
  openingSheet,
  selectionRef,
  sheetOpening,
  sheetState,
  workbookState,
  NO_WORKBOOK_STATE,
  accessibleLabels,
  editorSheets,
  type EditorFailure,
  type EditorMode,
  type EditorSnapshot,
  type XlsxEditorOptions,
} from './editor-model.ts';

export type { XlsxEditorInstance } from './editor-instance.ts';
import type { Opened } from './opened-workbook.ts';
import { workbookBytes } from './workbook-bytes.ts';
import { EditorListeners } from './editor-events.ts';

class Editor implements XlsxEditorInstance {
  private readonly options: XlsxEditorOptions;
  private readonly events = new EditorListeners(this, () => this.edits.dropRefusal);
  private state: EditorSnapshot = EMPTY_SNAPSHOT;
  private opened: Opened | null = null;
  private active = -1;
  private readonly fonts: EditorFonts;
  private numberLocale: NumberLocale;
  private host: HTMLElement | null = null;
  private els: SurfaceElements | null = null;
  private readonly timers: HostTimers;
  private frame: number | null = null;
  private layout: SheetLayout | null = null;
  private readonly conditionalCut = new ConditionalCut(() => this.publish());
  private readonly live = new LiveAnnouncer(() => this.els?.live, this.words.bind(this));
  private box: StageBox | null = null;
  private readonly wheelZoom = new WheelZoom();
  /** A cell to scroll into view at the next render, when a reveal found no layout. */
  private pendingReveal: GridPoint | null = null;
  private readonly scroll = new SurfaceScroll({
    scroller: () => this.els?.scroller ?? null,
    view: () => this.view(),
    scale: (zoom) => zoomProfile(zoom).pixelsPerInch / 72,
    scrolled: () => {
      this.pointer.syncResize();
      this.schedule();
    },
  });
  private loadToken = 0;
  private resize: ResizeObserver | null = null;
  private offPictures: (() => void) | null = null;
  private readonly pointer: GridPointer;
  private readonly cleanups: (() => void)[] = [];
  private destroyed = false;
  private readonly edits: CellEditing;
  private readonly sheetOps: SheetCommands;
  private readonly clipboard: WorkbookClipboard;
  private offStore: (() => void) | null = null;

  constructor(options: XlsxEditorOptions) {
    this.options = options;
    this.fonts = new EditorFonts(options.fonts);
    this.numberLocale = options.numberLocale ?? EN_US;
    if (options.mode === 'view') this.state = { ...EMPTY_SNAPSHOT, mode: 'view' };
    this.timers = hostTimers(() => this.host);
    registerEditorHost(this, () => this.host);
    this.edits = new CellEditing({
      target: () => this.editTarget(),
      viewing: () => this.state.mode === 'view',
      loading: () => this.state.status === 'loading',
      locale: () => this.numberLocale,
      t: this.t,
      surface: () => (this.els ? { grid: this.els.scroller } : null),
      say: (text) => this.live.say(text),
      silence: () => this.live.silence(),
      merges: () => this.mergesOf(this.active),
      cellFont: (cell) => cellFont(this.opened, this.active, cell, this.measurer),
      shownGeneral: () => this.shown(),
      confirm: (text) => this.live.confirm(text, !this.state.selectionSingle),
      publish: () => this.publish(),
      publishDraft: () => {
        this.state = {
          ...this.state,
          editText: this.edits.draft,
          editChanged: this.edits.changed,
          composing: this.edits.composing,
        };
        this.events.notify();
      },
      schedule: () => this.schedule(),
      render: () => this.render(),
      reveal: (cell) => this.reveal(cell),
      moveFrom: (cell, move, inRange) => this.moveFrom(cell, move, inRange),
      restore: (step, selection, sheet, label, changed) =>
        this.restore(step, selection, sheet, label, changed),
    });
    this.sheetOps = new SheetCommands({
      edits: this.edits,
      index: () => this.opened?.index ?? null,
      active: () => this.active,
      t: (key, params) => this.t(key, params),
      show: (index) => this.setActiveSheet(index),
      announce: (text) => this.live.say(text),
    });
    this.clipboard = new WorkbookClipboard({
      edits: this.edits,
      target: () => this.editTarget(),
      locale: () => this.numberLocale,
      merges: () => this.mergesOf(this.active),
      select: (anchor, focus) => this.select(anchor, focus),
      announce: (text) => this.live.say(text),
      repaint: () => this.schedule(),
      t: (key, params) => this.t(key, params),
      conditionalCut: this.conditionalCut,
    });
    registerEditing(this, this.edits);
    this.pointer = new GridPointer({
      painted: () => {
        const o = this.opened;
        const part = o ? xmlPartOf(o.pkg, o.index.sheets[this.active]?.partName) : undefined;
        return paintedGrid(this.els, this.box, this.layout, this.view(), part);
      },
      // A drag ends when the open, the shown sheet, the revision, the zoom, or the mode changes.
      state: () =>
        `${this.loadToken}:${this.active}:${this.opened?.store.revision ?? -1}:${this.view()?.zoom}:${this.state.mode}`,
      current: () => {
        const view = this.view();
        return view ? { sheet: this.active, selection: view.selection, focus: view.focus } : null;
      },
      setSelection: (selection, focus, reveal) => this.setSelection(selection, focus, reveal),
      select: (anchor, focus) => this.select(anchor, focus),
      extend: (to) => this.extend(to),
      commitOpenEdit: () => this.edits.commitOpenEdit().ok,
      focus: () => this.focus(),
      refusal: () => this.edits.refusal,
      showRefusal: (refusal) => void this.edits.refuse(refusal),
      resize: (axis, first, last, size) => this.resizeAxis(axis, first, last, size),
      sizeState: () => this.state,
    });
  }

  // ------------------------------------------------------------------ state

  snapshot(): EditorSnapshot {
    return this.state;
  }

  readonly subscribe: XlsxEditorInstance['subscribe'] = this.events.subscribe;
  readonly on: XlsxEditorInstance['on'] = this.events.on;

  // Without a translator, the English catalog, as DOCX Editor's core does (`createT(en)`).
  private readonly t = (key: string, params?: Readonly<Record<string, string | number>>): string =>
    (this.options.translate ?? englishText)(key, params);

  private get measurer(): FontMeasurer | null {
    return this.fonts.measurer;
  }

  private view(): SheetView | undefined {
    return this.opened?.views.get(this.active);
  }

  private publish(patch: Partial<EditorSnapshot> = {}): void {
    const o = this.opened;
    const view = this.view();
    const entry = o?.index.sheets[this.active];
    // While a load runs, the old workbook's cell state shows nothing that a command reads.
    const loading = (patch.status ?? this.state.status) === 'loading';
    const sheet =
      o && view && entry && !loading
        ? sheetState(o, entry, view.selection, this.numberLocale, this.edits.pending, this.shown())
        : NO_SHEET_STATE;
    this.state = {
      ...this.state,
      sheets: o?.sheets ?? [],
      activeSheet: o ? this.active : -1,
      selection: view?.selection ?? null,
      selectionRef: view ? selectionRef(view.selection) : null,
      ...this.conditionalCut.apply(sheet, o?.store, o?.index.sheets[this.active]?.partName),
      zoom: view?.zoom ?? 1,
      editing: this.edits.editing,
      editText: this.edits.draft,
      editChanged: this.edits.changed,
      composing: this.edits.composing,
      canUndo: o?.store.canUndo ?? false,
      canRedo: o?.store.canRedo ?? false,
      revision: o?.store.revision ?? 0,
      ...(o ? workbookState(o.pkg, o.index) : NO_WORKBOOK_STATE),
      refusal: this.edits.refusal,
      editDropped: this.edits.dropped,
      numberLocale: this.numberLocale,
      ...patch,
    };
    this.events.notify();
  }

  // ------------------------------------------------------------------ open

  async setFonts(fonts: XlsxEditorOptions['fonts']): Promise<EditorFailure | null> {
    if (this.fonts.holds(fonts)) return null;
    try {
      await this.fonts.replace(fonts);
    } catch (cause) {
      return { code: 'fonts', detail: String(cause) };
    }
    if (this.destroyed) return null;
    this.layout = null;
    this.registerFaces();
    this.publish();
    this.schedule();
    return null;
  }

  fontReport(): FontReport {
    return this.fonts.report;
  }

  setNumberLocale(locale: NumberLocale | undefined): void {
    const next = locale ?? EN_US;
    if (sameNumberLocale(next, this.numberLocale)) return;
    this.numberLocale = next;
    this.layout = null;
    this.publish();
    this.schedule();
  }

  async load(source: Uint8Array | ArrayBuffer): Promise<EditorSnapshot> {
    const token = (this.loadToken += 1);
    // A new workbook drops the open edit at once, and the old announcement goes.
    this.live.clear();
    this.edits.drop();
    this.publish({ status: 'loading', failure: null });
    try {
      await this.fonts.ensure();
    } catch (cause) {
      if (token === this.loadToken) this.failLoad({ code: 'fonts', detail: String(cause) });
      return this.state;
    }
    if (token !== this.loadToken || this.destroyed) return this.state;
    const wanted = () => token === this.loadToken && !this.destroyed;
    const read = await readForLoad(workbookBytes(source), this.host, wanted);
    if (!read) return this.state;
    if (!read.ok) {
      this.failLoad(read.failure);
      return this.state;
    }
    const index = workbookIndex(read.pkg);
    if (!index) {
      this.failLoad({ code: 'package', detail: 'no workbook part' });
      return this.state;
    }
    const sheets = editorSheets(index);
    const first = openingSheet(sheets, index.activeTab);
    this.edits.settleLoad();
    this.release();
    const win = this.host ? hostWindowOf(this.host) : undefined;
    const pictures = createPictureStore(read.pkg, win);
    this.offPictures = pictures.onReady(() => this.schedule());
    const store = new WorkbookDocumentStore(read.pkg);
    this.offStore = store.subscribe(() => this.onStoreChange(store));
    this.opened = { store, pkg: read.pkg, index, sheets, pictures, views: new Map() };
    if (first < 0) {
      this.active = -1;
      this.fail({ code: 'no-worksheet' });
      return this.state;
    }
    this.active = first;
    this.ensureView(first);
    this.publish({ status: 'ready', failure: null });
    this.registerFaces();
    this.schedule();
    return this.state;
  }

  close(): void {
    this.loadToken += 1;
    this.edits.drop();
    this.release();
    this.active = -1;
    this.publish({ status: 'empty', failure: null });
    this.schedule();
  }

  /**
   * A load that fails closes the previous workbook: the grid, the tabs, and the snapshot show
   * no workbook, so nothing shows that the editor would refuse to edit or save.
   */
  private failLoad(failure: EditorFailure): void {
    this.edits.settleLoad();
    this.release();
    this.active = -1;
    this.fail(failure);
  }

  private fail(failure: EditorFailure): void {
    this.publish({ status: 'failed', failure });
    this.schedule();
  }

  private release(): void {
    this.pointer.cancelResize();
    this.clipboard.reset();
    this.offStore?.();
    this.offStore = null;
    this.offPictures?.();
    this.offPictures = null;
    this.opened?.pictures.dispose();
    this.opened = null;
    this.layout = null;
    this.pendingReveal = null;
    forgetLongTexts();
  }

  private ensureView(sheet: number): SheetView {
    const o = this.opened!;
    let view = o.views.get(sheet);
    if (view) return view;
    const opening = sheetOpening(o.pkg, o.index.sheets[sheet]!.partName);
    view = openingView(opening, selectCell(sheet, opening.active, this.mergesOf(sheet)));
    o.views.set(sheet, view);
    return view;
  }

  private mergesOf(sheet: number): Merges {
    const o = this.opened;
    const part = xmlPartOf(o?.pkg ?? { parts: new Map() }, o?.index.sheets[sheet]?.partName);
    return part ? mergedRanges(part.root).ranges : [];
  }

  // ------------------------------------------------------------------ commands

  setActiveSheet(index: number): boolean {
    const s = this.opened?.sheets[index];
    // While another open runs, the shown workbook is about to go: no view command applies.
    if (!s || !s.visible || !s.worksheet || this.state.status !== 'ready') return false;
    if (index === this.active) return true;
    if (!this.commitEdit().ok) return false;
    this.showSheet(index);
    this.live.selected(true);
    this.publish();
    this.schedule();
    return true;
  }

  setZoom(zoom: number): void {
    const view = this.view();
    if (!view || this.state.status !== 'ready') return;
    const next = clampZoom(zoom);
    if (next === view.zoom) return;
    view.zoom = next;
    this.layout = null;
    // A zoom cancels a resize drag, and its guide goes now, not at the next pointer move.
    this.pointer.syncResize();
    this.scroll.restore();
    this.live.say(this.t('grid.zoom', { percent: Math.round(next * 100) }));
    this.publish();
    this.schedule();
  }

  zoomStep(direction: 1 | -1): void {
    const view = this.view();
    if (view) this.setZoom(zoomStepFrom(view.zoom, direction));
  }

  /** `reveal`, for the engine's own callers, is the cell to scroll to: `focus` by default. */
  select(anchor: GridPoint, focus?: GridPoint, reveal?: GridPoint): boolean {
    const view = this.view();
    const finite = (p: GridPoint) => Number.isFinite(p.row) && Number.isFinite(p.col);
    if (!view || this.state.status !== 'ready' || !finite(anchor)) return false;
    if ((focus && !finite(focus)) || !this.edits.commitOpenEdit().ok) return false;
    const merges = this.mergesOf(this.active);
    let sel = selectCell(this.active, anchor, merges);
    if (focus) sel = extendTo(sel, focus, merges);
    view.selection = sel;
    view.focus = focus ? clampPoint(focus) : sel.active;
    this.edits.clearRefusal();
    this.reveal(reveal ? clampPoint(reveal) : view.focus);
    this.live.selected();
    this.publish();
    this.schedule();
    return true;
  }

  setMode(mode: EditorMode): boolean {
    const change = mode !== this.state.mode;
    // A switch to view mode commits an open edit first. A refused commit keeps the editor open
    // with its text and shows the reason, and the switch is refused, as `select` refuses. While
    // an input method composes, it is refused with no word: the user did not ask for it.
    if (change && mode === 'view' && (this.edits.composing || !this.edits.commitOpenEdit().ok))
      return false;
    if (change) this.edits.clearRefusal();
    this.els?.scroller.setAttribute('aria-readonly', String(mode === 'view'));
    this.publish({ mode, modeRequests: this.state.modeRequests + 1 });
    // A switch to view mode cancels a resize drag, and its guide goes.
    this.pointer.syncResize();
    return true;
  }

  focus(): void {
    (this.edits.input ?? this.els?.scroller)?.focus({ preventScroll: true });
  }

  dismissRefusal(): void {
    if (!this.edits.refusal) return;
    this.edits.clearRefusal();
    this.publish();
  }

  // ------------------------------------------------------------------ editing

  /** The store revision changed: follow it, and paint the new values. */
  private onStoreChange(store: WorkbookDocumentStore): void {
    const o = this.opened;
    if (!o || o.store !== store) return;
    this.live.storeChanging(() => this.selectionText(true, o));
    const index = workbookIndex(store.package);
    if (!index) return;
    // A sheet op or its undo moves sheets: each view and the shown sheet follow their sheet.
    const shown = o.index.sheets[this.active]?.partName;
    const next = followSheets(o.index, index, o.views, this.active);
    this.opened = { ...o, pkg: store.package, index, sheets: next.sheets, views: next.views };
    this.active = next.active;
    if (index.sheets[this.active]?.partName !== shown && this.active >= 0) {
      this.ensureView(this.active);
      this.scroll.restore();
    }
    this.layout = null;
    this.pointer.syncResize();
    this.publish();
    this.schedule();
  }

  /** The workbook and sheet that an edit applies to, when one is ready. */
  private editTarget(): EditTarget | null {
    const o = this.opened;
    const view = this.view();
    const entry = o?.index.sheets[this.active];
    if (!o || !view || !entry || this.state.status !== 'ready') return null;
    const [{ store, pkg, index }, { selection, focus }] = [o, view];
    const revision = store.revision;
    return { store, pkg, index, sheet: this.active, entry, selection, focus, revision };
  }

  enterValue(text: string, at?: GridPoint): EditResult {
    return this.edits.enterValue(text, at);
  }

  clearValues(): EditResult {
    return this.edits.clearValues();
  }

  applyFormat(patch: FormatPatch): EditResult {
    return this.edits.applyFormat(patch);
  }

  clearFormats(): EditResult {
    return this.edits.clear('formats');
  }

  clearAll(): EditResult {
    return this.edits.clear('all');
  }

  toggleFormat(which: FormatToggle): EditResult {
    return this.edits.toggleFormat(which);
  }

  insertRows(): EditResult {
    return this.edits.structure('insertRows');
  }

  deleteRows(): EditResult {
    return this.edits.structure('deleteRows');
  }

  insertColumns(): EditResult {
    return this.edits.structure('insertColumns');
  }

  deleteColumns(): EditResult {
    return this.edits.structure('deleteColumns');
  }

  mergeCells(options: MergeOptions = {}): EditResult {
    const r = this.edits.merge(options);
    const range = r.ok ? this.view()?.selection.range : undefined;
    if (range) this.select(range.start, range.end, range.start);
    return r;
  }

  unmergeCells(): EditResult {
    return this.edits.merge('unmerge');
  }

  addSheet(name?: string): EditResult {
    return this.sheetOps.add(name);
  }

  renameSheet(index: number, name: string): EditResult {
    return this.sheetOps.rename(index, name);
  }

  deleteSheet(index: number): EditResult {
    return this.sheetOps.delete(index);
  }

  moveSheet(index: number, to: number): EditResult {
    return this.sheetOps.move(index, to);
  }

  copy(): CopyResult {
    return this.clipboard.copy();
  }

  cut(): CopyResult {
    return this.clipboard.cut();
  }

  paste(data: PasteData): EditResult {
    return this.clipboard.paste(data);
  }

  /** Escape on the grid: end a pending copy or cut (`endPending`). @internal */
  escape(): boolean {
    return this.clipboard.endPending();
  }

  setColumnWidth(first: number, last: number, width: number): EditResult {
    return this.resizeAxis('columns', first, last, width);
  }

  setRowHeight(first: number, last: number, height: number): EditResult {
    return this.resizeAxis('rows', first, last, height);
  }

  private resizeAxis(
    axis: ResizeAxis,
    first: number,
    last: number,
    size: number | 'fit'
  ): EditResult {
    if (size === 'fit') {
      const zoom = this.view()?.zoom;
      return runFit(this.edits, axis, [first, last], this.measurer, this.numberLocale, zoom);
    }
    const sheet = this.opened?.index.sheets[this.active]?.name;
    return runResize(this.edits, sheet, axis, first, last, size);
  }

  /** A structural command from the key table. */
  structure(name: StructuralOpName): void {
    this.edits.structure(name);
  }

  undo(): EditResult {
    return this.edits.undo();
  }

  redo(): EditResult {
    return this.edits.redo();
  }

  /** Show sheet `index`, at its kept scroll and selection, with no announcement. */
  private showSheet(index: number): void {
    this.edits.clearRefusal();
    this.active = index;
    this.pendingReveal = null;
    this.ensureView(index);
    this.layout = null;
    this.pointer.syncResize();
    this.scroll.restore();
  }

  /**
   * After an undo or a redo: take the selection that the history entry keeps, on its sheet,
   * and reveal its active cell, never the end of a band, as a heading press does. One
   * announcement names the step, the sheet when it was not shown, and the selection.
   */
  private restore(
    step: 'undo' | 'redo',
    s: HistorySelection | null,
    of: string | null,
    label: string | null,
    changed: string | null
  ): void {
    const sheets = this.opened?.sheets;
    const sheet = s ? (this.opened?.index.sheets.findIndex((e) => e.name === s.sheet) ?? -1) : -1;
    const shown = sheet >= 0 && sheets?.[sheet]?.visible === true && !!sheets[sheet]!.worksheet;
    const moved = shown && sheet !== this.active;
    if (moved) this.showSheet(sheet);
    const view = shown ? this.view() : undefined;
    if (s && view) {
      ({ selection: view.selection, focus: view.focus } = historyGridSelection(s, this.active));
      this.reveal(view.selection.active);
    }
    const text = view ? this.selectionText() : undefined;
    const cell = this.live.changed(() => this.selectionText(true));
    // A step on a sheet that was not shown names that sheet, also when the sheet is hidden or
    // the step kept no selection; the selection then stays where it is.
    const name = s?.sheet ?? of ?? undefined;
    const shownName = this.opened?.index.sheets[this.active]?.name;
    const named = moved || name !== shownName ? name : undefined;
    // A sheet op names its change and the sheet that shows after the step.
    const words = this.view() ? this.selectionText() : undefined;
    const shownStep =
      shownName !== undefined && words !== undefined
        ? { sheet: shownName, selection: words }
        : undefined;
    this.live.say(
      sheetStepAnnouncement(this.t, step, label, changed, shownStep) ??
        stepAnnouncement(this.t, step, text, named, cell)
    );
    this.publish();
    this.schedule();
  }

  save(): Uint8Array<ArrayBuffer> | null {
    // A failed or pending load shows no workbook, so it saves none, not the previous one.
    if (!this.opened || this.state.status !== 'ready') return this.edits.refuseSave(null);
    const open = this.edits.commitOpenEdit();
    if (!open.ok) return this.edits.refuseSave(open.refusal);
    return this.edits.saved(saveXlsx(this.opened.store.package));
  }

  saveInTasks(): Promise<SavedWorkbook | null> {
    const store = this.state.status === 'ready' ? (this.opened?.store ?? null) : null;
    const token = this.loadToken;
    const isCurrent = () => token === this.loadToken && !this.destroyed;
    return saveInSlices(store, this.edits, this.host, isCurrent);
  }

  beginEdit(text?: string, mode?: CellEditMode): EditResult {
    return this.edits.begin(text, mode);
  }

  commitEdit(move?: EditMove): EditResult {
    return move === undefined ? this.edits.commit() : this.edits.commitMoving(move);
  }

  setEditText(text: string, options?: EditTextOptions): EditResult {
    return this.edits.typeInBar(text, options?.composing === true);
  }

  selectReference(reference: string): EditResult {
    const o = this.opened;
    const shown = o && this.state.status === 'ready' ? { ...o, active: this.active } : null;
    return goTo(reference, this.edits, shown, this);
  }

  cancelEdit(): void {
    this.edits.cancel();
  }

  // ------------------------------------------------------------------ surface

  attach(element: HTMLElement): void {
    if (this.destroyed) return;
    if (this.host === element) return;
    this.detach();
    this.host = element;
    element.replaceChildren();
    const els = buildSurface(element);
    this.els = els;
    this.edits.attach(els.stage);
    bindSurface(this.cleanups, els, {
      editor: this,
      pointer: this.pointer,
      input: () => this.edits.input ?? null,
      beginEdit: (text, mode) => void this.beginEdit(text, mode),
      onScroll: () => this.scroll.onScroll(),
      exitHint: () => this.t('grid.exitHint'),
      live: this.live,
      // Mod+wheel and a trackpad pinch zoom (`WheelZoom`); a plain wheel scrolls.
      onWheel: (e) => {
        const step = this.wheelZoom.step(e);
        if (step === undefined) return;
        e.preventDefault();
        if (step !== 0) this.zoomStep(step);
      },
    });
    const win = hostWindowOf(element);
    if (typeof win.ResizeObserver === 'function') {
      this.resize = new win.ResizeObserver(() => this.schedule());
      this.resize.observe(els.scroller);
    }
    if (this.opened && this.measurer) {
      // Pictures decode in the window the surface lives in now.
      this.offPictures?.();
      this.opened.pictures.dispose();
      const pictures = createPictureStore(this.opened.pkg, win);
      this.opened = { ...this.opened, pictures };
      this.offPictures = pictures.onReady(() => this.schedule());
    }
    this.registerFaces();
    this.schedule();
  }

  detach(): void {
    this.pointer.cancelResize();
    this.edits.detach();
    this.timers.cancelAnimationFrame(this.frame);
    this.frame = null;
    for (const c of this.cleanups.splice(0)) c();
    this.resize?.disconnect();
    this.resize = null;
    this.fonts.unregister();
    this.live.clear(false);
    this.host?.replaceChildren();
    this.host = null;
    this.els = null;
  }

  destroy(): void {
    this.detach();
    this.release();
    this.destroyed = true;
    this.events.clear();
  }

  private registerFaces(): void {
    if (this.host) this.fonts.register(this.host.ownerDocument, () => this.schedule());
  }

  private schedule(): void {
    if (!this.els || this.frame !== null) return;
    this.frame = this.timers.requestAnimationFrame(() => {
      this.frame = null;
      this.render();
    });
  }

  /** The `General` text that the grid shows on the active sheet, for the decimal steps. */
  private shown = (o = this.opened) =>
    shownGeneralOf(o, this.active, this.measurer, this.view()?.zoom, this.numberLocale);

  /** Paint the window now. Called from an animation frame; tests call it through `flush`. */
  render(): void {
    const els = this.els;
    if (!els) return;
    const o = this.opened;
    const view = this.view();
    const entry = o?.index.sheets[this.active];
    els.scroller.setAttribute('aria-busy', this.state.status === 'loading' ? 'true' : 'false');
    if (!o || !view || !entry || !this.measurer) {
      clearSurface(els);
      return;
    }
    this.live.settle(() => this.selectionText(true));
    els.scroller.setAttribute('aria-label', this.t('grid.label', { sheet: entry.name }));
    // Names the exit from the grid before the user needs it (`grid-exit.ts`).
    els.scroller.setAttribute('aria-description', this.t('grid.exitDescription'));
    els.scroller.setAttribute('aria-readonly', String(this.state.mode === 'view'));
    const done = layoutWindow(view, {
      pkg: o.pkg,
      sheetName: entry.name,
      previous: this.layout,
      headWidth: this.box?.headWidth,
      width: els.scroller.clientWidth,
      height: els.scroller.clientHeight,
      measurer: this.measurer,
      profile: zoomProfile(view.zoom),
      locale: this.numberLocale,
    });
    if (!done) return;
    const { layout, box } = done;
    this.layout = layout;
    if (layout.conditionalCut) this.conditionalCut.mark(o.store, entry.partName);
    if (view.pendingTopLeft) this.schedule();
    this.box = box;
    const pending = this.pendingReveal;
    if (pending) {
      // A reveal asked for while there was no layout: apply it, and lay out the new window.
      this.pendingReveal = null;
      const at = { x: view.scrollX, y: view.scrollY };
      this.reveal(pending);
      if (view.scrollX !== at.x || view.scrollY !== at.y) return this.render();
    }
    this.scroll.size(surfaceExtent(layout, this.facts(), view.focus, box), els.extent);
    const gridW = Math.max(0, box.width - box.headWidth) / box.scale;
    const gridH = Math.max(0, box.height - box.headHeight) / box.scale;
    const map = viewStateOf(layout, view, gridW, gridH);
    paintWindow(els, layout, {
      box,
      map,
      pictures: o.pictures,
      labels: accessibleLabels(this.t),
      selection: view.selection.range,
      active: selectCell(this.active, view.selection.active, this.mergesOf(this.active)).range,
      pending: this.clipboard.pendingRange(entry.partName),
      headings: view.headings,
    });
    this.edits.place(this.active, map, box);
  }

  /** Scroll so a cell shows. */
  private reveal(cell: GridPoint): void {
    const view = this.view();
    const layout = this.layout;
    const box = this.box;
    // An edit that just committed dropped the layout: the next render reveals the cell.
    if (view && (!layout || !box)) this.pendingReveal = cell;
    if (!view || !layout || !box) return;
    const next = scrollToShow(
      viewStateOf(
        layout,
        view,
        (box.width - box.headWidth) / box.scale,
        (box.height - box.headHeight) / box.scale
      ),
      cell
    );
    view.scrollX = next.x;
    view.scrollY = next.y;
    // The extent grows to the cell first, so the jump is not cut short (`extentReaching`).
    if (!this.els) return this.scroll.restore();
    this.scroll.size(extentReaching(layout, this.facts(), view.focus, cell, box), this.els.extent);
  }

  /** The words of the selection for the live region, after its sheet when the sheet changed. */
  private words(sheetChanged = false) {
    const [sel, text] = [this.view()?.selection, this.selectionText()];
    const entry = this.opened?.index.sheets[this.active];
    if (!sel || text === undefined || !entry) return undefined;
    return selectionWords(this.t, sel, text, this.selectionText(true), entry.name, sheetChanged);
  }

  /** The words of the selection, or with `cell` of the active cell, of the shown sheet in `o`. */
  private selectionText(cell = false, o = this.opened): string | undefined {
    const [sel, t, at] = [this.view()?.selection, this.t, this.numberLocale];
    const entry = this.opened?.index.sheets[this.active];
    if (!sel || !o || !entry) return undefined;
    const rules = !this.conditionalCut.holds(o.store, entry.partName);
    const a = [at, t, rules, this.shown(o)] as const;
    if (cell) return cellAnnouncement(o.pkg, o.index, entry.partName, sel.active, ...a);
    return selectionAnnouncement(o.pkg, o.index, entry.partName, sel, this.merges, ...a);
  }

  // ------------------------------------------------------------------ input

  /** The hidden lines of the newest revision, also while the layout is stale or not built. */
  hidden(): HiddenLines {
    const o = this.opened;
    const part = o ? xmlPartOf(o.pkg, o.index.sheets[this.active]?.partName) : undefined;
    return part ? hiddenLines(sheetIndex(part)) : NONE_HIDDEN;
  }

  /** Visible rows in the window, for Page Up and Page Down. */
  get pageRows(): number {
    const rows = this.layout?.panes[this.layout.panes.length - 1]?.rows.length ?? 10;
    return Math.max(1, rows - 1);
  }

  // The members that the key table (`grid-keys.ts`) reads.

  get selection(): GridSelection {
    return this.view()!.selection;
  }

  get focusCell(): GridPoint {
    return this.view()!.focus;
  }

  get merges(): Merges {
    return this.mergesOf(this.active);
  }

  facts(): { lastRow: number; lastCol: number } {
    const part = xmlPartOf(this.opened!.pkg, this.opened!.index.sheets[this.active]?.partName);
    return part ? sheetFacts(sheetIndex(part)) : { lastRow: -1, lastCol: -1 };
  }

  dataEdge(from: GridPoint, dRow: number, dCol: number): GridPoint {
    const part = xmlPartOf(this.opened!.pkg, this.opened!.index.sheets[this.active]?.partName);
    return part ? dataEdge(sheetFacts(sheetIndex(part)), from, dRow, dCol, this.hidden()) : from;
  }

  private moveFrom(cell: GridPoint, move: readonly [number, number], inRange: boolean): void {
    const view = this.view();
    if (!view) return;
    const m = commitMove(view.selection, cell, move, inRange, this.merges, this.hidden());
    if (m.keep) this.setSelection({ ...view.selection, active: m.at }, view.focus, m.at);
    else this.select(m.at);
  }

  extend(to: GridPoint): void {
    const view = this.view();
    if (view) this.setSelection(extendTo(view.selection, to, this.merges), clampPoint(to), to);
  }

  selectAll(): void {
    const sel = selectBand(this.active, 'all', 0, 0);
    this.setSelection(sel, sel.active);
  }

  /**
   * Take a selection as `select` does: the old refusal clears, and the live region announces it.
   * Heading, corner, and Mod+A selections come here. `reveal` scrolls that cell into view: the key
   * table and a heading drag pass the cell that moved; Mod+A and the corner pass none.
   */
  setSelection(selection: GridSelection, focus: GridPoint, reveal?: GridPoint): void {
    const view = this.view();
    if (!view || !this.edits.commitOpenEdit().ok) return;
    view.selection = selection;
    view.focus = focus;
    this.edits.clearRefusal();
    if (reveal) this.reveal(reveal);
    this.live.selected();
    this.publish();
    this.schedule();
  }

  cycleSheet(direction: 1 | -1): void {
    this.sheetOps.cycle(direction);
  }
}

/**
 * Create a workbook editor. Mount it with `attach`, open a file with `load`.
 *
 * @public
 */
export function createXlsxEditor(options: XlsxEditorOptions = {}): XlsxEditorInstance {
  return new Editor(options);
}

/**
 * Paint an editor's window now, without waiting for an animation frame. For tests.
 *
 * @internal
 */
export function flushEditor(editor: XlsxEditorInstance): void {
  if (editor instanceof Editor) editor.render();
}
