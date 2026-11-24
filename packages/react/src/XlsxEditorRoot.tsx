/** @jsxImportSource react */
// Provider-first host for the workbook editor, after PPTX Editor's `PptxEditorRoot.tsx` (which
// follows docx-editor/packages/react/src/editor/DocxEditorRoot.tsx).
//
// `Root` renders no DOM of its own: it creates the editor in an effect (StrictMode-safe: a fresh
// instance per effect run, destroyed on cleanup), publishes it through context, and lets
// `Content` attach the grid surface wherever the host's tree puts it.

import {
  forwardRef,
  useEffect,
  useImperativeHandle,
  useLayoutEffect,
  useRef,
  useState,
  type CSSProperties,
  type ReactNode,
} from 'react';
import {
  createXlsxEditor,
  cycleRegion,
  emptySnapshotOf,
  followModeProp,
  relabelEditor,
  reportFontError,
  sameFontSource,
  sameNumberLocale,
  watchFontErrors,
  watchGridSetup,
  whenOpenSettles,
  type FontSource,
  type ModeProp,
  type NumberLocale,
  type EditorFailure,
  type EditorMode,
  type EditorSnapshot,
  type SavedWorkbook,
  type XlsxEditorInstance,
  type ColorMode,
} from '@xlsx-editor.dev/core/editor';
import type { TFunction, Translations } from '@xlsx-editor.dev/i18n';
import {
  CatalogContext,
  RootModeContext,
  ModeWaitsContext,
  LayoutPartsContext,
  TranslationContext,
  XlsxEditorContext,
  useEditorState,
  useLayoutPart,
  useTranslation,
  useXlsxEditor,
  type LayoutParts,
} from './context.ts';
import { isDevelopment } from './is-development.ts';
import { useCatalogScope } from './locale.tsx';
import {
  ChromeDarkContext,
  ScopedByAncestorContext,
  useChromeDark,
  useResolvedDark,
} from './color-mode.ts';

/** The faces that measure and paint text: bytes, or a function that loads them once. @public */
export type XlsxEditorFonts = readonly FontSource[] | (() => Promise<readonly FontSource[]>);

/** Props for `XlsxEditor.Root`. Only `document` identity reloads the workbook. @public */
export interface XlsxEditorRootProps {
  /**
   * The bytes of a workbook or a template package, as a `Uint8Array` or an `ArrayBuffer`. A new
   * array or buffer opens the new workbook. A change
   * to `undefined` closes the open workbook (`close`), so the editor shows no workbook, as DOCX
   * Editor mounts no document without one.
   */
  readonly document?: Uint8Array | ArrayBuffer | undefined;
  /**
   * The faces that measure and paint text. Omitted, the packaged substitute faces load
   * (`loadSubstituteFonts`); nothing else is fetched. As DOCX Editor's `fonts` prop, a new
   * value replaces the faces in place (`setFonts`): the workbook, the selection, and the undo
   * history stay. A list is compared by its faces (`sameFontSource`), so an inline list does
   * not load the faces again; a function is compared by identity, so keep it stable. When new
   * faces do not load, the old faces stay and `onFontError` fires with the cause.
   */
  readonly fonts?: XlsxEditorFonts;
  /**
   * The number locale of the shown values: separators and built-in format codes. A value that
   * differs (`sameNumberLocale`) applies in place (`setNumberLocale`); an inline object with
   * the same value changes nothing. An open cell edit keeps the number locale it opened with;
   * the next edit reads the new one. A separate prop from `locale` (the UI
   * strings): the engine formats numbers from explicit separators, not a BCP-47 tag, and ships
   * no regional data.
   */
  readonly numberLocale?: NumberLocale;
  /**
   * `'edit'` (default) or `'view'`. In `'view'`, every edit command is refused and the
   * toolbar's edit controls are disabled; selection, sheet tabs, zoom, and save stay
   * available. A change applies without a reload. For when a change applies and waits, see
   * "View mode" in the README.
   */
  readonly mode?: EditorMode;
  /**
   * UI strings; English when omitted. A missing or `null` entry falls back to the catalog of an
   * outer `LocaleProvider`, else to English. `_lang` sets the plural rules, and `_dir` or a
   * right-to-left `_lang` mirrors the chrome (`useChromeDir`).
   */
  readonly locale?: Translations;
  /**
   * Color theme of the chrome, as on DOCX Editor's root. `'dark'` puts `.dark` beside
   * `.xlsx-editor` on the viewport, and `'system'` follows the OS setting as it changes.
   * Defaults to `'light'`. The workbook canvas keeps the workbook's own colors. Changes apply
   * without a remount.
   */
  readonly colorMode?: ColorMode;
  /**
   * Fired once per editor, as DOCX Editor's `onReady` is: after the editor is published and any
   * `XlsxEditor.Content` in the same commit has attached it, and after the first open of
   * `document` settles (`snapshot().status` is `'ready'` or `'failed'`). Without `document`, it
   * fires right after the attach. A `select` or a zoom from it applies to the opened workbook.
   */
  readonly onReady?: (editor: XlsxEditorInstance) => void;
  /**
   * Fired after every change of the editor state, with the new snapshot: a content change and
   * a view change (selection, zoom, sheet) alike. To act on content changes only, as DOCX
   * Editor's `onChange` does, compare `snapshot.revision` with the last one. DOCX Editor
   * passes a document change and PPTX Editor the instance; this editor passes the snapshot,
   * because it is the one immutable state object that chrome renders, and it carries the
   * revision. The `change` event of `useEditorEvent` fires for content changes only.
   */
  readonly onChange?: (snapshot: EditorSnapshot) => void;
  /**
   * Fired when the fonts that measure and paint text do not load, as DOCX Editor's
   * `onFontError` is. When an open fails for its fonts, the snapshot also turns `failed` with
   * the failure code `fonts`; fires once for each failed open. When a new `fonts` value does
   * not load, the old faces stay and it fires too. The argument differs from DOCX Editor's
   * `EditorFontError`: it is the snapshot's `EditorFailure` with the code `fonts` and the cause
   * in `detail`, the same object that `snapshot().failure` holds, so a host reads one shape.
   * When some packaged substitute faces do not load and the others do (`fonts` omitted), the
   * workbook opens, and it fires with `substituteFailures`, as PPTX Editor reports them
   * (`fontReport()`). Without a handler, each failure prints one console warning.
   */
  readonly onFontError?: (failure: EditorFailure) => void;
  readonly children?: ReactNode;
}

/**
 * The imperative handle of `XlsxEditor` and `XlsxEditor.Root`. It stays small: it never
 * mirrors the commands of the editor instance, such as the sheet, clipboard, and resize commands.
 * Call them on `getEditor()`.
 *
 * It differs from DOCX Editor's `DocxEditorRef` in four ways, for these reasons:
 *
 * - No `exec`. Each command is a typed method of `XlsxEditorInstance`, and the engine has no
 *   command union to dispatch on. A toolbar slot runs through `useEditorCommand`.
 * - No `getDocumentHandle`. The workbook has no document handle API; read the state with
 *   `snapshot()` and edit through `getEditor()`.
 * - `save()` is synchronous, as PPTX Editor's is, because the package writer is synchronous.
 * - `load()` returns a promise of the snapshot after the open, where DOCX Editor's returns
 *   nothing: opening loads fonts and parses the package asynchronously, so the caller can wait.
 *
 * @public
 */
export interface XlsxEditorRef {
  /**
   * Open the bytes of a workbook or a template package. Resolves with the snapshot after the
   * open, or `null` before mount. An open cell edit is dropped without a commit, as
   * `XlsxEditorInstance.load` does.
   */
  load(bytes: Uint8Array | ArrayBuffer): Promise<EditorSnapshot | null>;
  /**
   * The workbook as bytes of the opened package's kind (`snapshot().workbookKind`: a workbook or a
   * template, macro-enabled when `snapshot().macroEnabled` is true), with every edit, or `null`
   * when none is open. An open cell edit is committed first; when the editor cannot store it,
   * `save` returns `null` and shows the reason.
   */
  save(): Uint8Array<ArrayBuffer> | null;
  /**
   * `save` in short tasks, so the page stays responsive while a large workbook saves. An edit
   * made while the save runs belongs to the next save. Resolves to `null` when it is refused.
   */
  saveInTasks(): Promise<SavedWorkbook | null>;
  /** The editor instance, with every editing command, or `null` before it mounts. */
  getEditor(): XlsxEditorInstance | null;
  /** Give the keyboard to the grid. */
  focus(): void;
  /** The editor's state now. Before it mounts, an empty snapshot with the `mode` prop. */
  snapshot(): EditorSnapshot;
}

/** Owns the editor instance. Takes a `ref` to an `XlsxEditorRef`. @public */
export const XlsxEditorRoot = forwardRef<XlsxEditorRef, XlsxEditorRootProps>(
  function XlsxEditorRoot(props, ref) {
    const { document: bytes, locale, mode = 'edit', colorMode, children } = props;
    const dark = useResolvedDark(colorMode);
    const [editor, setEditorState] = useState<XlsxEditorInstance | null>(null);
    const live = useRef<XlsxEditorInstance | null>(null);
    const callbacks = useRef({
      onReady: props.onReady,
      onChange: props.onChange,
      onFontError: props.onFontError,
    });
    callbacks.current = {
      onReady: props.onReady,
      onChange: props.onChange,
      onFontError: props.onFontError,
    };
    const created = useRef({ fonts: props.fonts, numberLocale: props.numberLocale, mode });
    // The `locale` prop merges onto the catalog of an outer `LocaleProvider`, else English. The
    // scope follows the content of the catalog, so an inline `locale={{...}}` builds no new
    // translator on each render.
    const scope = useCatalogScope(locale, undefined);
    const t: TFunction = scope.t;
    const tRef = useRef(t);
    tRef.current = t;
    const modeRef = useRef(mode);
    modeRef.current = mode;

    // A locale change relabels the surface once: the grid label, the accessible stale words, and
    // the cell editor label, which the engine translates through `tRef`. `t` follows the content
    // of the catalog, so an inline `locale={{...}}` does not repaint on each render.
    const labeled = useRef(t);
    useEffect(() => {
      if (!editor || labeled.current === t) return;
      labeled.current = t;
      relabelEditor(editor);
    }, [editor, t]);

    useImperativeHandle(
      ref,
      () => ({
        load: async (b) => (live.current ? live.current.load(b) : null),
        save: () => live.current?.save() ?? null,
        saveInTasks: async () => (live.current ? live.current.saveInTasks() : null),
        getEditor: () => live.current,
        focus: () => live.current?.focus(),
        snapshot: () => live.current?.snapshot() ?? emptySnapshotOf(modeRef.current),
      }),
      []
    );

    useEffect(() => {
      const { fonts, numberLocale, mode: initialMode } = created.current;
      const instance = createXlsxEditor({
        // Omitted, the engine loads the packaged substitute faces and keeps their failures.
        ...(fonts ? { fonts } : {}),
        mode: initialMode,
        ...(numberLocale ? { numberLocale } : {}),
        // The engine's keys are keys of the shared catalog.
        translate: (key, params) =>
          tRef.current(key as Parameters<TFunction>[0], params ? { ...params } : undefined),
      });
      live.current = instance;
      setEditorState(instance);
      const offFonts = watchFontErrors(instance, (f) =>
        reportFontError(f, callbacks.current.onFontError)
      );
      const off = instance.subscribe(() => callbacks.current.onChange?.(instance.snapshot()));
      return () => {
        offFonts();
        off();
        if (live.current === instance) live.current = null;
        instance.destroy();
      };
    }, []);

    const opened = useRef(false);
    useEffect(() => {
      if (!editor) return;
      if (bytes) void editor.load(bytes);
      else if (opened.current) editor.close();
      opened.current = bytes !== undefined;
    }, [editor, bytes]);

    // After the load effect, in the commit that published the editor: a `Content` in the tree
    // has attached, and the first open has started. `onReady` waits for that open to settle.
    useEffect(() => {
      if (!editor) return undefined;
      return whenOpenSettles(editor, () => callbacks.current.onReady?.(editor));
    }, [editor]);

    // A new `fonts` or `numberLocale` value replaces the current one in place, as DOCX
    // Editor's root does; the first values were passed at creation.
    const { fonts, numberLocale } = props;
    useEffect(() => {
      if (!editor || sameFontSource(fonts, created.current.fonts)) return;
      created.current = { ...created.current, fonts };
      void editor.setFonts(fonts).then((failure) => {
        if (failure) reportFontError(failure, callbacks.current.onFontError);
      });
    }, [editor, fonts]);
    useEffect(() => {
      if (!editor || sameNumberLocale(numberLocale, created.current.numberLocale)) return;
      created.current = { ...created.current, numberLocale };
      editor.setNumberLocale(numberLocale);
    }, [editor, numberLocale]);

    // The `mode` prop applies when it changes, as DOCX Editor's root applies it: an instance
    // call to `setMode` stays until the next prop change. A change that the open cell edit
    // refuses waits, with a notice, and applies when the edit closes (`followModeProp`). A new
    // editor ends the wait of the old one.
    const modeProp = useRef<ModeProp | null>(null);
    const [modeWaits, setModeWaits] = useState(false);
    useEffect(() => {
      if (!editor) return undefined;
      const follow = followModeProp(editor, setModeWaits);
      modeProp.current = follow;
      return () => {
        follow.dispose();
        if (modeProp.current === follow) modeProp.current = null;
      };
    }, [editor]);
    useEffect(() => {
      modeProp.current?.apply(mode);
    }, [editor, mode]);

    // In development, a layout that shows the grid without `StatusMessage` gets one warning:
    // refusals and the stale notice show only there. The children's effects run before this one.
    const parts = useRef<LayoutParts>({ grids: 0, statusMessages: 0 });
    useEffect(() => {
      const { grids, statusMessages } = parts.current;
      if (!editor || grids === 0 || statusMessages > 0 || !isDevelopment()) return;
      console.warn(
        'XlsxEditor.Root: the layout has no XlsxEditor.StatusMessage, so edit refusals and the ' +
          'stale-value notice are not shown. Add <XlsxEditor.StatusMessage /> to the layout.'
      );
    }, [editor]);

    return (
      <XlsxEditorContext.Provider value={editor}>
        <RootModeContext.Provider value={mode}>
          <ModeWaitsContext.Provider value={modeWaits}>
            <CatalogContext.Provider value={scope}>
              <TranslationContext.Provider value={t}>
                <LayoutPartsContext.Provider value={parts.current}>
                  <ChromeDarkContext.Provider value={dark}>{children}</ChromeDarkContext.Provider>
                </LayoutPartsContext.Provider>
              </TranslationContext.Provider>
            </CatalogContext.Provider>
          </ModeWaitsContext.Provider>
        </RootModeContext.Provider>
      </XlsxEditorContext.Provider>
    );
  }
);

/** The props of `XlsxEditor.Viewport`. @public */
export interface XlsxEditorViewportProps {
  /** Appended after the `xlsx-editor` and `xlsx-editor-viewport` classes. */
  readonly className?: string;
  /**
   * The name of the editor's landmark region. Default: "Workbook editor". `false` renders no
   * region, for a host whose own landmarks are inside.
   */
  readonly label?: string | false;
  /** Inline style, for example the editor's height. */
  readonly style?: CSSProperties;
  readonly children?: ReactNode;
}

/**
 * The editor's styling scope and landmark region: carries the `.xlsx-editor` class, and
 * `xlsx-editor-viewport`, which keeps its width from following its content. F6 and
 * Shift+F6 inside it move the keyboard among the menu bar, the toolbar, the formula bar, the
 * grid, and the footer.
 *
 * @public
 */
export function XlsxEditorViewport({ className, label, style, children }: XlsxEditorViewportProps) {
  const { t } = useTranslation();
  const editor = useXlsxEditor();
  const dark = useChromeDark();
  const region =
    label === false
      ? {}
      : { role: 'region' as const, 'aria-label': label ?? t('viewport.ariaLabel') };
  return (
    <div
      className={['xlsx-editor', dark ? 'dark' : '', 'xlsx-editor-viewport', className]
        .filter(Boolean)
        .join(' ')}
      {...(style ? { style } : {})}
      {...region}
      // F6 and Shift+F6 move the keyboard among the menu bar, the toolbar, the formula bar, the
      // grid, and the footer.
      onKeyDown={(e) => void cycleRegion(e, e.currentTarget, editor)}
    >
      <ScopedByAncestorContext.Provider value>{children}</ScopedByAncestorContext.Provider>
    </div>
  );
}

/** The props of `XlsxEditor.Content`. @public */
export interface XlsxEditorContentProps {
  /** Appended after the `xlsx-canvas` class. */
  readonly className?: string;
}

/** Whether the editor shows a workbook, so the grid has something to show. */
const showsWorkbook = (snapshot: EditorSnapshot): boolean => snapshot.status === 'ready';

/**
 * The grid: attaches the editor's surface in a layout effect. While no workbook is ready (empty,
 * loading, or failed), the grid is hidden, so a custom layout shows only `StatusMessage` in that
 * state, as the composed editor does. @public
 */
export function XlsxEditorContent({ className }: XlsxEditorContentProps) {
  const editor = useXlsxEditor();
  const ready = useEditorState(showsWorkbook);
  useLayoutPart('grids');
  const host = useRef<HTMLDivElement>(null);
  useLayoutEffect(() => {
    if (!editor || !host.current) return;
    editor.attach(host.current);
    return () => editor.detach();
  }, [editor]);
  // In development, name a missing stylesheet, which leaves the grid with no scroll container.
  // The check waits for the document load, so a stylesheet that loads late gives no warning.
  useEffect(() => {
    const element = host.current;
    if (!editor || !element || !isDevelopment()) return undefined;
    return watchGridSetup(element, 'XlsxEditor.Content', '@xlsx-editor.dev/react/styles.css');
  }, [editor]);
  const classes = ['xlsx-canvas', ready ? null : 'xlsx-canvas-idle', className];
  return <div ref={host} className={classes.filter(Boolean).join(' ')} />;
}
