// Provider-first host for the workbook editor: the Vue twin of the React adapter's
// `XlsxEditorRoot.tsx`, after PPTX Editor's Vue `components.ts`
// (pptx-editor/packages/vue/src/components.ts at d4502552), which follows DOCX Editor's Vue
// adapter.
//
// `XlsxEditorRoot` renders no DOM of its own: it creates the editor on mount, provides it, and
// lets `XlsxEditorContent` attach the grid surface wherever the host's template puts it. Props
// are declared once as a const; the public props type derives from it, so the two cannot drift
// (the parity check reads the runtime props). The rules that do not depend on a framework (the
// font failure report, the wait of `ready` for the first open, and the `mode` prop that waits
// for an open cell edit) come from core, so both adapters act the same.
//
// Differences from React: the callbacks are emits (`ready`, `change`, `fontError`), and
// `children` is the default slot. `ready` fires after the first open settles, as React's
// `onReady`; the root reads whether a `fontError` listener exists at each failure, as PPTX
// Editor's root does, so a failure with no listener prints one warning.

import {
  defineComponent,
  getCurrentInstance,
  h,
  inject,
  onBeforeUnmount,
  onMounted,
  onUnmounted,
  provide,
  shallowRef,
  watch,
  type ExtractPublicPropTypes,
  type PropType,
  type StyleValue,
} from 'vue';
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
  type ColorMode,
  type EditorFailure,
  type EditorMode,
  type EditorSnapshot,
  type FontSource,
  type ModeProp,
  type NumberLocale,
  type SavedWorkbook,
  type XlsxEditorInstance,
} from '@xlsx-editor.dev/core/editor';
import type { TFunction, Translations } from '@xlsx-editor.dev/i18n';
import {
  hasListener,
  layoutPartsKey,
  modeWaitsKey,
  provideXlsxEditor,
  rootModeKey,
  translatorFor,
  useLayoutPart,
  useTranslation,
  useXlsxEditor,
  xlsxTranslationKey,
  type LayoutParts,
} from './context.ts';
import { provideChromeDark, provideScopedByAncestor, useChromeDark } from './color-mode.ts';
import { isDevelopment } from './is-development.ts';

/** The faces that measure and paint text: bytes, or a function that loads them once. @public */
export type XlsxEditorFonts = readonly FontSource[] | (() => Promise<readonly FontSource[]>);

const xlsxEditorRootProps = {
  /**
   * The bytes of a workbook or a template package. A new array or buffer opens the new
   * workbook; a change to `undefined` closes the open workbook (`close`).
   */
  document: {
    type: [Uint8Array, ArrayBuffer] as PropType<Uint8Array | ArrayBuffer>,
    default: undefined,
  },
  /**
   * The faces that measure and paint text. Omitted, the packaged substitute faces load. A new
   * value replaces the faces in place (`setFonts`); a list is compared by its faces.
   */
  fonts: { type: [Array, Function] as PropType<XlsxEditorFonts>, default: undefined },
  /**
   * The number locale of the shown values: separators and built-in format codes. A value that
   * differs (`sameNumberLocale`) applies in place (`setNumberLocale`).
   */
  numberLocale: { type: Object as PropType<NumberLocale>, default: undefined },
  /** `'edit'` (default) or `'view'`. A change applies without a reload. */
  mode: { type: String as PropType<EditorMode>, default: undefined },
  /** UI strings; English when omitted. A `null` entry falls back to English. */
  locale: { type: Object as PropType<Translations>, default: undefined },
  /** The chrome color theme: `'light'` (default), `'dark'`, or `'system'`. */
  colorMode: { type: String as PropType<ColorMode>, default: undefined },
} as const;

/** Props of `XlsxEditorRoot`, the twin of React's `XlsxEditorRootProps`. @public */
export type XlsxEditorRootProps = ExtractPublicPropTypes<typeof xlsxEditorRootProps>;

/** The emits of `XlsxEditorRoot` and `XlsxEditor`. */
export const xlsxEditorRootEmits = {
  /** Once per editor, after the first open of `document` settles. */
  ready: (_editor: XlsxEditorInstance) => true,
  /**
   * After every change of the editor state, with the new snapshot: a content change and a view
   * change alike. The `change` event of `useEditorEvent` fires for content changes only.
   */
  change: (_snapshot: EditorSnapshot) => true,
  /** When the fonts that measure and paint text do not load. */
  fontError: (_failure: EditorFailure) => true,
};

/**
 * The imperative handle of `XlsxEditor` and `XlsxEditorRoot`, the twin of React's
 * `XlsxEditorRef`: a template ref on either component reads these members (`expose`). Every
 * member is safe to call before the editor mounts. The editing commands are on the instance
 * that `getEditor` returns.
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
   * template), with every edit, or `null` when none is open. An open cell edit is committed first;
   * when the editor cannot store it, `save` returns `null` and shows the reason.
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

/**
 * Owns the editor instance, as React's `XlsxEditor.Root`. Only a new `document` reloads the
 * workbook; the other props apply in place. Emits `ready` once per editor, `change` after every
 * change of the editor state, and `fontError` for each font failure. @public
 */
export const XlsxEditorRoot = defineComponent({
  name: 'XlsxEditorRoot',
  props: xlsxEditorRootProps,
  emits: xlsxEditorRootEmits,
  setup(props, { emit, expose, slots }) {
    const self = getCurrentInstance();
    // Read at each failure: a parent render can add or remove the listener.
    const reportTo = (): ((failure: EditorFailure) => void) | undefined =>
      hasListener(self?.vnode.props ?? null, 'fontError') ? (f) => emit('fontError', f) : undefined;
    const mode = (): EditorMode => props.mode ?? 'edit';
    const editor = shallowRef<XlsxEditorInstance | null>(null);
    provideXlsxEditor(editor);
    provide(rootModeKey, mode);
    const modeWaits = shallowRef(false);
    provide(modeWaitsKey, modeWaits);
    const parts: LayoutParts = { grids: 0, statusMessages: 0 };
    provide(layoutPartsKey, parts);
    provideChromeDark(() => props.colorMode);

    // The catalog's content, not its identity: an inline locale object makes a new value on
    // each parent render and must not relabel the grid.
    const t = shallowRef<TFunction>(translatorFor(props.locale));
    provide(xlsxTranslationKey, t);
    watch(
      () => (props.locale ? JSON.stringify(props.locale) : ''),
      () => {
        t.value = translatorFor(props.locale);
        // The grid label, the accessible stale words, and the cell editor label.
        if (editor.value) relabelEditor(editor.value);
      }
    );

    // The ref members read the live instance at each call, as React's `useImperativeHandle`.
    expose({
      load: async (bytes) => (editor.value ? editor.value.load(bytes) : null),
      save: () => editor.value?.save() ?? null,
      saveInTasks: async () => (editor.value ? editor.value.saveInTasks() : null),
      getEditor: () => editor.value,
      focus: () => editor.value?.focus(),
      snapshot: () => editor.value?.snapshot() ?? emptySnapshotOf(mode()),
    } satisfies XlsxEditorRef);

    let created: { fonts: XlsxEditorFonts | undefined; numberLocale: NumberLocale | undefined } = {
      fonts: undefined,
      numberLocale: undefined,
    };
    let opened = false;
    let modeProp: ModeProp | null = null;
    const stops: (() => void)[] = [];
    onMounted(() => {
      const fonts = props.fonts;
      const numberLocale = props.numberLocale;
      const instance = createXlsxEditor({
        // Omitted, the engine loads the packaged substitute faces and keeps their failures.
        ...(fonts ? { fonts } : {}),
        mode: mode(),
        ...(numberLocale ? { numberLocale } : {}),
        // The engine's keys are keys of the shared catalog.
        translate: (key, params) =>
          t.value(key as Parameters<TFunction>[0], params ? { ...params } : undefined),
      });
      created = { fonts, numberLocale };
      stops.push(watchFontErrors(instance, (f) => reportFontError(f, reportTo())));
      stops.push(instance.subscribe(() => emit('change', instance.snapshot())));
      // `XlsxEditorContent` attaches at once, in its own watcher of the editor.
      editor.value = instance;
      modeProp = followModeProp(instance, (waits) => (modeWaits.value = waits));
      modeProp.apply(mode());
      if (props.document) void instance.load(props.document);
      opened = props.document !== undefined;
      // After the attach and the start of the first open: `ready` waits for that open to settle.
      stops.push(whenOpenSettles(instance, () => emit('ready', instance)));
      // In development, a layout that shows the grid without `StatusMessage` gets one warning:
      // refusals and the stale notice show only there. The children mounted first.
      if (parts.grids > 0 && parts.statusMessages === 0 && isDevelopment())
        console.warn(
          'XlsxEditorRoot: the layout has no StatusMessage, so edit refusals and the ' +
            'stale-value notice are not shown. Add <StatusMessage /> to the layout.'
        );
    });
    // After the children unmounted, as DOCX Editor's root: Content detaches first.
    onUnmounted(() => {
      for (const stop of stops.splice(0)) stop();
      modeProp?.dispose();
      modeProp = null;
      const instance = editor.value;
      editor.value = null;
      instance?.destroy();
    });

    watch(
      () => props.document,
      (bytes) => {
        const instance = editor.value;
        if (!instance) return;
        if (bytes) void instance.load(bytes);
        else if (opened) instance.close();
        opened = bytes !== undefined;
      }
    );
    // The `mode` prop applies when it changes, as DOCX Editor's root applies it: an instance
    // call to `setMode` stays until the next prop change. A change that the open cell edit
    // refuses waits, with a notice, and applies when the edit closes (`followModeProp`).
    watch(mode, (next) => modeProp?.apply(next));
    // A new `fonts` or `numberLocale` value replaces the current one in place, as DOCX
    // Editor's root does; the first values were passed at creation.
    watch(
      () => props.fonts,
      (fonts) => {
        const instance = editor.value;
        if (!instance || sameFontSource(fonts, created.fonts)) return;
        created = { ...created, fonts };
        void instance.setFonts(fonts).then((failure) => {
          if (failure) reportFontError(failure, reportTo());
        });
      }
    );
    watch(
      () => props.numberLocale,
      (numberLocale) => {
        const instance = editor.value;
        if (!instance || sameNumberLocale(numberLocale, created.numberLocale)) return;
        created = { ...created, numberLocale };
        instance.setNumberLocale(numberLocale);
      }
    );
    return () => slots.default?.();
  },
});

const xlsxEditorViewportProps = {
  /**
   * Appended after the `xlsx-editor` and `xlsx-editor-viewport` classes. A host `class` also
   * works.
   */
  className: { type: String, default: undefined },
  /**
   * The name of the editor's landmark region. Default: "Workbook editor". `false` renders no
   * region, for a host whose own landmarks are inside.
   */
  label: { type: [String, Boolean] as PropType<string | false>, default: undefined },
  /** Inline style, for example the editor's height. A host `style` also works. */
  style: { type: [String, Object, Array] as PropType<StyleValue>, default: undefined },
} as const;

/** The props of `XlsxEditorViewport`, the twin of React's `XlsxEditorViewportProps`. @public */
export type XlsxEditorViewportProps = ExtractPublicPropTypes<typeof xlsxEditorViewportProps>;

/**
 * The editor's styling scope and landmark region: carries the `.xlsx-editor` class, and `dark`
 * with it when the root's `colorMode` resolves to dark. It also carries `xlsx-editor-viewport`,
 * which keeps its width from following its content. F6 and Shift+F6 inside it move the
 * keyboard among the toolbar, the formula bar, the grid, and the footer. @public
 */
export const XlsxEditorViewport = defineComponent({
  name: 'XlsxEditorViewport',
  props: xlsxEditorViewportProps,
  setup(props, { slots }) {
    const { t } = useTranslation();
    const editor = useXlsxEditor();
    const dark = useChromeDark();
    provideScopedByAncestor();
    return () =>
      h(
        'div',
        {
          class: ['xlsx-editor', dark.value ? 'dark' : '', 'xlsx-editor-viewport', props.className]
            .filter(Boolean)
            .join(' '),
          style: props.style,
          ...(props.label === false
            ? {}
            : { role: 'region', 'aria-label': props.label ?? t('viewport.ariaLabel') }),
          // F6 and Shift+F6 move the keyboard among the toolbar, the formula bar, the grid, and
          // the footer.
          onKeydown: (e: KeyboardEvent) =>
            void cycleRegion(e, e.currentTarget as HTMLElement, editor.value),
        },
        slots.default?.()
      );
  },
});

const xlsxEditorContentProps = {
  /** Appended after the `xlsx-canvas` class. A host `class` also works. */
  className: { type: String, default: undefined },
} as const;

/** The props of `XlsxEditorContent`, the twin of React's `XlsxEditorContentProps`. @public */
export type XlsxEditorContentProps = ExtractPublicPropTypes<typeof xlsxEditorContentProps>;

/** The grid: attaches the editor's surface to its element. @public */
export const XlsxEditorContent = defineComponent({
  name: 'XlsxEditorContent',
  props: xlsxEditorContentProps,
  setup(props) {
    const editor = useXlsxEditor();
    // Inside a root only: the root's layout check counts the grid.
    if (inject(layoutPartsKey, null)) useLayoutPart('grids');
    const host = shallowRef<HTMLDivElement | null>(null);
    let attached: XlsxEditorInstance | null = null;
    let stopSetup: (() => void) | null = null;
    const attach = (next: XlsxEditorInstance | null): void => {
      if (attached === next) return;
      attached?.detach();
      stopSetup?.();
      stopSetup = null;
      attached = null;
      const el = host.value;
      if (!next || !el) return;
      next.attach(el);
      attached = next;
      // In development, name a missing stylesheet, which leaves the grid with no scroll
      // container. The check waits for the document load.
      if (isDevelopment())
        stopSetup = watchGridSetup(el, 'XlsxEditorContent', '@xlsx-editor.dev/vue/styles.css');
    };
    onMounted(() => attach(editor.value));
    // At once when the root publishes the editor, so `ready` comes after the attach.
    watch(editor, (next) => attach(next), { flush: 'sync' });
    onBeforeUnmount(() => attach(null));
    return () =>
      h('div', {
        ref: host,
        class: `xlsx-canvas${props.className ? ` ${props.className}` : ''}`,
      });
  },
});
