// The composed editor: the menu bar, the toolbar, the formula bar, the grid, and a footer with
// the sheet tabs and the zoom. The Vue twin of the React adapter's `XlsxEditor.tsx`, after PPTX Editor's Vue
// sugar (pptx-editor/packages/vue/src/index.ts at d4502552).
//
// Differences from React: the callbacks are emits (`save` and `saveStart` too, as DOCX Editor's
// Vue sugar takes its callbacks), and `className` and `style` are the host's fallthrough `class`
// and `style`, which land on the viewport (`inheritAttrs: false`). The save control is enabled
// only while a `save` listener exists, as React's is only with `onSave`.

import {
  defineComponent,
  getCurrentInstance,
  h,
  mergeProps,
  shallowRef,
  type ExtractPublicPropTypes,
  type PropType,
} from 'vue';
import {
  emptySnapshotOf,
  type ColorMode,
  type EditorFailure,
  type EditorMode,
  type EditorSnapshot,
  type NumberLocale,
  type SavedWorkbook,
  type XlsxEditorInstance,
} from '@xlsx-editor.dev/core/editor';
import {
  XlsxEditorContent,
  XlsxEditorRoot,
  XlsxEditorViewport,
  type XlsxEditorFonts,
  type XlsxEditorRef,
} from './XlsxEditorRoot.ts';
import { StatusMessage, ZoomControl } from './chrome.tsx';
import { FormulaBar } from './FormulaBar.tsx';
import { SheetTabs } from './SheetTabs.tsx';
import { hasListener, useEditorState } from './context.ts';
import { XlsxEditorToolbar } from './toolbar/Toolbar.tsx';
import { XlsxEditorMenu, type XlsxEditorMenuProps } from './menu/Menu.tsx';
import { XlsxEditorContextMenu, type XlsxEditorContextMenuProps } from './menu/ContextMenu.tsx';
import type { Translations } from '@xlsx-editor.dev/i18n';
import { useChromeDir } from './locale.ts';

const xlsxEditorProps = {
  /** The bytes of a workbook or a template package, as on `XlsxEditorRoot`. */
  document: {
    type: [Uint8Array, ArrayBuffer] as PropType<Uint8Array | ArrayBuffer>,
    default: undefined,
  },
  /** The faces that measure and paint text, as on `XlsxEditorRoot`. */
  fonts: { type: [Array, Function] as PropType<XlsxEditorFonts>, default: undefined },
  /** The number locale of the shown values, as on `XlsxEditorRoot`. */
  numberLocale: { type: Object as PropType<NumberLocale>, default: undefined },
  /** `'edit'` (default) or `'view'`, as on `XlsxEditorRoot`. */
  mode: { type: String as PropType<EditorMode>, default: undefined },
  /** UI strings, as on `XlsxEditorRoot`. English when omitted. */
  locale: { type: Object as PropType<Translations>, default: undefined },
  /** The chrome color theme, as on `XlsxEditorRoot`. */
  colorMode: { type: String as PropType<ColorMode>, default: undefined },
  /**
   * Whether your page opens another workbook now, while the editor still shows the old one. While
   * it is true, the save control is disabled and Control or Command and S starts no save. See the
   * `opening` prop of `XlsxEditorToolbar`.
   */
  opening: { type: Boolean, default: undefined },
  /**
   * The menu bar above the toolbar: `true` (the default) shows it, `false` leaves it out, and an
   * object shows it with those props, which win over the ones that the editor passes (`onSave`,
   * `onSaveStart`, and `opening`).
   *
   * The default File > Open replaces the open workbook with the picked file and does not ask
   * first, also when the workbook has unsaved edits. To ask first, pass your own `onOpen` in the
   * object; `snapshot().revision` tells you whether the workbook changed since your last save.
   * To remove the row, set `menu` to `false` and compose the menu bar with a hidden `Open` row.
   */
  menu: {
    type: [Boolean, Object] as PropType<boolean | XlsxEditorMenuProps>,
    default: undefined,
  },
  /**
   * The grid context menu: `true` (the default) mounts it, `false` leaves it out, so a secondary
   * press on the grid shows the browser's own menu, and an object mounts it with those props.
   */
  contextMenu: {
    type: [Boolean, Object] as PropType<boolean | XlsxEditorContextMenuProps>,
    default: undefined,
  },
} as const;

/**
 * Props of the `XlsxEditor` component, the twin of React's `XlsxEditorProps`. Derived from the
 * component's runtime props, so the two cannot drift. The host's `class` and `style` land on the
 * viewport. @public
 */
export type XlsxEditorProps = ExtractPublicPropTypes<typeof xlsxEditorProps>;

const statusOf = (snapshot: EditorSnapshot) => snapshot.status;

/** The chrome and the grid, inside the viewport. */
const Body = defineComponent({
  name: 'XlsxEditorBody',
  props: {
    onSave: {
      type: Function as PropType<(bytes: Uint8Array<ArrayBuffer>, saved: SavedWorkbook) => void>,
      default: undefined,
    },
    onSaveStart: {
      type: Function as PropType<(saving: Promise<SavedWorkbook | null>) => void>,
      default: undefined,
    },
    opening: { type: Boolean, default: undefined },
    menu: {
      type: [Boolean, Object] as PropType<boolean | XlsxEditorMenuProps>,
      default: undefined,
    },
    contextMenu: {
      type: [Boolean, Object] as PropType<boolean | XlsxEditorContextMenuProps>,
      default: undefined,
    },
  },
  setup(props) {
    // The status alone: the body renders again only when it changes, not on each key of an edit.
    const status = useEditorState(statusOf);
    const dir = useChromeDir();
    return () => [
      props.menu !== false
        ? h(XlsxEditorMenu, {
            ...(props.onSave ? { onSave: props.onSave } : {}),
            ...(props.onSaveStart ? { onSaveStart: props.onSaveStart } : {}),
            ...(props.opening === undefined ? {} : { opening: props.opening }),
            ...(typeof props.menu === 'object' ? props.menu : {}),
          })
        : null,
      h(XlsxEditorToolbar, {
        ...(props.onSave ? { onSave: props.onSave } : {}),
        ...(props.onSaveStart ? { onSaveStart: props.onSaveStart } : {}),
        ...(props.opening === undefined ? {} : { opening: props.opening }),
      }),
      h(FormulaBar),
      h(StatusMessage),
      h(XlsxEditorContent, status.value === 'ready' ? {} : { className: 'xlsx-canvas-idle' }),
      h('div', { class: 'xlsx-footer', dir: dir.value }, [h(SheetTabs), h(ZoomControl)]),
      props.contextMenu !== false
        ? h(XlsxEditorContextMenu, typeof props.contextMenu === 'object' ? props.contextMenu : {})
        : null,
    ];
  },
});

/**
 * A workbook editor with its chrome, the twin of React's `XlsxEditor`. A template ref reads an
 * `XlsxEditorRef`. Emits `ready`, `change`, and `fontError` as `XlsxEditorRoot` does; `save`
 * with the bytes when the toolbar's save control or Control or Command and S is pressed; and
 * `saveStart` with the promise of the result when such a press starts a delivery to `save`.
 * Without a `save` listener, the save control is disabled with the engine's reason. @public
 */
export const XlsxEditor = defineComponent({
  name: 'XlsxEditor',
  // The root renders no element, so the host's attributes (`class`, `style`, `id`, `data-*`)
  // go onto the viewport, as React's `className` and `style` do.
  inheritAttrs: false,
  props: xlsxEditorProps,
  emits: {
    ready: (_editor: XlsxEditorInstance) => true,
    change: (_snapshot: EditorSnapshot) => true,
    fontError: (_failure: EditorFailure) => true,
    save: (_bytes: Uint8Array<ArrayBuffer>, _saved: SavedWorkbook) => true,
    saveStart: (_saving: Promise<SavedWorkbook | null>) => true,
  },
  setup(props, { attrs, emit, expose }) {
    const self = getCurrentInstance();
    const root = shallowRef<XlsxEditorRef | null>(null);
    // The same members as the root's handle, read through the root at each call.
    expose({
      load: async (bytes) => (root.value ? root.value.load(bytes) : null),
      save: () => root.value?.save() ?? null,
      saveInTasks: async () => (root.value ? root.value.saveInTasks() : null),
      getEditor: () => root.value?.getEditor() ?? null,
      focus: () => root.value?.focus(),
      snapshot: () => root.value?.snapshot() ?? emptySnapshotOf(props.mode ?? 'edit'),
    } satisfies XlsxEditorRef);
    const onSave = (bytes: Uint8Array<ArrayBuffer>, saved: SavedWorkbook): void =>
      emit('save', bytes, saved);
    const onSaveStart = (saving: Promise<SavedWorkbook | null>): void => emit('saveStart', saving);
    return () => {
      // Read at render, not once: a parent render can add or remove a listener.
      const listeners = self?.vnode.props ?? null;
      return h(
        XlsxEditorRoot,
        {
          ref: root,
          ...(props.document ? { document: props.document } : {}),
          ...(props.fonts ? { fonts: props.fonts } : {}),
          ...(props.numberLocale ? { numberLocale: props.numberLocale } : {}),
          ...(props.mode ? { mode: props.mode } : {}),
          ...(props.locale ? { locale: props.locale } : {}),
          ...(props.colorMode ? { colorMode: props.colorMode } : {}),
          // Forwarded only with a listener, so the root still warns when nobody listens.
          ...(hasListener(listeners, 'fontError')
            ? { onFontError: (failure: EditorFailure) => emit('fontError', failure) }
            : {}),
          onReady: (editor: XlsxEditorInstance) => emit('ready', editor),
          onChange: (snapshot: EditorSnapshot) => emit('change', snapshot),
        },
        {
          default: () =>
            h(XlsxEditorViewport, mergeProps(attrs), () =>
              h(Body, {
                // A `.once` save listener does not enable Save: after it runs, a press would save
                // for no listener. `XlsxEditorToolbar` takes `save` as a prop, which has no
                // `.once` either.
                ...(hasListener(listeners, 'save', false) ? { onSave } : {}),
                ...(hasListener(listeners, 'saveStart') ? { onSaveStart } : {}),
                ...(props.opening === undefined ? {} : { opening: props.opening }),
                ...(props.menu === undefined ? {} : { menu: props.menu }),
                ...(props.contextMenu === undefined ? {} : { contextMenu: props.contextMenu }),
              })
            ),
        }
      );
    };
  },
});
