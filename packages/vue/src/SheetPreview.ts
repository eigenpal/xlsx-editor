// `SheetPreview`: a preview of a sheet for a host's own surface, such as the card of a workbook in
// a chat. The twin of the React adapter's `SheetPreview.tsx`, with the same props and markup. It
// is a `span`, so a button or a link can hold it. It paints with `renderPreview` of the editor
// instance, and paints again when the content, the open workbook, the shown sheet, or its own
// size changes.

import {
  defineComponent,
  h,
  onBeforeUnmount,
  onMounted,
  ref,
  watch,
  type ExtractPublicPropTypes,
} from 'vue';
import type { EditorSnapshot } from '@xlsx-editor.dev/core/editor';
import { useEditorState, useXlsxEditor } from './context.ts';

const sheetPreviewProps = {
  /** The position of the sheet in `snapshot().sheets`. The shown sheet when absent. */
  sheet: { type: Number, default: undefined },
  /** The zoom, from 0.1 to 4, where 1 is 100%. The default is 0.6. */
  zoom: { type: Number, default: undefined },
  /**
   * The accessible name of the preview, which then has the `img` role. Without it, the preview is
   * hidden from assistive technology, for a preview inside a control that has its own name.
   */
  label: { type: String, default: undefined },
  /** Appended after the `xlsx-sheet-preview` class. Give the element its size here. */
  className: { type: String, default: undefined },
} as const;

/** Props for `SheetPreview`, the twin of React's `SheetPreviewProps`. @public */
export type SheetPreviewProps = ExtractPublicPropTypes<typeof sheetPreviewProps>;

/** What the preview shows: it paints again when this text changes. */
const contentKey = (s: EditorSnapshot): string =>
  `${s.status}:${s.revision}:${s.activeSheet}:${s.sheets.length}`;

/**
 * A preview of the top-left window of a sheet, painted as the grid paints it, with no headings
 * and no selection. Size it with `className`: it fills its own box. Place it inside
 * `XlsxEditorRoot`.
 *
 * @public
 */
export const SheetPreview = defineComponent({
  name: 'SheetPreview',
  props: sheetPreviewProps,
  setup(props) {
    const editor = useXlsxEditor();
    const key = useEditorState(contentKey);
    const box = ref<HTMLSpanElement | null>(null);
    const size = ref('');
    let observer: ResizeObserver | null = null;
    const paint = (): void => {
      const el = box.value;
      const instance = editor.value;
      if (!el || !instance) return;
      const options = {
        ...(props.sheet === undefined ? {} : { sheet: props.sheet }),
        ...(props.zoom === undefined ? {} : { zoom: props.zoom }),
      };
      if (!instance.renderPreview(el, options)) el.replaceChildren();
    };
    onMounted(() => {
      const el = box.value;
      const Observer = el?.ownerDocument.defaultView?.ResizeObserver;
      if (el && Observer) {
        observer = new Observer(([entry]) => {
          const rect = entry?.contentRect;
          if (rect) size.value = `${Math.round(rect.width)}x${Math.round(rect.height)}`;
        });
        observer.observe(el);
      }
      paint();
    });
    onBeforeUnmount(() => observer?.disconnect());
    watch([editor, key, size, () => props.sheet, () => props.zoom], paint, { flush: 'post' });
    return () =>
      h('span', {
        ref: box,
        class: ['xlsx-sheet-preview', props.className],
        ...(props.label ? { role: 'img', 'aria-label': props.label } : { 'aria-hidden': 'true' }),
      });
  },
});
