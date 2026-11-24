/** @jsxImportSource react */
// `XlsxEditor.SheetPreview`: a preview of a sheet for a host's own surface, such as the card of a
// workbook in a chat or a file list. It paints with `renderPreview` of the editor instance, and
// paints again when the content, the open workbook, the shown sheet, or its own size changes,
// so the host writes no effect. Place it anywhere inside `XlsxEditor.Root`, also outside
// `XlsxEditor.Viewport`. It is a `span`, so a button or a link can hold it.

import { useLayoutEffect, useRef, useState } from 'react';
import type { EditorSnapshot } from '@xlsx-editor.dev/core/editor';
import { useEditorState, useXlsxEditor } from './context.ts';

/** Props for `XlsxEditor.SheetPreview`. @public */
export interface SheetPreviewProps {
  /** The position of the sheet in `snapshot().sheets`. The shown sheet when absent. */
  readonly sheet?: number;
  /** The zoom, from 0.1 to 4, where 1 is 100%. The default is 0.6. */
  readonly zoom?: number;
  /**
   * The accessible name of the preview, which then has the `img` role. Without it, the preview is
   * hidden from assistive technology, for a preview inside a control that has its own name.
   */
  readonly label?: string;
  /** Appended after the `xlsx-sheet-preview` class. Give the element its size here. */
  readonly className?: string;
}

/** What the preview shows: it paints again when this text changes. */
const contentKey = (s: EditorSnapshot): string =>
  `${s.status}:${s.revision}:${s.activeSheet}:${s.sheets.length}`;

/**
 * A preview of the top-left window of a sheet, painted as the grid paints it, with no headings
 * and no selection. Size it with `className`: it fills its own box.
 *
 * @example
 * ```tsx
 * <XlsxEditor.Root document={bytes}>
 *   <XlsxEditor.SheetPreview sheet={0} className="card-preview" label="Sales" />
 * </XlsxEditor.Root>
 * ```
 *
 * @public
 */
export function SheetPreview({ sheet, zoom, label, className }: SheetPreviewProps) {
  const editor = useXlsxEditor();
  const key = useEditorState(contentKey);
  const box = useRef<HTMLSpanElement>(null);
  const [size, setSize] = useState('');
  useLayoutEffect(() => {
    const el = box.current;
    const Observer = el?.ownerDocument.defaultView?.ResizeObserver;
    if (!el || !Observer) return undefined;
    const observer = new Observer(([entry]) => {
      const rect = entry?.contentRect;
      if (rect) setSize(`${Math.round(rect.width)}x${Math.round(rect.height)}`);
    });
    observer.observe(el);
    return () => observer.disconnect();
  }, []);
  useLayoutEffect(() => {
    const el = box.current;
    if (!editor || !el) return;
    if (
      !editor.renderPreview(el, {
        ...(sheet === undefined ? {} : { sheet }),
        ...(zoom === undefined ? {} : { zoom }),
      })
    )
      el.replaceChildren();
  }, [editor, key, size, sheet, zoom]);
  return (
    <span
      ref={box}
      className={`xlsx-sheet-preview${className ? ` ${className}` : ''}`}
      {...(label ? { role: 'img', 'aria-label': label } : { 'aria-hidden': true })}
    />
  );
}
