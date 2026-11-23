/** @jsxImportSource react */
// The workbook, the narrow editor, and the layout stubs of the toolbar tests: happy-dom has no
// layout, so a test gives the boxes that the toolbar reads.

import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { expect } from 'bun:test';
import { render, waitFor } from '@testing-library/react';
import { buildXlsx, c, row } from '../../../../scripts/fixtures/builder.ts';
import { XlsxEditor, type FontSource, type XlsxEditorInstance } from '../index.ts';
import type { Translations } from '@xlsx-editor.dev/i18n';

const fontsDir = join(
  dirname(
    Bun.resolveSync('@docx-editor.dev/fonts/package.json', join(import.meta.dir, '../../../core'))
  ),
  'assets'
);
export const FONTS: FontSource[] = [
  {
    family: 'Carlito',
    bold: false,
    italic: false,
    bytes: new Uint8Array(readFileSync(join(fontsDir, 'Carlito-Regular.ttf'))),
    origin: 'substitute',
  },
];

export const workbook = (afterMerges = '') =>
  buildXlsx({
    sheets: [{ name: 'S', rows: row(1, [c('A1', 1), c('B1', 'two')]), afterMerges }],
    styles: { dxfs: ['<dxf/>'] },
  });

/** A toolbar layout for happy-dom, which has none: rectangles by element, the rest unchanged. */
export type Rect = Partial<DOMRect>;
export function stubLayout(rectOf: (el: HTMLElement) => Rect | null, barWidth = 0): () => void {
  const proto = HTMLElement.prototype as unknown as Record<string, unknown>;
  const rect = Element.prototype.getBoundingClientRect;
  const width = Object.getOwnPropertyDescriptor(HTMLElement.prototype, 'clientWidth')!;
  Element.prototype.getBoundingClientRect = function (this: HTMLElement) {
    const r = rectOf(this);
    return (
      r
        ? { x: 0, y: 0, top: 0, left: 0, right: 0, bottom: 0, width: 0, height: 0, ...r }
        : rect.call(this)
    ) as DOMRect;
  };
  Object.defineProperty(proto, 'clientWidth', {
    configurable: true,
    get(this: HTMLElement) {
      return this.classList.contains('xlsx-toolbar') ? barWidth : width.get!.call(this);
    },
  });
  return () => {
    Element.prototype.getBoundingClientRect = rect;
    Object.defineProperty(proto, 'clientWidth', width);
  };
}

/** Give `scrollHeight` from `heightOf` where it returns a number; gives the restore. */
export function stubScrollHeight(heightOf: (el: HTMLElement) => number | null): () => void {
  const proto = HTMLElement.prototype;
  const own = Object.getOwnPropertyDescriptor(proto, 'scrollHeight');
  const base = own ?? Object.getOwnPropertyDescriptor(Element.prototype, 'scrollHeight')!;
  Object.defineProperty(proto, 'scrollHeight', {
    configurable: true,
    get(this: HTMLElement) {
      return heightOf(this) ?? base.get!.call(this);
    },
  });
  return () => {
    if (own) Object.defineProperty(proto, 'scrollHeight', own);
    else delete (proto as unknown as Record<string, unknown>).scrollHeight;
  };
}

/**
 * The box of a host container that clips its overflow: `clip` pixels from the top of the window,
 * or from `clip[0]` to `clip[1]`.
 */
export type Clip = number | readonly [top: number, bottom: number];
export const clipBox = (clip: Clip): Rect =>
  typeof clip === 'number'
    ? { top: 0, bottom: clip, height: clip }
    : { top: clip[0], bottom: clip[1], height: clip[1] - clip[0] };

/** A group is ten pixels for each letter of its name. */
export const groupWidth = (el: HTMLElement): Rect | null =>
  el.dataset.group ? { width: (el.getAttribute('aria-label') ?? '').length * 10 } : null;

/** Open a workbook in a toolbar of `barWidth` pixels, with `locale`. */
export async function openNarrow(barWidth: number, locale?: Translations) {
  let viewer: XlsxEditorInstance | null = null;
  const ui = (l?: Translations) => (
    <div style={{ height: 400 }}>
      <XlsxEditor
        document={workbook()}
        fonts={FONTS}
        onReady={(v) => (viewer = v)}
        {...(l ? { locale: l } : {})}
      />
    </div>
  );
  const view = render(ui(locale));
  await waitFor(() => expect(view.getAllByRole('tab').length).toBe(1), { timeout: 5000 });
  return { view, viewer: viewer!, rerender: (l?: Translations) => view.rerender(ui(l)) };
}
