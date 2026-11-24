// `SheetPreview`, the twin of the React adapter's `sheet-preview.test.tsx`: paints a canvas once
// the workbook is ready, paints the edit after an edit, names itself with `label`, and is hidden
// from assistive technology without one.

import { afterAll, afterEach, beforeAll, expect, test } from 'bun:test';
import { h, nextTick } from 'vue';
import { waitFor } from '@testing-library/react';
import { buildXlsx, c, row } from '../../../scripts/fixtures/builder.ts';
import {
  SheetPreview,
  SheetTabs,
  XlsxEditorRoot,
  XlsxEditorViewport,
  type XlsxEditorInstance,
} from '../src/index.ts';
import { FONTS, mount, unmountAll } from './mount.ts';

// happy-dom lays nothing out: give the preview a size of 200 by 100.
const proto = HTMLElement.prototype;
const sizes = ['clientWidth', 'clientHeight'] as const;
const saved = sizes.map((k) => Object.getOwnPropertyDescriptor(Element.prototype, k));
beforeAll(() => {
  sizes.forEach((k, i) =>
    Object.defineProperty(proto, k, {
      configurable: true,
      get(this: HTMLElement) {
        if (this.classList.contains('xlsx-sheet-preview')) return i === 0 ? 200 : 100;
        return saved[i]?.get?.call(this) ?? 0;
      },
    })
  );
});
afterAll(() => sizes.forEach((k) => delete (proto as unknown as Record<string, unknown>)[k]));
afterEach(unmountAll);

const workbook = buildXlsx({ sheets: [{ name: 'S', rows: row(1, [c('A1', 1)]) }] });

test('paints after load and after an edit, with the label as its name', async () => {
  let editor: XlsxEditorInstance | null = null;
  const view = mount(() =>
    h(
      XlsxEditorRoot,
      { document: workbook, fonts: FONTS, onReady: (e: XlsxEditorInstance) => (editor = e) },
      {
        default: () => [
          h(SheetPreview, { label: 'Sheet S', className: 'card' }),
          h(XlsxEditorViewport, null, () => [h(SheetTabs)]),
        ],
      }
    )
  );
  const preview = view.getByRole('img', { name: 'Sheet S' });
  expect(preview.className).toBe('xlsx-sheet-preview card');
  await waitFor(() => expect(preview.querySelector('canvas')).not.toBeNull(), { timeout: 5000 });
  const calls: unknown[] = [];
  const paint = editor!.renderPreview.bind(editor!);
  editor!.renderPreview = (el, options) => (calls.push(options), paint(el, options));
  editor!.selectReference('B2');
  editor!.enterValue('7');
  await nextTick();
  await waitFor(() => expect(calls.length).toBeGreaterThan(0));
});

test('is hidden without a label, and empty before a workbook', () => {
  const view = mount(() =>
    h(XlsxEditorRoot, { fonts: FONTS }, { default: () => [h(SheetPreview, { sheet: 0, zoom: 1 })] })
  );
  const preview = view.container.querySelector('.xlsx-sheet-preview')!;
  expect(preview.getAttribute('aria-hidden')).toBe('true');
  expect(preview.childElementCount).toBe(0);
});
