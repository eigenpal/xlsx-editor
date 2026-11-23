// Shared helpers for the Vue adapter tests: one workbook, the substitute face, and a mount that
// gives the Testing Library queries of its host, as PPTX Editor's Vue tests mount with
// `createApp` and read the DOM with plain queries.
//
// A Vue render runs in the next microtask, not inside the event handler, so a test awaits
// `nextTick()` after an event or an editor call before it reads the DOM. The editor's state is
// synchronous and needs no wait.

import { expect } from 'bun:test';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fireEvent, queries, waitFor, within, type BoundFunctions } from '@testing-library/react';
import { createApp, type App, type VNodeChild } from 'vue';
import { buildXlsx, c, row } from '../../../scripts/fixtures/builder.ts';
import type { FontSource } from '../src/index.ts';

const fontsDir = join(
  dirname(
    Bun.resolveSync('@docx-editor.dev/fonts/package.json', join(import.meta.dir, '../../core'))
  ),
  'assets'
);

/** The substitute regular face, so text measures with no font download. */
export const FONTS: FontSource[] = [
  {
    family: 'Carlito',
    bold: false,
    italic: false,
    bytes: new Uint8Array(readFileSync(join(fontsDir, 'Carlito-Regular.ttf'))),
    origin: 'substitute',
  },
];

/** Three sheets: First (A1 text, A5 a formula), a hidden sheet, and Second. */
export const workbook = buildXlsx({
  sheets: [
    { name: 'First', rows: row(1, [c('A1', 'one')]) + row(5, [c('A5', 84, { f: 'A2*2' })]) },
    { name: 'Hidden', state: 'hidden', rows: row(1, [c('A1', 'secret')]) },
    { name: 'Second', rows: row(1, [c('A1', 'two')]) },
  ],
});

const mounted = new Set<{ app: App; host: HTMLElement }>();

/** A mounted Vue tree: its host element, the queries of that host, and `unmount`. */
export type Mounted = BoundFunctions<typeof queries> & {
  readonly container: HTMLElement;
  unmount(): void;
};

/**
 * Mount `render` in a new host element in the document. `unmountAll` removes it. `setup` can
 * install app-level options, such as a mixin that counts renders, before the mount.
 */
export function mount(render: () => VNodeChild, setup?: (app: App) => void): Mounted {
  const host = document.createElement('div');
  document.body.append(host);
  const app = createApp({ render });
  setup?.(app);
  const entry = { app, host };
  mounted.add(entry);
  app.mount(host);
  return {
    ...within(host),
    container: host,
    unmount: () => {
      if (!mounted.delete(entry)) return;
      app.unmount();
      host.remove();
    },
  };
}

/** Unmount every tree that `mount` mounted. Call it in `afterEach`. */
export function unmountAll(): void {
  for (const { app, host } of [...mounted]) {
    app.unmount();
    host.remove();
  }
  mounted.clear();
}

/** Wait until the view shows `count` sheet tabs. */
export async function tabsShow(view: Mounted, count = 2): Promise<void> {
  await waitFor(() => expect(view.getAllByRole('tab').length).toBe(count), { timeout: 5000 });
}

/** A field of the formula bar by its label. */
export const field = (view: Mounted, name: string) =>
  view.getByLabelText(name) as HTMLInputElement | HTMLTextAreaElement;

/**
 * Set a field's value and fire the event that a browser fires for it: `input` for a text field,
 * which a Vue text field listens to, and `change` for a select. React's tests fire `change` for
 * both, because React's `onChange` listens to `input` too.
 */
export function change(
  field: HTMLInputElement | HTMLTextAreaElement | HTMLSelectElement,
  value: string
): boolean {
  field.value = value;
  return field instanceof HTMLSelectElement ? fireEvent.change(field) : fireEvent.input(field);
}
