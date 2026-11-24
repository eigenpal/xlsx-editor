/** @jsxImportSource react */
// The demo (`App.tsx`) in happy-dom, for the tests of its flows: the open, the saves, the
// questions before unsaved edits are lost, and the work that overlaps a load. These
// tests took the steps of the browser check (`scripts/verify-demo.ts`) that need no browser
// engine. The check keeps the steps that need one: the paint, the real clipboard, a real
// download, and the layout of the page.
//
// The harness stands in for the browser services that the demo uses:
// - `fetch` serves the packaged substitute faces and the sample workbook from the repository. A
//   test can hold or fail the fetch of the sample, from before the demo mounts.
// - The file save picker keeps each written copy, as the stand-in picker of the browser check
//   does. A test can cancel the next pick or replace the picker.
// - `window.confirm` records each question, and whether a load ran when it asked, and answers it
//   with `dialogs.accept`.
// - `navigator.userActivation` says that a click is recent, as after a real press.
//
// The URL of the sample comes from `import.meta.env.BASE_URL` (`workbooks.ts`), which only Vite
// sets, so the harness gives the module its own URL. A test opens a fixture as a picked file.

import { mock, setDefaultTimeout } from 'bun:test';
import { readdirSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { act, fireEvent, render, waitFor, type RenderResult } from '@testing-library/react';
import { cellFormat } from '../../../../packages/core/src/cascade/index.ts';
import {
  attr,
  cellKey,
  parseCellRef,
  readXlsx,
  sheetByName,
  sheetIndex,
  workbookIndex,
  xmlPartOf,
  type XlsxPackage,
} from '../../../../packages/core/src/store/index.ts';
import { readZip, writeZip } from '../../../../packages/ooxml-core/src/zip.ts';
import { fixtureFile } from '../../../../scripts/fixtures/catalog.ts';

const ROOT = join(import.meta.dir, '../../../..');
export const FIXTURES_DIR = join(ROOT, 'fixtures');
const SAMPLE_FILE = join(ROOT, 'examples/vite/public/sample.xlsx');

/** The names of the committed fixtures, sorted: each `.xlsx` and `.xlsm` file (`fixtureFile`). */
export const fixtureNames = (): string[] =>
  readdirSync(FIXTURES_DIR)
    .filter((f) => /\.xls[xm]$/.test(f))
    .map((f) => f.replace(/\.xls[xm]$/, ''))
    .sort();

mock.module('../workbooks.ts', () => ({ SAMPLE_URL: '/sample.xlsx' }));

/** How long a test waits for the demo, in milliseconds. */
export const WAIT_MS = 20_000;

// A test of a flow waits for several opens and saves, each of which can take seconds.
setDefaultTimeout(90_000);

/** The bytes of a committed fixture. */
export const fixtureBytes = (name: string): Uint8Array<ArrayBuffer> =>
  new Uint8Array(readFileSync(join(FIXTURES_DIR, fixtureFile(name))));

/** The bytes of the committed sample workbook. */
export const sampleBytes = (): Uint8Array<ArrayBuffer> => new Uint8Array(readFileSync(SAMPLE_FILE));

/** The sample with no sheet in its workbook part: it reads, but the editor has nothing to show. */
export function sampleWithNoSheet(): Uint8Array {
  const zip = readZip(sampleBytes());
  if (!zip.ok) throw new Error(`the sample does not unzip (${zip.reason})`);
  const entries = new Map(zip.entries);
  const workbook = new TextDecoder().decode(entries.get('/xl/workbook.xml'));
  const empty = workbook.replace(/<sheets>[\s\S]*?<\/sheets>/, '<sheets/>');
  if (empty === workbook) throw new Error('the sample workbook part has no sheets element');
  entries.set('/xl/workbook.xml', new TextEncoder().encode(empty));
  return writeZip(entries);
}

/** A cell of `pkg`, with its resolved font weight, format code, and number format id. */
export function formatAt(pkg: XlsxPackage, sheet: string, ref: string) {
  const index = workbookIndex(pkg)!;
  const part = xmlPartOf(pkg, sheetByName(index, sheet)?.partName);
  const at = parseCellRef(ref)!;
  const cell = part ? sheetIndex(part).cells.get(cellKey(at.row, at.col)) : undefined;
  const { format } = cellFormat(pkg, index, attr(cell?.el, 's'));
  return {
    bold: format.font.bold,
    code: format.formatCode ?? 'General',
    numFmtId: format.numFmtId,
    cell,
  };
}

/** Reads a saved copy, and throws when it does not read. */
export function readSaved(bytes: Uint8Array): XlsxPackage {
  const read = readXlsx(bytes);
  if (!read.ok) throw new Error(`the saved copy does not read (${read.error.reason})`);
  return read.pkg;
}

/** The names of the sheets of a saved copy. */
export const sheetsOf = (bytes: Uint8Array): string[] =>
  workbookIndex(readSaved(bytes))?.sheets.map((s) => s.name) ?? [];

/** A pending answer that a test releases. */
interface Held {
  readonly match: string;
  readonly done: Promise<void>;
}

/** The browser services of one mounted demo, and the controls of a test over them. */
class Services {
  readonly saved: Uint8Array[] = [];
  readonly dialogs = { seen: [] as string[], whileLoading: [] as boolean[], accept: true };
  readonly fetched: string[] = [];
  readonly suggested: string[] = [];
  cancelNext = false;
  activation = true;
  private readonly held: Held[] = [];
  private readonly failing = new Set<string>();

  /** Holds each fetch whose URL holds `match` until the returned call. */
  hold(match: string): () => void {
    let release = () => {};
    const done = new Promise<void>((resolve) => (release = resolve));
    const entry = { match, done };
    this.held.push(entry);
    return () => {
      this.held.splice(this.held.indexOf(entry), 1);
      release();
    };
  }

  /** Fails each fetch whose URL holds `match` as a lost network does, until the returned call. */
  fail(match: string): () => void {
    this.failing.add(match);
    return () => this.failing.delete(match);
  }

  readonly fetch = async (input: RequestInfo | URL): Promise<Response> => {
    const url = String(input instanceof Request ? input.url : input);
    if (url.startsWith('file:')) return new Response(readFileSync(fileURLToPath(url)));
    this.fetched.push(url);
    for (const h of [...this.held]) if (url.includes(h.match)) await h.done;
    for (const m of this.failing) if (url.includes(m)) throw new TypeError('the network failed');
    if (url.endsWith('/sample.xlsx')) return new Response(sampleBytes());
    return new Response(null, { status: 404 });
  };

  /** The default file save picker: it keeps each written copy in `saved`. */
  readonly keepingPicker = async (options: { suggestedName: string }) => {
    this.suggested.push(options.suggestedName);
    if (this.cancelNext) {
      this.cancelNext = false;
      throw new DOMException('The user cancelled the picker.', 'AbortError');
    }
    return {
      createWritable: async () => {
        const chunks: Blob[] = [];
        return {
          write: async (data: Blob) => void chunks.push(data),
          close: async () =>
            void this.saved.push(new Uint8Array(await new Blob(chunks).arrayBuffer())),
        };
      },
    };
  };
}

/** The first `Element` that matches `selector`, or a thrown error that names it. */
function one<T extends Element>(root: ParentNode, selector: string): T {
  const el = root.querySelector<T>(selector);
  if (!el) throw new Error(`no element ${selector}`);
  return el;
}

const key = (el: Element, k: string, init: KeyboardEventInit = {}) =>
  act(() => {
    el.dispatchEvent(
      new KeyboardEvent('keydown', { key: k, bubbles: true, cancelable: true, ...init })
    );
  });

/** Lets the demo's tasks, timers, and frames run for `ms` milliseconds. */
export const settle = (ms = 0) => act(() => new Promise<void>((done) => setTimeout(done, ms)));

/**
 * happy-dom lays out nothing, so the grid's scroller gets the size of a window of the browser
 * check (1600 by 900 pixels, less the chrome), and the grid shows the cells of that window.
 */
function sizeTheGrid(): void {
  const proto = HTMLElement.prototype;
  for (const [name, size] of [
    ['clientWidth', 1200],
    ['clientHeight', 600],
  ] as const) {
    const own = Object.getOwnPropertyDescriptor(proto, name);
    if (own?.get && (own.get as { grid?: true }).grid) continue;
    const base = own?.get ?? (() => 0);
    const get = Object.assign(
      function (this: HTMLElement) {
        return this.classList.contains('xlsx-scroller') ? size : base.call(this);
      },
      { grid: true as const }
    );
    Object.defineProperty(proto, name, { configurable: true, get });
  }
}

/** The options of {@link mountDemo}. */
interface MountOptions {
  /** False takes the file save picker away, so a save downloads. */
  readonly picker?: boolean;
  /** Runs before the demo mounts, so a test can hold or fail the download of the sample. */
  readonly beforeStart?: (services: Services) => void;
  /** False returns at once, with no wait for the sample. */
  readonly waitForSample?: boolean;
}

/** Mounts the demo with its stand-in browser services, and waits for the sample. */
export async function mountDemo(options: MountOptions = {}) {
  sizeTheGrid();
  const services = new Services();
  options.beforeStart?.(services);
  const w = window as unknown as Record<string, unknown>;
  globalThis.fetch = services.fetch as typeof fetch;
  w.fetch = services.fetch;
  if (options.picker === false) delete w.showSaveFilePicker;
  else w.showSaveFilePicker = services.keepingPicker;
  w.confirm = (text: string) => {
    services.dialogs.seen.push(text);
    services.dialogs.whileLoading.push(document.querySelector('.xlsx-canvas-idle') !== null);
    return services.dialogs.accept;
  };
  Object.defineProperty(navigator, 'userActivation', {
    configurable: true,
    get: () => ({ isActive: services.activation }),
  });
  const { App } = await import('../App.tsx');
  const view: RenderResult = render(<App />);
  const demo = new Demo(view, services);
  if (options.waitForSample !== false) await demo.showsSheet('Welcome');
  return demo;
}

/** One mounted demo, with the steps that the tests of its flows take. */
export class Demo {
  constructor(
    readonly view: RenderResult,
    readonly services: Services
  ) {}

  get saved() {
    return this.services.saved;
  }

  get dialogs() {
    return this.services.dialogs;
  }

  /** The grid's scroller, which takes the keys of the grid. */
  get grid(): HTMLElement {
    return one(document, '.xlsx-scroller');
  }

  get nameBox(): HTMLInputElement {
    return this.view.getByLabelText('Cell reference') as HTMLInputElement;
  }

  get content(): HTMLTextAreaElement {
    return this.view.getByLabelText('Cell content') as HTMLTextAreaElement;
  }

  /** The file input of Open a file in the top bar: the top bar's stop in the F6 cycle. */
  get openInput(): HTMLInputElement {
    return one<HTMLInputElement>(document, '.demo-topbar input[type=file]');
  }

  get viewOnlyBox(): HTMLInputElement {
    return this.view.getByRole('checkbox', { name: 'View only' }) as HTMLInputElement;
  }

  /** The Save button of the top bar. */
  get headerSave(): HTMLButtonElement {
    return [...document.querySelectorAll<HTMLButtonElement>('.demo-topbar button')].find(
      (b) => b.title === 'Save'
    )!;
  }

  /** The Save button of the toolbar. */
  get toolbarSave(): HTMLElement {
    return one(document, '.xlsx-toolbar [data-slot="file.save"]');
  }

  /** A toolbar button by its accessible name. */
  toolbarButton(name: string): HTMLElement {
    return this.view.getByRole('button', { name }) as HTMLElement;
  }

  /** The sheet that the grid shows, from its accessible name. */
  shownSheet(): string | null {
    const label = document.querySelector('[role=grid]')?.getAttribute('aria-label') ?? '';
    return label.startsWith('Worksheet ') ? label.slice('Worksheet '.length) : null;
  }

  /** Waits until the grid shows the sheet `name` and the editor is ready. */
  async showsSheet(name: string): Promise<void> {
    await waitFor(
      () => {
        if (this.shownSheet() !== name) throw new Error(`the grid shows ${this.shownSheet()}`);
        if (document.querySelector('.xlsx-canvas-idle')) throw new Error('the editor loads');
      },
      { timeout: WAIT_MS }
    );
  }

  /** Whether the editor loads a workbook now. */
  loading(): boolean {
    return document.querySelector('.xlsx-canvas-idle') !== null;
  }

  /** Waits until the editor starts a load. */
  async loadStarts(): Promise<void> {
    await waitFor(
      () => {
        if (!this.loading()) throw new Error('no load runs');
      },
      { timeout: WAIT_MS, interval: 1 }
    );
  }

  /** Shows the sheet `name` with its tab. */
  async tab(name: string): Promise<void> {
    act(() => void fireEvent.click(this.view.getByRole('tab', { name })));
    await this.showsSheet(name);
  }

  /** The text of each alert on the page: the demo's own and the editor's. */
  alerts(): string[] {
    return [...document.querySelectorAll('[role=alert]')].map((a) => a.textContent ?? '');
  }

  /** The text of the demo's message status region. */
  notice(): string {
    return document.querySelector('.demo-message__status')?.textContent ?? '';
  }

  /** The status text of long work in the top bar. */
  busy(): string {
    return document.querySelector('.demo-busy')?.textContent ?? '';
  }

  /** Throws unless an alert holds `text` now, with no wait that lets a load go on. */
  alertNow(text: string): void {
    if (!this.alerts().some((a) => a.includes(text)))
      throw new Error(`no alert says "${text}": ${JSON.stringify(this.alerts())}`);
  }

  /** Waits until an alert holds `text`. */
  async alertSays(text: string): Promise<void> {
    await waitFor(
      () => {
        if (!this.alerts().some((a) => a.includes(text)))
          throw new Error(`no alert says "${text}": ${JSON.stringify(this.alerts())}`);
      },
      { timeout: WAIT_MS, interval: 1 }
    );
  }

  /** Waits until a status region holds `text`. */
  async statusSays(text: string): Promise<void> {
    await waitFor(
      () => {
        const all = [...document.querySelectorAll('[role=status]')].map((s) => s.textContent);
        if (!all.some((s) => s?.includes(text)))
          throw new Error(`no status says "${text}": ${JSON.stringify(all)}`);
      },
      { timeout: WAIT_MS, interval: 1 }
    );
  }

  /** Closes the message of the demo with its button. */
  closeMessage(): void {
    act(() => void fireEvent.click(this.view.getByRole('button', { name: 'Close the message' })));
  }

  /** Whether leaving the page now asks first: the demo's `beforeunload` handler cancels it. */
  leaveAsks(): boolean {
    const e = new Event('beforeunload', { cancelable: true });
    window.dispatchEvent(e);
    return e.defaultPrevented;
  }

  /** Presses a key on `el`, the grid by default. */
  key(k: string, init: KeyboardEventInit = {}, el: Element = document.activeElement ?? this.grid) {
    return key(el, k, init);
  }

  /** Presses Control+S on `el`, the grid by default. */
  saveKey(el: Element = this.grid, init: KeyboardEventInit = {}) {
    return key(el, 's', { ctrlKey: true, ...init });
  }

  /** Selects `ref` with the name box, and returns the text of the formula bar for it. */
  goTo(ref: string): string {
    const box = this.nameBox;
    act(() => {
      box.focus();
      fireEvent.change(box, { target: { value: ref } });
    });
    key(box, 'Enter');
    if (box.value !== ref) throw new Error(`the name box went to ${box.value}, not ${ref}`);
    return this.content.value;
  }

  /**
   * Types `text` into the active cell from the grid, as keys do: the first character opens the
   * cell editor, and the rest goes into it. With `enter`, Enter commits it.
   */
  typeInGrid(text: string, enter = true): void {
    act(() => this.grid.focus());
    key(this.grid, text[0]!);
    const input = one<HTMLTextAreaElement>(document, '.xlsx-cell-editor');
    act(() => {
      input.value = text;
      input.dispatchEvent(new InputEvent('input', { bubbles: true }));
    });
    if (enter) key(input, 'Enter');
  }

  /** Clicks Bold in the toolbar. */
  bold(): void {
    act(() => void fireEvent.click(this.toolbarButton('Bold')));
  }

  /** Whether the toolbar's Bold shows the selection as bold. */
  boldShown(): boolean {
    return this.toolbarButton('Bold').getAttribute('aria-pressed') === 'true';
  }

  /** The name of the open workbook in the top bar, with its extension. */
  workbookName(): string {
    const field = document.querySelector<HTMLInputElement>('.demo-topbar .demo-file__name');
    const name = field?.value ?? '';
    return name + (document.querySelector('.demo-topbar .demo-file__ext')?.textContent ?? '');
  }

  /** Picks a file with the file input of the top bar. */
  pickFile(name: string, bytes: Uint8Array | string): void {
    const file = new File([typeof bytes === 'string' ? bytes : new Uint8Array(bytes)], name);
    act(() => void fireEvent.change(this.openInput, { target: { files: [file] } }));
  }

  /** Picks the committed fixture `name` as a file, as a user opens a workbook. */
  pickFixture(name: string): void {
    this.pickFile(fixtureFile(name), fixtureBytes(name));
  }

  /** Picks the committed sample workbook as a file. */
  pickSample(): void {
    this.pickFile('sample.xlsx', sampleBytes());
  }

  /** Whether the grid shows its column headings. */
  headingsShown(): boolean {
    const heads = document.querySelector<HTMLElement>('.xlsx-column-heads');
    return !!heads && !heads.hidden && heads.style.display !== 'none';
  }

  /**
   * Saves through the top bar's Save, the toolbar's Save, or Control+S in the grid, and returns
   * the bytes of the one copy that it writes. A second copy fails the step.
   */
  async save(from: 'header' | 'toolbar' | 'keys' = 'header'): Promise<Uint8Array> {
    const before = this.saved.length;
    if (from === 'keys') this.saveKey();
    else act(() => void fireEvent.click(from === 'header' ? this.headerSave : this.toolbarSave));
    await waitFor(
      () => {
        if (this.saved.length <= before)
          throw new Error(`the ${from} Save wrote no copy: ${JSON.stringify(this.alerts())}`);
      },
      { timeout: WAIT_MS }
    );
    // One save writes one copy: a second handler of the same key or click would write two.
    await settle(50);
    const copies = this.saved.length - before;
    if (copies !== 1) throw new Error(`the ${from} Save wrote ${copies} copies`);
    return this.saved.at(-1)!;
  }

  /** Runs `opens` with a no to each question, and returns the number of questions it gave. */
  async askedOnOpen(opens: () => unknown, settleMs = 300): Promise<number> {
    const before = this.dialogs.seen.length;
    this.dialogs.accept = false;
    try {
      await opens();
      await settle(settleMs);
      await waitFor(
        () => {
          if (this.busy() !== '') throw new Error('an open runs');
        },
        { timeout: WAIT_MS }
      );
    } finally {
      this.dialogs.accept = true;
    }
    return this.dialogs.seen.length - before;
  }
}
