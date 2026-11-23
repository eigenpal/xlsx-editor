/** @jsxImportSource react */
// The host procedure of `packages/react/docs/toolbar-and-keyboard.md` for a number format dialog
// of its own: the start when the
// dialog opens, the close check through `useEditorState`, and the Apply step with
// `useEditorCommand('number.custom').execute`. A host dialog applies a code, shows a problem,
// and closes when the selection moves. `HostDialog` is the `NumberFormatHostDialog` sample of
// that guide: change the two together.

import { afterEach, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { useCallback, useEffect, useLayoutEffect, useRef, useState } from 'react';
import { act, cleanup, fireEvent, render, waitFor } from '@testing-library/react';
import { buildXlsx, c, row } from '../../../../scripts/fixtures/builder.ts';
import { translatorFor } from '../context.ts';
import {
  refusalText,
  numberFormatDialogApplyStep,
  numberFormatDialogCurrent,
  numberFormatDialogStart,
  useEditorCommand,
  useEditorState,
  useTranslation,
  useXlsxEditor,
  XlsxEditor,
  type FontSource,
  type NumberFormatDialogStart,
  type XlsxEditorInstance,
} from '../index.ts';

const fontsDir = join(
  dirname(
    Bun.resolveSync('@docx-editor.dev/fonts/package.json', join(import.meta.dir, '../../../core'))
  ),
  'assets'
);
const FONTS: FontSource[] = [
  {
    family: 'Carlito',
    bold: false,
    italic: false,
    bytes: new Uint8Array(readFileSync(join(fontsDir, 'Carlito-Regular.ttf'))),
    origin: 'substitute',
  },
];

afterEach(cleanup);

// The DOM test environment has no modal dialog: `showModal` and `close` set and clear `open`.
const dialogProto = HTMLDialogElement.prototype as unknown as Record<string, unknown>;
if (typeof dialogProto.showModal !== 'function')
  dialogProto.showModal = function (this: HTMLDialogElement) {
    this.setAttribute('open', '');
  };
if (typeof dialogProto.close !== 'function')
  dialogProto.close = function (this: HTMLDialogElement) {
    this.removeAttribute('open');
  };

function NumberFormatHostDialog({
  start,
  close,
}: {
  start: NumberFormatDialogStart;
  close: () => void;
}) {
  const editor = useXlsxEditor()!;
  const { t } = useTranslation();
  const command = useEditorCommand('number.custom');
  const [code, setCode] = useState(start.initial);
  const [error, setError] = useState<string | null>(null);
  const ref = useRef<HTMLDialogElement>(null);
  // The browser focuses the field, keeps Tab in the dialog, and returns focus on close.
  useLayoutEffect(() => {
    const dialog = ref.current!;
    dialog.showModal();
    return () => dialog.close();
  }, []);
  const current = useEditorState(
    useCallback(() => numberFormatDialogCurrent(editor, start), [editor, start])
  );
  useEffect(() => {
    if (!current) close();
  }, [current, close]);
  function apply() {
    const step = numberFormatDialogApplyStep(editor, start, code);
    if (step.kind === 'problem') {
      setError(t(step.problem.key, step.problem.params));
    } else if (step.kind !== 'apply' || command.execute(code)) {
      close();
    } else {
      // The editor refused the edit: show its reason.
      const snapshot = editor.snapshot();
      setError(
        snapshot.refusal
          ? refusalText(snapshot.refusal, t, snapshot.numberLocale)
          : t('dialogs.numberFormat.notApplied')
      );
    }
  }
  return (
    <dialog
      ref={ref}
      aria-modal="true"
      aria-label={t('dialogs.numberFormat.title')}
      onCancel={(e) => {
        e.preventDefault();
        close();
      }}
    >
      <input
        aria-label={t('dialogs.numberFormat.code')}
        value={code}
        onChange={(e) => setCode(e.currentTarget.value)}
      />
      {error ? <p role="alert">{error}</p> : null}
      <button type="button" onClick={apply}>
        {t('common.apply')}
      </button>
    </dialog>
  );
}

function Host() {
  const editor = useXlsxEditor();
  const [start, setStart] = useState<NumberFormatDialogStart | null>(null);
  const close = useCallback(() => setStart(null), []);
  return (
    <>
      <button type="button" onClick={() => setStart(editor && numberFormatDialogStart(editor))}>
        Host open
      </button>
      {start ? <NumberFormatHostDialog start={start} close={close} /> : null}
    </>
  );
}

async function open() {
  let viewer: XlsxEditorInstance | null = null;
  const saved: Uint8Array[] = [];
  const view = render(
    <XlsxEditor.Root
      document={buildXlsx({ sheets: [{ name: 'S', rows: row(1, [c('A1', 1234.5)]) }] })}
      fonts={FONTS}
      onReady={(v) => (viewer = v)}
    >
      <XlsxEditor.Viewport>
        <XlsxEditor.Toolbar onSave={(b) => saved.push(b)} />
        <Host />
        <XlsxEditor.Content />
        <XlsxEditor.SheetTabs />
      </XlsxEditor.Viewport>
    </XlsxEditor.Root>
  );
  await waitFor(() => expect(view.getAllByRole('tab').length).toBe(1), { timeout: 5000 });
  return { view, viewer: viewer!, saved };
}

test('a host dialog that follows the toolbar guide applies a code, and shows a problem', async () => {
  const { view, viewer } = await open();
  fireEvent.click(view.getByRole('button', { name: 'Host open' }));
  const field = view.getByRole('textbox', { name: 'Format code' });
  fireEvent.change(field, { target: { value: '"open' } });
  fireEvent.click(view.getByRole('button', { name: 'Apply' }));
  expect(view.getByRole('alert').textContent).toBe(
    'The format code cannot be read. Check the sections, the quotes, and the square brackets.'
  );
  fireEvent.change(field, { target: { value: '0.0%' } });
  fireEvent.click(view.getByRole('button', { name: 'Apply' }));
  expect(view.queryByRole('dialog', { name: 'Number format' })).toBeNull();
  expect(viewer.snapshot().activeFormat?.numberFormat).toMatchObject({ id: 164, code: '0.0%' });
});

test('a host dialog that follows the toolbar guide closes when the selection moves', async () => {
  const { view, viewer } = await open();
  fireEvent.click(view.getByRole('button', { name: 'Host open' }));
  expect(view.getByRole('dialog', { name: 'Number format' })).toBeTruthy();
  act(() => {
    viewer.select({ row: 3, col: 2 });
  });
  expect(view.queryByRole('dialog', { name: 'Number format' })).toBeNull();
  expect(viewer.snapshot().canUndo).toBe(false);
});

test('a host dialog that follows the toolbar guide shows the reason of an engine refusal', async () => {
  const { view, viewer } = await open();
  fireEvent.click(view.getByRole('button', { name: 'Host open' }));
  fireEvent.change(view.getByRole('textbox', { name: 'Format code' }), {
    target: { value: '0.0' },
  });
  // A host-side check that the start does not see: the slot stays enabled, execute refuses, and
  // the snapshot holds the refusal, as the editor's own refusals do.
  const original = viewer.applyFormat.bind(viewer);
  const snapshot = viewer.snapshot.bind(viewer);
  const refusal = { code: 'view-mode' } as const;
  viewer.applyFormat = () => ({ ok: false, refusal }) as never;
  viewer.snapshot = () => ({ ...snapshot(), refusal }) as never;
  fireEvent.click(view.getByRole('button', { name: 'Apply' }));
  viewer.applyFormat = original;
  viewer.snapshot = snapshot;
  expect(view.getByRole('dialog', { name: 'Number format' })).toBeTruthy();
  expect(view.getByRole('alert').textContent).toBe(
    refusalText(refusal, translatorFor(undefined), viewer.snapshot().numberLocale)
  );
});

test('Control and S in a host dialog with aria-modal does not save', async () => {
  const { view, saved } = await open();
  fireEvent.click(view.getByRole('button', { name: 'Host open' }));
  const field = view.getByRole('textbox', { name: 'Format code' });
  fireEvent.keyDown(field, { key: 's', ctrlKey: true });
  expect(saved.length).toBe(0);
  expect(view.getByRole('dialog', { name: 'Number format' })).toBeTruthy();
});

test('Escape on a host dialog that follows the toolbar guide closes it and writes nothing', async () => {
  const { view, viewer } = await open();
  fireEvent.click(view.getByRole('button', { name: 'Host open' }));
  const dialog = view.getByRole('dialog', { name: 'Number format' });
  fireEvent(dialog, new Event('cancel', { cancelable: true }));
  expect(view.queryByRole('dialog', { name: 'Number format' })).toBeNull();
  expect(viewer.snapshot().canUndo).toBe(false);
});
