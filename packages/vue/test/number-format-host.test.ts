// The host procedure of the toolbar guide for a number format dialog of its own, the twin of the
// React adapter's `number-format-host.test.tsx`: the start when the dialog opens, the close
// check through `useEditorState`, and the Apply step with
// `useEditorCommand('number.custom').execute`. A host dialog applies a code, shows a problem,
// and closes when the selection moves. `NumberFormatHostDialog` is the Vue form of the guide's
// sample, with the same steps.

import { afterEach, expect, test } from 'bun:test';
import { fireEvent } from '@testing-library/react';
import {
  defineComponent,
  h,
  nextTick,
  onBeforeUnmount,
  onMounted,
  shallowRef,
  watch,
  type PropType,
} from 'vue';
import { buildXlsx, c, row } from '../../../scripts/fixtures/builder.ts';
import { translatorFor } from '../src/context.ts';
import {
  refusalText,
  numberFormatDialogApplyStep,
  numberFormatDialogCurrent,
  numberFormatDialogStart,
  useEditorCommand,
  useEditorState,
  useTranslation,
  useXlsxEditor,
  SheetTabs,
  XlsxEditorContent,
  XlsxEditorRoot,
  XlsxEditorToolbar,
  XlsxEditorViewport,
  type NumberFormatDialogStart,
  type XlsxEditorInstance,
} from '../src/index.ts';
import { FONTS, change, mount, tabsShow, unmountAll } from './mount.ts';

afterEach(unmountAll);

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

const NumberFormatHostDialog = defineComponent({
  props: {
    start: { type: Object as PropType<NumberFormatDialogStart>, required: true },
    close: { type: Function as PropType<() => void>, required: true },
  },
  setup(props) {
    const editor = useXlsxEditor();
    const { t } = useTranslation();
    const command = useEditorCommand('number.custom');
    const code = shallowRef(props.start.initial);
    const error = shallowRef<string | null>(null);
    const dialog = shallowRef<HTMLDialogElement | null>(null);
    // The browser focuses the field, keeps Tab in the dialog, and returns focus on close.
    onMounted(() => dialog.value!.showModal());
    onBeforeUnmount(() => dialog.value?.close());
    const current = useEditorState(() => numberFormatDialogCurrent(editor.value!, props.start));
    watch(current, (next) => {
      if (!next) props.close();
    });
    function apply(): void {
      const step = numberFormatDialogApplyStep(editor.value!, props.start, code.value);
      if (step.kind === 'problem') {
        error.value = t(step.problem.key, step.problem.params);
      } else if (step.kind !== 'apply' || command.execute(code.value)) {
        props.close();
      } else {
        // The editor refused the edit: show its reason.
        const snapshot = editor.value!.snapshot();
        error.value = snapshot.refusal
          ? refusalText(snapshot.refusal, t, snapshot.numberLocale)
          : t('dialogs.numberFormat.notApplied');
      }
    }
    return () =>
      h(
        'dialog',
        {
          ref: dialog,
          'aria-modal': 'true',
          'aria-label': t('dialogs.numberFormat.title'),
          onCancel: (e: Event) => {
            e.preventDefault();
            props.close();
          },
        },
        [
          h('input', {
            'aria-label': t('dialogs.numberFormat.code'),
            value: code.value,
            onInput: (e: Event) => (code.value = (e.currentTarget as HTMLInputElement).value),
          }),
          error.value ? h('p', { role: 'alert' }, error.value) : null,
          h('button', { type: 'button', onClick: apply }, t('common.apply')),
        ]
      );
  },
});

const Host = defineComponent({
  setup() {
    const editor = useXlsxEditor();
    const start = shallowRef<NumberFormatDialogStart | null>(null);
    const close = (): void => {
      start.value = null;
    };
    return () => [
      h(
        'button',
        {
          type: 'button',
          onClick: () => (start.value = editor.value && numberFormatDialogStart(editor.value)),
        },
        'Host open'
      ),
      start.value ? h(NumberFormatHostDialog, { start: start.value, close }) : null,
    ];
  },
});

async function open() {
  let viewer: XlsxEditorInstance | null = null;
  const saved: Uint8Array[] = [];
  const view = mount(() =>
    h(
      XlsxEditorRoot,
      {
        document: buildXlsx({ sheets: [{ name: 'S', rows: row(1, [c('A1', 1234.5)]) }] }),
        fonts: FONTS,
        onReady: (v: XlsxEditorInstance) => (viewer = v),
      },
      {
        default: () =>
          h(XlsxEditorViewport, null, () => [
            h(XlsxEditorToolbar, { onSave: (b: Uint8Array) => saved.push(b) }),
            h(Host),
            h(XlsxEditorContent),
            h(SheetTabs),
          ]),
      }
    )
  );
  await tabsShow(view, 1);
  return { view, viewer: viewer!, saved };
}

test('a host dialog that follows the toolbar guide applies a code, and shows a problem', async () => {
  const { view, viewer } = await open();
  fireEvent.click(view.getByRole('button', { name: 'Host open' }));
  await nextTick();
  const field = view.getByRole('textbox', { name: 'Format code' });
  change(field as HTMLInputElement, '"open');
  await nextTick();
  fireEvent.click(view.getByRole('button', { name: 'Apply' }));
  await nextTick();
  expect(view.getByRole('alert').textContent).toBe(
    'The format code cannot be read. Check the sections, the quotes, and the square brackets.'
  );
  change(field as HTMLInputElement, '0.0%');
  await nextTick();
  fireEvent.click(view.getByRole('button', { name: 'Apply' }));
  await nextTick();
  expect(view.queryByRole('dialog', { name: 'Number format' })).toBeNull();
  expect(viewer.snapshot().activeFormat?.numberFormat).toMatchObject({ id: 164, code: '0.0%' });
});

test('a host dialog that follows the toolbar guide closes when the selection moves', async () => {
  const { view, viewer } = await open();
  fireEvent.click(view.getByRole('button', { name: 'Host open' }));
  await nextTick();
  expect(view.getByRole('dialog', { name: 'Number format' })).toBeTruthy();
  viewer.select({ row: 3, col: 2 });
  await nextTick();
  expect(view.queryByRole('dialog', { name: 'Number format' })).toBeNull();
  expect(viewer.snapshot().canUndo).toBe(false);
});

test('a host dialog that follows the toolbar guide shows the reason of an engine refusal', async () => {
  const { view, viewer } = await open();
  fireEvent.click(view.getByRole('button', { name: 'Host open' }));
  await nextTick();
  change(view.getByRole('textbox', { name: 'Format code' }) as HTMLInputElement, '0.0');
  await nextTick();
  // A host-side check that the start does not see: the slot stays enabled, execute refuses, and
  // the snapshot holds the refusal, as the editor's own refusals do.
  const original = viewer.applyFormat.bind(viewer);
  const snapshot = viewer.snapshot.bind(viewer);
  const refusal = { code: 'view-mode' } as const;
  viewer.applyFormat = () => ({ ok: false, refusal }) as never;
  viewer.snapshot = () => ({ ...snapshot(), refusal }) as never;
  fireEvent.click(view.getByRole('button', { name: 'Apply' }));
  await nextTick();
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
  await nextTick();
  const field = view.getByRole('textbox', { name: 'Format code' });
  fireEvent.keyDown(field, { key: 's', ctrlKey: true });
  await nextTick();
  expect(saved.length).toBe(0);
  expect(view.getByRole('dialog', { name: 'Number format' })).toBeTruthy();
});

test('Escape on a host dialog that follows the toolbar guide closes it and writes nothing', async () => {
  const { view, viewer } = await open();
  fireEvent.click(view.getByRole('button', { name: 'Host open' }));
  await nextTick();
  const dialog = view.getByRole('dialog', { name: 'Number format' });
  fireEvent(dialog, new Event('cancel', { cancelable: true }));
  await nextTick();
  expect(view.queryByRole('dialog', { name: 'Number format' })).toBeNull();
  expect(viewer.snapshot().canUndo).toBe(false);
});
