// The Vue twin of packages/react/src/__tests__/chrome-a11y.test.tsx: the element in each
// pressed, selected, focused, or disabled state matches the forced-colors rule of that state,
// and the element out of that state does not. The stylesheet checks are shared, so they run in
// the React test only.

import { afterEach, describe, expect, test } from 'bun:test';
import { fireEvent, waitFor } from '@testing-library/react';
import { h, nextTick } from 'vue';
import { buildXlsx, c, row } from '../../../scripts/fixtures/builder.ts';
import { XlsxEditor } from '../src/index.ts';
import { FORCED_STATES } from '../../core/src/editor/__tests__/forced-colors-states.ts';
import { FONTS, mount, unmountAll } from './mount.ts';
import { workbook } from './toolbar-stubs.ts';

afterEach(unmountAll);

async function open(bytes = workbook()) {
  const view = mount(() =>
    h('div', { style: { height: '400px' } }, [h(XlsxEditor, { document: bytes, fonts: FONTS })])
  );
  await waitFor(() => expect(view.getAllByRole('tab').length).toBeGreaterThan(0), {
    timeout: 5000,
  });
  return view;
}

/** Fire `event` and wait for Vue to render. */
async function act(event: () => unknown): Promise<void> {
  event();
  await nextTick();
}

describe('the chrome in each state matches its forced-colors rule', () => {
  const { pressed, selectedTab, selectedOption, focusedOption, disabledButton } = FORCED_STATES;

  test('a pressed toggle, the selected tab, and a disabled button', async () => {
    const view = await open(
      buildXlsx({
        sheets: [
          { name: 'S', rows: row(1, [c('A1', 1)]) },
          { name: 'T', rows: row(1, [c('A1', 2)]) },
        ],
      })
    );
    const bold = view.getByRole('button', { name: 'Bold' });
    expect(bold.matches(pressed.selector)).toBe(false);
    await act(() => fireEvent.click(bold));
    expect(bold.getAttribute('aria-pressed')).toBe('true');
    expect(bold.matches(pressed.selector)).toBe(true);
    expect(view.getByRole('tab', { name: 'S' }).matches(selectedTab.selector)).toBe(true);
    expect(view.getByRole('tab', { name: 'T' }).matches(selectedTab.selector)).toBe(false);
    const { focusedTab, focusedSelectedTab } = FORCED_STATES;
    const [first, second] = [
      view.getByRole('tab', { name: 'S' }),
      view.getByRole('tab', { name: 'T' }),
    ];
    first.focus();
    expect(first.matches(focusedTab.selector)).toBe(true);
    expect(first.matches(focusedSelectedTab.selector)).toBe(true);
    second.focus();
    expect(second.matches(focusedTab.selector)).toBe(true);
    expect(second.matches(focusedSelectedTab.selector)).toBe(false);
    expect(view.getByRole('button', { name: 'Redo' }).matches(disabledButton.selector)).toBe(true);
    expect(view.getByRole('button', { name: 'Undo' }).matches(disabledButton.selector)).toBe(false);
  });

  test('the selected and the focused list option', async () => {
    const view = await open();
    await act(() => fireEvent.click(view.getByRole('button', { name: /^Number format/ })));
    const general = view.getByRole('option', { name: 'General' });
    const number = view.getByRole('option', { name: 'Number' });
    expect(general.getAttribute('aria-selected')).toBe('true');
    expect(general.matches(selectedOption.selector)).toBe(true);
    expect(number.matches(selectedOption.selector)).toBe(false);
    number.focus();
    expect(number.matches(focusedOption.selector)).toBe(true);
    expect(general.matches(focusedOption.selector)).toBe(false);
    // The item under the list takes the ring of a focused option, never the selected fill.
    const more = view.getByRole('button', { name: 'More number formats…' });
    expect(more.matches(FORCED_STATES.focusedListMore.selector)).toBe(false);
    more.focus();
    expect(more.matches(FORCED_STATES.focusedListMore.selector)).toBe(true);
    expect(more.matches(selectedOption.selector)).toBe(false);
    // The selected option with focus: its ring takes the text color of its fill.
    general.focus();
    expect(general.matches(FORCED_STATES.focusedSelectedOption.selector)).toBe(true);
    expect(number.matches(FORCED_STATES.focusedSelectedOption.selector)).toBe(false);
  });

  test('the selected color swatch and the selected no fill', async () => {
    const { swatch, selectedSwatch, selectedNoFill, focusedSwatch } = FORCED_STATES;
    const view = await open();
    await act(() => fireEvent.click(view.getByRole('button', { name: 'Font color options' })));
    const red = view.getByRole('button', { name: 'Red' });
    expect(red.matches(swatch.selector)).toBe(true);
    expect(red.matches(selectedSwatch.selector)).toBe(false);
    await act(() => fireEvent.click(red));
    await act(() => fireEvent.click(view.getByRole('button', { name: 'Font color options' })));
    expect(view.getByRole('button', { name: 'Red' }).matches(selectedSwatch.selector)).toBe(true);
    expect(view.getByRole('button', { name: 'Blue' }).matches(selectedSwatch.selector)).toBe(false);
    await act(() => fireEvent.click(view.getByRole('button', { name: 'Fill color options' })));
    // A cell with no fill: the no fill option is the selected one.
    const none = view.getByRole('button', { name: 'No fill' });
    expect(none.getAttribute('aria-pressed')).toBe('true');
    expect(none.matches(selectedNoFill.selector)).toBe(true);
    expect(none.matches(selectedSwatch.selector)).toBe(false);
    expect(none.matches(focusedSwatch.selector)).toBe(false);
    none.focus();
    expect(none.matches(focusedSwatch.selector)).toBe(true);
  });

  test('a focused color swatch differs from the selected one', async () => {
    const { selectedSwatch, focusedSwatch } = FORCED_STATES;
    const view = await open();
    await act(() => fireEvent.click(view.getByRole('button', { name: 'Font color options' })));
    await act(() => fireEvent.click(view.getByRole('button', { name: 'Red' })));
    await act(() => fireEvent.click(view.getByRole('button', { name: 'Font color options' })));
    const red = view.getByRole('button', { name: 'Red' });
    const blue = view.getByRole('button', { name: 'Blue' });
    blue.focus();
    expect(blue.matches(focusedSwatch.selector)).toBe(true);
    expect(blue.matches(selectedSwatch.selector)).toBe(false);
    expect(red.matches(selectedSwatch.selector)).toBe(true);
    expect(red.matches(focusedSwatch.selector)).toBe(false);
    // The selected swatch keeps the line inside its edge while it has focus.
    red.focus();
    expect(red.matches(focusedSwatch.selector)).toBe(true);
    expect(red.matches(selectedSwatch.selector)).toBe(true);
  });

  test('the focused and the disabled sheet menu item', async () => {
    const { focusedMenuItem, disabledMenuItem } = FORCED_STATES;
    const view = await open();
    await act(() => fireEvent.contextMenu(view.getByRole('tab', { name: 'S' })));
    // The only sheet cannot be deleted.
    const remove = view.getByRole('menuitem', { name: 'Delete' });
    const rename = view.getByRole('menuitem', { name: 'Rename' });
    expect(remove.matches(disabledMenuItem.selector)).toBe(true);
    expect(rename.matches(disabledMenuItem.selector)).toBe(false);
    rename.focus();
    expect(rename.matches(focusedMenuItem.selector)).toBe(true);
    expect(rename.matches(FORCED_STATES.focusedMenuRow.selector)).toBe(true);
    remove.focus();
    expect(remove.matches(focusedMenuItem.selector)).toBe(false);
    expect(remove.matches(FORCED_STATES.focusedMenuRow.selector)).toBe(false);
  });
});
