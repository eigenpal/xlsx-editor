/** @jsxImportSource react */
// The number format commands of the toolbar in happy-dom (`bun run test:dom`): the list offers
// the currency, accounting, long date, fraction, and scientific entries, the currency symbol
// comes from the number locale, and the decimal step buttons run the engine command or show
// the engine's reason.

import { afterEach, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { act, cleanup, fireEvent, render, waitFor } from '@testing-library/react';
import { buildXlsx, c, row } from '../../../../scripts/fixtures/builder.ts';
import {
  XlsxEditor,
  type FontSource,
  type NumberLocale,
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

async function open(numberLocale?: NumberLocale) {
  let viewer: XlsxEditorInstance | null = null;
  const view = render(
    <div style={{ height: 400 }}>
      <XlsxEditor
        document={buildXlsx({ sheets: [{ name: 'S', rows: row(1, [c('A1', 1234.5)]) }] })}
        fonts={FONTS}
        onReady={(v) => (viewer = v)}
        {...(numberLocale ? { numberLocale } : {})}
      />
    </div>
  );
  await waitFor(() => expect(view.getAllByRole('tab').length).toBe(1), { timeout: 5000 });
  return { view, viewer: viewer! };
}

describe('the number format commands of the toolbar', () => {
  test('the list offers each entry, and a pick shows it on the trigger', async () => {
    const { view, viewer } = await open();
    fireEvent.click(view.getByRole('button', { name: /^Number format/ }));
    const names = view.getAllByRole('option').map((o) => o.textContent);
    expect(names).toEqual([
      'General',
      'Number',
      'Currency',
      'Accounting',
      'Percentage',
      'Scientific',
      'Fraction',
      'Short date',
      'Long date',
      'Time',
      'Text',
    ]);
    // An entry says only its name: no title, so no code becomes its accessible description.
    for (const o of view.getAllByRole('option')) {
      expect(o.getAttribute('title')).toBeNull();
      expect(o.getAttribute('aria-describedby')).toBeNull();
    }
    fireEvent.click(view.getByRole('option', { name: 'Currency' }));
    expect(viewer.snapshot().activeFormat?.numberFormat.code).toBe('"$"#,##0.00');
    expect(view.getByRole('button', { name: /^Number format/ }).textContent).toContain('Currency');
  });

  test('currency takes the symbol of the number locale', async () => {
    const { view, viewer } = await open({ decimal: ',', group: '.', currencySymbol: '€' });
    fireEvent.click(view.getByRole('button', { name: /^Number format/ }));
    fireEvent.click(view.getByRole('option', { name: 'Currency' }));
    expect(viewer.snapshot().activeFormat?.numberFormat.code).toBe('"€"#,##0.00');
  });

  test('the decimal buttons step the code, and show the reason when they cannot', async () => {
    const { view, viewer } = await open();
    // General shows 1234.5: an increase shows two places.
    fireEvent.click(view.getByRole('button', { name: 'Increase decimal places' }));
    expect(viewer.snapshot().activeFormat?.numberFormat.code).toBe('0.00');
    fireEvent.click(view.getByRole('button', { name: 'Decrease decimal places' }));
    fireEvent.click(view.getByRole('button', { name: 'Decrease decimal places' }));
    expect(viewer.snapshot().activeFormat?.numberFormat.code).toBe('0');
    const decrease = view.getByRole('button', { name: 'Decrease decimal places' });
    expect(decrease.getAttribute('aria-disabled')).toBe('true');
    expect(decrease.getAttribute('title')).toBe(
      'Decrease decimal places: This number format has no decimal places to remove.'
    );
  });

  test('a stepped entry shows the name of its entry, not a custom code', async () => {
    const { view, viewer } = await open();
    fireEvent.click(view.getByRole('button', { name: /^Number format/ }));
    fireEvent.click(view.getByRole('option', { name: 'Percentage' }));
    fireEvent.click(view.getByRole('button', { name: 'Decrease decimal places' }));
    expect(viewer.snapshot().activeFormat?.numberFormat.code).toBe('0.0%');
    const trigger = view.getByRole('button', { name: /^Number format/ });
    expect(trigger.textContent).toContain('Percentage');
    expect(trigger.textContent).not.toContain('Custom');
    // The accessible name and the title give the places, and never the code.
    expect(trigger.getAttribute('aria-label')).toBe('Number format: Percentage, 1 decimal place');
    expect(trigger.getAttribute('title')).toStartWith('Number format: Percentage, 1 decimal place');
    fireEvent.click(view.getByRole('button', { name: 'Increase decimal places' }));
    fireEvent.click(view.getByRole('button', { name: 'Increase decimal places' }));
    expect(trigger.getAttribute('aria-label')).toBe('Number format: Percentage, 3 decimal places');
    expect(trigger.getAttribute('title')).not.toContain('0.000%');
    // An exact entry names its places; a custom code is its own visible text.
    fireEvent.click(trigger);
    fireEvent.click(view.getByRole('option', { name: 'Percentage' }));
    expect(trigger.getAttribute('title')).toStartWith(
      'Number format: Percentage, 2 decimal places'
    );
    act(() => {
      expect(viewer.applyFormat({ numberFormat: { code: '0.0" kg"' } }).ok).toBe(true);
    });
    // A custom code names its places too, and its title gives the whole text.
    expect(trigger.getAttribute('aria-label')).toBe(
      'Number format: Custom: 0.0" kg", 1 decimal place'
    );
    expect(trigger.getAttribute('title')).toStartWith(
      'Number format: Custom: 0.0" kg", 1 decimal place'
    );
    act(() => {
      expect(viewer.applyFormat({ numberFormat: { code: '#,##0.00' } }).ok).toBe(true);
    });
    expect(trigger.getAttribute('aria-label')).toBe(
      'Number format: Custom: #,##0.00, 2 decimal places'
    );
    // A code with no step names no places.
    act(() => {
      expect(viewer.applyFormat({ numberFormat: { code: 'yyyy-mm' } }).ok).toBe(true);
    });
    expect(trigger.getAttribute('aria-label')).toBe('Number format: Custom: yyyy-mm');
    expect(trigger.getAttribute('title')).toStartWith('Number format: Custom: yyyy-mm');
  });

  test('the name of an entry states its decimal places, and an entry with none does not', async () => {
    const { view } = await open();
    const trigger = () => view.getByRole('button', { name: /^Number format/ });
    const pick = (name: string) => {
      fireEvent.click(trigger());
      fireEvent.click(view.getByRole('option', { name }));
    };
    pick('Number');
    expect(trigger().getAttribute('aria-label')).toBe('Number format: Number, 2 decimal places');
    pick('Currency');
    expect(trigger().getAttribute('aria-label')).toBe('Number format: Currency, 2 decimal places');
    pick('Short date');
    expect(trigger().getAttribute('aria-label')).toBe('Number format: Short date');
    // The title gives the value for every entry, also one with no places.
    expect(trigger().getAttribute('title')).toStartWith('Number format: Short date');
    pick('General');
    expect(trigger().getAttribute('aria-label')).toBe('Number format: General');
    expect(trigger().getAttribute('title')).toStartWith('Number format: General');
  });

  test('a disabled list gives its value and then its reason on a new line', async () => {
    const { view, viewer } = await open();
    act(() => {
      expect(viewer.selectReference('A:A').ok).toBe(true);
    });
    const trigger = view.getByRole('button', { name: /^Number format/ });
    expect(trigger.getAttribute('aria-disabled')).toBe('true');
    expect(trigger.getAttribute('title')).toBe(
      'Number format: General\nThe selection is too large to format. Select at most one whole row of cells.'
    );
  });

  test('the keyboard opens the list on the entry of a stepped code, and does not pick it', async () => {
    const { view } = await open();
    fireEvent.click(view.getByRole('button', { name: /^Number format/ }));
    fireEvent.click(view.getByRole('option', { name: 'Percentage' }));
    fireEvent.click(view.getByRole('button', { name: 'Increase decimal places' }));
    const trigger = view.getByRole('button', { name: /^Number format/ });
    fireEvent.keyDown(trigger, { key: 'ArrowDown' });
    const entry = view.getByRole('option', { name: 'Percentage' });
    expect(entry.ownerDocument.activeElement).toBe(entry);
    expect(entry.getAttribute('aria-selected')).toBe('false');
    // The entry is the current family, which a screen reader names, but it is not picked.
    expect(entry.getAttribute('aria-current')).toBe('true');
    expect(view.queryAllByRole('option', { selected: true })).toHaveLength(0);
  });

  test('a held Enter on a decimal button steps once', async () => {
    const { view } = await open();
    const increase = view.getByRole('button', { name: 'Increase decimal places' });
    // The first press is not a repeat: the browser clicks the button.
    expect(fireEvent.keyDown(increase, { key: 'Enter' })).toBe(true);
    // A repeat of the held key is cancelled, so the browser does not click again.
    expect(fireEvent.keyDown(increase, { key: 'Enter', repeat: true })).toBe(false);
  });

  test('a currency symbol that is not valid leaves currency and accounting out', async () => {
    const { view } = await open({ decimal: '.', group: ',', currencySymbol: '\u0001' });
    fireEvent.click(view.getByRole('button', { name: /^Number format/ }));
    const names = view.getAllByRole('option').map((o) => o.textContent);
    expect(names).not.toContain('Currency');
    expect(names).not.toContain('Accounting');
    expect(names).toContain('Number');
  });
});
