// Formats every measured number format case (scripts/fixtures/cases.ts) and compares the text
// with the measured text (fixtures/measured.json), in the number locale that each
// fixture's reference record names (fixtures/references.json): its separators and the built-in
// codes that it replaces. The codes come from the measured texts.
//
// Colors: the reference PDFs paint `[Red]` text red (1 0 0) and `[Blue]` text blue (0 0 1)
// (number-formats.pdf); format.test.ts keeps the color of each section.

import { describe, expect, test } from 'bun:test';
import {
  caseCell,
  FORMAT_CASES,
  FORMAT_CASES_1904,
  FORMAT_CASES_MORE,
  type FormatCase,
} from '../../../../../../scripts/fixtures/cases.ts';
import { measuredCells } from '../../../../../../scripts/fixtures/measured.ts';
import { referenceLocale } from '../../../../../../scripts/fidelity/references.ts';
import { builtinFormatCode, dateToSerial, formatValue, type NumberLocale } from '../index.ts';

/** The number locale of a fixture's reference PDF. */
function referenceOf(fixture: string): NumberLocale {
  const locale = referenceLocale(fixture);
  if (!locale) throw new Error(`${fixture}: no reference record`);
  return locale;
}

function run(fixture: string, cases: readonly FormatCase[], date1904: boolean): void {
  describe(`measured cases of ${fixture}`, () => {
    const measured = measuredCells(fixture);
    const reference = referenceOf(fixture);
    cases.forEach((k, i) => {
      const ref = caseCell(i).ref;
      const code = k.code ?? builtinFormatCode(k.id!, reference)!;
      test(`${ref}: ${JSON.stringify(k.value)} through ${JSON.stringify(code)}`, () => {
        const value = k.isoDate ? dateToSerial(String(k.value), date1904)! : k.value;
        expect(value).toBeDefined();
        const out = formatValue(value, code, { locale: reference, date1904 });
        expect(out.problem).toBeUndefined();
        const shown = measured[ref]!.text;
        if (!out.fill) {
          expect(out.text).toBe(shown);
          return;
        }
        // The measured text holds the fill run for its column width; compare around it.
        const { char, index } = out.fill;
        const tail = out.text.length - index;
        expect(shown.slice(0, index)).toBe(out.text.slice(0, index));
        expect(shown.slice(shown.length - tail)).toBe(out.text.slice(index));
        const run = shown.slice(index, shown.length - tail);
        expect(run.length).toBeGreaterThan(0);
        for (const c of run) expect(c).toBe(char);
      });
    });
  });
}

run('number-formats', FORMAT_CASES, false);
run('number-formats-more', FORMAT_CASES_MORE, false);
run('dates-1904', FORMAT_CASES_1904, true);
