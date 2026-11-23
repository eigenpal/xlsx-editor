/** @jsxImportSource react */
// Each committed fixture in the demo, in happy-dom (`demo-harness.tsx`): it opens from the
// workbook list with no question, and each of its visible worksheets shows. A value typed into an
// empty cell of its opening sheet saves, the saved copy keeps the value and every sheet, and the
// copy opens again from a picked file with the value. This was the fixture step of the browser
// check (`scripts/verify-demo.ts`). The paint of each fixture is measured by `bun run fidelity`.

import { afterEach, describe, expect, test } from 'bun:test';
import { cleanup } from '@testing-library/react';
import { cellValue, sheetStale } from '../../../../packages/core/src/cascade/index.ts';
import { editorSheets, openingSheet } from '../../../../packages/core/src/editor/editor-model.ts';
import {
  cellKey,
  parseCellRef,
  readXlsx,
  sheetByName,
  sheetIndex,
  workbookIndex,
  xmlPartOf,
  type XlsxPackage,
} from '../../../../packages/core/src/store/index.ts';
import { SHEET_GUIDE } from '../../../../scripts/sample-workbook/showcase-sheets.ts';
import { SAMPLE } from '../sample.ts';
import { fixtureBytes, fixtureNames, mountDemo, readSaved, type Demo } from './demo-harness.tsx';

/**
 * The cell that the test edits on the opening sheet of each fixture. The test refuses a fixture
 * that has a cell there, so the edit never replaces fixture content.
 */
const FIXTURE_EDIT_REF = 'AB300';
const [SAMPLE_SHEET] = SHEET_GUIDE[0]!;
const TEXT = 'Edited in the demo';

/** The stored text of the cell `ref` of `sheet` in `pkg`, or `undefined` for no cell. */
function textAt(pkg: XlsxPackage, sheet: string, ref: string): string | undefined {
  const index = workbookIndex(pkg)!;
  const part = xmlPartOf(pkg, sheetByName(index, sheet)?.partName);
  const at = parseCellRef(ref)!;
  const cell = part ? sheetIndex(part).cells.get(cellKey(at.row, at.col)) : undefined;
  if (!part || !cell) return undefined;
  const { value } = cellValue(pkg, index, cell.el, sheetStale(index, part.root));
  return value.kind === 'string' ? value.text : JSON.stringify(value);
}

/** The names of the sheets of `pkg`. */
const namesOf = (pkg: XlsxPackage) => workbookIndex(pkg)?.sheets.map((s) => s.name);

afterEach(cleanup);

/**
 * Opens the sample and waits until it shows. A saved copy opens on the sheet of its fixture, so
 * the sample opens before it: the grid then shows that sheet only after the copy loads.
 */
async function openSample(demo: Demo): Promise<void> {
  demo.selectWorkbook(SAMPLE);
  await demo.showsSheet(SAMPLE_SHEET);
}

describe('each committed fixture in the demo', () => {
  for (const fixture of fixtureNames())
    test(fixture, async () => {
      const read = readXlsx(fixtureBytes(fixture));
      if (!read.ok) throw new Error(`the fixture does not read (${read.error.reason})`);
      const index = workbookIndex(read.pkg)!;
      const sheets = editorSheets(index);
      const first = sheets[openingSheet(sheets, index.activeTab)];
      if (!first) throw new Error('the fixture has no sheet to open');
      // A fixture that opens on a sheet with the name of the sample's first sheet cannot be told
      // apart from the sample.
      expect(first.name).not.toBe(SAMPLE_SHEET);
      expect(textAt(read.pkg, first.name, FIXTURE_EDIT_REF)).toBeUndefined();
      const demo = await mountDemo();
      const asked = demo.dialogs.seen.length;
      demo.selectWorkbook(fixture);
      await demo.showsSheet(first.name);
      for (const { name } of sheets.filter((s) => s.visible && s.worksheet)) await demo.tab(name);
      await demo.tab(first.name);

      // The edit: a value typed into an empty cell, committed with Enter.
      demo.goTo(FIXTURE_EDIT_REF);
      demo.typeInGrid(TEXT);
      expect(demo.goTo(FIXTURE_EDIT_REF)).toBe(TEXT);
      const bytes = await demo.save('toolbar');
      const saved = readSaved(bytes);
      expect(namesOf(saved)).toEqual(namesOf(read.pkg));
      expect(textAt(saved, first.name, FIXTURE_EDIT_REF)).toBe(TEXT);
      demo.closeMessage();

      // The saved copy opens again in the demo, from a picked file, with the value.
      await openSample(demo);
      demo.pickFile(`${fixture}.xlsx`, bytes);
      await demo.showsSheet(first.name);
      expect(demo.goTo(FIXTURE_EDIT_REF)).toBe(TEXT);
      // No step left an unsaved edit, so no open asked.
      expect(demo.dialogs.seen.length).toBe(asked);
    });
});
