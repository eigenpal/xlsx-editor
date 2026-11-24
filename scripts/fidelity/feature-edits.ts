// The feature edits: one recipe for each support matrix row that edits or keeps workbook content
// (`docs/site/data/support-matrix.ts`). Each recipe edits a fixture of its row through the public
// editor API, saves, reopens, and finds the edit again (`feature-edit-kit.ts`). A row that keeps
// content edits next to it, and its check finds the content as it was.
//
// `bun test scripts/fidelity/feature-edits.test.ts` runs every recipe, and `bun
// scripts/fidelity/feature-edits.ts --out <dir>` runs them and writes one edited copy per row,
// `<row>.xlsx` (`<row>.xlsm` on a macro-enabled fixture). A maintainer opens each copy, and
// `fixtures/feature-edits.json` records the result (CONTRIBUTING.md, "Save check").
//
// Same form as pptx-editor/scripts/fidelity/feature-edits.ts at 30dc48ea (EigenPal). Differences:
// a recipe names no id (the key is its row), and the run returns a list of problems. `--out`
// prints the sheet count of each copy, which the save check compares, and writes no copy of a
// recipe that fails.

import { mkdirSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { parseArgs } from 'node:util';
import { contentTypeOf, workbookIndex } from '../../packages/core/src/store/index.ts';
import { isMacroEnabledType } from '../../packages/core/src/store/package/main-part-type.ts';
import {
  all,
  featureCopyName,
  has,
  lacks,
  point,
  removed,
  runFeatureEdit,
  text,
  type FeatureEdit,
} from './feature-edit-kit.ts';

const WORKBOOK = /^\/xl\/workbook\.xml$/;
const STYLES = /^\/xl\/styles\.xml$/;
/** The parts that the first shared string of a workbook adds or changes. */
const SHARED_STRINGS = [
  /^\/xl\/sharedStrings\.xml$/,
  /^\/xl\/_rels\/workbook\.xml\.rels$/,
  /^\/\[Content_Types\]\.xml$/,
];
/** The parts that a dropped calculation chain changes. */
const CALC_CHAIN_DROP = [
  WORKBOOK,
  removed(/^\/xl\/calcChain\.xml$/),
  /^\/xl\/_rels\/workbook\.xml\.rels$/,
  /^\/\[Content_Types\]\.xml$/,
];

/** A defined name as the saved workbook part writes it. */
const definedName = (name: string, formula: string) =>
  `<definedName name="${name}">${formula}</definedName>`;

/** The recipe of a format edit on `colors-fonts`, B1, through a toolbar slot or the editor. */
const formatEdit = (
  edit: FeatureEdit['edit'],
  check: (
    format: ReturnType<Parameters<FeatureEdit['check']>[0]['cellFormat']>
  ) => string | undefined
): FeatureEdit => ({
  fixture: 'colors-fonts',
  edit: (t) => {
    t.select('B1');
    edit(t);
  },
  check: (d) => check(d.cellFormat('Colors', 'B1')),
  parts: [STYLES],
});

/**
 * The recipe of a drawing row on `grid-structure`: an edit beside the drawing, and the drawing
 * part and its media stay as they were. The pictures and shapes rows share one copy.
 */
const keptDrawingEdit = (kept: FeatureEdit['check']): FeatureEdit => ({
  fixture: 'grid-structure',
  edit: (t) => t.run(t.editor.enterValue('edited', point('A7'))),
  check: (d) => all(text(d.cellText('Grid', 'A7'), 'edited'), kept(d)),
  parts: SHARED_STRINGS,
});

export const FEATURE_EDITS: Readonly<Record<string, FeatureEdit>> = {
  'editing.cell-values': {
    fixture: 'basic-values',
    // A typed percent reads as a number and sets the number format of its form.
    edit: (t) => t.run(t.editor.enterValue('12%', point('C3'))),
    check: (d) =>
      all(
        has(d.cellXml('Values', 'C3'), '<v>0.12</v>'),
        d.cellFormat('Values', 'C3').numFmtId === 9 ? undefined : 'C3 is not built-in format 9'
      ),
    parts: [STYLES, WORKBOOK],
  },
  'editing.cell-editor': {
    fixture: 'basic-values',
    mount: true,
    edit: (t) => {
      t.select('C3');
      t.run(t.editor.beginEdit('Typed text'));
      t.run(t.editor.commitEdit('down'));
    },
    check: (d) => text(d.cellText('Values', 'C3'), 'Typed text'),
    parts: [...SHARED_STRINGS, WORKBOOK],
  },
  'editing.formula-bar': {
    fixture: 'basic-values',
    mount: true,
    // The content field of the formula bar sets the text of the open edit.
    edit: (t) => {
      t.select('C3');
      t.run(t.editor.beginEdit());
      t.run(t.editor.setEditText('=A2+1'));
      t.run(t.editor.commitEdit());
    },
    check: (d) => has(d.cellXml('Values', 'C3'), '<f>A2+1</f>'),
    parts: [WORKBOOK],
  },
  'editing.formulas': {
    fixture: 'formulas',
    edit: (t) => t.run(t.editor.enterValue('=SUM(A1:A5)', point('C1'))),
    check: (d) => has(d.cellXml('Formulas', 'C1'), '<f>SUM(A1:A5)</f>'),
    parts: CALC_CHAIN_DROP,
  },
  'editing.undo-redo': {
    fixture: 'basic-values',
    // Three edits, three undo steps, and two redo steps: the first two edits stay, the third
    // does not. Each step changes the result.
    edit: (t) => {
      t.run(t.editor.enterValue('11', point('C3')));
      t.run(t.editor.enterValue('22', point('C4')));
      t.run(t.editor.enterValue('33', point('C5')));
      t.run(t.editor.undo());
      t.run(t.editor.undo());
      t.run(t.editor.undo());
      t.run(t.editor.redo());
      t.run(t.editor.redo());
    },
    check: (d) =>
      all(
        has(d.cellXml('Values', 'C3'), '<v>11</v>'),
        has(d.cellXml('Values', 'C4'), '<v>22</v>'),
        lacks(d.cellXml('Values', 'C5'), '33')
      ),
    parts: [WORKBOOK],
  },
  'editing.sheet-headings': {
    fixture: 'basic-values',
    // The fixture hides its headings (`showRowColHeaders="0"`). View > Headings shows them, and
    // the default of the attribute is true, so the saved view has no attribute.
    edit: (t) => t.toolbar('view.headings'),
    check: (d) =>
      all(
        has(d.sheetXml('Values'), '<sheetView'),
        lacks(d.sheetXml('Values'), 'showRowColHeaders')
      ),
  },
  'editing.clear': {
    fixture: 'colors-fonts',
    // Clear formats keeps the text of B1; clear contents keeps the cell format of B2.
    edit: (t) => {
      t.select('B1');
      t.toolbar('clear.formats');
      t.select('B2');
      t.toolbar('clear.contents');
    },
    check: (d) =>
      all(
        text(d.cellText('Colors', 'B1'), 'Sample'),
        lacks(d.cellXml('Colors', 'B1'), ' s="'),
        has(d.cellXml('Colors', 'B2'), 's="2"'),
        lacks(d.cellXml('Colors', 'B2'), 'Sample')
      ),
  },
  'editing.font-style': formatEdit(
    (t) => {
      t.toolbar('font.bold');
      t.toolbar('font.italic');
      t.toolbar('font.underline');
    },
    (f) => has(f.font, '<b/>', '<i/>', '<u/>')
  ),
  'editing.font-name-size': formatEdit(
    (t) => {
      t.toolbar('font.family', 'Arial');
      t.toolbar('font.size', '14');
    },
    (f) => has(f.font, '<name val="Arial"/>', '<sz val="14"/>')
  ),
  'editing.font-color': formatEdit(
    (t) => t.toolbar('font.color', 'C00000'),
    (f) => has(f.font, '<color rgb="FFC00000"/>')
  ),
  'editing.fill': formatEdit(
    // A color that no fill of the fixture has, so the save writes a new fill.
    (t) => t.toolbar('fill.color', '7030A0'),
    (f) => has(f.fill, 'patternType="solid"', '<fgColor rgb="FF7030A0"/>')
  ),
  'editing.borders': formatEdit(
    (t) => t.toolbar('borders.outline'),
    (f) =>
      has(
        f.border,
        '<left style="thin">',
        '<right style="thin">',
        '<top style="thin">',
        '<bottom style="thin">'
      )
  ),
  'editing.alignment': formatEdit(
    (t) => {
      t.toolbar('alignment.center');
      t.toolbar('alignment.wrap');
    },
    (f) => has(f.xf, 'horizontal="center"', 'wrapText="1"')
  ),
  'editing.number-format': formatEdit(
    (t) => t.run(t.editor.applyFormat({ numberFormat: { id: 4 } })),
    (f) => (f.numFmtId === 4 ? undefined : `the format id is ${f.numFmtId}, not 4`)
  ),
  'editing.custom-number-format': formatEdit(
    (t) => t.toolbar('number.custom', '0.000;[Red]-0.000'),
    (f) =>
      f.numFmtCode === '0.000;[Red]-0.000' && f.numFmtId >= 164
        ? undefined
        : `the format is ${f.numFmtId} ${JSON.stringify(f.numFmtCode)}`
  ),
  'editing.number-format-list': {
    fixture: 'number-formats-more',
    edit: (t) => {
      t.select('B4');
      t.toolbar('number.format', 'scientific');
      t.select('B5');
      t.toolbar('number.format', 'currency');
    },
    check: (d) =>
      all(
        d.cellFormat('Formats', 'B4').numFmtId === 11 ? undefined : 'B4 is not built-in format 11',
        d.cellFormat('Formats', 'B5').numFmtCode?.includes('$')
          ? undefined
          : 'B5 has no currency code'
      ),
    parts: [STYLES],
  },
  'editing.decimal-places': {
    fixture: 'number-formats-more',
    // B4 shows `#,##0 ;(#,##0)` (built-in 37): one more place in each section.
    edit: (t) => {
      t.select('B4');
      t.toolbar('number.increaseDecimals');
    },
    check: (d) =>
      d.cellFormat('Formats', 'B4').numFmtCode === '#,##0.0 ;(#,##0.0)'
        ? undefined
        : `B4 has ${JSON.stringify(d.cellFormat('Formats', 'B4').numFmtCode)}`,
    parts: [STYLES],
  },
  'editing.clipboard': {
    fixture: 'clipboard-edit',
    // A copy of a row with a formula and cell formats, pasted two rows down.
    edit: (t) => {
      t.select('A2:D2');
      const copy = t.editor.copy();
      t.run(copy);
      if (!copy.ok) return;
      t.select('A6');
      t.run(t.editor.paste({ token: copy.content.token, text: copy.content.text }));
    },
    check: (d) =>
      all(
        text(d.cellText('Clip', 'A6'), 'Pens'),
        has(d.cellXml('Clip', 'D6'), '<f>B6*C6</f>', 's="2"')
      ),
    parts: [...SHARED_STRINGS, WORKBOOK],
  },
  'editing.resize': {
    fixture: 'resize-edit',
    edit: (t) => {
      t.run(t.editor.setColumnWidth(4, 4, 30));
      t.run(t.editor.setRowHeight(5, 5, 50));
    },
    check: (d) =>
      all(
        has(d.sheetXml('Sizes'), '<col min="5" max="5" width="30" customWidth="1"/>'),
        has(d.sheetXml('Sizes'), '<row r="6" ht="50" customHeight="1"')
      ),
  },
  'editing.sheets': {
    fixture: 'sheet-ops',
    // Add a sheet after Summary, rename the hidden sheet Data that a defined name and formulas
    // name, delete the hidden sheet Scratch, and move Summary to the end.
    edit: (t) => {
      t.run(t.editor.addSheet('Added'));
      t.run(t.editor.renameSheet(2, 'Inputs'));
      t.run(t.editor.deleteSheet(3));
      t.run(t.editor.moveSheet(0, 2));
    },
    check: (d) =>
      all(
        d.sheetNames.join(',') === 'Added,Inputs,Summary'
          ? undefined
          : `the sheets are ${d.sheetNames.join(', ')}`,
        has(d.partXml('/xl/workbook.xml'), 'Inputs!$B$2', 'SUM(Inputs!$B$2:$B$3)'),
        has(d.sheetXml('Summary'), '<f>Inputs!B2</f>', '<f>#REF!*2</f>')
      ),
    parts: [
      WORKBOOK,
      /^\/xl\/_rels\/workbook\.xml\.rels$/,
      /^\/\[Content_Types\]\.xml$/,
      /^\/xl\/worksheets\/sheet\d+\.xml$/,
    ],
  },
  'editing.merge-cells': {
    fixture: 'grid-structure',
    // Merge an empty range, and unmerge one that the file holds.
    edit: (t) => {
      t.select('E2:F3');
      t.run(t.editor.mergeCells());
      t.select('F12');
      t.run(t.editor.unmergeCells());
    },
    check: (d) =>
      all(
        has(d.sheetXml('Grid'), '<mergeCell ref="E2:F3"/>', '<mergeCell ref="B8:D10"/>'),
        lacks(d.sheetXml('Grid'), '<mergeCell ref="F12:F14"/>')
      ),
  },
  'editing.rows': {
    fixture: 'structure',
    // Insert a row above row 3: the cells below, the sums, the references of another sheet,
    // and the defined names move down.
    edit: (t) => {
      t.select('A3');
      t.run(t.editor.insertRows());
    },
    check: (d) =>
      all(
        text(d.cellText('Data', 'A4'), 'South'),
        has(d.cellXml('Data', 'B7'), '<f>SUM(B2:B6)</f>'),
        has(d.sheetXml('Refs'), '<f>SUM(Data!B2:C6)</f>'),
        has(d.partXml('/xl/workbook.xml'), definedName('Totals', 'Data!$D$2:$D$6'))
      ),
    parts: [WORKBOOK, /^\/xl\/worksheets\/sheet2\.xml$/, /^\/xl\/drawings\/drawing1\.xml$/],
  },
  'editing.columns': {
    fixture: 'structure',
    // Delete column B: the cells to its right move left, and its references become #REF!.
    edit: (t) => {
      t.select('B1');
      t.run(t.editor.deleteColumns());
    },
    check: (d) =>
      all(
        has(d.cellXml('Data', 'B1'), 'Q2'),
        has(d.cellXml('Data', 'B2'), '<v>20</v>'),
        has(d.sheetXml('Data'), '<f>SUM($B$2:$B$5)</f>', '<f>#REF!</f>'),
        has(
          d.partXml('/xl/workbook.xml'),
          definedName('Totals', 'Data!$C$2:$C$5'),
          definedName('Gone', 'Data!#REF!')
        )
      ),
    parts: [WORKBOOK, /^\/xl\/worksheets\/sheet2\.xml$/, /^\/xl\/drawings\/drawing1\.xml$/],
  },
  'save.preserved-content': {
    fixture: 'preserved-parts',
    // An edit beside the kept content: the extension data, the conditional formatting, the
    // hyperlinks, and the binary part stay as they were.
    edit: (t) => t.run(t.editor.enterValue('next', point('C1'))),
    check: (d) =>
      all(
        text(d.cellText('Kept', 'C1'), 'next'),
        has(
          d.sheetXml('Kept'),
          '<fx:kept value="1"/>',
          '<conditionalFormatting sqref="B1">',
          '<hyperlink ref="A1" r:id="rIdLink1"/>'
        ),
        d.pkg.parts.has('/xl/fixtureData.bin') ? undefined : 'the binary part is gone'
      ),
    parts: SHARED_STRINGS,
  },
  'save.hyperlinks': {
    fixture: 'structure',
    // Insert a row above the hyperlink of A5: the link moves to A6 with its cell.
    edit: (t) => {
      t.select('A2');
      t.run(t.editor.insertRows());
    },
    check: (d) =>
      all(
        has(d.sheetXml('Data'), '<hyperlink ref="A6" location="Refs!A1" display="West"/>'),
        has(d.partXml('/xl/workbook.xml'), definedName('Totals', 'Data!$D$3:$D$6'))
      ),
    parts: [WORKBOOK, /^\/xl\/worksheets\/sheet2\.xml$/, /^\/xl\/drawings\/drawing1\.xml$/],
  },
  'save.macro-enabled': {
    fixture: 'macro-enabled',
    // A value edit in a macro-enabled workbook. The save keeps the macro-enabled content type in
    // the `Override` element of the main part (ECMA-376 Part 2, 7.2.3.2.5, `CT_Override`). The
    // main part keeps the root element of the Workbook Part (Part 1, 12.3.23).
    edit: (t) => t.run(t.editor.enterValue('95', point('C3'))),
    check: (d) =>
      all(
        has(d.cellXml('Budget', 'C3'), '<v>95</v>'),
        isMacroEnabledType(contentTypeOf(d.pkg, workbookIndex(d.pkg)?.mainPart ?? ''))
          ? undefined
          : 'the main part is not macro-enabled'
      ),
    // The workbook has a formula, so the save asks for a full calculation on load (design D3).
    parts: [WORKBOOK],
  },
  'save.calc-chain': {
    fixture: 'formulas',
    edit: (t) => t.run(t.editor.enterValue('=A1+1', point('C1'))),
    check: (d) =>
      all(
        has(d.cellXml('Formulas', 'C1'), '<f>A1+1</f>'),
        d.pkg.parts.has('/xl/calcChain.xml') ? 'the calculation chain is kept' : undefined,
        lacks(d.partXml('/[Content_Types].xml'), 'calcChain')
      ),
    parts: CALC_CHAIN_DROP,
  },
  'save.full-calc-on-load': {
    fixture: 'wrapped-only-formula',
    // The only formula is inside `mc:AlternateContent`, so only the formula scan sets the flag.
    edit: (t) => t.run(t.editor.enterValue('7', point('B3'))),
    check: (d) => has(d.partXml('/xl/workbook.xml'), 'fullCalcOnLoad="1"'),
    parts: [WORKBOOK],
  },
  'drawings.pictures': keptDrawingEdit((d) =>
    all(
      has(d.partXml('/xl/drawings/drawing1.xml'), '<xdr:pic>'),
      d.pkg.parts.has('/xl/media/image1.png') ? undefined : 'the picture is gone'
    )
  ),
  'drawings.shapes': keptDrawingEdit((d) =>
    has(d.partXml('/xl/drawings/drawing1.xml'), '<xdr:sp macro="" textlink="">')
  ),
};

if (import.meta.main) {
  const { values } = parseArgs({ options: { out: { type: 'string' } } });
  if (values.out) mkdirSync(values.out, { recursive: true });
  let failed = 0;
  for (const [row, recipe] of Object.entries(FEATURE_EDITS)) {
    const run = await runFeatureEdit(recipe);
    if (run.problems.length > 0) failed++;
    console.log(
      `${row}: ${run.problems.length > 0 ? `FAIL ${run.problems.join('; ')}` : `ok, sheets: ${run.sheetCount}`}`
    );
    if (values.out && run.problems.length === 0)
      writeFileSync(join(values.out, featureCopyName(row, recipe.fixture)), run.bytes);
  }
  process.exitCode = failed ? 1 : 0;
}
