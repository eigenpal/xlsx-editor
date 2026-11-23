// The demo's sample workbook: one sheet per group of features that XLSX Editor shows and
// edits. The demo opens it at start. Output: examples/vite/public/sample.xlsx.
//
//   bun run sample:build
//
// Every feature that the sample uses is a supported or partial row of the support matrix
// (`docs/site/data/support-matrix.ts`). It has no conditional formatting, chart, table, pivot
// table, data validation, external link, or macro. All content is written here and made up;
// the only picture is generated (showcase-image.ts).
//
// Same role as pptx-editor/scripts/sample-deck/showcase.ts at 79a5d22d (EigenPal), for a workbook.
// Difference: PPTX Editor's demo has two decks, a short `sample.pptx` and the feature tour
// `showcase.pptx`. This demo has one workbook, the feature tour, so the showcase builder writes
// `sample.xlsx`, the file that the demo opens at start.

import { buildXlsx, type FixtureWorkbook } from '../fixtures/builder.ts';
import { StringTable, StyleTable } from './showcase-kit.ts';
import { buildFormatsSheet, buildSalesSheet, buildWelcomeSheet } from './showcase-sheets.ts';
import { buildLayoutSheet, buildScheduleSheet, buildTextSheet } from './showcase-layout.ts';

/** The default font of the sample: `x:font` children. */
const SAMPLE_FONT =
  '<sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/>';

/** The sample as fixture input. */
function sampleWorkbook(): FixtureWorkbook {
  const styles = new StyleTable();
  const strings = new StringTable();
  const sheets = [
    buildWelcomeSheet,
    buildSalesSheet,
    buildFormatsSheet,
    buildTextSheet,
    buildLayoutSheet,
    buildScheduleSheet,
  ].map((sheet) => sheet(styles, strings));
  // The sample keeps its own fonts: the fixture builder's defaults are the open faces of the
  // reference PDFs, and the sample is not a fixture.
  return {
    sheets,
    rawStrings: strings.items,
    styles: styles.styles(),
    defaultFont: SAMPLE_FONT,
    theme: { minor: 'Calibri', major: 'Calibri' },
  };
}

/** The bytes of `examples/vite/public/sample.xlsx`. */
export function buildSample(): Uint8Array {
  return buildXlsx(sampleWorkbook());
}
