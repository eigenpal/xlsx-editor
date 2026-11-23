// Print the painted page of the fidelity harness to PDF (`run.ts`), as PPTX Editor's
// `printDeck` prints its harness page: explicit calls of the page around `page.pdf`, in
// try/finally. Chromium keeps the glyphs of a canvas as PDF text only when the canvas is painted
// during `beforeprint` (`page/main.ts`), so `harnessPrint` arms that paint and `harnessEndPrint`
// reports how often it ran. A PDF is refused unless it ran exactly once: no paint means the PDF
// holds an earlier paint as an image, and two mean that the PDF can hold either.

/** The part of a Playwright page that printing uses. */
export interface PrintablePage {
  evaluate<R>(pageFunction: () => R | Promise<R>): Promise<R>;
  pdf(options: { preferCSSPageSize: boolean; printBackground: boolean }): Promise<Uint8Array>;
}

/** Print the rendered fixture of the harness page, and return the PDF bytes. */
export async function printPainted(page: PrintablePage): Promise<Uint8Array> {
  if (!(await page.evaluate(() => window.harnessPrint())))
    throw new Error(
      'the page has no rendered fixture to print; call `harnessRender` and check that it returns ok first'
    );
  let pdf: Uint8Array;
  let paints = 0;
  try {
    pdf = await page.pdf({ preferCSSPageSize: true, printBackground: true });
  } finally {
    paints = await page.evaluate(() => window.harnessEndPrint());
  }
  if (paints !== 1)
    throw new Error(
      `the print paint ran ${paints} times while Chromium printed, not once; arm the paint with \`harnessPrint\` and paint once in the \`beforeprint\` listener (\`page/main.ts\`)`
    );
  return new Uint8Array(pdf);
}
