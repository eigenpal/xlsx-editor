<p align="center">
  <a href="https://www.xlsx-editor.dev/">
    <picture>
      <source media="(prefers-color-scheme: dark)" srcset="https://raw.githubusercontent.com/eigenpal/xlsx-editor/main/.github/assets/readme-logo-dark.svg" />
      <source media="(prefers-color-scheme: light)" srcset="https://raw.githubusercontent.com/eigenpal/xlsx-editor/main/.github/assets/readme-logo-light.svg" />
      <img src="https://raw.githubusercontent.com/eigenpal/xlsx-editor/main/.github/assets/readme-logo-light.svg" alt="XLSX Editor by EigenPal" width="320" height="90" />
    </picture>
  </a>
</p>

# @xlsx-editor.dev/editor-api

Read and write an `.xlsx` workbook from code. The object model has the shape of the Office.js workbook object model: `Workbook`, `WorksheetCollection`, `Worksheet`, and `Range`. The package runs on the XLSX Editor engine, on a server, from XLSX bytes.

The package is not Office.js and does not run in an add-in host. It implements a small, documented subset of that object model, and every type in it is authored here. It uses the EigenPal Pro License. See [License](#license).

## Install

The package is private: it is not published to npm. Import it from a workspace of this repository, after you build the packages with `bun run build:packages`.

For task guides and API details, see the [Editing API](https://xlsx-editor.dev/docs/editor-api) page of the developer documentation.

## Edit a workbook on a server

The following example reads two cells and writes a header row. It also writes a column that depends on the two cells, then saves:

```ts
import { readFile, writeFile } from 'node:fs/promises';
import { XlsxEditor } from '@xlsx-editor.dev/editor-api';

const runtime = await XlsxEditor.createServer(await readFile('book.xlsx'));
try {
  await runtime.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem('Data');
    const units = sheet.getRange('B2:B3').load('values');
    await context.sync();
    sheet.getRange('A1:C1').values = [['Region', 'Units', 'Checked']];
    const checked = units.values.map(([value]) => [value !== '']);
    sheet.getRange('C2:C3').values = checked;
    await context.sync();
  });
  await writeFile('edited.xlsx', await runtime.save());
} finally {
  runtime.dispose();
}
```

## What the subset does

- Calls on proxy objects are queued. Nothing reaches the workbook until `await context.sync()`, and a sync is atomic: all of its calls apply, or none of them.
- A sync reads the workbook as it was before the sync. A load of `values` or `numberFormat` after a write in the same sync is refused. Sync the writes, then load.
- A read gives the value that the file stores. The API does not calculate formulas, so it refuses a `values` read of a range that holds a formula cell.
- A write takes a finite number, a boolean, text, `""` to empty a cell, or `null` to leave it. A value for a formula cell, a formula, and text that reads as another value are refused with `NotSupported`. `numberFormat` is read-only.
- The API refuses a write back of loaded values when a loaded string reads as another value. Examples are `#DIV/0!`, `-`, `TRUE`, and `00123`. Give `null` for each cell that you do not change.
- The value writes of one sync reach at most 10,000 cells. Split a larger range into parts, and sync after each part.

## Errors

Each refusal is an `XlsxEditorError` with a stable `code`. Test for one with `isXlsxEditorError(error)`, then branch on its `code`, never on its message. A refused sync applies nothing of its batch. The [Errors](https://xlsx-editor.dev/docs/editor-api#errors) section of the developer documentation lists each code and its cause.

## Office.js compatibility

`compat/manifest.json` lists each supported member, and `compat/runtime-notes.json` records how each one behaves where the subset differs. Every other member of the Office.js object model is absent from the types. The subset is not full Office.js compatibility. `compat/README.md` describes the files and the update process.

## License

This package uses the EigenPal Pro License. See [LICENSE.md](LICENSE.md). Evaluation is permitted. Production use requires a commercial agreement. Contact [licensing@eigenpal.com](mailto:licensing@eigenpal.com).
