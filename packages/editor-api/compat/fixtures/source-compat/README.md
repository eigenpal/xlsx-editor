# Office.js compatibility fixtures

These files hold Office.js workbook code, written against the `XlsxEditor` namespace, to check API compatibility at compile time. They do not run, and they do not ship in the package.

EigenPal wrote every line. The fixtures follow the call patterns of the Office.js documentation: `load` and `context.sync()`, `getItem` and `getItemOrNullObject`, `getRange`, and two-dimensional `values` and `numberFormat` arrays. Unlike the PPTX Editor fixtures, they copy no published samples, so this directory has no third-party license file.

## Adaptations

The fixtures use `runtime.run` for batches instead of a host `run` function, because the package has no global host. They add exports, type annotations, and local test values. They read `numberFormat` and do not set it, because the package publishes it read-only.
