// Copied from the `configObject` in pptx-editor/scripts/api-extractor.mjs at e22f3f3b (EigenPal).
// Differences: the object is a function in its own module, and `scripts/api-extractor.mjs` passes
// the paths, which resolve from the repository root; `apiReport.includeForgottenExports` is true;
// and `extractorMessageReporting` has no `ae-forgotten-export` key, because its `default` level
// `none` covers it. The compiler options and the other message levels are equal to the source.
//
// The API Extractor configuration of one published entry. `includeForgottenExports` writes each
// type that a public signature reaches and the entry does not export into the report, so
// `api:check` fails when the shape of such a type changes.

/**
 * @param {{ root: string, dts: string, reportFileName: string, reportFolder: string,
 *   reportTempFolder: string }} options
 * @returns {object} The `configObject` of `ExtractorConfig.prepare`.
 */
export function apiReportConfigObject({
  root,
  dts,
  reportFileName,
  reportFolder,
  reportTempFolder,
}) {
  return {
    projectFolder: root,
    mainEntryPointFilePath: dts,
    compiler: {
      overrideTsconfig: { compilerOptions: { skipLibCheck: true, lib: ['ES2022', 'DOM'] } },
    },
    apiReport: {
      enabled: true,
      reportFileName,
      reportFolder,
      reportTempFolder,
      includeForgottenExports: true,
    },
    docModel: { enabled: false },
    dtsRollup: { enabled: false },
    tsdocMetadata: { enabled: false },
    messages: {
      extractorMessageReporting: {
        // `scripts/api-extractor.mjs` judges forgotten exports against its baseline. A warning
        // level here would fail each run, because Extractor counts the warning before the callback.
        default: { logLevel: 'none' },
      },
      compilerMessageReporting: { default: { logLevel: 'warning' } },
      tsdocMessageReporting: { default: { logLevel: 'none' } },
    },
  };
}
