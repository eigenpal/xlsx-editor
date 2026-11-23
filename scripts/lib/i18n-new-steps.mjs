// The steps that `bun run i18n:new <lang>` prints. They follow "Add a locale" in
// packages/i18n/README.md, and scripts/__tests__/i18n-new-steps.test.ts keeps the two equal.

/**
 * Lists what a contributor does after `i18n:new` writes the locale files.
 *
 * @param {string} lang The BCP 47 tag of the new locale.
 * @returns {string[]} The steps, in order.
 */
export function newLocaleSteps(lang) {
  return [
    `Open packages/i18n/${lang}.json`,
    'Replace null values with translations (null keys fall back to English)',
    `Add '${lang}' to the entries of @xlsx-editor.dev/i18n in scripts/lib-packages.mjs`,
    'Run: bun run build:packages',
    'Run: node scripts/api-extractor.mjs --local --record-forgotten-exports',
    'Run: bun run i18n:status',
    'Commit the locale files, scripts/lib-packages.mjs, scripts/api-forgotten-exports.baseline.json, and docs/api/i18n/, and open a pull request',
  ];
}
