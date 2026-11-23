// The rules of `scripts/check-package-artifacts.mjs` that read one string or one manifest, so a
// unit test can hold each rule without a build.

// Static and dynamic imports, re-exports, side-effect imports, and requires. The request must be
// a whole package specifier, so a variable named `from` does not match. Minified output writes
// a side-effect import with no space: `import"name"`.
const BARE_REQUEST =
  /(?:\bfrom\s*|\bimport\s*\(\s*|\bimport\s*|\brequire\(\s*)(['"])((?:@[\w.-]+\/)?\w[\w.:-]*(?:\/[\w.@-]+)*)\1/g;
// A template literal with no substitution is a static request too.
const TEMPLATE_REQUEST =
  /\b(?:import|require)\s*\(\s*`((?:@[\w.-]+\/)?\w[\w.:-]*(?:\/[\w.@-]+)*)`/g;
// A triple-slash directive makes a declaration file read the types of a package.
const TYPE_REFERENCE =
  /^\s*\/\/\/\s*<reference\s+types\s*=\s*(['"])((?:@[\w.-]+\/)?\w[\w.:-]*(?:\/[\w.@-]+)*)\1/gm;

/**
 * @param {string} content The text of a built file.
 * @returns {string[]} Each bare package request in the file.
 */
export function bareImportRequests(content) {
  return [
    ...[...content.matchAll(BARE_REQUEST)].map((match) => match[2]),
    ...[...content.matchAll(TEMPLATE_REQUEST)].map((match) => match[1]),
    ...[...content.matchAll(TYPE_REFERENCE)].map((match) => match[2]),
  ];
}

// Static imports, re-exports, side-effect imports, and dynamic imports of a relative file.
const RELATIVE_REQUEST = /(?:\bfrom\s*|\bimport\s*\(\s*|\bimport\s*)(['"])(\.{1,2}\/[^'"]+)\1/g;

/**
 * @param {string} content The text of a built file.
 * @returns {string[]} Each relative ESM request in the file.
 */
export function relativeImportRequests(content) {
  return [...content.matchAll(RELATIVE_REQUEST)].map((match) => match[2]);
}

// DOCX Editor's core manifest has neither field, and its adapter and i18n manifests have both.
const NO_LISTING_FIELDS = new Set(['@xlsx-editor.dev/core']);

/**
 * @param {{ name: string, keywords?: unknown, homepage?: unknown }} manifest
 * @returns {string[]} Each listing field that the manifest leaves out.
 */
export function listingFieldProblems(manifest) {
  if (NO_LISTING_FIELDS.has(manifest.name)) return [];
  const problems = [];
  if (!Array.isArray(manifest.keywords) || manifest.keywords.length === 0)
    problems.push(`${manifest.name} has no keywords`);
  if (typeof manifest.homepage !== 'string' || !manifest.homepage.startsWith('https://'))
    problems.push(`${manifest.name} has no https homepage`);
  return problems;
}

/**
 * @param {string} request A bare package request.
 * @returns {string} The package name: `@scope/name` or `name`.
 */
export function packageNameOf(request) {
  const parts = request.split('/');
  return request.startsWith('@') ? parts.slice(0, 2).join('/') : parts[0];
}

/**
 * A request for a subpath that the package does not export fails on the consumer's disk.
 *
 * @param {string} request A bare package request.
 * @param {Map<string, { exports?: Record<string, unknown> }>} manifests The published manifests.
 * @returns {string | null} The problem, or `null` when another package owns the request or the
 * export map holds it.
 */
export function unexportedRequestProblem(request, manifests) {
  const name = packageNameOf(request);
  const manifest = manifests.get(name);
  if (!manifest) return null;
  const subpath = request === name ? '.' : `.${request.slice(name.length)}`;
  if (Object.hasOwn(manifest.exports ?? {}, subpath)) return null;
  return `${request}, which ${name} does not export`;
}

/**
 * A source path in a published declaration names a file that the tarball does not hold.
 *
 * @param {string} content The text of a built declaration file.
 * @returns {boolean} Whether the text holds a path into the source of a workspace package.
 */
export function exposesWorkspacePath(content) {
  return (
    /(?:\.\.\/)+core\/src\//.test(content) || /(?<![\w.-])packages\/[^/\s]+\/src\//.test(content)
  );
}

/**
 * A bundler drops an import that it reads as free of side effects, so a stylesheet that the
 * `sideEffects` field does not list leaves the consumer's build with no error.
 *
 * @param {{ name: string, sideEffects?: unknown, exports?: Record<string, unknown> }} manifest
 * @returns {string | null} The problem, or `null` when the field is the shipped value.
 */
export function sideEffectsProblem(manifest) {
  const hasStylesheet = Object.keys(manifest.exports ?? {}).some((key) => key.endsWith('.css'));
  const expected = hasStylesheet ? '["*.css"]' : 'false';
  const actual = JSON.stringify(manifest.sideEffects);
  if (actual === expected) return null;
  return `${manifest.name} sets sideEffects to ${actual}, expected ${expected}`;
}

/**
 * The fix for a required file that a tarball leaves out. Git ignores the notices, so a script
 * writes them after the build. A file that is on disk but not in the tarball is missing from
 * `files`. A tracked file, such as the README or a license, that is not on disk is restored.
 *
 * @param {string} file The required path, relative to the package root.
 * @param {boolean} onDisk Whether the file is in the package directory.
 * @param {string} packageDir The directory name of the package under `packages/`.
 * @returns {string} The fix, with a leading space. A tracked file that is not on disk is restored
 *   from git.
 */
export function missingFileRemedy(file, onDisk, packageDir) {
  if (onDisk) return ` Add it to \`files\` in packages/${packageDir}/package.json.`;
  if (file === 'THIRD_PARTY_NOTICES.md')
    return ' Run `bun run notices:generate` after `bun run build:packages`.';
  return ` Restore packages/${packageDir}/${file} from git.`;
}

/**
 * The `licenses/` entries that a tarball must hold. npm never publishes system files such as
 * `.DS_Store`, and no notice is a dotfile, so a dotfile on disk is not required.
 *
 * @param {string[]} names The file names in the package's `licenses/` directory.
 * @returns {string[]} The required paths, relative to the package root.
 */
export function requiredLicenseFiles(names) {
  return names.filter((name) => !name.startsWith('.')).map((name) => `licenses/${name}`);
}

// The Ecma copyright line inside one legal comment: the match may not cross a `*/`.
const ECMA_LEGAL_COMMENT = /\/\*!(?:(?!\*\/)[^])*?© Ecma International(?:(?!\*\/)[^])*?\*\//;

/**
 * @param {string} content The text of a built file.
 * @returns {boolean} Whether one legal comment (`/*!`) in the file holds the Ecma copyright line.
 */
export function hasEcmaLegalComment(content) {
  return ECMA_LEGAL_COMMENT.test(content);
}
