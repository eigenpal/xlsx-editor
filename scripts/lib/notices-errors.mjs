// The texts of the module load, thrown error, write, and delete failures of
// `scripts/generate-third-party-notices.mjs`, kept here so unit tests can check the fix that each
// failure names. The script keeps its other error texts, which its own test runs. DOCX Editor and PPTX Editor keep these texts in the script;
// this repository adds the fixes, so it tests them (scripts/__tests__/notices-errors.test.ts).

/** `message` with one final period. */
export function withPeriod(message) {
  return message.replace(/\.?$/, '.');
}

/** The message of a thrown value, which need not be an `Error`. */
function messageOf(error) {
  return String(error?.message ?? error);
}

/**
 * A notice module that failed to load. The module, or a tracked file that it reads when it
 * loads, is broken: an edit can be fixed, and a damaged file can be restored from git.
 *
 * @param {string} file The module, relative to the repository root.
 * @param {string[]} reads The tracked files that the module reads when it loads.
 * @param {unknown} error The load error.
 * @returns {string} The message, with the files to fix.
 */
export function moduleLoadError(file, reads, error) {
  const files = [file, ...reads].map((name) => `\`${name}\``).join(' or ');
  return (
    `${file} failed to load: ${withPeriod(messageOf(error))} Fix ${files}, or restore them ` +
    'from git, then run this script again.'
  );
}

/**
 * An error that one package's notices threw. The HarfBuzz notice errors name their own fix. Any
 * other error is a missing install or build output, or a broken tracked file.
 *
 * @param {string} name The package name.
 * @param {unknown} error The thrown error.
 * @returns {string} The message, with its fix.
 */
export function packageError(name, error) {
  const message = `${name}: ${withPeriod(messageOf(error))}`;
  if (messageOf(error).startsWith('HarfBuzz')) return message;
  return (
    `${message} Run \`bun install\` and \`bun run build:packages\`, then run this script again. ` +
    'If the error names a tracked file, fix it or restore it from git.'
  );
}

/**
 * A notice file that could not be written.
 *
 * @param {string} file The notice file, relative to the repository root.
 * @param {unknown} error The write error.
 * @returns {string} The message, with its fix.
 */
export function writeError(file, error) {
  return (
    `Could not write ${file}: ${withPeriod(messageOf(error))} Make the file and its directory ` +
    'writable, then run this script again.'
  );
}

/**
 * A notice file that a failed run could not delete.
 *
 * @param {string} file The notice file, relative to the repository root.
 * @param {unknown} error The delete error.
 * @returns {string} The message, with its fix.
 */
export function deleteError(file, error) {
  return (
    `Could not delete ${file}: ${withPeriod(messageOf(error))} Delete it, then run this script ` +
    'again. Until then, the file can name licenses that the build does not ship.'
  );
}
