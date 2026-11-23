// The naming gate of AGENTS.md "Naming": committed text never names a proprietary office app, a
// vendor of one, or a trademarked feature name. After PPTX Editor's
// scripts/check-public-terminology.mjs at 8dc88c4d, with three changes: it reads every file that
// git tracks or would track (not a fixed list of folders), it takes exact, reviewed exceptions
// from scripts/public-terminology-allow.json for names that the file format defines, and it
// passes the exact npm package names of the build tools (NPM_PACKAGES).
//
// Compressed fixtures (`.xlsx`) and images are not read. Fixtures are built by scripts, which the
// gate reads.
//
// A pass does not prove the rule. Names that are also common words (the word processor's name,
// for example) cannot be matched safely, so read your diff by hand.
import { execFileSync } from 'node:child_process';
import { existsSync, lstatSync, readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = dirname(dirname(fileURLToPath(import.meta.url)));

// Build the names from parts, so this file does not hold them as text.
const join2 = (...parts) => parts.join('');
// Spaces (also no-break spaces), hyphens, and underscores between the parts of a name.
const SEP = '(?:[\\s_\\u2010-\\u2011-]|%20)?';
const INSENSITIVE = new RegExp(
  [
    // Inside identifiers too (`is<Name>Compat`), but not the common word that starts with it.
    `${join2('ex', 'cel')}(?!l|s\\b|sior)`,
    join2('power', SEP, 'point'),
    `${join2('power', SEP)}(query|bi|pivot)\\b`,
    // Only the namespace URIs that the format defines may name the vendor host.
    `(?<!schemas[.-])${join2('micro', SEP, 'soft')}`,
    `\\b${join2('ms', 'ft')}\\b`,
    `\\b${join2('ms', SEP, 'office')}`,
    // Vendor protocol and specification IDs, docs hosts, and products.
    `(?<![-\\w])${join2('ms', '-')}(oi\\d+|oe\\d+|xlsx|xlsb|xls|ovba|cfb|offcrypto|odata)\\b`,
    `\\b${join2('ms', 'dn')}\\b`,
    `\\b(support|templates)\\.${join2('off', 'ice')}\\.com`,
    join2('office', '\\.live\\.com'),
    `\\b${join2('aka', '\\.ms')}\\b`,
    join2('sheets', '\\.google\\.com'),
    `${join2('office', SEP)}(scripts|97|xp)\\b`,
    join2('one', SEP, 'drive'),
    join2('share', SEP, 'point'),
    `(star|neo)${join2('off', 'ice')}`,
    `${join2('colla', 'bora')}\\b`,
    `\\b${join2('soft', 'maker')}`,
    `${join2('office', SEP, '(365|20\\d\\d)')}`,
    `\\b[om]365\\b`,
    `${join2('google', SEP)}(sheets?|spreadsheets?)\\b`,
    join2('docs', '\\.google\\.com'),
    `${join2('google', SEP)}(docs|slides|workspace|drive)\\b`,
    `(?<![a-z])${join2('i', 'work')}(?![a-z])`,
    join2('quat', 'tro'),
    `\\b${join2('wps', SEP)}spreadsheets?\\b`,
    join2('smart', 'sheet'),
    `\\b${join2('zo', 'ho')}\\b`,
    join2('libre', SEP, 'office'),
    join2('open', SEP, 'office'),
    `\\b${join2('only', '[-_]?', 'office')}`,
    `\\b${join2('s', 'office')}\\b`,
    `${join2('apple', "('s)?", SEP, 'numbers')}`,
    `\\b${join2('g', 'numeric')}`,
    `\\b${join2('wps', SEP, 'office')}\\b`,
    `\\b${join2('king', 'soft')}`,
    `${join2('lo', 'tus')}${SEP}(1-?2-?3|notes|symphony)`,
    join2('visual', SEP, 'basic'),
  ].join('|'),
  'i'
);
// The control name, case-sensitive, so the identifier `activeX` and the words "active x" pass.
const SENSITIVE = new RegExp(
  `${join2('Active', 'X')}|${join2('ACTIVE', 'X')}|${join2('Active', ' X')}\\b`
);
// Namespace URIs that the format defines may name the vendor, its apps, and its controls. They are
// removed before the test, so the rest of the line is still checked.
const NAMESPACE = new RegExp(
  `urn:schemas-${join2('micro', 'soft')}-com:[\\w:.-]*|https?://schemas\\.${join2('micro', 'soft')}\\.com/[\\w/.-]*`,
  'gi'
);
// npm package names that the build tools depend on (AGENTS.md, "Naming": npm package names are an
// exception). Exact names only, so a lockfile upgrade needs no allowlist entry and prose stays checked.
const NPM_PACKAGES = new RegExp(
  `@${join2('micro', 'soft')}/(?:api-extractor(?:-model)?|tsdoc(?:-config)?)(?![\\w-])`,
  'g'
);
const FORBIDDEN = {
  test: (text) => {
    const rest = text.replace(NAMESPACE, '').replace(NPM_PACKAGES, '');
    return INSENSITIVE.test(rest) || SENSITIVE.test(rest);
  },
};

// Verbatim third-party material (the Ecma files and the legal texts), and the allowlist itself.
const SKIPPED = [
  /^reference\/ecma-376\//,
  /^LICENSE$/,
  /^CLA\.md$/,
  /(^|\/)LICENSE\.md$/,
  /^scripts\/public-terminology-allow\.json$/,
];
// Every file is read as UTF-8 text (a PDF as Latin-1), except these compressed or image formats. A PDF is read
// too, so its producer and title metadata are checked.
const BINARY =
  /\.(png|jpe?g|gif|webp|ico|xlsx|xlsm|xltx|docx|pptx|zip|gz|tgz|woff2?|ttf|otf|wasm)$/i;

const pathViolations = [];
let allowed;
const used = new Set();

function listFiles() {
  const out = execFileSync(
    'git',
    ['ls-files', '--cached', '--others', '--exclude-standard', '-z'],
    {
      cwd: root,
      encoding: 'utf8',
    }
  );
  return (
    out
      .split('\0')
      .filter((file) => file && !SKIPPED.some((pattern) => pattern.test(file)))
      .filter((file) => {
        // A file name is committed text too, whatever the file type.
        // A part path that the format defines goes in the allowlist with the line `(file name)`.
        if (FORBIDDEN.test(file)) {
          const key = `${file}\0(file name)`;
          if (allowed.has(key)) used.add(key);
          else pathViolations.push(`${file} (file name)`);
        }
        return !BINARY.test(file);
      })
      .filter((file) => !SKIPPED.some((pattern) => pattern.test(file)))
      // A symbolic link or a directory is not text to read.
      .filter((file) => existsSync(join(root, file)) && lstatSync(join(root, file)).isFile())
  );
}

function loadAllowed() {
  const path = join(root, 'scripts', 'public-terminology-allow.json');
  if (!existsSync(path)) return new Set();
  const entries = JSON.parse(readFileSync(path, 'utf8')).allowed ?? [];
  return new Set(entries.map((entry) => `${entry.file}\0${entry.line}`));
}

allowed = loadAllowed();
const files = listFiles();
const violations = [...pathViolations];
for (const file of files) {
  readFileSync(join(root, file), /\.pdf$/i.test(file) ? 'latin1' : 'utf8')
    .split('\n')
    .forEach((line, index) => {
      if (!FORBIDDEN.test(line)) return;
      const key = `${file}\0${line.trim()}`;
      if (allowed.has(key)) used.add(key);
      else violations.push(`${file}:${index + 1}`);
    });
}
const stale = [...allowed].filter((key) => !used.has(key)).map((key) => key.replace('\0', ': '));

if (violations.length || stale.length) {
  if (violations.length)
    console.error(`Public terminology check failed:\n${violations.join('\n')}`);
  if (stale.length)
    console.error(`Stale entries in scripts/public-terminology-allow.json:\n${stale.join('\n')}`);
  process.exitCode = 1;
} else {
  console.info(`Public terminology check passed (${files.length} files).`);
}
