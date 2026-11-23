// Reads the links of the READMEs and guides that npm installed, so `check:consumer-install` proves
// that each link reaches a file and a heading in the published packages. A link to another file of
// the repository, such as a docs site page, resolves in the source checkout.
import { existsSync, readFileSync, readdirSync, statSync } from 'node:fs';
import path from 'node:path';

const REPOSITORY = 'https://github.com/eigenpal/xlsx-editor/blob/main/';

// The text outside fenced code, with each code span replaced, so that code is not read as a link.
function proseOf(markdown) {
  let fence = null;
  const lines = [];
  for (const line of markdown.split(/\r?\n/)) {
    const marker = line.match(/^\s*(`{3,}|~{3,})/);
    if (fence) {
      if (marker && marker[1][0] === fence[0] && marker[1].length >= fence.length) fence = null;
    } else if (marker) {
      fence = marker[1];
    } else {
      lines.push(line);
    }
  }
  return lines;
}

/**
 * @param {string} markdown The text of a Markdown file.
 * @returns {Set<string>} The anchor of each heading, with the suffix of a repeated heading.
 */
export function headingAnchors(markdown) {
  const anchors = new Set();
  const seen = new Map();
  for (const line of proseOf(markdown)) {
    const heading = line.match(/^#{1,6}\s+(.+?)\s*#*\s*$/);
    if (!heading) continue;
    const base = heading[1]
      .replaceAll('`', '')
      .toLowerCase()
      .replace(/[^\p{L}\p{M}\p{N}\p{Pc}\- ]/gu, '')
      .replaceAll(' ', '-');
    const count = seen.get(base) ?? 0;
    seen.set(base, count + 1);
    anchors.add(count === 0 ? base : `${base}-${count}`);
  }
  return anchors;
}

/**
 * @param {string} markdown The text of a Markdown file.
 * @returns {string[]} The target of each inline link.
 */
export function markdownLinks(markdown) {
  const prose = proseOf(markdown)
    .join('\n')
    .replace(/(`+)[\s\S]*?\1/g, 'CODE');
  return [...prose.matchAll(/\[[^\]]*\]\(([^)\s]+)(?:\s+"[^"]*")?\)/g)].map((match) => match[1]);
}

function markdownFiles(directory) {
  if (!existsSync(directory)) return [];
  return readdirSync(directory, { withFileTypes: true }).flatMap((entry) => {
    const file = path.join(directory, entry.name);
    if (entry.isDirectory()) return markdownFiles(file);
    return entry.name.endsWith('.md') ? [file] : [];
  });
}

// The installed file that a link names, or the reason that it names none. A link to this
// repository on GitHub drops its query, such as `?plain=1`, and resolves its `.` and `..`
// segments first, as a browser does, so `docs/../packages/core/` names the installed package.
function linkTarget(target, file, packageRoot, roots, repositoryRoot) {
  if (target.startsWith(REPOSITORY)) {
    const repositoryPath = path.posix.normalize(target.slice(REPOSITORY.length).split('?')[0]);
    if (repositoryPath.startsWith('packages/')) {
      const [slug, ...rest] = repositoryPath.slice('packages/'.length).split('/');
      const root = roots.get(slug);
      if (!root) return { problem: `names no published package` };
      return { root, file: path.join(root, rest.join('/')) };
    }
    const resolved = path.resolve(repositoryRoot, repositoryPath);
    const inside = path.relative(repositoryRoot, resolved);
    if (inside.startsWith('..') || path.isAbsolute(inside)) {
      return { problem: 'leaves the repository' };
    }
    return { root: repositoryRoot, file: resolved };
  }
  if (/^[a-z][a-z\d+.-]*:/i.test(target)) return null;
  const resolved = path.resolve(path.dirname(file), target);
  const inside = path.relative(packageRoot, resolved);
  if (inside.startsWith('..') || path.isAbsolute(inside)) {
    return { problem: 'leaves the package' };
  }
  return { root: packageRoot, file: resolved };
}

// The problems of the links of one file. Each found Markdown page of the repository outside the
// packages goes to `pages`. With `repositoryOnly`, the file is such a page: it links other site
// pages by their site path, so only its links to this repository on GitHub are read.
function linkProblems(file, name, context, pages, repositoryOnly) {
  const { packageRoot, roots, repositoryRoot } = context;
  const problems = [];
  const found = [];
  for (const link of markdownLinks(readFileSync(file, 'utf8'))) {
    if (repositoryOnly && !link.startsWith(REPOSITORY)) continue;
    const [target, anchor] = link.split('#');
    const at =
      target === '' ? { file } : linkTarget(target, file, packageRoot, roots, repositoryRoot);
    if (at === null) continue;
    if (at.problem) {
      problems.push(`${name}: the link ${link} ${at.problem}`);
      continue;
    }
    if (!existsSync(at.file) || !statSync(at.file).isFile()) {
      const where = at.root === repositoryRoot ? 'repository' : 'installed';
      problems.push(`${name}: the link ${link} names no ${where} file`);
      continue;
    }
    found.push(at.file);
    if (at.root === repositoryRoot && /\.mdx?$/.test(at.file)) pages.add(at.file);
    if (anchor && !headingAnchors(readFileSync(at.file, 'utf8')).has(anchor)) {
      problems.push(`${name}: the link ${link} names no heading`);
    }
  }
  return { problems, found };
}

/**
 * @param {Map<string, string>} roots The installed directory of each package, by its folder name
 * in `packages/`.
 * @param {string} repositoryRoot The source checkout. A link to a repository file outside
 * `packages/`, such as a docs site page, resolves in it. The links of such a page to this
 * repository on GitHub are read too, one level deep. A link of such a page into `packages/`
 * resolves in the installed package, as in a guide. A file that exists in the checkout passes,
 * also when git ignores it.
 * @returns {string[]} Each link that does not resolve, and each guide that the README omits.
 */
export function installedDocProblems(roots, repositoryRoot) {
  const problems = [];
  const pages = new Set();
  for (const packageRoot of roots.values()) {
    const context = { packageRoot, roots, repositoryRoot };
    const readme = path.join(packageRoot, 'README.md');
    const guides = markdownFiles(path.join(packageRoot, 'docs'));
    const linked = new Set();
    for (const file of [readme, ...guides]) {
      const name = path.relative(path.dirname(packageRoot), file);
      if (!existsSync(file)) {
        problems.push(`${name}: the file is not in the package`);
        continue;
      }
      const result = linkProblems(file, name, context, pages, false);
      problems.push(...result.problems);
      if (file === readme) for (const target of result.found) linked.add(target);
    }
    for (const guide of guides) {
      if (!linked.has(guide)) {
        const name = path.relative(path.dirname(packageRoot), guide);
        problems.push(`${name}: the package README has no link to this guide`);
      }
    }
  }
  const context = { packageRoot: repositoryRoot, roots, repositoryRoot };
  for (const page of pages) {
    const name = path.relative(repositoryRoot, page);
    problems.push(...linkProblems(page, name, context, new Set(), true).problems);
  }
  return problems;
}
