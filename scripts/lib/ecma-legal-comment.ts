// Copied from docx-editor/scripts/ecma-legal-comment.mjs at 35a29d6d (EigenPal), which holds the
// `/*!` block of pptx-editor/scripts/geometry/build-presets.ts (PPTX Editor at d4502552). Differences:
// it is TypeScript in `scripts/lib/` and finds the NOTICE with Bun's `import.meta.dir`, because
// the two generators that import it are TypeScript scripts that Bun runs, as in PPTX Editor; it
// returns no final newline, because those generators (the preset geometry and the predefined
// function names) place the comment in their own line layout; and its comment names only the
// bundlers that keep legal comments.
//
// The Ecma notice must travel with every copy of a generated table. A legal comment (`/*!`)
// stays in the built package, and bundlers that keep legal comments, such as esbuild and Bun,
// keep it in an app bundle. A minified Vite build removes it (reference/README.md, "Terms").
import { readFileSync } from 'node:fs';
import { join } from 'node:path';

/** `text` wrapped at `width` columns, each line prefixed for a block comment. */
function commentLines(text: string, width = 96): string {
  const out: string[] = [];
  for (const paragraph of text.split('\n')) {
    let line = '';
    for (const word of paragraph.split(/\s+/).filter(Boolean)) {
      if (line && line.length + word.length + 1 > width) {
        out.push(line);
        line = word;
      } else line = line ? `${line} ${word}` : word;
    }
    out.push(line);
  }
  return out.map((line) => (line ? ` * ${line}` : ' *')).join('\n');
}

/** A `/*!` comment: `preamble`, then the full Ecma notice from `reference/ecma-376/NOTICE`. */
export function ecmaLegalComment(preamble: string): string {
  const notice = readFileSync(join(import.meta.dir, '../../reference/ecma-376/NOTICE'), 'utf8');
  const start = notice.indexOf('COPYRIGHT NOTICE');
  if (start < 0) throw new Error('reference/ecma-376/NOTICE has no COPYRIGHT NOTICE heading.');
  const text = notice.slice(start).trim();
  if (text.includes('*/')) throw new Error('The notice cannot hold a comment terminator.');
  return `/*!\n${commentLines(preamble)}\n *\n${commentLines(text)}\n */`;
}
