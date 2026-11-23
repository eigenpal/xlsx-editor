#!/usr/bin/env node
// Copied from docx-editor/scripts/check-shipped-shaper.mjs at c6529c8b (EigenPal). Differences:
// core ships ESM only, so the check of the CommonJS chunk is left out; the failure message drops
// the DOCX Editor issue number; and the file list reads only the
// `.js` chunks.
//
// Assert the SHIPPED bundle, not the config that produced it.
//
//   bun run check:shipped-shaper     # after `bun run build:packages`
//
// A past DOCX Editor defect was a property of core's OUTPUT: a bare `module` specifier reaching a
// consumer's bundler. Every guard added with the fix checks an input — the tsup config
// object, the vendor's source, the plugin's own application count. None of them reads
// `dist/`, so the one thing a consumer actually installs was the one thing nothing
// asserted. These are the greps that would have caught the bug, and would catch a
// regression in a dependency bump, an esbuild upgrade, or a config refactor that still
// type-checks.
import { existsSync, readFileSync, readdirSync, statSync } from 'node:fs';
import path from 'node:path';

const root = path.resolve(import.meta.dirname, '..');
const dist = path.join(root, 'packages/core/dist');
const failures = [];

if (!existsSync(dist)) {
  console.error('packages/core/dist is missing — run `bun run build:packages` first.');
  process.exit(1);
}

const esm = readdirSync(dist).filter((file) => file.endsWith('.js'));
const read = (file) => readFileSync(path.join(dist, file), 'utf8');

// 1. The bug itself. A browser bundler cannot resolve Node's `module`, and the guard that
//    would skip it is a runtime value, so the specifier must not survive at all.
const MODULE_SPECIFIER = /(?:from\s*["']module["']|import\(\s*["']module["']\s*\))/;
for (const file of esm) {
  if (MODULE_SPECIFIER.test(read(file))) {
    failures.push(`${file} still imports Node's \`module\`; browser builds will fail`);
  }
}

// 2. The binary the inlined loader looks for, beside the bundle that looks for it.
const wasm = path.join(dist, 'harfbuzz.wasm');
if (!existsSync(wasm) || statSync(wasm).size === 0) {
  failures.push('dist/harfbuzz.wasm is missing or empty; every browser consumer 404s at runtime');
}

// 3. The escape hatch, actually wired. The build asserts the patch was APPLIED; this asserts
//    it survived treeshaking and minification into the emitted glue.
const glue = esm.filter((file) => read(file).includes('harfbuzz.wasm'));
if (glue.length === 0) {
  failures.push('no ESM chunk references harfbuzz.wasm; the runtime was not inlined');
} else if (
  !glue.some((file) =>
    /\w+\(new URL\("harfbuzz\.wasm",\s*import\.meta\.url\)\.href\)/.test(read(file))
  )
) {
  failures.push(
    'the emitted glue reads harfbuzz.wasm directly; setHarfBuzzWasmUrl is wired to nothing'
  );
}

if (failures.length > 0) {
  console.error('Shipped shaper checks failed:');
  for (const failure of failures) console.error(`  - ${failure}`);
  process.exit(1);
}
console.log(`check:shipped-shaper: core's dist ships a browser-resolvable shaper`);
