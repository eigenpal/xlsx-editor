// Build the synthetic fixtures under `fixtures/`.
//
//   bun run fixtures:build            write every fixture
//   bun run fixtures:build --check    fail when a committed fixture differs from its build
//
// Each fixture also needs a reference PDF next to it (`AGENTS.md`, "Fidelity").

import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { buildXlsx } from './builder.ts';
import { FIXTURES } from './catalog.ts';

const OUT = 'fixtures';
if (
  new Date(0).getTimezoneOffset() !== 0 ||
  new Date(Date.UTC(1980, 0, 2, 12)).getTimezoneOffset() !== 0
)
  throw new Error('Build fixtures in UTC (TZ=UTC): the zip entry time reads local fields.');
const check = process.argv.includes('--check');
const stale: string[] = [];
mkdirSync(OUT, { recursive: true });
for (const fixture of FIXTURES) {
  const bytes = buildXlsx(fixture.workbook());
  const path = `${OUT}/${fixture.name}.xlsx`;
  if (check) {
    if (!existsSync(path) || Buffer.compare(readFileSync(path), Buffer.from(bytes)) !== 0)
      stale.push(path);
  } else writeFileSync(path, bytes);
}
if (stale.length) {
  console.error(
    `Fixtures differ from their build:\n  ${stale.join('\n  ')}\nFix: bun run fixtures:build`
  );
  process.exit(1);
}
console.log(`${check ? 'Checked' : 'Built'} ${FIXTURES.length} fixtures in ${OUT}/.`);
