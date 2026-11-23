// Build the demo's sample workbook (showcase.ts) into `examples/vite/public/sample.xlsx`.
//
//   bun run sample:build
//
// `showcase.test.ts` fails when the committed file differs from its build. Same role as
// pptx-editor/scripts/sample-deck/build.ts at 79a5d22d (EigenPal). Like
// `scripts/fixtures/build.ts`, it runs in UTC, because the zip writer reads local time fields.

import { writeFileSync } from 'node:fs';
import { join, relative } from 'node:path';
import { buildSample } from './showcase.ts';

if (
  new Date(0).getTimezoneOffset() !== 0 ||
  new Date(Date.UTC(1980, 0, 2, 12)).getTimezoneOffset() !== 0
)
  throw new Error('Build the sample in UTC (TZ=UTC): the zip entry time reads local fields.');
const path = join(import.meta.dir, '../../examples/vite/public/sample.xlsx');
const bytes = buildSample();
writeFileSync(path, bytes);
console.log(`Wrote ${relative(process.cwd(), path)} (${bytes.length} bytes).`);
