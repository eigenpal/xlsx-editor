// `bun scripts/fidelity/glyph-share.ts <fixture>`: the share of the differing pixels of a
// fidelity run that lie inside the glyph boxes of the reference PDF's text. Run
// `bun run fidelity` first: this reads `<fixture>/ours.png` and `diff.png` in the report folder of
// that run (`reportDir`: `fidelity-report/`, or `XLSX_FIDELITY_REPORT_DIR`). The run writes
// `diff.png` only when some pixel differs, so with `ours.png` and no `diff.png` the tool reports
// 0 differing pixels. It warns when `ours.png` is older than the reference PDF. A wrong fixture
// name or a missing input prints its line and exits with code 2, as the run does
// (`helperInputs`).
//
// A glyph box spans the advance of a text item, from one em above its baseline to `DESCENT_EM`
// below it, widened by the line position threshold on its left and right sides. The share shows where the
// differing pixels lie: a share near 1 puts them on the drawn glyphs, not on the grid, the
// borders, or the fills. It does not show their cause; `rasterizer-control.ts` separates paint
// from the rasterizers.

import { existsSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { PNG } from 'pngjs';
import { differingPixelsIn, type PixelBox } from './compare.ts';
import { pdfItems } from './pdf-lines.ts';
import { FIXTURES } from '../fixtures/catalog.ts';
import { helperInputs } from './helper-inputs.ts';
import { FIXTURE_DIR } from './references.ts';
import { THRESHOLDS } from './thresholds.ts';

/** The depth of a glyph box below the baseline, as a share of the font size. */
const DESCENT_EM = 0.3;

const inputs = helperInputs(
  'glyph-share.ts',
  process.argv[2],
  ['ours.png'],
  FIXTURES.map((f) => f.name),
  FIXTURE_DIR
);
if (inputs.errors.length) {
  for (const line of inputs.errors) console.error(line);
  process.exit(2);
}
if (inputs.warning) console.warn(inputs.warning);
const { name, report } = inputs;
if (!existsSync(join(report, 'diff.png'))) {
  console.log(`${name}: 0 differing pixels (the run wrote no diff.png)`);
  process.exit(0);
}
const diff = PNG.sync.read(readFileSync(join(report, 'diff.png')));
const { items } = await pdfItems(inputs.pdf);
const k = THRESHOLDS.compareDpi / 72;
const pad = THRESHOLDS.maxLineShiftPt;
const boxes: PixelBox[] = items.map((i) => [
  (i.x - pad) * k,
  (i.y - i.size) * k,
  (i.x + i.width + pad) * k,
  (i.y + DESCENT_EM * i.size) * k,
]);
const { inside, outside, covered } = differingPixelsIn(diff, boxes);
const total = inside + outside;
const share = total ? inside / total : 0;
const area = covered / (diff.width * diff.height);
console.log(
  `${name}: ${total} differing pixels, ${inside} inside ${boxes.length} glyph boxes ` +
    `(${(share * 100).toFixed(1)}%), ${outside} outside; the boxes cover ` +
    `${(area * 100).toFixed(1)}% of the page`
);
