// The substitute fonts of `@docx-editor.dev/fonts`, loaded into a HarfBuzz measurer for tests.

import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { createHarfBuzzMeasurer, initializeHarfBuzz } from '../text/harfbuzz-measurer.ts';
import type { TextMeasurer } from '../text/measurer.ts';

const fontsDir = join(
  dirname(Bun.resolveSync('@docx-editor.dev/fonts/package.json', import.meta.dir)),
  'assets'
);

const FACES: readonly [string, string, boolean, boolean][] = [
  ['Carlito', 'Carlito-Regular.ttf', false, false],
  ['Carlito', 'Carlito-Bold.ttf', true, false],
  ['Carlito', 'Carlito-Italic.ttf', false, true],
  ['Liberation Sans', 'LiberationSans-Regular.ttf', false, false],
];

/** A measurer with Carlito (for Calibri) and Liberation Sans. */
export async function substituteMeasurer(): Promise<TextMeasurer> {
  await initializeHarfBuzz();
  return createHarfBuzzMeasurer(
    FACES.map(([family, file, bold, italic]) => ({
      family,
      bold,
      italic,
      bytes: new Uint8Array(readFileSync(join(fontsDir, file))),
      origin: 'substitute',
    }))
  );
}
