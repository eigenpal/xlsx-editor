// The input check of the helper tools `glyph-share.ts` and `rasterizer-control.ts`: the fixture
// name, the reference PDF of the fixture, and the rasters of a run in the report folder. Each
// problem gives one line. A tool prints it and exits with code 2, as the run does for an input
// that it cannot read.

import { existsSync, statSync } from 'node:fs';
import { join } from 'node:path';
import { reportDir, runFirstText } from './report.ts';

type Env = Readonly<Record<string, string | undefined>>;

/** What a helper tool reads, or why it cannot start. */
export interface HelperInputs {
  /** The fixture name, checked against the catalog. */
  readonly name: string;
  /** The report folder of the fixture; empty when the fixture name is refused. */
  readonly report: string;
  /** The reference PDF of the fixture; empty when the fixture name is refused. */
  readonly pdf: string;
  /** The lines to print before the tool stops with code 2. Empty when the inputs exist. */
  readonly errors: readonly string[];
  /** A raster of the run that is older than the reference PDF. */
  readonly warning?: string;
}

/**
 * Check the inputs of the helper tool `tool` for the fixture `name`. `rasters` are the files of the
 * run in the report folder of the fixture; the first one is the raster of the run that a newer
 * reference PDF makes stale.
 */
export function helperInputs(
  tool: string,
  name: string | undefined,
  rasters: readonly string[],
  fixtures: readonly string[],
  fixtureDir: string,
  env: Env = process.env
): HelperInputs {
  if (!name || !fixtures.includes(name))
    return {
      name: name ?? '',
      report: '',
      pdf: '',
      errors: [
        ...(name ? [`unknown fixture: ${name}`] : []),
        `usage: bun scripts/fidelity/${tool} <fixture>`,
        `fixtures: ${fixtures.join(', ')}`,
      ],
    };
  const report = join(reportDir(env), name);
  const pdf = join(fixtureDir, `${name}.pdf`);
  if (!existsSync(pdf))
    return { name, report, pdf, errors: [`no reference PDF for ${name} in ${fixtureDir}`] };
  const missing = rasters.find((file) => !existsSync(join(report, file)));
  if (missing)
    return { name, report, pdf, errors: [`no ${join(report, missing)}: ${runFirstText(env)}`] };
  const first = rasters[0];
  const stale = first && statSync(join(report, first)).mtimeMs < statSync(pdf).mtimeMs;
  return {
    name,
    report,
    pdf,
    errors: [],
    ...(stale && {
      warning: `${name}: ${first} is older than the reference PDF: ${runFirstText(env, 'again')}`,
    }),
  };
}
