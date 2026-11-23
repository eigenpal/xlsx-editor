// The workbooks that the demo lists: the sample workbook first, then the committed fixture
// workbooks, served by the dev server from `fixtures/`. The demo fetches only these files, from
// its own origin.

import { SAMPLE } from './sample.ts';

const urls = import.meta.glob('../../../fixtures/*.xlsx', {
  query: '?url',
  import: 'default',
  eager: true,
}) as Record<string, string>;

/** The URL of the sample workbook, in `public/`. */
const SAMPLE_URL = `${import.meta.env.BASE_URL}sample.xlsx`;

/** Fixture name to the URL that the dev server gives it. */
export const FIXTURES: ReadonlyMap<string, string> = new Map(
  Object.entries(urls)
    .map(([path, url]) => [path.replace(/^.*\/(.+)\.xlsx$/, '$1'), url] as const)
    .sort(([a], [b]) => a.localeCompare(b))
);

// The sample's name is not a fixture name, so the list has one entry for each.
if (FIXTURES.has(SAMPLE)) throw new Error(`A fixture may not be named "${SAMPLE}".`);

/** Every workbook that the demo can open by name: the sample, then the fixtures. */
export const WORKBOOKS: ReadonlyMap<string, string> = new Map([[SAMPLE, SAMPLE_URL], ...FIXTURES]);
