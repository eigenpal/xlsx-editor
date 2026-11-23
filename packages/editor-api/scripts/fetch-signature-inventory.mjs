#!/usr/bin/env node
// After PPTX Editor packages/editor-api/scripts/fetch-signature-inventory.mjs at d4502552, an
// adaptation of DOCX Editor packages/editor-api/scripts/fetch-signature-inventory.mjs at commit
// 5993bd223945cabdaf3900557d9fb59bfd8c2d8c. Changes: the workbook namespace of Office.js, and
// the inventory keeps only the classes and interfaces that `compat/manifest.json` selects, with
// every member of each, instead of the whole namespace. The workbook namespace is far larger than
// the presentation namespace, and the conformance check reads only the selected classes. Each row
// also leaves out `requirementSet`: the requirement set names carry the namespace name, which
// committed text does not write (AGENTS.md "Naming"), and no check reads them.
//
// Explicit maintenance command: it downloads the pinned `@types/office-js` release named in
// `compat/provenance.json` and checks its integrity hash. Tests and builds stay offline.
import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createHash } from 'node:crypto';
import { verifySubresourceIntegrity } from './lib/integrity.mjs';
import { extractFileFromTarGzip } from './lib/tar.mjs';
import { extractUpstreamInventory } from './lib/signature-inventory.mjs';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const provenance = JSON.parse(await fs.readFile(path.join(root, 'compat/provenance.json'), 'utf8'));
const manifest = JSON.parse(await fs.readFile(path.join(root, 'compat/manifest.json'), 'utf8'));
const upstream = provenance.upstreamPackage;
const response = await fetch(upstream.tarballUrl, { signal: AbortSignal.timeout(60_000) });
if (!response.ok) throw new Error(`Upstream download failed: ${response.status}`);
const bytes = Buffer.from(await response.arrayBuffer());
if (!verifySubresourceIntegrity(bytes, upstream.integrity))
  throw new Error('Upstream integrity mismatch');
const source = extractFileFromTarGzip(bytes, 'index.d.ts');
if (!source) throw new Error('Upstream index.d.ts missing');

/** The owners the manifest selects, as inventory row ids name them. */
const owners = new Set(
  Object.entries(manifest.symbols ?? {}).map(([name, selection]) =>
    selection.namespace ? `${selection.namespace}.${name}` : name
  )
);
const endpoints = extractUpstreamInventory(source.toString('utf8'))
  .filter((row) => owners.has(row.uid.split('#')[0]))
  .map(({ requirementSet: _set, ...row }) => row);
for (const owner of owners)
  if (!endpoints.some((row) => row.uid.split('#')[0] === owner))
    throw new Error(`The pinned release has no ${owner}`);
const header = {
  schemaVersion: 1,
  upstream: {
    name: upstream.name,
    version: upstream.version,
    integrity: upstream.integrity,
    sourceUrl: upstream.sourceRepository.sourceUrl,
    declarationSha256: createHash('sha256').update(source).digest('hex'),
  },
};
await fs.writeFile(
  path.join(root, 'compat/reference/workbook.inventory.json'),
  // One endpoint per line keeps the fixture practical to diff.
  JSON.stringify(header, null, 2).slice(0, -2) +
    ',\n  "endpoints": [\n' +
    endpoints.map((entry) => `    ${JSON.stringify(entry)}`).join(',\n') +
    '\n  ]\n}\n'
);
console.log(`Pinned ${endpoints.length} endpoints from ${upstream.name}@${upstream.version}.`);
