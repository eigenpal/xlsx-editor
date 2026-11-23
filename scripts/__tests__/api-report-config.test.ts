import { afterEach, describe, expect, test } from 'bun:test';
import { Extractor, ExtractorConfig } from '@microsoft/api-extractor';
import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { apiReportConfigObject } from '../lib/api-report-config.mjs';

let root: string | undefined;

afterEach(() => {
  if (root) rmSync(root, { recursive: true, force: true });
  root = undefined;
});

function invoke(directory: string, localBuild: boolean) {
  const config = ExtractorConfig.prepare({
    configObject: apiReportConfigObject({
      root: directory,
      dts: join(directory, 'index.d.ts'),
      reportFileName: 'index.api.md',
      reportFolder: join(directory, 'reports'),
      reportTempFolder: join(directory, 'temp'),
    }) as never,
    configObjectFullPath: join(directory, 'api-extractor.json'),
    packageJsonFullPath: join(directory, 'package.json'),
  });
  return Extractor.invoke(config, {
    localBuild,
    messageCallback: (message) => {
      message.handled = true;
    },
  });
}

describe('apiReportConfigObject', () => {
  test('a change to a type that the entry does not export is drift', () => {
    root = mkdtempSync(join(tmpdir(), 'api-report-config-'));
    mkdirSync(join(root, 'reports'));
    writeFileSync(join(root, 'package.json'), '{"name":"fixture","version":"1.0.0"}');
    writeFileSync(
      join(root, 'index.d.ts'),
      "import type { ZipLimits } from './limits';\nexport declare function read(limits?: ZipLimits): void;\n"
    );
    const declare = (member: string) =>
      writeFileSync(join(root!, 'limits.d.ts'), `export interface ZipLimits {\n  ${member}\n}\n`);
    declare('maxEntries: number;');
    invoke(root, true);
    expect(invoke(root, false).apiReportChanged).toBe(false);
    declare('maxParts: string;');
    expect(invoke(root, false).apiReportChanged).toBe(true);
  });
});
