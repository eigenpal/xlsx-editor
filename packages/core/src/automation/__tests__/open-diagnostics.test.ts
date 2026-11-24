// The public reason and limit of each refusal of the package reader.

import { describe, expect, test } from 'bun:test';
import type { PackageRejection } from '../../store/package/package.ts';
import { openDiagnosticOf } from '../open-diagnostics.ts';

const xml = (detail: string): PackageRejection =>
  ({ reason: 'xml', partName: '/secret.xml', detail }) as PackageRejection;

describe('openDiagnosticOf', () => {
  test('each reader refusal maps to its code, and each cap to its limit', () => {
    const cases: [PackageRejection, string, string?][] = [
      [{ reason: 'zip', detail: 'too-many-entries' }, 'open.zip-entry-limit', 'zip.maxEntries'],
      [
        { reason: 'zip', detail: 'too-large', limit: 'zip.maxTotalBytes' },
        'open.zip-size-limit',
        'zip.maxTotalBytes',
      ],
      [
        { reason: 'zip', detail: 'too-large', limit: 'zip.maxRatio' },
        'open.zip-ratio-limit',
        'zip.maxRatio',
      ],
      [{ reason: 'zip', detail: 'bad-name' }, 'open.zip-invalid-entry-name'],
      [{ reason: 'zip', detail: 'inflate-error' }, 'open.zip-corrupt'],
      [{ reason: 'zip', detail: 'overlapping-entries' }, 'open.zip-corrupt'],
      [{ reason: 'zip', detail: 'invalid-limits' }, 'open.invalid-limits'],
      [{ reason: 'zip', detail: 'something-else' }, 'open.failed'],
      [xml('too-large'), 'open.xml-part-size-limit', 'xml.maxBytes'],
      [xml('too-deep'), 'open.xml-depth-limit', 'xml.maxDepth'],
      [xml('too-many-elements'), 'open.xml-element-limit', 'xml.maxElements'],
      [xml('dtd-forbidden'), 'open.xml-dtd-forbidden'],
      [xml('entity-forbidden'), 'open.xml-dtd-forbidden'],
      [xml('parse-error'), 'open.xml-malformed'],
      [xml('unbound-prefix'), 'open.xml-malformed'],
      [xml('no-root'), 'open.xml-malformed'],
      [xml('bad-name'), 'open.xml-malformed'],
      [xml('invalid-limits'), 'open.invalid-limits'],
      [xml('something-else'), 'open.failed'],
      [{ reason: 'no-content-types' }, 'open.content-types-invalid'],
      [{ reason: 'not-a-workbook', detail: '/secret.xml' }, 'open.missing-workbook'],
      [
        {
          reason: 'workbook',
          detail: 'sheet-name-too-long',
          limit: 'workbook.maxSheetNameLength',
        },
        'open.sheet-name-limit',
        'workbook.maxSheetNameLength',
      ],
    ];
    for (const [error, code, limit] of cases) {
      const answer = openDiagnosticOf(error);
      expect(answer.code).toBe(code as never);
      expect(answer.limit).toBe(limit as never);
      expect(JSON.stringify(answer)).not.toContain('secret');
    }
  });
});
