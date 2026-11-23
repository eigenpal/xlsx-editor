// Keeps docs/preserved-content.md complete: every child of the listed schema types is either a
// typed kind of the engine or named on that page (task 2.10).

import { describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { kindOf, type NodeKind } from '../package/node-kinds.ts';
import { NS } from '../package/namespaces.ts';

const ROOT = `${import.meta.dir}/../../../../..`;
const XSD = readFileSync(`${ROOT}/reference/ecma-376/part4/transitional-schemas/sml.xsd`, 'utf8');
const DOC = readFileSync(`${ROOT}/docs/preserved-content.md`, 'utf8');

/** The child element names of a complex type, from the schema text. */
function childNames(type: string): string[] {
  const start = XSD.indexOf(`<xsd:complexType name="${type}">`);
  const end = XSD.indexOf('</xsd:complexType>', start);
  return [...XSD.slice(start, end).matchAll(/<xsd:element name="(\w+)"/g)].map((m) => m[1]!);
}

/** The section of the page that lists a schema type. */
function listed(type: string): Set<string> {
  const at = DOC.indexOf(`(\`${type}\`)`);
  if (at < 0) return new Set();
  const end = DOC.indexOf('\n\n', DOC.indexOf('\n\n', at) + 2);
  return new Set([...DOC.slice(at, end).matchAll(/^- `(\w+)`$/gm)].map((m) => m[1]!));
}

const TYPES: readonly [string, NodeKind][] = [
  ['CT_Workbook', 'workbook'],
  ['CT_Worksheet', 'worksheet'],
  ['CT_SheetView', 'sheetView'],
  ['CT_Row', 'row'],
  ['CT_Cell', 'cell'],
  ['CT_Stylesheet', 'styleSheet'],
  ['CT_Sst', 'sst'],
  ['CT_Rst', 'stringItem'],
];

describe('preserved content inventory', () => {
  for (const [type, kind] of TYPES)
    test(`${type}: every child is typed or listed as preserved`, () => {
      const names = childNames(type);
      expect(names.length).toBeGreaterThan(0);
      const missing = names.filter(
        (name) => kindOf(NS.x, name, kind) === 'generic' && !listed(type).has(name)
      );
      const stale = [...listed(type)].filter((name) => kindOf(NS.x, name, kind) !== 'generic');
      expect({ missing, stale }).toEqual({ missing: [], stale: [] });
    });
});
