/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See LICENSE.md in this package.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// After PPTX Editor's `packages/editor-api/scripts/__tests__/conformance.test.ts` at d4502552.
// Differences: the probes use members of the workbook subset.
//
// The generated Office.js conformance assertions and the source-compatibility fixtures. The
// compiler runs offline on committed files; nothing here fetches the network.
import { afterAll, describe, expect, test } from 'bun:test';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import {
  ASSERTIONS_PATH,
  generateConformance,
  generateFromCompat,
  qualifyTypeText,
} from '../generate-conformance.mjs';
import { TYPECHECK_TIMEOUT_MS, typecheckProject } from '../lib/typecheck-compat.mjs';
import manifest from '../../compat/manifest.json';
import runtimeNotes from '../../compat/runtime-notes.json';
import inventory from '../../compat/reference/workbook.inventory.json';

type Row = { uid: string; [key: string]: unknown };

const packageRoot = path.resolve(import.meta.dir, '..', '..');
const compatTsconfig = path.join(packageRoot, 'compat', 'tsconfig.json');
const rows = inventory.endpoints as Row[];
const rowOf = (uid: string) => structuredClone(rows.find((r) => r.uid === uid)!);
const temporary: string[] = [];

afterAll(() => {
  for (const directory of temporary) fs.rmSync(directory, { recursive: true, force: true });
});

/**
 * Generate assertions for `members` from `endpoints` in a temporary project, and compile it.
 * The assertions name the real package, or `declarations` when it is given. Answers the
 * diagnostics of the generated file, each as the alias or function that the diagnostic line
 * declares, then the message.
 */
function compileAgainstPackage(
  members: Record<string, string[]>,
  endpoints: Row[],
  declarations?: string
): string[] {
  const directory = fs.mkdtempSync(path.join(os.tmpdir(), 'xlsx-conformance-'));
  temporary.push(directory);
  if (declarations !== undefined)
    fs.writeFileSync(path.join(directory, 'declarations.ts'), declarations);
  const symbols = Object.fromEntries(
    Object.entries(members).map(([name, list]) => [name, { members: list }])
  );
  const { assertionsSource } = generateConformance({
    inventory: { endpoints },
    manifest: { symbols },
    packageImport:
      declarations === undefined ? path.join(packageRoot, 'src', 'index.ts') : './declarations.ts',
    typeAssertImport: path.join(packageRoot, 'compat', 'type-assert.ts'),
  });
  const file = path.join(directory, 'assertions.ts');
  fs.writeFileSync(file, assertionsSource);
  fs.writeFileSync(
    path.join(directory, 'tsconfig.json'),
    JSON.stringify({
      extends: path.join(packageRoot, 'tsconfig.json'),
      compilerOptions: { noEmit: true },
      include: ['assertions.ts'],
    })
  );
  const lines = assertionsSource.split('\n');
  return typecheckProject(path.join(directory, 'tsconfig.json'), { scope: 'project' }).map(
    (diagnostic) => {
      const [, row, message] = /assertions\.ts:(\d+):\d+ - ([^\n]*)/.exec(diagnostic) ?? [];
      let declared: RegExpExecArray | null = null;
      for (let index = Number(row) - 1; index >= 0 && !declared; index--)
        declared = /^(?:export function|type) (\w+)/.exec(lines[index] ?? '');
      const text = (message ?? diagnostic).replace(/Namespace '"[^"]*"'/, 'The package');
      return `${declared?.[1] ?? '?'}: ${text}`;
    }
  );
}

describe('the conformance assertions', () => {
  test(
    'the compat project (generated assertions and fixtures) type-checks with zero diagnostics',
    () => {
      expect(typecheckProject(compatTsconfig, { scope: 'project' })).toEqual([]);
    },
    TYPECHECK_TIMEOUT_MS
  );

  test('the committed generated file is current', () => {
    expect(fs.readFileSync(ASSERTIONS_PATH, 'utf8')).toBe(generateFromCompat().assertionsSource);
  });

  test('every selected member is asserted or named in conformanceExceptions', () => {
    const result = generateFromCompat();
    const selected = Object.values(
      manifest.symbols as Record<string, { members: string[]; isFunction?: boolean }>
    ).reduce(
      (count, selection) => count + (selection.isFunction ? 1 : selection.members.length),
      0
    );
    expect(result.assertions + result.skipped.length).toBe(selected);
    expect(result.skipped.map((entry: { uid: string }) => entry.uid)).toEqual(
      manifest.conformanceExceptions.map((entry) => entry.uid)
    );
    for (const { uid } of result.skipped)
      expect(result.assertionsSource).toContain(
        `// Not asserted: manifest.json conformanceExceptions "${uid}".`
      );
  });

  const members = {
    Range: ['address', 'values'],
    Worksheet: ['getRange'],
    WorksheetCollection: ['items'],
  };
  const uids = ['Range#address', 'Range#values', 'Worksheet#getRange', 'WorksheetCollection#items'];

  test(
    'the pinned reference compiles against the real package',
    () => {
      expect(compileAgainstPackage(members, uids.map(rowOf))).toEqual([]);
    },
    TYPECHECK_TIMEOUT_MS
  );

  test(
    'a reference that the package does not match fails its assertion, in either direction',
    () => {
      const [address, values, getRange, items] = uids.map(rowOf);
      // A writable reference against a read-only property.
      Object.assign(address!, { readonly: false, writeType: 'string' });
      // A narrower type than the package declares: the check is not one-directional.
      Object.assign(values!, { type: 'string [ ] [ ]', writeType: 'string [ ] [ ]' });
      // A required parameter where the package has an optional one.
      (
        getRange!.overloads as { parameters: { optional: boolean; type: string }[] }[]
      )[0]!.parameters[0]!.optional = false;
      (
        getRange!.overloads as { parameters: { optional: boolean; type: string }[] }[]
      )[0]!.parameters[0]!.type = 'string';
      // A type name that the package does not export.
      items!.type = 'Sheet [ ]';
      expect(compileAgainstPackage(members, [address!, values!, getRange!, items!])).toEqual([
        "_assert_Range_address: Type 'false' does not satisfy the constraint 'true'.",
        "_write_Range_address: Cannot assign to 'address' because it is a read-only property.",
        "_assert_Range_values: Type 'false' does not satisfy the constraint 'true'.",
        "_assert_Worksheet_getRange: Type 'false' does not satisfy the constraint 'true'.",
        "Ref_WorksheetCollection_items: The package has no exported member 'Sheet'.",
        "_assert_WorksheetCollection_items: Type 'false' does not satisfy the constraint 'true'.",
      ]);
    },
    TYPECHECK_TIMEOUT_MS
  );

  test(
    'a declaration that drifts from the reference fails its assertion',
    () => {
      // The reference is unchanged. The package side drops `readonly` from `address`, narrows
      // `values` to strings, and makes the address of `getRange` required.
      const declarations = `
        export declare class Range {
          address: string;
          values: string[][];
        }
        export declare class Worksheet {
          getRange(address: string): Range;
        }
      `;
      const drifted = { Range: ['address', 'values'], Worksheet: ['getRange'] };
      expect(compileAgainstPackage(drifted, uids.slice(0, 3).map(rowOf), declarations)).toEqual([
        "_assert_Range_address: Type 'false' does not satisfy the constraint 'true'.",
        "_assert_Range_values: Type 'false' does not satisfy the constraint 'true'.",
        "_assert_Worksheet_getRange: Type 'false' does not satisfy the constraint 'true'.",
      ]);
    },
    TYPECHECK_TIMEOUT_MS
  );
});

describe('generateConformance', () => {
  test('qualifies package type names and leaves string literals and type parameters alone', () => {
    expect(qualifyTypeText('"Worksheet" | Worksheet | null')).toBe(
      '"Worksheet" | XlsxEditor.Worksheet | null'
    );
    expect(qualifyTypeText('( context : RequestContext ) => Promise < T >', ['T'])).toBe(
      '(context: XlsxEditor.RequestContext) => Promise<T>'
    );
  });

  test('refuses a conformance exception that names a member the manifest does not select', () => {
    expect(() =>
      generateConformance({
        inventory: { endpoints: [rowOf('Range#values')] },
        manifest: {
          symbols: { Range: { members: ['values'] } },
          conformanceExceptions: [{ uid: 'Range#format', reason: 'Not selected.' }],
        },
      })
    ).toThrow('is not a selected member');
  });

  test('refuses a selected member that the reference inventory does not list', () => {
    expect(() =>
      generateConformance({
        inventory: { endpoints: [] },
        manifest: { symbols: { Range: { members: ['values'] } } },
      })
    ).toThrow('has no row for Range#values');
  });
});

describe('the runtime notes', () => {
  test('each note names a selected member, with a status and its behavior', () => {
    const selected = new Set<string>();
    for (const [name, selection] of Object.entries(
      manifest.symbols as Record<string, { members: string[]; namespace?: string }>
    )) {
      const prefix = selection.namespace ? `${selection.namespace}.` : '';
      for (const member of selection.members) selected.add(`${prefix}${name}#${member}`);
    }
    const notes = runtimeNotes as Record<string, { status: string; notes: string }>;
    expect(Object.keys(notes).length).toBeGreaterThan(0);
    for (const [uid, note] of Object.entries(notes)) {
      expect(selected.has(uid)).toBe(true);
      expect(['equivalent', 'partial']).toContain(note.status);
      expect(note.notes.length).toBeGreaterThan(0);
    }
  });
});
