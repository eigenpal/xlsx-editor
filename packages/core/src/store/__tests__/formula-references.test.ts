// A typed formula names only sheets, external workbook indexes, and tables that the workbook
// holds. Another name is refused with `unknown-formula-reference`, so it is never saved as a
// link to a workbook that is not there. A sheet name with a space must be quoted.

import { describe, expect, test } from 'bun:test';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { buildXlsx, c, row } from '../../../../../scripts/fixtures/builder.ts';
import { readXlsx } from '../package/package.ts';
import { WorkbookDocumentStore } from '../store/document-store.ts';
import type { WorkbookOp } from '../store/ops.ts';

const TABLE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml';
const tableXml =
  '<?xml version="1.0" encoding="UTF-8"?><table xmlns="http://schemas.openxmlformats.org/' +
  'spreadsheetml/2006/main" id="1" name="Sales" displayName="Sales" ref="A1:B2">' +
  '<tableColumns count="2"><tableColumn id="1" name="Q"/><tableColumn id="2" name="N"/>' +
  '</tableColumns></table>';

/** A workbook with the sheets `S` and `My sheet`, one external reference, and table `Sales`. */
function storeOf(): WorkbookDocumentStore {
  const bytes = buildXlsx({
    sheets: [
      { name: 'S', rows: row(1, [c('A1', 1)]) },
      { name: 'My sheet', rows: row(1, [c('A1', 2)]) },
    ],
    extraParts: [['/xl/tables/table1.xml', TABLE, strToU8(tableXml)]],
  });
  const files = unzipSync(bytes);
  const wb = strFromU8(files['xl/workbook.xml']!);
  files['xl/workbook.xml'] = strToU8(
    wb.replace(
      '</sheets>',
      '</sheets><externalReferences><externalReference r:id="rIdX"/></externalReferences>'
    )
  );
  const read = readXlsx(zipSync(files));
  if (!read.ok) throw new Error(JSON.stringify(read.error));
  return new WorkbookDocumentStore(read.pkg);
}

const formula = (text: string, ref = 'B1'): WorkbookOp => ({
  op: 'setCellValue',
  sheet: 'S',
  ref,
  value: { type: 'formula', text },
});

function result(ops: WorkbookOp[]): string {
  const r = storeOf().transact(ops);
  return r.ok ? 'ok' : r.rejection.code;
}

describe('the references of a typed formula', () => {
  test.each([
    ["'My sheet'!A1+S!A1", 'ok'],
    ['s!A1', 'ok'],
    ['SUM(S:S!A1)', 'ok'],
    ['[1]Other!A1', 'ok'],
    ['SUM(Sales[N])', 'ok'],
    ['Nope!A1', 'unknown-formula-reference'],
    ['SUM(S:Nope!A1)', 'unknown-formula-reference'],
    ['[2]Other!A1', 'unknown-formula-reference'],
    ['[0]S!A1', 'unknown-formula-reference'],
    ["'[Book.xlsx]S'!A1", 'unknown-formula-reference'],
    ['[Book.xlsx]S!A1', 'unknown-formula-reference'],
    ['SUM(Missing[N])', 'unknown-formula-reference'],
    ['Nope!Total', 'unknown-formula-reference'],
    ['My sheet!A1', 'unknown-formula-reference'],
  ])('%s is %s', (text, expected) => {
    expect(result([formula(text)])).toBe(expected);
  });

  test('a batch with an unknown reference is refused as a whole, and writes nothing', () => {
    const store = storeOf();
    const before = store.revision;
    const r = store.transact([formula('S!A1', 'B1'), formula('Nope!A1', 'C1')]);
    expect(r.ok ? 'ok' : r.rejection.code).toBe('unknown-formula-reference');
    expect(store.revision).toBe(before);
  });
});
