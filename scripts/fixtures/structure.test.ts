// The structure fixture: its edited copy deletes row 4 and inserts column C of Data, and each
// rule of design D4 that the fixture covers holds in the saved bytes that the save check opens.

import { describe, expect, test } from 'bun:test';
import {
  attr,
  cellKey,
  child,
  children,
  NS,
  parseCellRef,
  readXlsx,
  sheetIndex,
  textOf,
  tokenizeFormula,
  workbookIndex,
  xmlPartOf,
  type Element,
  type XlsxPackage,
} from '../../packages/core/src/store/index.ts';
import { fixtureCopies } from '../save-check/records.ts';
import { FIXTURES } from './catalog.ts';

function open(bytes: Uint8Array): XlsxPackage {
  const read = readXlsx(bytes);
  if (!read.ok) throw new Error(JSON.stringify(read.error));
  return read.pkg;
}

const fixture = FIXTURES.find((f) => f.name === 'structure')!;
const pkg = open(fixtureCopies(fixture).edited);
const DATA = '/xl/worksheets/sheet1.xml';
const REFS = '/xl/worksheets/sheet2.xml';
const root = (part: string) => xmlPartOf(pkg, part)!.root;

function cell(part: string, ref: string): Element | undefined {
  const a = parseCellRef(ref)!;
  return sheetIndex(xmlPartOf(pkg, part)!).cells.get(cellKey(a.row, a.col))?.el;
}
const f = (part: string, ref: string) => child(cell(part, ref), 'x:f');
const text = (part: string, ref: string) => textOf(f(part, ref));

describe('the structure fixture, edited', () => {
  test('cells move, a reference to the deleted row is #REF!, and ranges shrink or grow', () => {
    expect(textOf(child(cell(DATA, 'B4'), 'x:v'))).toBe('70');
    expect(cell(DATA, 'C2')).toBeUndefined();
    expect(text(DATA, 'B5')).toBe('SUM(B2:B4)');
    expect(text(DATA, 'D5')).toBe('SUM($D$2:$D$4)');
    expect(text(DATA, 'E5')).toBe('#REF!');
    expect(text(REFS, 'B1')).toBe('Data!E2');
    expect(text(REFS, 'B2')).toBe('SUM(Data!B2:D4)');
    expect(text(REFS, 'B3')).toBe('Data!#REF!');
    expect(text(REFS, 'B4')).toBe('Data!A4');
    expect(workbookIndex(pkg)!.fullCalcOnLoad).toBe(true);
  });

  test('one shared formula of Data stays shared; the other unshares the child of row 4', () => {
    expect([text(DATA, 'E2'), attr(f(DATA, 'E2'), 'ref')]).toEqual(['B2+D2', 'E2:E4']);
    for (const r of ['E3', 'E4']) expect(attr(f(DATA, r), 'si')).toBe('0');
    expect([text(DATA, 'I2'), attr(f(DATA, 'I2'), 'ref')]).toEqual(['B3*2', 'I2:I4']);
    expect([text(DATA, 'I3'), attr(f(DATA, 'I3'), 't')]).toEqual(['#REF!*2', undefined]);
    expect([attr(f(DATA, 'I4'), 't'), attr(f(DATA, 'I4'), 'si')]).toEqual(['shared', '1']);
    expect(cell(DATA, 'I5')).toBeUndefined();
  });

  test('the plain cross-sheet formulas of Refs move or become #REF!', () => {
    expect(text(REFS, 'C1')).toBe('Data!B2*2');
    expect(text(REFS, 'C2')).toBe('Data!B3*2');
    expect(text(REFS, 'C3')).toBe('Data!#REF!*2');
    expect(text(REFS, 'C4')).toBe('Data!B4*2');
  });

  test('the merged range shrinks, the column entry splits, and the hyperlink moves', () => {
    const data = root(DATA);
    const merges = children(child(data, 'x:mergeCells'), 'x:mergeCell');
    expect(merges.map((m) => attr(m, 'ref'))).toEqual(['G3:G4']);
    const cols = children(child(data, 'x:cols'), 'x:col').map((c) => [
      attr(c, 'min'),
      attr(c, 'max'),
    ]);
    expect(cols).toEqual([
      ['1', '1'],
      ['2', '2'],
      ['4', '5'],
    ]);
    const link = child(child(data, 'x:hyperlinks'), 'x:hyperlink')!;
    expect([attr(link, 'ref'), attr(link, 'location')]).toEqual(['A4', 'Refs!A1']);
  });

  test('defined names move or become #REF!', () => {
    const wb = root('/xl/workbook.xml');
    const names = children(child(wb, 'x:definedNames'), 'x:definedName').map((n) => [
      attr(n, 'name'),
      textOf(n),
    ]);
    expect(names).toEqual([
      ['Gone', 'Data!#REF!'],
      ['Totals', 'Data!$E$2:$E$4'],
    ]);
  });

  test('drawing anchors move by their kind', () => {
    const anchors = root('/xl/drawings/drawing1.xml').children.filter(
      (n): n is Element => n.type === 'element'
    );
    const at = (a: Element, tag: string) => {
      const p = child(a, { ns: NS.xdr, local: tag })!;
      return ['col', 'row'].map((k) => Number(textOf(child(p, { ns: NS.xdr, local: k }))));
    };
    expect([at(anchors[0]!, 'from'), at(anchors[0]!, 'to')]).toEqual([
      [1, 8],
      [4, 11],
    ]);
    expect(at(anchors[1]!, 'from')).toEqual([5, 8]);
    expect(attr(child(anchors[2]!, { ns: NS.xdr, local: 'pos' }), 'y')).toBe('3175000');
  });
});

describe('every fixture', () => {
  // A shared formula whose text holds a sheet-qualified reference fails to open in the save
  // check (structure-case.ts). No fixture copy that the check opens holds one.
  const qualifiedShared = (bytes: Uint8Array): string[] => {
    const p = open(bytes);
    const out: string[] = [];
    for (const sheet of workbookIndex(p)!.sheets) {
      const part = sheet.partName ? xmlPartOf(p, sheet.partName) : undefined;
      if (!part) continue;
      for (const entry of sheetIndex(part).cells.values()) {
        const fx = child(entry.el, 'x:f');
        if (attr(fx, 't') !== 'shared' || attr(fx, 'ref') === undefined) continue;
        const read = tokenizeFormula(textOf(fx));
        if (!read.ok || read.tokens.some((t) => 'sheet' in t && t.sheet))
          out.push(`${sheet.name}!${textOf(fx)}`);
      }
    }
    return out;
  };

  test('holds no shared formula with a sheet-qualified reference, saved or edited', () => {
    for (const fx of FIXTURES) {
      const copies = fixtureCopies(fx);
      expect([fx.name, qualifiedShared(copies.saved)]).toEqual([fx.name, []]);
      expect([fx.name, qualifiedShared(copies.edited)]).toEqual([fx.name, []]);
    }
  });
});
