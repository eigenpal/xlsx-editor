// The measured result of each measurement case (scripts/fixtures/cases.ts): the displayed text,
// the fill and font colors, the font name, the bold state, and the bottom border of a cell.
// A maintainer records them in `fixtures/measured.json` with the hash of the fixture that was
// measured, through `record` below. Tests compare the engine with these values, and a stale
// record fails its test, like a stale reference PDF.
//
//   bun scripts/fixtures/measured.ts refs           list `fixture<TAB>ref` to measure
//   bun scripts/fixtures/measured.ts record < tsv   record the measured values

import { createHash } from 'node:crypto';
import { existsSync, readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import {
  FORMAT_CASES,
  FORMAT_CASES_1904,
  FORMAT_CASES_MORE,
  STYLE_CASES,
  caseCell,
} from './cases.ts';
import { STYLE_RESOLUTION_CASES } from './fill-style-cases.ts';
import { fixtureFile } from './catalog.ts';

const FIXTURES = join(import.meta.dir, '..', '..', 'fixtures');
export const MEASURED = join(FIXTURES, 'measured.json');

/** The bottom border of a cell: its `RRGGBB` color and its `ST_BorderStyle` value (18.18.3). */
export interface MeasuredBorder {
  readonly color: string;
  readonly style: string;
}

/** The measured values of one cell. Colors are `RRGGBB`. */
export interface MeasuredCell {
  readonly text: string;
  readonly fill?: string;
  readonly font?: string;
  readonly fontName?: string;
  readonly bold?: boolean;
  /** The bottom border, when the cell has one. */
  readonly bottom?: MeasuredBorder;
}

export interface MeasuredFixture {
  readonly workbook: string;
  readonly cells: Readonly<Record<string, MeasuredCell>>;
}

/** The measured cells, by fixture. */
export const MEASURED_REFS: Readonly<Record<string, readonly string[]>> = {
  'number-formats': FORMAT_CASES.map((_, i) => caseCell(i).ref),
  'number-formats-more': FORMAT_CASES_MORE.map((_, i) => caseCell(i).ref),
  'dates-1904': FORMAT_CASES_1904.map((_, i) => caseCell(i).ref),
  'colors-fonts': STYLE_CASES.map((k) => k.ref),
  'cell-styles': STYLE_RESOLUTION_CASES.map((k) => k.ref),
};

const sha256 = (bytes: Uint8Array): string => createHash('sha256').update(bytes).digest('hex');
export const workbookHash = (name: string): string =>
  sha256(new Uint8Array(readFileSync(join(FIXTURES, fixtureFile(name)))));

export function readMeasured(): Record<string, MeasuredFixture> {
  if (!existsSync(MEASURED)) return {};
  return (
    JSON.parse(readFileSync(MEASURED, 'utf8')) as { fixtures: Record<string, MeasuredFixture> }
  ).fixtures;
}

/** The measured cells of a fixture, or an error when the record is missing or stale. */
export function measuredCells(name: string): Readonly<Record<string, MeasuredCell>> {
  const all = readMeasured();
  const record = Object.hasOwn(all, name) ? all[name] : undefined;
  if (!record) throw new Error(`${name}: no measured record in fixtures/measured.json`);
  if (record.workbook !== workbookHash(name))
    throw new Error(`${name}: the fixture changed since it was measured (stale)`);
  return record.cells;
}

/** `R,G,B` as `RRGGBB`, or undefined. */
function hex(rgb: string | undefined): string | undefined {
  if (!rgb) return undefined;
  const parts = rgb.split(',').map((p) => Number(p.trim()));
  if (parts.length !== 3 || parts.some((p) => !Number.isInteger(p) || p < 0 || p > 255))
    return undefined;
  return parts.map((p) => p.toString(16).toUpperCase().padStart(2, '0')).join('');
}

/** `R,G,B|style` as a border, where style is an `ST_BorderStyle` value, or undefined. */
function border(value: string | undefined): MeasuredBorder | undefined {
  const [rgb, style] = (value ?? '').split('|');
  const color = hex(rgb);
  return color && style && /^[a-zA-Z]+$/.test(style.trim())
    ? { color, style: style.trim() }
    : undefined;
}

if (import.meta.main) {
  const [command] = process.argv.slice(2);
  if (command === 'refs') {
    for (const [name, refs] of Object.entries(MEASURED_REFS))
      for (const ref of refs) console.log(`${name}\t${ref}`);
  } else if (command === 'record') {
    // Lines: fixture, ref, text, fill r,g,b, font r,g,b, font name, bold, bottom r,g,b|style
    const input = readFileSync(0, 'utf8');
    const out: Record<string, { workbook: string; cells: Record<string, MeasuredCell> }> = {};
    for (const line of input.split('\n')) {
      if (!line.trim()) continue;
      const [name, ref, text, fill, font, fontName, bold, bottom] = line.split('\t');
      if (!name || !ref || !Object.hasOwn(MEASURED_REFS, name))
        throw new Error(`bad line: ${line}`);
      out[name] ??= { workbook: workbookHash(name), cells: {} };
      const cell: MeasuredCell = {
        text: text ?? '',
        ...(hex(fill) ? { fill: hex(fill) } : {}),
        ...(hex(font) ? { font: hex(font) } : {}),
        ...(fontName ? { fontName } : {}),
        ...(bold ? { bold: bold === 'true' } : {}),
        ...(border(bottom) ? { bottom: border(bottom) } : {}),
      };
      out[name].cells[ref] = cell;
    }
    const merged = { ...readMeasured(), ...out };
    const sorted = Object.fromEntries(Object.entries(merged).sort(([a], [b]) => (a < b ? -1 : 1)));
    writeFileSync(MEASURED, `${JSON.stringify({ version: 1, fixtures: sorted }, null, 2)}\n`);
    console.log(`recorded ${Object.keys(out).join(', ')}`);
  } else {
    console.error('usage: refs | record < measured.tsv');
    process.exit(1);
  }
}
