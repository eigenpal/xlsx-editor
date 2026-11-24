// The file kind of the open package gives the extension and the media type of a save. A
// macro-enabled package keeps its macro-enabled extension, so the saved file opens cleanly.

import { describe, expect, test } from 'bun:test';
import { FILE_KINDS, KIND_EXTENSION, OPEN_ACCEPT, fileKindOf } from './file-kinds.ts';

describe('file kinds', () => {
  test('the kind and the macro-enabled flag give the file kind', () => {
    expect(fileKindOf({ workbookKind: null, macroEnabled: false })).toBeNull();
    expect(fileKindOf({ workbookKind: 'workbook', macroEnabled: false })).toBe('workbook');
    expect(fileKindOf({ workbookKind: 'template', macroEnabled: false })).toBe('template');
    expect(fileKindOf({ workbookKind: 'workbook', macroEnabled: true })).toBe('macroWorkbook');
    expect(fileKindOf({ workbookKind: 'template', macroEnabled: true })).toBe('macroTemplate');
  });

  test('each file kind has its extension, and the file input and the pattern take them all', () => {
    const exts = Object.values(FILE_KINDS).map((k) => k.ext);
    expect(exts).toEqual(['.xlsx', '.xltx', '.xlsm', '.xltm']);
    expect(OPEN_ACCEPT.split(',')).toEqual(exts);
    for (const ext of exts) expect(KIND_EXTENSION.test(`book${ext.toUpperCase()}`)).toBe(true);
    expect(KIND_EXTENSION.test('book.xlam')).toBe(false);
  });
});
