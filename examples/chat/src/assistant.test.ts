// The scripted assistant: each command maps to one editor action, and a workbook context
// with no workbook always drafts one first.

import { describe, expect, test } from 'bun:test';
import { FIRST_WORKBOOK_TEXT, reply, type WorkbookContext } from './assistant.ts';

const BOOK: WorkbookContext = {
  hasWorkbook: true,
  sheets: [
    { index: 0, name: 'Welcome' },
    { index: 1, name: 'Sales' },
    { index: 2, name: 'Sales 2026' },
  ],
  activeSheet: 'Sales',
  selection: 'B5',
};

describe('reply', () => {
  test('drafts a workbook when the chat has none', () => {
    const answer = reply('go to Sales', { ...BOOK, hasWorkbook: false });
    expect(answer).toEqual({ text: FIRST_WORKBOOK_TEXT, workbook: true });
  });

  test('writes a value to a cell, in both word orders', () => {
    expect(reply('set b5 to 1200', BOOK).action).toEqual({
      kind: 'write',
      ref: 'B5',
      text: '1200',
    });
    expect(reply('put "Done" in C3', BOOK).action).toEqual({
      kind: 'write',
      ref: 'C3',
      text: 'Done',
    });
  });

  test('goes to the sheet that the text names, the longest name first', () => {
    expect(reply('go to Sales', BOOK).action).toEqual({ kind: 'sheet', index: 1 });
    expect(reply('open sales 2026 please', BOOK).action).toEqual({ kind: 'sheet', index: 2 });
    // A name inside a longer word is no match.
    expect(reply('show the salesman list', BOOK).action).toEqual({ kind: 'open' });
  });

  test('selects a cell or a range', () => {
    expect(reply('select B5:E8', BOOK).action).toEqual({ kind: 'select', ref: 'B5:E8' });
  });

  test('maps the commands of the panel', () => {
    expect(reply('bold', BOOK).action).toEqual({ kind: 'format', toggle: 'bold' });
    expect(reply('add a sheet', BOOK).action).toEqual({ kind: 'addSheet' });
    expect(reply('undo', BOOK).action).toEqual({ kind: 'undo' });
    expect(reply('zoom in', BOOK).action).toEqual({ kind: 'zoom', direction: 1 });
    expect(reply('make it view only', BOOK).action).toEqual({ kind: 'mode', mode: 'view' });
    expect(reply('allow edits', BOOK).action).toEqual({ kind: 'mode', mode: 'edit' });
    expect(reply('download it', BOOK).action).toEqual({ kind: 'download' });
  });

  test('answers with no action when it does not know the command', () => {
    const answer = reply('what is this?', BOOK);
    expect(answer.action).toBeUndefined();
    expect(answer.text).toContain('Sales');
  });
});
