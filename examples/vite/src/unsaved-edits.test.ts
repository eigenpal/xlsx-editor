// The unsaved-edit state of the demo (`unsaved-edits.ts`).

import { describe, expect, test } from 'bun:test';
import { unsavedEdits, type EditFacts } from './unsaved-edits.ts';

/** An unsaved-edit state over snapshot facts that the test sets. */
function withFacts(start: Partial<EditFacts>) {
  let facts: EditFacts = {
    status: 'ready',
    revision: 0,
    editing: false,
    editChanged: false,
    ...start,
  };
  const state = unsavedEdits(() => facts);
  return { state, set: (next: Partial<EditFacts>) => (facts = { ...facts, ...next }) };
}

describe('unsavedEdits', () => {
  test('an edit after a clean mark is unsaved, and a new clean mark saves it', () => {
    const { state, set } = withFacts({});
    state.markClean(0);
    expect(state.hasUnsavedEdits()).toBe(false);
    set({ revision: 1 });
    expect(state.hasUnsavedEdits()).toBe(true);
    state.markClean(1);
    expect(state.hasUnsavedEdits()).toBe(false);
  });

  test('an open workbook with no edit is not unsaved before any clean mark', () => {
    expect(withFacts({}).state.hasUnsavedEdits()).toBe(false);
  });

  test('an open cell edit with changed text is unsaved, before any commit', () => {
    const { state, set } = withFacts({});
    state.markClean(0);
    // An open edit with unchanged text loses nothing.
    set({ editing: true });
    expect(state.hasUnsavedEdits()).toBe(false);
    set({ editChanged: true });
    expect(state.hasUnsavedEdits()).toBe(true);
    // A cancel drops the text: nothing is unsaved again.
    set({ editing: false, editChanged: false });
    expect(state.hasUnsavedEdits()).toBe(false);
  });

  test('a reopen that carried edits stays unsaved at its first revision', () => {
    const { state, set } = withFacts({ revision: 4 });
    set({ revision: 0 });
    state.markClean(0, true);
    expect(state.hasUnsavedEdits()).toBe(true);
  });

  test('no open workbook leaves no edit to ask about, until a workbook opens', () => {
    const { state, set } = withFacts({});
    state.markClean(0, true);
    set({ revision: 3 });
    expect(state.hasUnsavedEdits()).toBe(true);
    set({ status: 'failed', revision: 0 });
    expect(state.hasUnsavedEdits()).toBe(false);
    set({ status: 'loading' });
    expect(state.hasUnsavedEdits()).toBe(false);
  });

  test('a reopen that carries edits keeps them protected until it ends', () => {
    let facts: EditFacts = { status: 'loading', revision: 0, editing: false, editChanged: false };
    const state = unsavedEdits(() => facts);
    expect(state.hasUnsavedEdits()).toBe(false);
    state.holdEdits(true);
    expect(state.hasUnsavedEdits()).toBe(true);
    state.holdEdits(false);
    facts = { ...facts, status: 'ready' };
    state.markClean(0, true);
    expect(state.hasUnsavedEdits()).toBe(true);
  });
});
