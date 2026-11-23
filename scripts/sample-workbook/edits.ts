// The edits of the sample workbook's edited copy. The sample save check (`save-check.ts`) makes
// them, and the demo save tests (`examples/vite/src/__tests__/demo-save.test.tsx`) make the Bold
// edit in the demo, so both check the same cell.
// This module has no side effects, unlike `save-check.ts`, which sets the time zone of the
// process.

/** The format edit of the edited copy: the Bold edit that the demo save tests make. */
export const SAMPLE_BOLD_EDIT = { sheet: 'Sales', ref: 'A5' } as const;

/** The value edit of the edited copy, which the local check reads back. */
export const SAMPLE_EDIT = { sheet: 'Sales', ref: 'B5', value: 4321 } as const;
