// The problems that the cascade lane reports. A problem never stops a render: the cascade
// falls back to a default, and the caller records the problem with the cell reference.

/** The reason of a cascade problem. */
export type CascadeProblemCode =
  // Stored values (values.ts).
  | 'shared-string-out-of-range'
  | 'bad-number'
  | 'bad-boolean'
  | 'unknown-cell-type'
  // Cell formats (cell-format.ts).
  | 'cell-format-out-of-range'
  | 'number-format-id-malformed'
  | 'font-out-of-range'
  | 'fill-out-of-range'
  | 'border-out-of-range'
  | 'gradient-rectangle-out-of-range'
  // Number format codes (number-format/), which fall back to `General`.
  | 'format-code-too-long'
  | 'format-too-many-sections'
  | 'format-too-many-brackets'
  | 'format-code-unparsable';

/** A problem of the cascade lane. `detail` is a short, capped excerpt of the file value. */
export interface CascadeProblem {
  readonly code: CascadeProblemCode;
  readonly detail?: string;
}
