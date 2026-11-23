/** @jsxImportSource react */
// The icons of the demo shell. The path data is in `icon-paths.ts`, which the Vue demo shares.
// The JSX import source pragma on the first line is added, as every React file here has it.

import { PATHS, type IconName } from './icon-paths.ts';

/** A shell icon: decorative, so the control around it carries the name. */
export function Icon({ name }: { name: IconName }) {
  return (
    <svg viewBox="0 -960 960 960" aria-hidden="true" focusable="false">
      <path d={PATHS[name]} />
    </svg>
  );
}
