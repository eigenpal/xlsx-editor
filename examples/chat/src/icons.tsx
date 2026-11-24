/** @jsxImportSource react */
// Material Symbols paths (Apache-2.0) for the chat shell, drawn as inline SVG (AGENTS.md,
// "Icons"). The React adapter does not export its own icon component, so the chat draws its
// icons here, at the size and fill of the editor's toolbar icons.

const PATHS = {
  arrowUp: 'M440-160v-487L216-423l-56-57 320-320 320 320-56 57-224-224v487h-80Z',
  /** Two corners that point out: Expand. */
  expand: 'M200-200v-240h80v160h160v80H200Zm480-480v160h80v-240H520v80h160Z',
  /** Two corners that point in: Show chat. */
  collapse: 'M440-440v240h-80v-160H200v-80h240Zm80-80v-240h80v160h160v80H520Z',
  editSquare:
    'M200-120q-33 0-56.5-23.5T120-200v-560q0-33 23.5-56.5T200-840h357l-80 80H200v560h560v-278l80-80v358q0 33-23.5 56.5T760-120H200Zm160-240v-170l367-367 170 170-367 367H360Zm80-80h56l232-232-28-28-29-28-231 231v57Z',
  spark:
    'M480-80q-6 0-11-4t-7-10q-17-67-51-126t-83-108q-49-49-108-83T94-462q-6-2-10-7t-4-11q0-6 4-11t10-7q67-17 126-51t108-83q49-49 83-108t51-126q2-6 7-10t11-4q6 0 10.5 4t6.5 10q18 67 52 126t83 108q49 49 108 83t126 51q6 2 10 7t4 11q0 6-4 11t-10 7q-67 17-126 51t-108 83q-49 49-83 108T498-94q-2 6-7 10t-11 4Z',
  /** table: a grid of cells, the workbook's mark. */
  table:
    'M200-120q-33 0-56.5-23.5T120-200v-560q0-33 23.5-56.5T200-840h560q33 0 56.5 23.5T840-760v560q0 33-23.5 56.5T760-120H200Zm240-240H200v160h240v-160Zm80 0v160h240v-160H520Zm-80-80v-160H200v160h240Zm80 0h240v-160H520v160ZM200-680h560v-80H200v80Z',
  close:
    'm256-200-56-56 224-224-224-224 56-56 224 224 224-224 56 56-224 224 224 224-56 56-224-224-224 224Z',
  back: 'M560-240 320-480l240-240 56 56-184 184 184 184-56 56Z',
  chevronRight: 'M504-480 320-664l56-56 240 240-240 240-56-56 184-184Z',
  download:
    'M480-320 280-520l56-58 104 104v-326h80v326l104-104 56 58-200 200ZM240-160q-33 0-56.5-23.5T160-240v-120h80v120h480v-120h80v120q0 33-23.5 56.5T720-160H240Z',
  lock: 'M240-80q-33 0-56.5-23.5T160-160v-400q0-33 23.5-56.5T240-640h40v-80q0-83 58.5-141.5T480-920q83 0 141.5 58.5T680-720v80h40q33 0 56.5 23.5T800-560v400q0 33-23.5 56.5T720-80H240Zm0-80h480v-400H240v400Zm240-120q33 0 56.5-23.5T560-360q0-33-23.5-56.5T480-440q-33 0-56.5 23.5T400-360q0 33 23.5 56.5T480-280ZM360-640h240v-80q0-50-35-85t-85-35q-50 0-85 35t-35 85v80Z',
} as const;

export type IconName = keyof typeof PATHS;

/** A decorative icon: the control around it carries the name. */
export function Icon({ name }: { name: IconName }) {
  return (
    <svg viewBox="0 -960 960 960" aria-hidden="true" focusable="false">
      <path d={PATHS[name]} />
    </svg>
  );
}
