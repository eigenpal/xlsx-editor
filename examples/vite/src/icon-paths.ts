// The path data of the demo shell icons, for the React icons (`icons.tsx`) and the Vue icon
// (examples/vue/src/Icon.vue). Material Symbols path data (Apache-2.0, Google): Outlined,
// weight 400, grade 0, 24 px, viewBox "0 -960 960 960", taken from
// pptx-editor/packages/core/src/editor/icon-paths.ts at 5b6ead87 (EigenPal). Save uses the
// editor's own `file_download`, as the toolbar Save does. Never hand-draw or edit a path here:
// every entry is an official Material Symbol.

import { CHROME_ICON_PATHS } from '@xlsx-editor.dev/core/editor';

/** The path data of each shell icon. */
export const PATHS = {
  folder_open:
    'M160-160q-33 0-56.5-23.5T80-240v-480q0-33 23.5-56.5T160-800h240l80 80h320q33 0 56.5 23.5T880-640H447l-80-80H160v480l96-320h684L837-217q-8 26-29.5 41.5T760-160H160Zm84-80h516l72-240H316l-72 240Zm0 0 72-240-72 240Zm-84-400v-80 80Z',
  visibility:
    'M480-320q75 0 127.5-52.5T660-500q0-75-52.5-127.5T480-680q-75 0-127.5 52.5T300-500q0 75 52.5 127.5T480-320Zm0-72q-45 0-76.5-31.5T372-500q0-45 31.5-76.5T480-608q45 0 76.5 31.5T588-500q0 45-31.5 76.5T480-392Zm0 192q-146 0-266-81.5T40-500q54-137 174-218.5T480-800q146 0 266 81.5T920-500q-54 137-174 218.5T480-200Zm0-300Zm0 220q113 0 207.5-59.5T832-500q-50-101-144.5-160.5T480-720q-113 0-207.5 59.5T128-500q50 101 144.5 160.5T480-280Z',
  table_chart:
    'M760-120H200q-33 0-56.5-23.5T120-200v-560q0-33 23.5-56.5T200-840h560q33 0 56.5 23.5T840-760v560q0 33-23.5 56.5T760-120ZM200-640h560v-120H200v120Zm100 80H200v360h100v-360Zm360 0v360h100v-360H660Zm-80 0H380v360h200v-360Z',
  close:
    'm256-200-56-56 224-224-224-224 56-56 224 224 224-224 56 56-224 224 224 224-56 56-224-224-224 224Z',
  text_fields:
    'M280-160v-520H80v-120h520v120H400v520H280Zm360 0v-320H520v-120h360v120H760v320H640Z',
  file_download: CHROME_ICON_PATHS.file_download[0],
} as const;

/** The name of a shell icon. */
export type IconName = keyof typeof PATHS;
