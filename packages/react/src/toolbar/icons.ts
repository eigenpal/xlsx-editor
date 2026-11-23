// The icons that the toolbar draws beyond the registry's: the "More" button, which is not a
// chrome slot, the carets of fields and menu buttons, and the check mark of the selected list
// option. It reads core's Material Symbols set (`icon-paths.ts`) at run time, so this package
// holds no copy of those icons. As in pptx-editor/packages/react/src/FontFamilyControl.tsx and
// Toolbar.tsx at d4502552 (EigenPal), a field takes `expand_more` and a menu button takes
// `arrow_drop_down`.

import { CHROME_ICON_PATHS } from '@xlsx-editor.dev/core/editor';

export const ICON_MORE: readonly string[] = CHROME_ICON_PATHS.more_horiz;
export const ICON_FIELD_CARET: readonly string[] = CHROME_ICON_PATHS.expand_more;
export const ICON_MENU_CARET: readonly string[] = CHROME_ICON_PATHS.arrow_drop_down;
export const ICON_CHECK: readonly string[] = CHROME_ICON_PATHS.check;
