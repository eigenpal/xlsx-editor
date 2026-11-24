// `Icon`: an inline Material Symbol, the icon set of the chrome, for a host's own controls beside
// the editor. The twin of the React adapter's `Icon.tsx`, with the same props and markup.

import { defineComponent, h, type ExtractPublicPropTypes, type PropType } from 'vue';
import { CHROME_ICON_PATHS, type ChromeIconName } from '@xlsx-editor.dev/core/editor';

const iconProps = {
  /** An icon of the chrome, such as `'undo'` or `'file_download'`. */
  name: { type: String as PropType<ChromeIconName>, default: undefined },
  /**
   * SVG path data in the 24 px Material Symbols view box (`0 -960 960 960`), in place of `name`.
   * Use it for an icon that the chrome does not have.
   */
  paths: { type: Array as PropType<readonly string[]>, default: undefined },
  /** Appended after the `xlsx-icon` class, which sets 18 by 18 pixels in the text color. */
  className: { type: String, default: undefined },
} as const;

/** Props for `Icon`, the twin of React's `IconProps`. @public */
export type IconProps = ExtractPublicPropTypes<typeof iconProps>;

/** The paths of a chrome icon. An unknown name, also a key such as `constructor`, draws nothing. */
const chromePaths = (name: string | undefined): readonly string[] =>
  name && Object.hasOwn(CHROME_ICON_PATHS, name) ? CHROME_ICON_PATHS[name as ChromeIconName] : [];

/**
 * An inline Material Symbol in the size and color of the toolbar icons. It is decorative
 * (`aria-hidden`): give the control around it its name.
 *
 * @public
 */
export const Icon = defineComponent({
  name: 'Icon',
  props: iconProps,
  setup(props) {
    return () => {
      const d = props.paths ?? chromePaths(props.name);
      return h(
        'svg',
        {
          class: ['xlsx-icon', props.className],
          viewBox: '0 -960 960 960',
          'aria-hidden': 'true',
          focusable: 'false',
        },
        d.map((path) => h('path', { d: path }))
      );
    };
  },
});
