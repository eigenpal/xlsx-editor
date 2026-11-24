/** @jsxImportSource react */
// `Icon`: an inline Material Symbol, the icon set of the chrome, for a host's own controls beside
// the editor, such as the buttons of a panel header. After `Icon` of
// pptx-editor/packages/react/src/chrome.tsx at cfb7c234 (EigenPal). Differences: `name` is typed
// as `ChromeIconName`, so an unknown name is a type error and draws nothing, with no run-time
// warning; the default class is `xlsx-icon`, which the core stylesheet sizes; and the `svg` is
// not focusable.

import type { ChromeIconName } from '@xlsx-editor.dev/core/editor';
import { CHROME_ICON_PATHS } from '@xlsx-editor.dev/core/editor';

/** Props for `Icon`. @public */
export interface IconProps {
  /** An icon of the chrome, such as `'undo'` or `'file_download'`. */
  readonly name?: ChromeIconName;
  /**
   * SVG path data in the 24 px Material Symbols view box (`0 -960 960 960`), in place of `name`.
   * Use it for an icon that the chrome does not have.
   */
  readonly paths?: readonly string[];
  /** Appended after the `xlsx-icon` class, which sets 18 by 18 pixels in the text color. */
  readonly className?: string;
}

/** The paths of a chrome icon. An unknown name, also a key such as `constructor`, draws nothing. */
const chromePaths = (name: string | undefined): readonly string[] =>
  name && Object.hasOwn(CHROME_ICON_PATHS, name) ? CHROME_ICON_PATHS[name as ChromeIconName] : [];

/**
 * An inline Material Symbol in the size and color of the toolbar icons. It is decorative
 * (`aria-hidden`): give the control around it its name.
 *
 * @example
 * ```tsx
 * <XlsxEditor.Toolbar.Action label="Download" icon={<Icon name="file_download" />} onSelect={save} />
 * ```
 *
 * @public
 */
export function Icon({ name, paths, className }: IconProps) {
  const d = paths ?? chromePaths(name);
  return (
    <svg
      className={className ? `xlsx-icon ${className}` : 'xlsx-icon'}
      viewBox="0 -960 960 960"
      aria-hidden="true"
      focusable="false"
    >
      {d.map((path, i) => (
        <path key={i} d={path} />
      ))}
    </svg>
  );
}
