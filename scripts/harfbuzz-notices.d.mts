// Copied from pptx-editor/scripts/harfbuzz-notices.d.mts at d4502552 (EigenPal), with no difference.
import manifest from '../packages/core/licenses/harfbuzz-manifest.json';

export const HARFBUZZ_MANIFEST: typeof manifest;
export const HARFBUZZ_NOTICE: string;
export function validateHarfBuzzNotices(options?: {
  coreRoot?: string;
  dependencyRoot?: string;
}): void;
