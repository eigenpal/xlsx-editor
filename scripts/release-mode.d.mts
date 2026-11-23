// Adapted from pptx-editor/scripts/release-mode.d.mts at 8e843068 (EigenPal). Differences: no
// canary declarations, and `prereleaseTag`.
export interface PublicPackage {
  name: string;
  version: string;
}

export declare function pendingChangesets(root: string): string[];
export declare function publicPackages(root: string): PublicPackage[];
export declare function prereleaseTag(root: string): string | null;
export declare function releaseMode(root: string): { mode: 'pending' | 'publish' };
