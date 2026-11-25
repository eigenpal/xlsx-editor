export interface PlannedRelease {
  kind: 'publish' | 'tag-only';
  name: string;
  version: string;
  access?: string;
  tag?: string;
}

export declare function isStable(version: string): boolean;
export declare function retargetPlan(
  plan: PlannedRelease[][],
  preTag: string | null,
  hasStableVersion: (name: string) => boolean
): PlannedRelease[][];
