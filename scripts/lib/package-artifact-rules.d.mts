export function bareImportRequests(content: string): string[];
export function packageNameOf(request: string): string;
export function relativeImportRequests(content: string): string[];
export function listingFieldProblems(manifest: {
  name: string;
  keywords?: unknown;
  homepage?: unknown;
}): string[];
export function unexportedRequestProblem(
  request: string,
  manifests: Map<string, { exports?: Record<string, unknown> }>
): string | null;
export function exposesWorkspacePath(content: string): boolean;
export function sideEffectsProblem(manifest: {
  name: string;
  sideEffects?: unknown;
  exports?: Record<string, unknown>;
}): string | null;
export function missingFileRemedy(file: string, onDisk: boolean, packageDir: string): string;
export function requiredLicenseFiles(names: string[]): string[];
export function hasEcmaLegalComment(content: string): boolean;
