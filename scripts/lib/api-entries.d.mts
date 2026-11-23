export function unlistedEntries(
  pkg: { readonly entries: readonly string[] },
  exportsMap: Record<string, unknown>
): string[];
export function orphanReports(
  packages: readonly { readonly slug: string; readonly entries: readonly string[] }[],
  reports: ReadonlyMap<string, readonly string[]>
): string[];
export function packageArgumentProblem(
  args: readonly string[],
  packages: readonly { name: string }[]
): string | null;
export function reportsBelow(apiRoot: string): Map<string, string[]>;
