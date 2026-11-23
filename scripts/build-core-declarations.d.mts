// XLSX Editor only: the types of scripts/build-core-declarations.mjs, for its test. DOCX Editor's
// root tsconfig does not compile its script tests, so it ships no such file.

/** Private workspace packages that core inlines; their types are bundled. */
export declare const INLINED: readonly string[];

/** One declaration entry for each published subpath. */
export declare function publishedEntries(
  manifest: { exports: Record<string, unknown> },
  paths: Record<string, string[]>
): { specifier: string; name: string; source: string }[];

export declare function declarationCandidates(target: string): string[];
export declare function packageName(specifier: string): string;
