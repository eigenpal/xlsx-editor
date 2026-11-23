// Types for scripts/lib/docs-style.mjs, so its test type-checks. XLSX Editor addition.
export function proseLines(markdown: string): string[];
export function checkDocsStyle(markdown: string): { line: number; message: string }[];
export function checkPackageReadme(
  markdown: string,
  manifest: { name?: string; private?: boolean }
): string[];
