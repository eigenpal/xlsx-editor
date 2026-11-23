// Copied from docx-editor/scripts/declaration-options.d.mts at f098ae68 (EigenPal). No differences.
export declare function declarationCompilerOptions(
  configUrl: string | URL,
  extra?: Record<string, unknown>
): Record<string, unknown> & { paths: Record<string, string[]>; pathsBase: string };
