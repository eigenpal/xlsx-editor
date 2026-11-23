// XLSX Editor only: the types of scripts/lib/declaration-files.mjs that its test uses. DOCX
// Editor's root tsconfig does not compile its script tests, so it ships no such file.

export declare function packageName(specifier: string): string;
export declare function declarationCandidates(target: string): string[];
export declare function defaultTypeRoots(dir: string): string[];
export declare function typescript7Compiler(packageDir: string): string;
export declare function packageDocumentationOf(source: string): string | null;
