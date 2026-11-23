// The name of the sample workbook in the demo's workbook list. It has its own module with no
// `import.meta.glob`, so `scripts/verify-demo.ts` and the demo tests can import it under Bun.

/** The name of the sample workbook that the demo opens at start (`bun run sample:build`). */
export const SAMPLE = 'sample';
