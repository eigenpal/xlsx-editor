// The key of the sample workbook in the demo. It has its own module with no `import.meta.env`,
// so `scripts/verify-demo.ts` and the demo tests can import it under Bun.

/** The key of the sample workbook that the demo opens at start (`bun run sample:build`). */
export const SAMPLE = 'sample';
