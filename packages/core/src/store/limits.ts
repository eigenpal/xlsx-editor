// The resource limits of a workbook read. They come from the shared OOXML layer, so the three
// editors refuse the same oversized input with the same reasons.

export { DEFAULT_LIMITS, resolveLimits, type ResourceLimits } from '@eigenpal/ooxml-core';
