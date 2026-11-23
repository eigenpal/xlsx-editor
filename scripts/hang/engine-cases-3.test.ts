// Share 3 of 6 of the engine run of the hostile workbooks (`engine-run.ts`).

import { describe } from 'bun:test';
import { engineTests } from './engine-tests.ts';

describe('the hostile workbooks through the engine, share 3 of 6', () => engineTests(3, 6));
