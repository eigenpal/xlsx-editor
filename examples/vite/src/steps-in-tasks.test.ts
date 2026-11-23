// The demo's step loop: it gives the result of the steps, yields to the page between slices,
// passes on what a step throws, and stops the steps that are no longer wanted.

import { describe, expect, test } from 'bun:test';
import { runInTasks, whileWanted } from './steps-in-tasks.ts';

/** The steps that `busySteps` did. */
let stepsDone = 0;

/** Steps that each take `ms` milliseconds, `count` times, and then return `count`. */
function* busySteps(count: number, ms: number): Generator<void, number> {
  for (let i = 0; i < count; i++) {
    const end = performance.now() + ms;
    while (performance.now() < end) {
      // Busy work of one step.
    }
    stepsDone++;
    yield;
  }
  return count;
}

describe('runInTasks', () => {
  test('gives the result of the steps', async () => {
    expect(await runInTasks(busySteps(3, 0))).toBe(3);
  });

  test('gives another task a turn before the steps end', async () => {
    stepsDone = 0;
    const done = runInTasks(busySteps(20, 5));
    const atTurn = new Promise<number>((resolve) => {
      const channel = new MessageChannel();
      channel.port1.onmessage = () => {
        channel.port1.close();
        resolve(stepsDone);
      };
      channel.port2.postMessage(null);
    });
    expect(await atTurn).toBeLessThan(20);
    expect(await done).toBe(20);
  });

  test('rejects with what a step throws', async () => {
    function* failing(): Generator<void, number> {
      yield;
      throw new Error('step failed');
    }
    await expect(runInTasks(failing())).rejects.toThrow('step failed');
  });
});

describe('whileWanted', () => {
  test('gives the result while the work is wanted', async () => {
    expect(await runInTasks(whileWanted(busySteps(3, 0), () => true))).toBe(3);
  });

  test('ends with null and closes the steps when the work is no longer wanted', async () => {
    let closed = false;
    function* steps(): Generator<void, number> {
      try {
        for (;;) yield;
      } finally {
        closed = true;
      }
    }
    expect(await runInTasks(whileWanted(steps(), () => false))).toBeNull();
    expect(closed).toBe(true);
  });
});
