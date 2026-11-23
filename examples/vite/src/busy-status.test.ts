import { describe, expect, test } from 'bun:test';
import { busyStatus } from './busy-status.ts';

const wait = (ms: number) => new Promise((resolve) => setTimeout(resolve, ms));

describe('the busy status of the demo', () => {
  test('shows nothing for short work', async () => {
    const shown: (string | null)[] = [];
    await busyStatus((t) => shown.push(t)).during('Saving', Promise.resolve(1));
    expect(shown).toEqual([null]);
  });

  test('shows the text during long work and clears it at the end', async () => {
    const shown: (string | null)[] = [];
    await busyStatus((t) => shown.push(t)).during('Saving', wait(400));
    expect(shown).toEqual(['Saving', null]);
  });

  test('clears the text when the work fails', async () => {
    const shown: (string | null)[] = [];
    const failing = wait(400).then(() => Promise.reject(new Error('no')));
    await busyStatus((t) => shown.push(t))
      .during('Saving', failing)
      .catch(() => undefined);
    expect(shown).toEqual(['Saving', null]);
  });

  test('keeps the text until the last work ends', async () => {
    const shown: (string | null)[] = [];
    const status = busyStatus((t) => shown.push(t));
    await Promise.all([status.during('Opening', wait(350)), status.during('Saving', wait(500))]);
    expect(shown).toEqual(['Saving', null]);
  });

  test('goes back to the work that still runs', async () => {
    const shown: (string | null)[] = [];
    const status = busyStatus((t) => shown.push(t));
    const opening = status.during('Opening', wait(700));
    await wait(350);
    await status.during('Saving', wait(100));
    await opening;
    expect(shown).toEqual(['Opening', 'Saving', 'Opening', null]);
  });
});
