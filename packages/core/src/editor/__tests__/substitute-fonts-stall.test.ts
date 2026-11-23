import { describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import {
  stallLimitedFetch,
  SUBSTITUTE_FONT_LATER_STALL_PART,
  SUBSTITUTE_FONT_STALL_MS,
  SUBSTITUTE_FONT_TOTAL_MS,
} from '../substitute-fonts.ts';

/** A fetch whose body sends `chunks`, one each `gapMs`, and then stays open when `stall` is set. */
function bodyFetch(chunks: readonly string[], gapMs: number, stall = false) {
  return async (_input: unknown, init?: RequestInit) => {
    const enc = new TextEncoder();
    let at = 0;
    const body = new ReadableStream<Uint8Array>({
      async pull(controller) {
        if (at < chunks.length) {
          await Bun.sleep(gapMs);
          controller.enqueue(enc.encode(chunks[at]!));
          at += 1;
          return;
        }
        if (!stall) return controller.close();
        await new Promise<void>((_, reject) => {
          init?.signal?.addEventListener('abort', () => reject(init.signal!.reason));
        });
      },
    });
    return new Response(body, { status: 200 });
  };
}

/** A fetch that never answers until its signal aborts. */
const silentFetch = (_input: unknown, init?: RequestInit) =>
  new Promise<Response>((_, reject) => {
    init?.signal?.addEventListener('abort', () => reject(init.signal!.reason));
  });

/** A fetch that answers at once with `text`. */
const quickFetch = (text: string) => async () => new Response(text, { status: 200 });

describe('stallLimitedFetch', () => {
  test('fails a request that gets no response within the limit', async () => {
    const limited = stallLimitedFetch(silentFetch, 20);
    await expect(limited('https://fonts.invalid/a.ttf')).rejects.toThrow('no bytes for 20 ms');
  });

  test('fails a body that stops sending bytes', async () => {
    const limited = stallLimitedFetch(bodyFetch(['ab'], 1, true), 20);
    await expect(limited('https://fonts.invalid/a.ttf')).rejects.toThrow('no bytes for 20 ms');
  });

  test('keeps a slow body that still sends bytes, longer than the limit in all', async () => {
    const limited = stallLimitedFetch(bodyFetch(['aa', 'bb', 'cc', 'dd', 'ee'], 15), 40);
    const response = await limited('https://fonts.invalid/a.ttf');
    expect(await response.text()).toBe('aabbccddee');
    expect(response.status).toBe(200);
  });

  test('passes on an abort of the caller', async () => {
    const caller = new AbortController();
    const limited = stallLimitedFetch(silentFetch, 10_000);
    const pending = limited('https://fonts.invalid/a.ttf', { signal: caller.signal });
    caller.abort(new Error('closed by the host'));
    await expect(pending).rejects.toThrow('closed by the host');
  });

  test('keeps the status of a failed response', async () => {
    const limited = stallLimitedFetch(async () => new Response('gone', { status: 404 }), 100);
    expect((await limited('https://fonts.invalid/a.ttf')).status).toBe(404);
  });

  test('ends a fetch and a body that ignore the abort', async () => {
    const deaf = () => new Promise<Response>(() => {});
    await expect(stallLimitedFetch(deaf, 20)('https://fonts.invalid/a')).rejects.toThrow(
      'no bytes for 20 ms'
    );
    const deafBody = async () =>
      new Response(new ReadableStream({ pull: () => new Promise(() => {}) }));
    await expect(stallLimitedFetch(deafBody, 20)('https://fonts.invalid/b')).rejects.toThrow(
      'no bytes for 20 ms'
    );
  });

  test('runs no more downloads than the limit, and a queued face does not time out', async () => {
    let active = 0;
    let most = 0;
    const slow = async () => {
      active += 1;
      most = Math.max(most, active);
      await Bun.sleep(15);
      active -= 1;
      return new Response('face');
    };
    // Six groups of 15 ms pass the 40 ms limit in all, and every download still sends in time.
    const limited = stallLimitedFetch(slow, 40, 2);
    const all = await Promise.all(Array.from({ length: 12 }, (_, i) => limited(`https://f/${i}`)));
    expect(all.map((r) => r.status)).toEqual(Array.from({ length: 12 }, () => 200));
    expect(most).toBe(2);
  });

  test('fails the whole queue after one limit when no byte ever comes', async () => {
    let calls = 0;
    const counted = (input: unknown, init?: RequestInit) => {
      calls += 1;
      return silentFetch(input, init);
    };
    const limited = stallLimitedFetch(counted, 30, 2);
    const started = performance.now();
    const results = await Promise.allSettled(
      Array.from({ length: 12 }, (_, i) => limited(`https://f/${i}`))
    );
    expect(results.every((r) => r.status === 'rejected')).toBe(true);
    // Six groups would take 180 ms. One limit ends them all.
    expect(performance.now() - started).toBeLessThan(120);
    // The second group shows that no byte comes, so the faces behind it never start.
    expect(calls).toBe(4);
  });

  test('keeps the healthy faces that wait behind stalled ones', async () => {
    let calls = 0;
    const mixed = (input: unknown, init?: RequestInit) => {
      calls += 1;
      return String(input).includes('stall') ? silentFetch(input, init) : quickFetch('face')();
    };
    const limited = stallLimitedFetch(mixed, 30, 2);
    const urls = ['ok-0', 'stall-1', 'stall-2', 'ok-3', 'ok-4'];
    const results = await Promise.allSettled(urls.map((u) => limited(`https://f/${u}`)));
    expect(results.map((r) => r.status)).toEqual([
      'fulfilled',
      'rejected',
      'rejected',
      'fulfilled',
      'fulfilled',
    ]);
    expect(calls).toBe(5);
  });

  test('shortens the limit after the first stall when stalled faces fill the slots', async () => {
    const mixed = (input: unknown, init?: RequestInit) =>
      String(input).includes('stall') ? silentFetch(input, init) : quickFetch('face')();
    const limited = stallLimitedFetch(mixed, 200, 2);
    const urls = Array.from({ length: 12 }, (_, i) => `https://f/${i % 2 ? 'stall' : 'ok'}-${i}`);
    const started = performance.now();
    const results = await Promise.allSettled(urls.map((u) => limited(u)));
    expect(results.filter((r) => r.status === 'fulfilled')).toHaveLength(6);
    // A full limit for each group of stalled faces would take 600 ms or more.
    expect(performance.now() - started).toBeLessThan(500);
  });

  test('fails a silent download beside one that still sends bytes', async () => {
    const mixed = (input: unknown, init?: RequestInit) =>
      String(input).includes('silent')
        ? silentFetch(input, init)
        : bodyFetch(
            Array.from({ length: 16 }, () => 'a'),
            20
          )(input, init);
    // The slow body runs for 320 ms, and its gaps stay inside the shorter limit of 50 ms.
    const limited = stallLimitedFetch(mixed, 200, 2);
    const started = performance.now();
    const silent = limited('https://f/silent').catch(() => performance.now() - started);
    const slow = limited('https://f/slow');
    expect(await silent).toBeLessThan(300);
    expect(await (await slow).text()).toBe('a'.repeat(16));
  });

  test('downloads the healthy faces that wait behind one group of stalled faces', async () => {
    const mixed = (input: unknown, init?: RequestInit) =>
      String(input).includes('stall') ? silentFetch(input, init) : quickFetch('face')();
    const limited = stallLimitedFetch(mixed, 30, 2);
    const urls = ['stall-0', 'stall-1', 'ok-2', 'ok-3', 'ok-4', 'ok-5'];
    const results = await Promise.allSettled(urls.map((u) => limited(`https://f/${u}`)));
    expect(results.map((r) => r.status)).toEqual([
      'rejected',
      'rejected',
      'fulfilled',
      'fulfilled',
      'fulfilled',
      'fulfilled',
    ]);
  });

  test('ends a download that sends bytes for longer than the total limit', async () => {
    const endless = bodyFetch(
      Array.from({ length: 1000 }, () => 'a'),
      10
    );
    const limited = stallLimitedFetch(endless, 100, 2, 80);
    await expect(limited('https://f/slow')).rejects.toThrow('the downloads passed 80 ms');
  });

  test('takes a queued download out when its caller aborts', async () => {
    const caller = new AbortController();
    const limited = stallLimitedFetch(silentFetch, 10_000, 1);
    const first = limited('https://f/first', { signal: caller.signal });
    const second = limited('https://f/second', { signal: caller.signal });
    caller.abort(new Error('closed by the host'));
    const results = await Promise.allSettled([first, second]);
    expect(results.map((r) => (r as PromiseRejectedResult).reason?.message)).toEqual([
      'closed by the host',
      'closed by the host',
    ]);
  });
});

describe('the font download limits of the React layout guide', () => {
  test('give the values of the constants', () => {
    const readme = readFileSync(
      new URL('../../../../react/docs/layout-and-props.md', import.meta.url),
      'utf8'
    ).replaceAll('\n', ' ');
    const seconds = (ms: number) => `${ms / 1000} seconds`;
    expect(readme).toContain(`no byte for ${seconds(SUBSTITUTE_FONT_STALL_MS)}`);
    expect(readme).toContain(
      `the limit for the others is ${seconds(SUBSTITUTE_FONT_STALL_MS * SUBSTITUTE_FONT_LATER_STALL_PART)}`
    );
    expect(readme).toContain(`end after ${seconds(SUBSTITUTE_FONT_TOTAL_MS)}`);
  });
});
