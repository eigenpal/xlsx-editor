import { describe, expect, test } from 'bun:test';
import {
  MEMORY_CAP_BYTES,
  parseProcessTable,
  treeBytes,
  treeOf,
  watchBrowser,
  watchProcess,
  type Watch,
} from './watchdog.ts';

const TABLE = `
  10     1   100
  11    10   200
  12    11   300
  20     1  9000
  bad line
`;

describe('the process table of the hang watchdog', () => {
  test('reads each line with three numbers and leaves out the others', () => {
    expect(parseProcessTable(TABLE).map((r) => r.pid)).toEqual([10, 11, 12, 20]);
  });

  test('takes only the root and its descendants', () => {
    const tree = treeOf(parseProcessTable(TABLE), 10);
    expect(tree.map((r) => r.pid).sort()).toEqual([10, 11, 12]);
    expect(treeBytes(tree)).toBe(600 * 1024);
  });

  test('gives no process for a root that ended', () => {
    expect(treeOf(parseProcessTable(TABLE), 99)).toEqual([]);
  });

  test('ends on a parent cycle in the table', () => {
    const rows = [
      { pid: 1, ppid: 2, rssKb: 1 },
      { pid: 2, ppid: 1, rssKb: 1 },
    ];
    expect(treeOf(rows, 1)).toHaveLength(2);
  });

  test('keeps the memory cap of one case at 3 GiB', () => {
    expect(MEMORY_CAP_BYTES).toBe(3 * 1024 ** 3);
  });
});

describe('the watch over a browser', () => {
  /** Waits until `check` is true, for up to 5 seconds. */
  async function until(check: () => boolean): Promise<void> {
    const end = Date.now() + 5000;
    while (!check() && Date.now() < end) await Bun.sleep(10);
  }

  test('stops a process that passes the time of the case', async () => {
    const child = Bun.spawn(['sleep', '30']);
    const watch = watchBrowser(child.pid, 0, 10);
    await child.exited;
    watch.stop();
    expect(watch.tripped()).toContain('the case passed');
    expect(watch.trippedMemory()).toBe(false);
    expect(child.signalCode).toBe('SIGKILL');
  });

  test('stops a process that passes the memory cap', async () => {
    const child = Bun.spawn(['sleep', '30']);
    const watch = watchBrowser(child.pid, 60_000, 10, 1);
    await child.exited;
    watch.stop();
    expect(watch.trippedMemory()).toBe(true);
    expect(watch.peakBytes()).toBeGreaterThan(1);
  });

  test('stops a process once, and keeps the first reason', async () => {
    const child = Bun.spawn(['sleep', '30']);
    // The memory cap trips first. The ticks after it pass the time of the case.
    const watch = watchBrowser(child.pid, 200, 10, 1);
    await child.exited;
    await Bun.sleep(400);
    watch.stop();
    expect(watch.trippedMemory()).toBe(true);
  });

  test('leaves a process alone that stays inside its limits', async () => {
    const child = Bun.spawn(['sleep', '30']);
    const watch = watchBrowser(child.pid, 60_000, 10);
    await until(() => watch.peakBytes() > 0);
    watch.stop();
    expect(watch.tripped()).toBeNull();
    expect(child.killed).toBe(false);
    child.kill();
    await child.exited;
  });

  /**
   * A watch with its own clock and an instant read of the process table, so each read happens
   * at one clock time and the clock moves only between reads.
   */
  function clockedWatch(pid: number) {
    const clock = { ms: 0 };
    const watch = watchProcess(pid, {
      deadlineMs: 1000,
      everyMs: 10,
      now: () => clock.ms,
      readTable: async () => [{ pid, ppid: 1, rssKb: 1 }],
    });
    return { clock, watch };
  }

  /**
   * Moves the clock 100 ms at a time, with reads between the steps, for at most 30 steps. A watch
   * that counts the time between reads stops the process after about 11 steps.
   */
  async function runUntilTripped(clock: { ms: number }, watch: Watch): Promise<void> {
    for (let step = 0; step < 30 && watch.tripped() === null; step += 1) {
      clock.ms += 100;
      await Bun.sleep(30);
    }
  }

  test('does not count the time of a late read, as after a sleep of the machine', async () => {
    const child = Bun.spawn(['sleep', '30']);
    const { clock, watch } = clockedWatch(child.pid);
    await until(() => watch.peakBytes() > 0);
    // The machine sleeps for a minute between two reads.
    clock.ms += 60_000;
    await Bun.sleep(100);
    expect(watch.tripped()).toBeNull();
    // Time that passes between reads on plan counts.
    await runUntilTripped(clock, watch);
    watch.stop();
    expect(watch.tripped()).toBe('the work passed 1000 ms');
    await child.exited;
  });

  test('counts the time after a restart that came during a late read', async () => {
    const child = Bun.spawn(['sleep', '30']);
    const { clock, watch } = clockedWatch(child.pid);
    await until(() => watch.peakBytes() > 0);
    clock.ms += 60_000;
    // A new unit of work starts before the late read: the sleep is not added to its time.
    watch.restart();
    await runUntilTripped(clock, watch);
    watch.stop();
    expect(watch.tripped()).toBe('the work passed 1000 ms');
    await child.exited;
  });

  test('does not stop a process after a stop that came during a read', async () => {
    const child = Bun.spawn(['sleep', '30']);
    // Every read trips the time limit. The first read starts at once and takes some ms.
    const watch = watchProcess(child.pid, { deadlineMs: 0, everyMs: 10 });
    await Bun.sleep(2);
    watch.stop();
    await Bun.sleep(300);
    expect(watch.tripped()).toBeNull();
    expect(child.killed).toBe(false);
    child.kill();
    await child.exited;
  });

  test('starts the time again on a restart, and names the unit of work', async () => {
    const child = Bun.spawn(['sleep', '30']);
    const watch = watchProcess(child.pid, { deadlineMs: 500, everyMs: 10, unit: 'the workbook' });
    for (let i = 0; i < 8; i += 1) {
      await Bun.sleep(100);
      watch.restart();
    }
    // 800 ms passed, but never 500 ms since the last restart.
    expect(watch.tripped()).toBeNull();
    await child.exited;
    watch.stop();
    expect(watch.tripped()).toBe('the workbook passed 500 ms');
  });
});
