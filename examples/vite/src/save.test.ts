// The save of the demo: a failed save shows an error, and a good save after it clears it; the
// saved name gets one `-edited` end, and a name that the user gave is the file name; a browser with no picker downloads the copy.

import { describe, expect, spyOn, test } from 'bun:test';
import {
  downloadNotice,
  keptSaveHolds,
  SAVE_FAILED,
  SAVE_READY,
  SAVE_REPEAT_MS,
  savedName,
  saveFileName,
  savedNoticeStays,
  saveRefused,
  saveRepeatGuard,
  saveFile,
  type SavePicker,
} from './save.ts';

const OPTIONS = { suggestedName: 'a-edited.xlsx', types: [] };

describe('the save of the demo', () => {
  test('a picker that the browser refuses after the click keeps the bytes for one more press', async () => {
    const late: SavePicker = () => Promise.reject(new DOMException('late', 'SecurityError'));
    const outcome = await saveFile(late, new Uint8Array(3), OPTIONS, true);
    expect(outcome).toEqual({ written: false, error: null, notice: SAVE_READY, late: true });
  });

  test('a security refusal with a recent click, or after the picker, is a failure', async () => {
    const error = spyOn(console, 'error').mockImplementation(() => {});
    const refused = () => Promise.reject(new DOMException('frame', 'SecurityError'));
    const active = () => true;
    const download = () => {};
    const sandboxed: SavePicker = refused;
    const failed = { written: false, error: SAVE_FAILED, notice: null };
    expect(await saveFile(sandboxed, new Uint8Array(3), OPTIONS, true, download, active)).toEqual(
      failed
    );
    const writeRefused: SavePicker = async () => ({ createWritable: refused });
    expect(
      await saveFile(writeRefused, new Uint8Array(3), OPTIONS, true, download, () => false)
    ).toEqual(failed);
    error.mockRestore();
  });

  test('a failed save shows an error, and the next good save clears it', async () => {
    const cause = new Error('disk full');
    const failing: SavePicker = () => Promise.reject(cause);
    // The cause goes to the console; the user sees the short text.
    const logged = spyOn(console, 'error').mockImplementation(() => {});
    const failed = await saveFile(failing, new Uint8Array(3), OPTIONS, true);
    expect(logged).toHaveBeenCalledWith('The save of the demo failed:', cause);
    logged.mockRestore();
    expect(failed).toEqual({ written: false, error: SAVE_FAILED, notice: null });
    expect(SAVE_FAILED).not.toContain('disk full');
    const written: Blob[] = [];
    const good: SavePicker = () =>
      Promise.resolve({
        name: 'a-edited.xlsx',
        createWritable: () =>
          Promise.resolve({
            write: (b: Blob) => (written.push(b), Promise.resolve()),
            close: () => Promise.resolve(),
          }),
      });
    const saved = await saveFile(good, new Uint8Array(3), OPTIONS, true);
    expect(saved).toEqual({
      written: true,
      error: null,
      notice: 'Saved 3 bytes to a-edited.xlsx.',
    });
    expect(written.length).toBe(1);
    // The notice shows the picked name without its bidirectional format controls.
    const renamed: SavePicker = (options) =>
      good(options).then((handle) => ({ ...handle, name: 'a‮xslx.exe' }));
    expect((await saveFile(renamed, new Uint8Array(3), OPTIONS, true)).notice).toBe(
      'Saved 3 bytes to axslx.exe.'
    );
    // A cancel of the picker shows nothing.
    const abort = Object.assign(new Error('cancel'), { name: 'AbortError' });
    const cancelled = await saveFile(() => Promise.reject(abort), new Uint8Array(3), OPTIONS, true);
    expect(cancelled).toEqual({ written: false, error: null, notice: null });
  });

  test('a refused save in the snapshot clears the saved notice', () => {
    expect(saveRefused({ refusal: { command: 'save' } })).toBe(true);
    expect(saveRefused({ refusal: {} })).toBe(false);
    expect(saveRefused({ refusal: null })).toBe(false);
  });

  test('the saved name gets one -edited end', () => {
    expect(savedName('book.xlsx', '.xlsx')).toBe('book-edited.xlsx');
    expect(savedName('book-edited.xlsx', '.xlsx')).toBe('book-edited.xlsx');
    expect(savedName('tmpl.xltx', '.xltx')).toBe('tmpl-edited.xltx');
    expect(savedName('macros-edited.xlsm', '.xlsm')).toBe('macros-edited.xlsm');
    expect(savedName('', '.xlsx')).toBe('workbook-edited.xlsx');
  });

  test('a name that the user gave is the file name, with the extension of the kind', () => {
    expect(saveFileName('book.xlsx', null, '.xlsx')).toBe('book-edited.xlsx');
    expect(saveFileName('book.xlsx', 'Budget 2026', '.xlsx')).toBe('Budget 2026.xlsx');
    expect(saveFileName('tmpl.xltx', 'Plan', '.xltx')).toBe('Plan.xltx');
  });

  test('the "Saved" notice stays until the revision moves past the saved one', () => {
    expect(savedNoticeStays(4, 4)).toBe(true);
    expect(savedNoticeStays(4, 5)).toBe(false);
    // A new open starts at revision 0, so the old notice goes too.
    expect(savedNoticeStays(4, 0)).toBe(false);
    expect(savedNoticeStays(null, 4)).toBe(false);
  });

  test('a save with no picker downloads the bytes under the suggested name and type', async () => {
    const type = 'application/test';
    const options = {
      suggestedName: 'a-edited.xlsx',
      types: [{ description: 'Workbook', accept: { [type]: ['.xlsx'] } }],
    };
    const got: { blob: Blob; name: string }[] = [];
    const outcome = await saveFile(
      undefined,
      new Uint8Array([1, 2, 3]),
      options,
      true,
      (blob, name) => got.push({ blob, name })
    );
    // The page cannot know that the file arrives, so the edits stay unsaved.
    expect(outcome).toEqual({
      written: false,
      error: null,
      notice: downloadNotice('a-edited.xlsx', true),
    });
    expect(outcome.notice).toContain('This page still asks before you leave.');
    expect(got.map((g) => g.name)).toEqual(['a-edited.xlsx']);
    expect(got[0]!.blob.type).toBe(type);
    expect([...new Uint8Array(await got[0]!.blob.arrayBuffer())]).toEqual([1, 2, 3]);
  });

  describe('saveRepeatGuard', () => {
    /** A guard on a clock that the test moves, with a count of the saves that started. */
    function guarded(result: () => boolean = () => true) {
      const state = { clock: 0, saves: 0 };
      const guard = saveRepeatGuard(() => state.clock);
      const press = (revision = 4) =>
        guard(revision, async () => {
          state.saves += 1;
          return result();
        });
      return { state, press, guard };
    }

    test('a repeat within the time of a started download starts no save', async () => {
      const { state, press } = guarded();
      await press();
      state.clock += SAVE_REPEAT_MS - 1;
      await press();
      expect(state.saves).toBe(1);
    });

    test('a dropped press does not extend the time', async () => {
      const { state, press } = guarded();
      await press();
      state.clock += SAVE_REPEAT_MS - 1;
      await press();
      state.clock += 1;
      await press();
      expect(state.saves).toBe(2);
    });

    test('an edit between two saves is no repeat', async () => {
      const { state, press } = guarded();
      await press(4);
      state.clock += 10;
      await press(5);
      expect(state.saves).toBe(2);
    });

    test('a press while a save runs starts no save', async () => {
      const { state, guard } = guarded();
      let finish: (started: boolean) => void = () => {};
      const first = guard(4, () => {
        state.saves += 1;
        return new Promise<boolean>((resolve) => (finish = resolve));
      });
      await guard(4, async () => {
        state.saves += 1;
        return true;
      });
      finish(true);
      await first;
      expect(state.saves).toBe(1);
    });

    test('a retry right after a failed save starts a save', async () => {
      const { state, press } = guarded(() => false);
      await press();
      await press();
      expect(state.saves).toBe(2);
    });

    test('five retries 600 ms apart each start a save while each one fails', async () => {
      const { state, press } = guarded(() => false);
      await press();
      for (let retry = 0; retry < 5; retry += 1) {
        state.clock += 600;
        await press();
      }
      expect(state.saves).toBe(6);
    });

    test('a save that throws never blocks a retry', async () => {
      const { state, guard } = guarded();
      const failing = guard(4, async () => {
        state.saves += 1;
        throw new Error('the save failed');
      });
      await expect(failing).rejects.toThrow('the save failed');
      await guard(4, async () => {
        state.saves += 1;
        return true;
      });
      expect(state.saves).toBe(2);
    });
  });

  test('the download notice names the question only over unsaved edits', async () => {
    expect(downloadNotice('a.xlsx', true)).toBe(
      'Started the download of a.xlsx. This page still asks before you leave.'
    );
    expect(downloadNotice('a.xlsx', false)).toBe('Started the download of a.xlsx.');
    // A bidirectional format control in a file name does not reach the notice.
    expect(downloadNotice('a‮xslx.exe', false)).toBe('Started the download of axslx.exe.');
    // With no unsaved edit, the page does not ask, so the outcome does not say that it asks.
    const clean = await saveFile(undefined, new Uint8Array(3), OPTIONS, false, () => {});
    expect(clean).toEqual({
      written: false,
      error: null,
      notice: downloadNotice('a-edited.xlsx', false),
    });
  });

  test('a download that throws shows the error of a failed save', async () => {
    const logged = spyOn(console, 'error').mockImplementation(() => {});
    const outcome = await saveFile(undefined, new Uint8Array(1), OPTIONS, true, () => {
      throw new Error('blocked');
    });
    logged.mockRestore();
    expect(outcome).toEqual({ written: false, error: SAVE_FAILED, notice: null });
  });
});

describe('the bytes that wait for a press of Save', () => {
  const kept = { loadAt: 2, revision: 5 };
  const now = { load: 2, revision: 5, replacing: false };

  test('stay while the workbook does not change', () => {
    expect(keptSaveHolds(kept, now)).toBe(true);
  });

  test('go after an edit', () => {
    expect(keptSaveHolds(kept, { ...now, revision: 6 })).toBe(false);
  });

  test('go after a headings switch or another open, which load the editor again', () => {
    expect(keptSaveHolds(kept, { ...now, load: 3 })).toBe(false);
  });

  test('go while an open replaces the workbook', () => {
    expect(keptSaveHolds(kept, { ...now, replacing: true })).toBe(false);
  });
});
