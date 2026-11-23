// The end of a corpus child that its runner cannot stop. The guard thread (`parent-guard.ts`)
// calls it when the runner is no longer the parent or the child is far above the memory cap, and
// the child (`worker.ts`) calls it when a write to the runner fails. It has no other import, so
// the guard thread starts fast.

import { lstatSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { basename, dirname, resolve } from 'node:path';

/**
 * Removes the work folder `dir`, then stops the process group of the child. A runner that got
 * SIGKILL cannot remove the folder, which holds work files of a workbook. Only a folder of the
 * form that the runner makes goes: in the temporary folder, with the `xlsx-corpus-` name, and
 * not a link.
 */
export function endOrphan(dir: string | undefined): void {
  try {
    const folder = dir ? resolve(dir) : undefined;
    if (
      folder &&
      basename(folder).startsWith('xlsx-corpus-') &&
      dirname(folder) === resolve(tmpdir()) &&
      lstatSync(folder).isDirectory()
    ) {
      rmSync(folder, { recursive: true, force: true });
    }
  } catch {
    // The stop below matters more than the folder.
  }
  try {
    process.kill(-process.pid, 'SIGKILL');
  } catch {
    // The child does not lead a group (it was not started by the runner): stop it alone.
    process.kill(process.pid, 'SIGKILL');
  }
}
