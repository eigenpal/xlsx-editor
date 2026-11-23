// A mark that holds until the task that set it ends. A browser runs a microtask checkpoint
// between the listeners of one task, such as the blur and the focus of one focus change, so a
// microtask cannot mark one task. A timer of 0 ms runs after the task, and it clears the mark.

import type { HostTimers } from './host-realm.ts';

/** A mark that holds until the task that set it ends: a timer of 0 ms clears it. */
export class TaskMark {
  #held = false;
  #timer: number | null = null;

  /** @param timers - The timers of the window that runs the task. */
  constructor(private readonly timers: HostTimers) {}

  /** Whether the task that runs now, or a task before the timer ran, set the mark. */
  get held(): boolean {
    return this.#held;
  }

  /** Set the mark until the task ends. */
  note(): void {
    this.#held = true;
    this.timers.clearTimeout(this.#timer);
    this.#timer = this.timers.setTimeout(() => (this.#held = false), 0);
  }
}
