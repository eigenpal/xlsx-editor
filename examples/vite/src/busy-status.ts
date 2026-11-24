// The status text of the demo while a long open or save runs. Short work shows nothing: the text
// comes only after a delay, so a small workbook does not make the bar flicker. The text names the
// newest work that still runs.

import { useState } from 'react';

/** The time before the text shows. */
const SHOW_AFTER_MS = 300;

/** Starts and ends the status text of long work. */
export interface BusyStatus {
  /** Runs `work` with `text` as the status until it ends. */
  during<T>(text: string, work: Promise<T>): Promise<T>;
}

/** Tracks the running work over `show`, which receives the newest text, or null at the end. */
export function busyStatus(show: (text: string | null) => void): BusyStatus {
  const running: { readonly text: string }[] = [];
  let timer: ReturnType<typeof setTimeout> | undefined;
  let shown: string | null = null;
  const showNewest = () => {
    const text = running.at(-1)?.text ?? null;
    if (text !== shown) show((shown = text));
  };
  return {
    async during(text, work) {
      const entry = { text };
      running.push(entry);
      if (shown !== null) showNewest();
      else timer ??= setTimeout(showNewest, SHOW_AFTER_MS);
      try {
        return await work;
      } finally {
        running.splice(running.indexOf(entry), 1);
        if (running.length > 0) {
          if (shown !== null) showNewest();
        } else {
          clearTimeout(timer);
          timer = undefined;
          shown = null;
          show(null);
        }
      }
    },
  };
}

/** The status text to show, and the status object, made once for the component. */
export function useBusyStatus(): readonly [string | null, BusyStatus] {
  const [text, setText] = useState<string | null>(null);
  const [status] = useState(() => busyStatus(setText));
  return [text, status];
}
