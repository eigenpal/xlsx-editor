/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See LICENSE.md in this package.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// After PPTX Editor's `packages/editor-api/src/runtime/runtime.ts` at d4502552, which adapts DOCX
// Editor's `packages/editor-api/src/runtime/runtime.ts` at 5993bd223945. Differences: the
// capabilities of the workbook subset, and `run` takes a callback only (no objects to adopt,
// because the subset has no tracked objects).
//
// A runtime: one workbook host, many runs. Runs are ISOLATED, not serialized: every `run` gets
// its own context and queue, so two runs cannot interleave into one batch.

import type { AutomationHost } from '@xlsx-editor.dev/core/automation';
import { Workbook } from '../model/workbook.ts';
import { fail, failArgument } from './errors.ts';
import { hostRefusal, RequestContext, type RuntimeSession } from './request-context.ts';

/**
 * What a `run` callback is given, and what it may answer with.
 *
 * @public
 */
export type RunCallback<T> = (context: RequestContext) => Promise<T>;

/**
 * Capabilities of a runtime, frozen for its lifetime.
 *
 * @public
 */
export interface WorkbookCapabilities {
  /** The runtime has a host for a workbook. */
  readonly document: boolean;
  /** `save()` is offered. */
  readonly save: boolean;
}

/**
 * A runtime: one workbook host, many runs.
 *
 * Disposal is final: {@link XlsxEditorRuntime.dispose} releases the host once and is safe to call
 * again, and every later `run` fails with `RuntimeDisposed`.
 *
 * @public
 */
export interface XlsxEditorRuntime {
  /** What the workbook host behind this runtime can do. */
  readonly capabilities: WorkbookCapabilities;
  /**
   * Run one batch of work against the workbook, and answer with the callback's value. Actions
   * that no `context.sync()` sent before the callback returns are dropped.
   */
  run<T>(callback: RunCallback<T>): Promise<T>;
  /** Release the host. Idempotent. */
  dispose(): void;
}

/**
 * A runtime over XLSX bytes: what `XlsxEditor.createServer` answers.
 *
 * @public
 */
export interface XlsxEditorServerRuntime extends XlsxEditorRuntime {
  /** The current workbook as a new XLSX byte array that the caller owns. */
  save(): Promise<Uint8Array>;
}

export function createRuntime(host: AutomationHost): XlsxEditorServerRuntime {
  const capabilities: WorkbookCapabilities = Object.freeze({
    document: host.capabilities.document,
    save: host.capabilities.save,
  });
  let disposed = false;
  const assertLive = (target?: string): void => {
    if (disposed) fail({ code: 'RuntimeDisposed', ...(target === undefined ? {} : { target }) });
  };
  const session: RuntimeSession = {
    host,
    assertLive,
    openWorkbook: (context) => new Workbook(context),
  };

  async function run<T>(callback: RunCallback<T>): Promise<T> {
    assertLive('run');
    if (typeof callback !== 'function') failArgument('run', 'a callback');
    const context = RequestContext.begin(session);
    let completed = false;
    try {
      const result = await callback(context);
      completed = true;
      return result;
    } finally {
      RequestContext.finish(context, completed);
    }
  }

  return {
    capabilities,
    run,
    dispose() {
      if (disposed) return;
      disposed = true;
      host.dispose();
    },
    async save(): Promise<Uint8Array> {
      assertLive('save');
      if (!capabilities.save) fail({ code: 'NotSupported', target: 'save' });
      const saved = host.save();
      if (!saved.ok) throw hostRefusal(saved.error, 'save');
      return saved.bytes;
    },
  };
}
