/**
 * Wraps a Worker so that a DataCloneError on postMessage surfaces with
 * enough context to debug it, instead of failing deep inside comlink with
 * no indication of which payload key was the culprit.
 *
 * IMPORTANT: this guard no longer sanitizes-and-retries. The previous
 * version dropped functions from the payload and silently retried, which
 * masked real rendering bugs (the render worker got a gutted payload and
 * silently drew nothing). That behavior is removed.
 *
 * Fast path: direct passthrough, zero overhead.
 * Error path: walk the payload, find the first non-cloneable value with
 * its key path, log one diagnostic, then rethrow so the caller sees it.
 */
export declare function guardWorkerPostMessage(worker: Worker): Worker;
