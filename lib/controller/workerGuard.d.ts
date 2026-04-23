/**
 * Wraps a Worker so that postMessage survives a payload containing
 * non-structured-cloneable values (functions, symbols, class instances with
 * method properties). On DataCloneError the payload is sanitized and posted
 * again, so a single bad cell can't kill the canvas render pipeline.
 *
 * The fast path is a direct passthrough — no per-message overhead when the
 * payload is already cloneable, which is the common case.
 */
export declare function guardWorkerPostMessage(worker: Worker): Worker;
