/**
 * Wraps a Worker so that postMessage survives a payload containing
 * non-structured-cloneable values (functions, symbols, class instances with
 * method properties). On DataCloneError the payload is sanitized and posted
 * again, so a single bad cell can't kill the canvas render pipeline.
 *
 * The fast path is a direct passthrough — no per-message overhead when the
 * payload is already cloneable, which is the common case.
 */

type PostMessageArgs =
  | [message: unknown]
  | [message: unknown, transfer: Transferable[]]
  | [message: unknown, options: StructuredSerializeOptions];

function isDataCloneError(err: unknown): boolean {
  return (
    typeof DOMException !== 'undefined' &&
    err instanceof DOMException &&
    err.name === 'DataCloneError'
  );
}

function sanitize(value: unknown, seen: WeakSet<object>): unknown {
  if (value === null) return null;
  const t = typeof value;
  if (t === 'undefined') return undefined;
  if (t === 'function' || t === 'symbol') return undefined;
  if (t !== 'object') return value;

  const obj = value as object;
  if (seen.has(obj)) return undefined;
  seen.add(obj);

  if (
    obj instanceof ArrayBuffer ||
    ArrayBuffer.isView(obj) ||
    obj instanceof Date ||
    obj instanceof RegExp ||
    obj instanceof Blob ||
    obj instanceof File ||
    (typeof MessagePort !== 'undefined' && obj instanceof MessagePort)
  ) {
    return obj;
  }

  if (obj instanceof Map) {
    const out = new Map();
    for (const [k, v] of obj) {
      const sk = sanitize(k, seen);
      const sv = sanitize(v, seen);
      if (sk !== undefined && sv !== undefined) out.set(sk, sv);
    }
    return out;
  }
  if (obj instanceof Set) {
    const out = new Set();
    for (const v of obj) {
      const sv = sanitize(v, seen);
      if (sv !== undefined) out.add(sv);
    }
    return out;
  }

  if (Array.isArray(obj)) {
    return obj.map((v) => sanitize(v, seen));
  }

  const plain: Record<string, unknown> = {};
  for (const k of Object.keys(obj)) {
    const v = sanitize((obj as Record<string, unknown>)[k], seen);
    if (v !== undefined) plain[k] = v;
  }
  return plain;
}

let warnedOnce = false;

export function guardWorkerPostMessage(worker: Worker): Worker {
  const original = worker.postMessage.bind(worker);
  worker.postMessage = ((...args: PostMessageArgs) => {
    try {
      (original as (...a: PostMessageArgs) => void)(...args);
    } catch (err) {
      if (!isDataCloneError(err)) throw err;
      const [message, second] = args;
      const sanitized = sanitize(message, new WeakSet());
      if (!warnedOnce) {
        warnedOnce = true;
        // eslint-disable-next-line no-console
        console.warn(
          '[excel-collab] worker.postMessage: payload contained ' +
            'non-cloneable values; sanitized and retried. ' +
            '(Subsequent occurrences will be silent.)',
          err,
        );
      }
      if (second === undefined) {
        (original as (m: unknown) => void)(sanitized);
      } else {
        (original as (...a: PostMessageArgs) => void)(
          sanitized,
          second as Transferable[] & StructuredSerializeOptions,
        );
      }
    }
  }) as Worker['postMessage'];
  return worker;
}
