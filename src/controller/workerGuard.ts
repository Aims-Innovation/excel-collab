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

function describeNonCloneable(value: unknown): string | null {
  if (value === null) return null;
  const t = typeof value;
  if (t === 'function') {
    const fn = value as { name?: string; toString?: () => string };
    return `function ${fn.name || '(anonymous)'} ${
      typeof fn.toString === 'function' ? fn.toString().slice(0, 80) : ''
    }`;
  }
  if (t === 'symbol') return 'symbol';
  return null;
}

function findFirstNonCloneable(
  value: unknown,
  path: string,
  seen: WeakSet<object>,
  maxDepth: number,
): { path: string; description: string } | null {
  if (maxDepth <= 0) return null;

  const direct = describeNonCloneable(value);
  if (direct !== null) return { path, description: direct };

  if (value === null || typeof value !== 'object') return null;
  const obj = value as object;
  if (seen.has(obj)) return null;
  seen.add(obj);

  if (
    obj instanceof ArrayBuffer ||
    ArrayBuffer.isView(obj) ||
    obj instanceof Date ||
    obj instanceof RegExp ||
    obj instanceof Blob ||
    obj instanceof File
  ) {
    return null;
  }

  if (obj instanceof Map) {
    let i = 0;
    for (const [k, v] of obj) {
      const found =
        findFirstNonCloneable(k, `${path}.key[${i}]`, seen, maxDepth - 1) ||
        findFirstNonCloneable(v, `${path}.val[${i}]`, seen, maxDepth - 1);
      if (found) return found;
      i++;
    }
    return null;
  }
  if (obj instanceof Set) {
    let i = 0;
    for (const v of obj) {
      const found = findFirstNonCloneable(
        v,
        `${path}[${i}]`,
        seen,
        maxDepth - 1,
      );
      if (found) return found;
      i++;
    }
    return null;
  }

  if (Array.isArray(obj)) {
    for (let i = 0; i < obj.length; i++) {
      const found = findFirstNonCloneable(
        obj[i],
        `${path}[${i}]`,
        seen,
        maxDepth - 1,
      );
      if (found) return found;
    }
    return null;
  }

  for (const k of Object.keys(obj)) {
    const found = findFirstNonCloneable(
      (obj as Record<string, unknown>)[k],
      `${path}.${k}`,
      seen,
      maxDepth - 1,
    );
    if (found) return found;
  }
  return null;
}

export function guardWorkerPostMessage(worker: Worker): Worker {
  const original = worker.postMessage.bind(worker);
  let diagnosedOnce = false;
  worker.postMessage = ((...args: PostMessageArgs) => {
    try {
      (original as (...a: PostMessageArgs) => void)(...args);
    } catch (err) {
      if (!isDataCloneError(err)) throw err;
      if (!diagnosedOnce) {
        diagnosedOnce = true;
        const [message] = args;
        const culprit = findFirstNonCloneable(
          message,
          'payload',
          new WeakSet(),
          32,
        );
        // eslint-disable-next-line no-console
        console.error(
          '[excel-collab] worker.postMessage: DataCloneError. ' +
            'The render worker received a non-cloneable payload and will not render. ' +
            'First non-cloneable value found at:',
          culprit ? culprit.path : '(not found within traversal depth)',
          culprit ? `type: ${culprit.description}` : '',
          '\nRaw payload:',
          args[0],
        );
      }
      throw err;
    }
  }) as Worker['postMessage'];
  return worker;
}
