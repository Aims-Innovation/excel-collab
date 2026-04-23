import { guardWorkerPostMessage } from '../workerGuard';

function makeFakeWorker(behavior: (msg: unknown) => void): Worker {
  return {
    postMessage: behavior,
  } as unknown as Worker;
}

describe('guardWorkerPostMessage', () => {
  let warnSpy: jest.SpyInstance;
  beforeEach(() => {
    warnSpy = jest.spyOn(console, 'warn').mockImplementation(() => {});
  });
  afterEach(() => {
    warnSpy.mockRestore();
  });

  test('passes cloneable payloads straight through', () => {
    const received: unknown[] = [];
    const worker = makeFakeWorker((m) => received.push(m));
    guardWorkerPostMessage(worker);

    const payload = { a: 1, b: [2, 3], c: { d: 'x' } };
    worker.postMessage(payload);

    expect(received).toEqual([payload]);
    expect(warnSpy).not.toHaveBeenCalled();
  });

  test('sanitizes and retries on DataCloneError', () => {
    const received: unknown[] = [];
    let firstCall = true;
    const worker = makeFakeWorker((m) => {
      if (firstCall) {
        firstCall = false;
        const err = new DOMException('fn not cloneable', 'DataCloneError');
        throw err;
      }
      received.push(m);
    });
    guardWorkerPostMessage(worker);

    const fn = () => 42;
    const payload = {
      keep: 'ok',
      drop: fn,
      nested: { inner: fn, also: 'kept' },
      list: [1, fn, 2],
    };
    worker.postMessage(payload);

    expect(received).toHaveLength(1);
    expect(received[0]).toEqual({
      keep: 'ok',
      nested: { also: 'kept' },
      list: [1, undefined, 2],
    });
    expect(warnSpy).toHaveBeenCalledTimes(1);
  });

  test('rethrows non-DataCloneError exceptions unchanged', () => {
    const worker = makeFakeWorker(() => {
      throw new Error('boom');
    });
    guardWorkerPostMessage(worker);

    expect(() => worker.postMessage({ any: 'thing' })).toThrow('boom');
  });

  test('handles circular references without stack overflow', () => {
    const received: unknown[] = [];
    let firstCall = true;
    const worker = makeFakeWorker((m) => {
      if (firstCall) {
        firstCall = false;
        throw new DOMException('circular', 'DataCloneError');
      }
      received.push(m);
    });
    guardWorkerPostMessage(worker);

    const a: Record<string, unknown> = { name: 'a' };
    const b: Record<string, unknown> = { name: 'b', ref: a };
    a.ref = b;
    worker.postMessage(a);

    expect(received).toHaveLength(1);
    const out = received[0] as Record<string, unknown>;
    expect(out.name).toBe('a');
    expect((out.ref as Record<string, unknown>).name).toBe('b');
  });

  test('preserves transfer list on sanitized retry', () => {
    const calls: Array<{ msg: unknown; second: unknown }> = [];
    let firstCall = true;
    const worker = makeFakeWorker(function (
      this: unknown,
      ...args: unknown[]
    ) {
      if (firstCall) {
        firstCall = false;
        throw new DOMException('x', 'DataCloneError');
      }
      calls.push({ msg: args[0], second: args[1] });
    } as unknown as (m: unknown) => void);
    guardWorkerPostMessage(worker);

    const transferable = new ArrayBuffer(4);
    (worker.postMessage as (m: unknown, t: Transferable[]) => void)(
      { buf: transferable, fn: () => 1 },
      [transferable],
    );

    expect(calls).toHaveLength(1);
    expect(calls[0].second).toEqual([transferable]);
    expect((calls[0].msg as { buf: ArrayBuffer }).buf).toBe(transferable);
  });
});
