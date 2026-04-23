import { guardWorkerPostMessage } from '../workerGuard';

function makeFakeWorker(behavior: (msg: unknown) => void): Worker {
  return {
    postMessage: behavior,
  } as unknown as Worker;
}

describe('guardWorkerPostMessage', () => {
  let errorSpy: jest.SpyInstance;
  beforeEach(() => {
    errorSpy = jest.spyOn(console, 'error').mockImplementation(() => {});
  });
  afterEach(() => {
    errorSpy.mockRestore();
  });

  test('passes cloneable payloads straight through', () => {
    const received: unknown[] = [];
    const worker = makeFakeWorker((m) => received.push(m));
    guardWorkerPostMessage(worker);

    const payload = { a: 1, b: [2, 3], c: { d: 'x' } };
    worker.postMessage(payload);

    expect(received).toEqual([payload]);
    expect(errorSpy).not.toHaveBeenCalled();
  });

  test('rethrows DataCloneError with enriched diagnostics (no silent retry)', () => {
    const worker = makeFakeWorker(() => {
      throw new DOMException('fn not cloneable', 'DataCloneError');
    });
    guardWorkerPostMessage(worker);

    const fn = () => 42;
    const payload = {
      sheetData: { r1c1: { value: 'kept' } },
      autoFilter: {
        range: { row: 0, col: 0, rowCount: 1, colCount: 1 },
        onReady: fn,
      },
    };

    expect(() => worker.postMessage(payload)).toThrow(DOMException);
    expect(errorSpy).toHaveBeenCalledTimes(1);
    const msg = String(errorSpy.mock.calls[0][1]);
    expect(msg).toContain('payload.autoFilter.onReady');
  });

  test('rethrows non-DataCloneError exceptions unchanged and does not log', () => {
    const worker = makeFakeWorker(() => {
      throw new Error('boom');
    });
    guardWorkerPostMessage(worker);

    expect(() => worker.postMessage({ any: 'thing' })).toThrow('boom');
    expect(errorSpy).not.toHaveBeenCalled();
  });

  test('diagnostic fires only once across repeated failures', () => {
    const worker = makeFakeWorker(() => {
      throw new DOMException('x', 'DataCloneError');
    });
    guardWorkerPostMessage(worker);

    for (let i = 0; i < 5; i++) {
      try {
        worker.postMessage({ fn: () => i });
      } catch {
        /* expected */
      }
    }
    expect(errorSpy).toHaveBeenCalledTimes(1);
  });

  test('diagnostic reports symbol values', () => {
    const worker = makeFakeWorker(() => {
      throw new DOMException('x', 'DataCloneError');
    });
    guardWorkerPostMessage(worker);

    expect(() =>
      worker.postMessage({ meta: { tag: Symbol('y-ref') } }),
    ).toThrow();
    const loggedPath = String(errorSpy.mock.calls[0][1]);
    expect(loggedPath).toContain('payload.meta.tag');
  });

  test('handles circular references without stack overflow', () => {
    const worker = makeFakeWorker(() => {
      throw new DOMException('x', 'DataCloneError');
    });
    guardWorkerPostMessage(worker);

    const a: Record<string, unknown> = { name: 'a' };
    const b: Record<string, unknown> = { name: 'b', ref: a };
    a.ref = b;

    expect(() => worker.postMessage(a)).toThrow();
  });
});
