import { getRandomColor } from './util';

type NameSpaceType =
  | 'react'
  | 'model'
  | 'test'
  | 'collaboration'
  | 'controller'
  | 'perf'
  | 'render'
  | 'input';

/**
 * Opt-in dev logging.
 *
 * Enable all namespaces:
 *   localStorage.setItem('debug', '1')
 *
 * Enable just one or two:
 *   localStorage.setItem('debug', 'perf,render')
 *
 * The special value '1' (or 'true') enables everything. Anything else
 * is parsed as a comma-separated list of namespaces.
 */
function readDebugFlag(): Set<string> | 'all' | null {
  try {
    const raw = window.localStorage.getItem('debug');
    if (!raw) return null;
    if (raw === '1' || raw === 'true' || raw === '*') return 'all';
    return new Set(
      raw
        .split(',')
        .map((s) => s.trim())
        .filter(Boolean),
    );
  } catch {
    return null;
  }
}

export class Debug {
  namespace: NameSpaceType;

  static readonly colorMap: Map<NameSpaceType, string> = new Map<
    NameSpaceType,
    string
  >();
  constructor(namespace: NameSpaceType) {
    this.namespace = namespace;
  }
  init = () => {
    this.setColor();
    return this.log;
  };
  log = (...rest: unknown[]): void => {
    if (!this.enable()) {
      return;
    }
    const { namespace } = this;
    const color = Debug.colorMap.get(namespace);
    const result = [`%c ${namespace}:`, `color:${color};`, ...rest];
    console.log(...result);
  };
  enable() {
    const flag = readDebugFlag();
    if (!flag) return false;
    if (flag === 'all') return true;
    return flag.has(this.namespace);
  }
  setColor() {
    if (!Debug.colorMap.has(this.namespace)) {
      Debug.colorMap.set(this.namespace, getRandomColor());
    }
  }
}

export const modelLog = new Debug('model').init();
export const controllerLog = new Debug('controller').init();
export const reactLog = new Debug('react').init();
export const collaborationLog = new Debug('collaboration').init();
export const perfLog = new Debug('perf').init();
export const renderLog = new Debug('render').init();
export const inputLog = new Debug('input').init();

/**
 * Measure a synchronous or async block of work. No-ops when 'perf' is
 * not enabled, so it's safe to leave in hot paths.
 *
 *   await perfMeasure('canvas.render', () => mainCanvas.render(data))
 */
export async function perfMeasure<T>(
  label: string,
  fn: () => T | Promise<T>,
): Promise<T> {
  const log = new Debug('perf');
  if (!log.enable()) return fn();
  const start = performance.now();
  try {
    return await fn();
  } finally {
    const elapsed = performance.now() - start;
    perfLog(`${label} ${elapsed.toFixed(1)}ms`);
  }
}
