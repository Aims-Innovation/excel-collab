import { Model } from '../model';
import type { IController } from '../types';
import { Controller } from './Controller';
import method from '../canvas/worker';
import { Doc } from 'yjs';
import { wrap } from 'comlink';
import { guardWorkerPostMessage } from './workerGuard';
import { collaborationLog } from '../util/debug';

function getMockHooks() {
  return {
    worker: method as any,
    doc: new Doc(),
  };
}

export function initController(
  options: {
    worker: Worker;
    doc: Doc;
  } = getMockHooks(),
): IController {
  const worker =
    options.worker instanceof Worker
      ? wrap(guardWorkerPostMessage(options.worker))
      : (options.worker as any);

  const hooks = {
    worker,
    doc: options.doc,
  };
  collaborationLog('initController', {
    docGuid: options.doc?.guid,
    docClientId: options.doc?.clientID,
    // Existence of _observers tells us yjs internals are already wired.
    // A consumer that passes a freshly-constructed Doc will have this
    // populated; a consumer that constructed Doc against a *different*
    // yjs copy won't — in that case 'initController' prints but the
    // consumer's later doc.on('update') binds to a different instance.
    docHasObservers: !!(options.doc as unknown as { _observers?: unknown })
      ._observers,
  });
  const model = new Model(hooks);
  const controller = new Controller(model, hooks);
  return controller;
}
