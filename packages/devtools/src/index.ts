import type { SpreadsheetEventMap, WorkbookEngine, WorkbookModel } from "@excelsior/core";

export interface WorkbookDevtoolsEvent<TKey extends keyof SpreadsheetEventMap = keyof SpreadsheetEventMap> {
  name: TKey;
  payload: SpreadsheetEventMap[TKey];
}

export interface WorkbookDevtoolsLogger {
  debug?: (...args: unknown[]) => void;
}

export interface AttachWorkbookDevtoolsOptions {
  onEvent?: (event: WorkbookDevtoolsEvent) => void;
  logger?: WorkbookDevtoolsLogger;
}

export interface WorkbookDevtoolsSession {
  events: WorkbookDevtoolsEvent[];
  snapshot: () => WorkbookModel;
  stop: () => void;
}

const EVENT_NAMES: Array<keyof SpreadsheetEventMap> = [
  "engine:created",
  "engine:disposed",
  "command:completed",
  "command:failed",
  "cell:updated",
  "selection:changed",
  "security:blocked-input",
  "formula:failed"
];

export const attachWorkbookDevtools = (
  engine: WorkbookEngine,
  options: AttachWorkbookDevtoolsOptions = {}
): WorkbookDevtoolsSession => {
  const events: WorkbookDevtoolsEvent[] = [];
  const unsubscribeCallbacks = EVENT_NAMES.map((name) =>
    engine.on(name, (payload) => {
      const event: WorkbookDevtoolsEvent = { name, payload };
      events.push(event);
      options.onEvent?.(event);
      options.logger?.debug?.("[Excelsior Devtools]", name, payload);
    })
  );

  return {
    events,
    snapshot: () => engine.getSnapshot(),
    stop: () => {
      for (const unsubscribe of unsubscribeCallbacks) {
        unsubscribe();
      }
    }
  };
};