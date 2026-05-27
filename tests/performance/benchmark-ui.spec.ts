import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

vi.mock("@excelsior/styles", () => ({}));

const createSpreadsheetMock = vi.fn(() => {
  const listeners = new Map<string, Array<(payload: { operations: unknown[]; durationMs: number }) => void>>();
  const renderer = {
    render: () => {}
  };

  const emit = (eventName: string, payload: { operations: unknown[]; durationMs: number }) => {
    for (const listener of listeners.get(eventName) ?? []) {
      listener(payload);
    }
  };

  const engine = {
    on: (eventName: string, listener: (payload: { operations: unknown[]; durationMs: number }) => void) => {
      const current = listeners.get(eventName) ?? [];
      current.push(listener);
      listeners.set(eventName, current);
      return () => undefined;
    },
    getActiveSheet: () => ({ id: "sheet-1" }),
    setCellValue: () => {
      renderer.render();
      emit("command:completed", {
        operations: [{ op: "add" }],
        durationMs: 0.05
      });
      return [];
    },
    updateCells: (input: { updates: unknown[] }) => {
      renderer.render();
      emit("command:completed", {
        operations: input.updates,
        durationMs: 0.02
      });
      return [];
    }
  };

  return {
    engine,
    renderer,
    destroy: () => undefined
  };
});

vi.mock("@excelsior/vanilla", () => ({
  createSpreadsheet: createSpreadsheetMock
}));

class FakeWorker {
  private readonly listeners = new Set<(event: MessageEvent<unknown>) => void>();

  postMessage(message: { requestId: string }): void {
    queueMicrotask(() => {
      for (const listener of this.listeners) {
        listener({
          data: {
            kind: "progress",
            requestId: message.requestId,
            progress: {
              phase: "aggregate",
              completed: 1,
              total: 1
            }
          }
        } as MessageEvent<unknown>);
      }

      for (const listener of this.listeners) {
        listener({
          data: {
            kind: "result",
            requestId: message.requestId,
            durationMs: 2,
            pivotSheet: {
              name: "Pivot",
              rowCount: 2,
              columnCount: 2,
              cells: {
                "0:0": { value: "Region", computedValue: "Region" },
                "0:1": { value: "Revenue", computedValue: "Revenue" },
                "1:0": { value: "North", computedValue: "North" },
                "1:1": { value: 42, computedValue: 42 }
              },
              merges: [],
              columns: {},
              rows: {}
            }
          }
        } as MessageEvent<unknown>);
      }
    });
  }

  addEventListener(_type: string, listener: (event: MessageEvent<unknown>) => void): void {
    this.listeners.add(listener);
  }

  removeEventListener(_type: string, listener: (event: MessageEvent<unknown>) => void): void {
    this.listeners.delete(listener);
  }

  terminate(): void {}
}

const waitFor = async (predicate: () => boolean, message: string) => {
  for (let attempt = 0; attempt < 30; attempt += 1) {
    if (predicate()) {
      return;
    }

    await Promise.resolve();
    await new Promise((resolve) => globalThis.setTimeout(resolve, 0));
  }

  throw new Error(message);
};

describe("benchmark ui", () => {
  const originalWorker = globalThis.Worker;

  beforeEach(() => {
    vi.resetModules();
    createSpreadsheetMock.mockClear();
    document.body.innerHTML = '<div id="app"></div>';
    globalThis.Worker = FakeWorker as unknown as typeof Worker;
  });

  it("renders budget status for mutation and pivot flows in CI", async () => {
    await import("../../apps/benchmark/src/main");

    document.querySelector<HTMLButtonElement>("#mutate-sequential")?.click();
    document.querySelector<HTMLButtonElement>("#mutate-batch")?.click();

    await waitFor(
      () => document.querySelector<HTMLElement>("[data-benchmark-budget-status='pass']") !== null,
      "Timed out waiting for mutation benchmark budget status."
    );

    document.querySelector<HTMLButtonElement>("#pivot-client")?.click();
    await waitFor(
      () => document.querySelector<HTMLElement>("[data-pivot-budget-status='pending']") !== null,
      "Timed out waiting for pivot client benchmark to update the summary."
    );
    await waitFor(
      () => (document.querySelector("#metrics")?.textContent ?? "").includes("pivot client"),
      "Timed out waiting for pivot client benchmark to finish."
    );
    await waitFor(
      () => document.querySelector<HTMLButtonElement>("#pivot-worker")?.disabled === false,
      "Timed out waiting for pivot controls to be re-enabled."
    );

    document.querySelector<HTMLButtonElement>("#pivot-worker")?.click();
    await waitFor(
      () => (document.querySelector("#metrics")?.textContent ?? "").includes("pivot worker"),
      "Timed out waiting for pivot worker benchmark to finish."
    );

    document.querySelector<HTMLButtonElement>("#pivot-server")?.click();

    await waitFor(
      () => document.querySelector<HTMLElement>("[data-pivot-budget-status='pass']") !== null,
      "Timed out waiting for pivot benchmark budget status."
    );

    expect(document.querySelector("#metrics")?.textContent).toContain("pivot server");
    expect(createSpreadsheetMock).toHaveBeenCalledTimes(1);
  });

  afterEach(() => {
    globalThis.Worker = originalWorker;
  });
});