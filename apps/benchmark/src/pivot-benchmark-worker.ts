import {
  buildPivotSheetAsync,
  type PivotBuildProgress,
  type PivotSheetInput,
  type WorkbookDataInput,
  type WorkbookModel
} from "@excelsior/core";

type PivotBenchmarkWorkerRequest = {
  kind: "buildPivot";
  requestId: string;
  workbook: WorkbookModel;
  input: PivotSheetInput;
};

type PivotBenchmarkWorkerResponse =
  | {
      kind: "progress";
      requestId: string;
      progress: PivotBuildProgress;
    }
  | {
      kind: "result";
      requestId: string;
      durationMs: number;
      pivotSheet: WorkbookDataInput;
    }
  | {
      kind: "error";
      requestId: string;
      code?: string;
      message: string;
    };

type PivotBenchmarkWorkerScope = typeof globalThis & {
  postMessage: (message: PivotBenchmarkWorkerResponse) => void;
};

const workerScope = self as PivotBenchmarkWorkerScope;

const postResponse = (response: PivotBenchmarkWorkerResponse): void => {
  workerScope.postMessage(response);
};

workerScope.addEventListener("message", (event: MessageEvent<PivotBenchmarkWorkerRequest>) => {
  if (event.data.kind !== "buildPivot") {
    return;
  }

  void (async () => {
    const startedAt = performance.now();

    try {
      const pivotSheet = await buildPivotSheetAsync(event.data.workbook, event.data.input, {
        chunkSize: 500,
        yieldControl: async () => Promise.resolve(),
        onProgress: (progress) => {
          postResponse({
            kind: "progress",
            requestId: event.data.requestId,
            progress
          });
        }
      });

      postResponse({
        kind: "result",
        requestId: event.data.requestId,
        durationMs: performance.now() - startedAt,
        pivotSheet
      });
    } catch (error) {
      postResponse({
        kind: "error",
        requestId: event.data.requestId,
        code: error instanceof Error ? error.name : undefined,
        message: error instanceof Error ? error.message : "Pivot worker execution failed."
      });
    }
  })();
});

export {};