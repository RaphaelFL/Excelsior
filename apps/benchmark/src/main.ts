import "@excelsior/styles";
import {
  ServerSideRowModel,
  WorkbookEngine,
  type DataRequest,
  type PivotBuildProgress,
  type PivotSheetInput,
  type WorkbookDataInput,
  type WorkbookModel
} from "@excelsior/core";
import { createSpreadsheet } from "@excelsior/vanilla";

const BENCHMARK_ROW_COUNT = 5000;
const BENCHMARK_COLUMN_COUNT = 200;
const BENCHMARK_MUTATION_BASE_COLUMN = 2;
const MUTATION_SCALES = [1000, 10000, 100000] as const;
const PIVOT_ROW_SCALES = [5000, 20000, 50000] as const;
const PIVOT_CHUNK_SIZE = 500;

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

const app = document.querySelector<HTMLDivElement>("#app");

if (!app) {
  throw new Error("Benchmark root was not found.");
}

app.innerHTML = `
  <div style="min-height:100vh;padding:24px;background:linear-gradient(180deg,#f5f1e8,#eee4d1);font-family:Aptos,Segoe UI Variable,Segoe UI,sans-serif;">
    <div style="max-width:1280px;margin:0 auto;display:grid;gap:18px;">
      <div style="display:flex;justify-content:space-between;align-items:center;gap:16px;flex-wrap:wrap;">
        <div>
          <h1 style="margin:0 0 6px;font-size:38px;color:#16202b;">Excelsior Benchmark</h1>
          <p style="margin:0;color:#556170;">Carga sintética para comparar mutações sequenciais contra batch real da engine.</p>
        </div>
        <div style="display:flex;gap:12px;flex-wrap:wrap;align-items:center;justify-content:flex-end;">
          <label style="display:grid;gap:6px;color:#556170;font-size:13px;">
            Escala
            <select id="mutation-scale" style="padding:10px 14px;border-radius:999px;border:1px solid rgba(22,32,43,.14);background:rgba(255,255,255,.84);color:#16202b;">
              ${MUTATION_SCALES.map((scale) => `<option value="${scale}">${scale.toLocaleString("pt-BR")} updates</option>`).join("")}
            </select>
          </label>
          <button id="run-suite" style="padding:12px 18px;border:none;border-radius:999px;background:#7c3aed;color:#fff;cursor:pointer;">Rodar suite completa</button>
          <button id="mutate-sequential" style="padding:12px 18px;border:none;border-radius:999px;background:#0f766e;color:#fff;cursor:pointer;">Rodar sequencial</button>
          <button id="mutate-batch" style="padding:12px 18px;border:none;border-radius:999px;background:#1d4ed8;color:#fff;cursor:pointer;">Rodar batch</button>
          <button id="clear-metrics" style="padding:12px 18px;border:1px solid rgba(22,32,43,.14);border-radius:999px;background:rgba(255,255,255,.84);color:#16202b;cursor:pointer;">Limpar métricas</button>
        </div>
      </div>
      <div style="display:flex;gap:10px;flex-wrap:wrap;align-items:center;">
        <span style="color:#556170;font-size:13px;">Presets:</span>
        ${MUTATION_SCALES.map(
          (scale) =>
            `<button data-scale-preset="${scale}" style="padding:8px 14px;border:1px solid rgba(22,32,43,.12);border-radius:999px;background:rgba(255,255,255,.84);color:#16202b;cursor:pointer;">${scale.toLocaleString("pt-BR")}</button>`
        ).join("")}
      </div>
      <div style="display:grid;gap:14px;padding:18px;border-radius:22px;background:rgba(255,255,255,.72);box-shadow:0 18px 45px rgba(31,42,55,.14);">
        <div style="display:flex;justify-content:space-between;align-items:center;gap:16px;flex-wrap:wrap;">
          <div>
            <h2 style="margin:0 0 6px;font-size:24px;color:#16202b;">Pivot pesado</h2>
            <p style="margin:0;color:#556170;">Compara materialização client-side no main thread, off-main-thread em worker dedicado e contrato server-side remoto.</p>
          </div>
          <div style="display:flex;gap:12px;flex-wrap:wrap;align-items:center;justify-content:flex-end;">
            <label style="display:grid;gap:6px;color:#556170;font-size:13px;">
              Linhas fonte
              <select id="pivot-scale" style="padding:10px 14px;border-radius:999px;border:1px solid rgba(22,32,43,.14);background:rgba(255,255,255,.84);color:#16202b;">
                ${PIVOT_ROW_SCALES.map((scale) => `<option value="${scale}">${scale.toLocaleString("pt-BR")} linhas</option>`).join("")}
              </select>
            </label>
            <button id="run-pivot-suite" style="padding:12px 18px;border:none;border-radius:999px;background:#b45309;color:#fff;cursor:pointer;">Rodar suite pivot</button>
            <button id="pivot-client" style="padding:12px 18px;border:none;border-radius:999px;background:#0f766e;color:#fff;cursor:pointer;">Pivot client</button>
            <button id="pivot-worker" style="padding:12px 18px;border:none;border-radius:999px;background:#7c3aed;color:#fff;cursor:pointer;">Pivot worker</button>
            <button id="pivot-server" style="padding:12px 18px;border:none;border-radius:999px;background:#1d4ed8;color:#fff;cursor:pointer;">Pivot server</button>
          </div>
        </div>
        <div id="pivot-summary" style="overflow:auto;"></div>
      </div>
      <div id="benchmark-host" style="height:72vh;"></div>
      <div id="benchmark-summary" style="overflow:auto;border-radius:18px;background:rgba(255,255,255,.78);box-shadow:0 18px 45px rgba(31,42,55,.14);padding:18px;"></div>
      <pre id="metrics" style="margin:0;padding:18px;border-radius:18px;background:rgba(255,255,255,.78);box-shadow:0 18px 45px rgba(31,42,55,.14);color:#556170;"></pre>
    </div>
  </div>
`;

const host = document.querySelector<HTMLDivElement>("#benchmark-host");
const scaleSelect = document.querySelector<HTMLSelectElement>("#mutation-scale");
const runSuiteButton = document.querySelector<HTMLButtonElement>("#run-suite");
const sequentialButton = document.querySelector<HTMLButtonElement>("#mutate-sequential");
const batchButton = document.querySelector<HTMLButtonElement>("#mutate-batch");
const clearMetricsButton = document.querySelector<HTMLButtonElement>("#clear-metrics");
const summary = document.querySelector<HTMLDivElement>("#benchmark-summary");
const pivotScaleSelect = document.querySelector<HTMLSelectElement>("#pivot-scale");
const runPivotSuiteButton = document.querySelector<HTMLButtonElement>("#run-pivot-suite");
const pivotClientButton = document.querySelector<HTMLButtonElement>("#pivot-client");
const pivotWorkerButton = document.querySelector<HTMLButtonElement>("#pivot-worker");
const pivotServerButton = document.querySelector<HTMLButtonElement>("#pivot-server");
const pivotSummary = document.querySelector<HTMLDivElement>("#pivot-summary");
const metrics = document.querySelector<HTMLPreElement>("#metrics");
const presetButtons = Array.from(document.querySelectorAll<HTMLButtonElement>("[data-scale-preset]"));

if (
  !host ||
  !scaleSelect ||
  !runSuiteButton ||
  !sequentialButton ||
  !batchButton ||
  !clearMetricsButton ||
  !summary ||
  !pivotScaleSelect ||
  !runPivotSuiteButton ||
  !pivotClientButton ||
  !pivotWorkerButton ||
  !pivotServerButton ||
  !pivotSummary ||
  !metrics
) {
  throw new Error("Benchmark UI did not initialize.");
}

const pivotWorker = new Worker(new URL("./pivot-benchmark-worker.ts", import.meta.url), {
  type: "module"
});

globalThis.addEventListener(
  "beforeunload",
  () => {
    pivotWorker.terminate();
  },
  { once: true }
);

const getCellKey = (row: number, col: number): string => `${row}:${col}`;

const getMutationCount = (): number => Number(scaleSelect.value);

const getMutationTarget = (index: number): { row: number; col: number } => ({
  row: index % BENCHMARK_ROW_COUNT,
  col: BENCHMARK_MUTATION_BASE_COLUMN + Math.floor(index / BENCHMARK_ROW_COUNT)
});

const describeMutationWindow = (mutationCount: number): string => {
  const lastTarget = getMutationTarget(mutationCount - 1);
  return `window: rows 0-${Math.min(BENCHMARK_ROW_COUNT - 1, mutationCount - 1)}, cols ${BENCHMARK_MUTATION_BASE_COLUMN}-${lastTarget.col}`;
};

const cells: Record<string, { value: string; computedValue: string }> = {};
for (let row = 0; row < 2000; row += 1) {
  cells[`${row}:0`] = { value: `Item ${row + 1}`, computedValue: `Item ${row + 1}` };
  cells[`${row}:1`] = { value: String(row * 17), computedValue: String(row * 17) };
}

const startedAt = performance.now();
const instance = createSpreadsheet(host, {
  data: [
    {
      name: "Benchmark",
      rowCount: BENCHMARK_ROW_COUNT,
      columnCount: BENCHMARK_COLUMN_COUNT,
      cells
    }
  ]
});

type EngineMetricsSnapshot = {
  commandCount: number;
  operationCount: number;
  engineDurationMs: number;
};

type RendererMetricsSnapshot = {
  renderCount: number;
  renderDurationMs: number;
};

type BenchmarkRunResult = {
  mode: "sequencial" | "batch";
  label: string;
  mutationCount: number;
  durationMs: number;
  commandCount: number;
  operationCount: number;
  engineDurationMs: number;
  renderCount: number;
  renderDurationMs: number;
  renderedCellCount: number;
};

type BenchmarkSummaryEntry = {
  sequencial?: BenchmarkRunResult;
  batch?: BenchmarkRunResult;
};

type PivotBenchmarkMode = "client" | "worker" | "server";

type PivotBenchmarkRunResult = {
  mode: PivotBenchmarkMode;
  label: string;
  sourceRowCount: number;
  durationMs: number;
  progressEvents: number;
  resultRowCount: number;
  resultColumnCount: number;
  remote: boolean;
};

type PivotBenchmarkSummaryEntry = Partial<Record<PivotBenchmarkMode, PivotBenchmarkRunResult>>;

type BudgetStatus = {
  status: "pending" | "pass" | "fail";
  label: string;
};

const MUTATION_BUDGETS = {
  maxBatchCommandCount: 1,
  maxBatchDurationRatio: 1.25,
  maxBatchRenderRatio: 1.1
} as const;

const PIVOT_BUDGETS = {
  maxWorkerDurationRatio: 1.25,
  maxServerDurationRatio: 1.5,
  minProgressEvents: 1
} as const;

let engineCommandCount = 0;
let engineOperationCount = 0;
let engineDurationMs = 0;
let rendererRenderCount = 0;
let rendererRenderDurationMs = 0;

instance.engine.on("command:completed", ({ operations, durationMs }) => {
  engineCommandCount += 1;
  engineOperationCount += operations.length;
  engineDurationMs += durationMs;
});

const originalRender = instance.renderer.render;
Object.defineProperty(instance.renderer, "render", {
  configurable: true,
  writable: true,
  value: () => {
    const startedAt = performance.now();
    rendererRenderCount += 1;
    originalRender();
    rendererRenderDurationMs += performance.now() - startedAt;
  }
});

const metricsLog: string[] = [`mount: ${(performance.now() - startedAt).toFixed(2)}ms`];
const renderMetrics = (): void => {
  metrics.textContent = metricsLog.join("\n");
};

const appendMetric = (label: string, durationMs: number): void => {
  metricsLog.push(`${label}: ${durationMs.toFixed(2)}ms`);
  renderMetrics();
};

const appendTextMetric = (value: string): void => {
  metricsLog.push(value);
  renderMetrics();
};

const benchmarkSummary = new Map<number, BenchmarkSummaryEntry>();
const pivotBenchmarkSummary = new Map<number, PivotBenchmarkSummaryEntry>();

const formatMetric = (value: number | undefined, suffix = "ms"): string =>
  value === undefined ? "-" : `${value.toFixed(2)}${suffix}`;

const formatCount = (value: number | undefined): string => (value === undefined ? "-" : value.toLocaleString("pt-BR"));

const evaluateMutationBudget = (entry: BenchmarkSummaryEntry | undefined): BudgetStatus => {
  if (!entry?.sequencial || !entry.batch) {
    return {
      status: "pending",
      label: "Aguardando seq + batch"
    };
  }

  const durationPass = entry.batch.durationMs <= entry.sequencial.durationMs * MUTATION_BUDGETS.maxBatchDurationRatio;
  const commandPass = entry.batch.commandCount <= MUTATION_BUDGETS.maxBatchCommandCount;
  const renderPass = entry.batch.renderCount <= Math.max(1, entry.sequencial.renderCount * MUTATION_BUDGETS.maxBatchRenderRatio);

  if (durationPass && commandPass && renderPass) {
    return {
      status: "pass",
      label: "Dentro do limiar"
    };
  }

  return {
    status: "fail",
    label: "Fora do limiar"
  };
};

const evaluatePivotBudget = (entry: PivotBenchmarkSummaryEntry | undefined): BudgetStatus => {
  if (!entry?.client || !entry.worker || !entry.server) {
    return {
      status: "pending",
      label: "Aguardando client + worker + server"
    };
  }

  const workerDurationPass = entry.worker.durationMs <= entry.client.durationMs * PIVOT_BUDGETS.maxWorkerDurationRatio;
  const serverDurationPass = entry.server.durationMs <= entry.client.durationMs * PIVOT_BUDGETS.maxServerDurationRatio;
  const workerProgressPass = entry.worker.progressEvents >= PIVOT_BUDGETS.minProgressEvents;
  const serverProgressPass = entry.server.progressEvents >= PIVOT_BUDGETS.minProgressEvents;
  const resultShapePass = entry.worker.resultRowCount > 0 && entry.worker.resultColumnCount > 0 && entry.server.remote;

  if (workerDurationPass && serverDurationPass && workerProgressPass && serverProgressPass && resultShapePass) {
    return {
      status: "pass",
      label: "Dentro do limiar"
    };
  }

  return {
    status: "fail",
    label: "Fora do limiar"
  };
};

const renderSummary = (): void => {
  const rows = MUTATION_SCALES.map((scale) => {
    const entry = benchmarkSummary.get(scale);
    const sequencial = entry?.sequencial;
    const batch = entry?.batch;
    const gain = sequencial && batch ? sequencial.durationMs - batch.durationMs : undefined;
    const budget = evaluateMutationBudget(entry);

    return `
      <tr>
        <td style="padding:10px 12px;border-bottom:1px solid rgba(22,32,43,.08);font-weight:600;">${scale.toLocaleString("pt-BR")}</td>
        <td style="padding:10px 12px;border-bottom:1px solid rgba(22,32,43,.08);">${formatMetric(sequencial?.durationMs)}</td>
        <td style="padding:10px 12px;border-bottom:1px solid rgba(22,32,43,.08);">${formatMetric(batch?.durationMs)}</td>
        <td style="padding:10px 12px;border-bottom:1px solid rgba(22,32,43,.08);">${formatMetric(gain)}</td>
        <td style="padding:10px 12px;border-bottom:1px solid rgba(22,32,43,.08);">${formatCount(sequencial?.commandCount)}</td>
        <td style="padding:10px 12px;border-bottom:1px solid rgba(22,32,43,.08);">${formatCount(batch?.commandCount)}</td>
        <td style="padding:10px 12px;border-bottom:1px solid rgba(22,32,43,.08);">${formatCount(sequencial?.operationCount)}</td>
        <td style="padding:10px 12px;border-bottom:1px solid rgba(22,32,43,.08);">${formatCount(batch?.operationCount)}</td>
        <td style="padding:10px 12px;border-bottom:1px solid rgba(22,32,43,.08);">${formatCount(sequencial?.renderCount)}</td>
        <td style="padding:10px 12px;border-bottom:1px solid rgba(22,32,43,.08);">${formatCount(batch?.renderCount)}</td>
        <td style="padding:10px 12px;border-bottom:1px solid rgba(22,32,43,.08);font-weight:600;color:${budget.status === "fail" ? "#b91c1c" : budget.status === "pass" ? "#166534" : "#556170"};" data-benchmark-budget-status="${budget.status}">${budget.label}</td>
      </tr>
    `;
  }).join("");

  summary.innerHTML = `
    <div style="display:grid;gap:12px;">
      <div>
        <h2 style="margin:0 0 6px;font-size:20px;color:#16202b;">Resumo por escala</h2>
        <p style="margin:0;color:#556170;">Comparativo acumulado entre execução sequencial e batch real da engine.</p>
      </div>
      <table style="width:100%;border-collapse:collapse;font-size:13px;color:#24303d;min-width:880px;">
        <thead>
          <tr style="text-align:left;background:rgba(15,118,110,.08);">
            <th style="padding:10px 12px;">Escala</th>
            <th style="padding:10px 12px;">Seq ms</th>
            <th style="padding:10px 12px;">Batch ms</th>
            <th style="padding:10px 12px;">Ganho</th>
            <th style="padding:10px 12px;">Seq cmds</th>
            <th style="padding:10px 12px;">Batch cmds</th>
            <th style="padding:10px 12px;">Seq ops</th>
            <th style="padding:10px 12px;">Batch ops</th>
            <th style="padding:10px 12px;">Seq renders</th>
            <th style="padding:10px 12px;">Batch renders</th>
            <th style="padding:10px 12px;">Limiar</th>
          </tr>
        </thead>
        <tbody>${rows}</tbody>
      </table>
    </div>
  `;
};

const renderPivotSummary = (): void => {
  const rows = PIVOT_ROW_SCALES.map((scale) => {
    const entry = pivotBenchmarkSummary.get(scale);
    const client = entry?.client;
    const worker = entry?.worker;
    const server = entry?.server;
    const workerGain = client && worker ? client.durationMs - worker.durationMs : undefined;
    const serverGain = client && server ? client.durationMs - server.durationMs : undefined;
    const budget = evaluatePivotBudget(entry);

    return `
      <tr>
        <td style="padding:10px 12px;border-bottom:1px solid rgba(22,32,43,.08);font-weight:600;">${scale.toLocaleString("pt-BR")}</td>
        <td style="padding:10px 12px;border-bottom:1px solid rgba(22,32,43,.08);">${formatMetric(client?.durationMs)}</td>
        <td style="padding:10px 12px;border-bottom:1px solid rgba(22,32,43,.08);">${formatMetric(worker?.durationMs)}</td>
        <td style="padding:10px 12px;border-bottom:1px solid rgba(22,32,43,.08);">${formatMetric(server?.durationMs)}</td>
        <td style="padding:10px 12px;border-bottom:1px solid rgba(22,32,43,.08);">${formatMetric(workerGain)}</td>
        <td style="padding:10px 12px;border-bottom:1px solid rgba(22,32,43,.08);">${formatMetric(serverGain)}</td>
        <td style="padding:10px 12px;border-bottom:1px solid rgba(22,32,43,.08);">${formatCount(client?.progressEvents)}</td>
        <td style="padding:10px 12px;border-bottom:1px solid rgba(22,32,43,.08);">${formatCount(worker?.progressEvents)}</td>
        <td style="padding:10px 12px;border-bottom:1px solid rgba(22,32,43,.08);">${formatCount(server?.progressEvents)}</td>
        <td style="padding:10px 12px;border-bottom:1px solid rgba(22,32,43,.08);font-weight:600;color:${budget.status === "fail" ? "#b91c1c" : budget.status === "pass" ? "#166534" : "#556170"};" data-pivot-budget-status="${budget.status}">${budget.label}</td>
      </tr>
    `;
  }).join("");

  pivotSummary.innerHTML = `
    <div style="display:grid;gap:12px;">
      <div>
        <h3 style="margin:0 0 6px;font-size:18px;color:#16202b;">Resumo pivot por escala</h3>
        <p style="margin:0;color:#556170;">Client usa createPivotSheetAsync no main thread, worker executa o mesmo plano fora da UI e server usa o contrato remoto pivotSheet.</p>
      </div>
      <table style="width:100%;border-collapse:collapse;font-size:13px;color:#24303d;min-width:980px;">
        <thead>
          <tr style="text-align:left;background:rgba(29,78,216,.08);">
            <th style="padding:10px 12px;">Linhas</th>
            <th style="padding:10px 12px;">Client ms</th>
            <th style="padding:10px 12px;">Worker ms</th>
            <th style="padding:10px 12px;">Server ms</th>
            <th style="padding:10px 12px;">Ganho worker</th>
            <th style="padding:10px 12px;">Ganho server</th>
            <th style="padding:10px 12px;">Client progress</th>
            <th style="padding:10px 12px;">Worker progress</th>
            <th style="padding:10px 12px;">Server progress</th>
            <th style="padding:10px 12px;">Limiar</th>
          </tr>
        </thead>
        <tbody>${rows}</tbody>
      </table>
    </div>
  `;
};

const updateSummary = (result: BenchmarkRunResult): void => {
  const entry = benchmarkSummary.get(result.mutationCount) ?? {};
  entry[result.mode] = result;
  benchmarkSummary.set(result.mutationCount, entry);
  renderSummary();
};

const updatePivotSummary = (result: PivotBenchmarkRunResult): void => {
  const entry = pivotBenchmarkSummary.get(result.sourceRowCount) ?? {};
  entry[result.mode] = result;
  pivotBenchmarkSummary.set(result.sourceRowCount, entry);
  renderPivotSummary();
};

const buildNextValue = (run: number, index: number): number => run * 1000 + index * 3;

let mutationRun = 0;
const lastSequentialDurations = new Map<number, number>();

const getEngineMetricsSnapshot = (): EngineMetricsSnapshot => ({
  commandCount: engineCommandCount,
  operationCount: engineOperationCount,
  engineDurationMs
});

const getRendererMetricsSnapshot = (): RendererMetricsSnapshot => ({
  renderCount: rendererRenderCount,
  renderDurationMs: rendererRenderDurationMs
});

const getEngineMetricsDelta = (before: EngineMetricsSnapshot): Pick<
  BenchmarkRunResult,
  "commandCount" | "operationCount" | "engineDurationMs"
> => ({
  commandCount: engineCommandCount - before.commandCount,
  operationCount: engineOperationCount - before.operationCount,
  engineDurationMs: engineDurationMs - before.engineDurationMs
});

const getRendererMetricsDelta = (before: RendererMetricsSnapshot): Pick<
  BenchmarkRunResult,
  "renderCount" | "renderDurationMs"
> => ({
  renderCount: rendererRenderCount - before.renderCount,
  renderDurationMs: rendererRenderDurationMs - before.renderDurationMs
});

const getRenderedCellCount = (): number => host.querySelectorAll(".excelsior-cell").length;

const setScale = (value: number): void => {
  scaleSelect.value = String(value);
};

const setControlsDisabled = (disabled: boolean): void => {
  scaleSelect.disabled = disabled;
  runSuiteButton.disabled = disabled;
  sequentialButton.disabled = disabled;
  batchButton.disabled = disabled;
  clearMetricsButton.disabled = disabled;
  pivotScaleSelect.disabled = disabled;
  runPivotSuiteButton.disabled = disabled;
  pivotClientButton.disabled = disabled;
  pivotWorkerButton.disabled = disabled;
  pivotServerButton.disabled = disabled;
  for (const presetButton of presetButtons) {
    presetButton.disabled = disabled;
  }
};

const getPivotScale = (): number => Number(pivotScaleSelect.value);

const createPivotBenchmarkSheet = (sourceRowCount: number): WorkbookDataInput => {
  const cells: Record<string, { value: string | number; computedValue: string | number }> = {
    "0:0": { value: "Region", computedValue: "Region" },
    "0:1": { value: "Quarter", computedValue: "Quarter" },
    "0:2": { value: "Channel", computedValue: "Channel" },
    "0:3": { value: "Revenue", computedValue: "Revenue" },
    "0:4": { value: "Units", computedValue: "Units" }
  };
  const regions = ["North", "South", "East", "West"];
  const quarters = ["Q1", "Q2", "Q3", "Q4"];
  const channels = ["Retail", "Online", "Partner"];

  for (let row = 1; row <= sourceRowCount; row += 1) {
    cells[`${row}:0`] = { value: regions[row % regions.length]!, computedValue: regions[row % regions.length]! };
    cells[`${row}:1`] = { value: quarters[row % quarters.length]!, computedValue: quarters[row % quarters.length]! };
    cells[`${row}:2`] = { value: channels[row % channels.length]!, computedValue: channels[row % channels.length]! };
    cells[`${row}:3`] = { value: row * 11, computedValue: row * 11 };
    cells[`${row}:4`] = { value: (row % 9) + 1, computedValue: (row % 9) + 1 };
  }

  return {
    name: `Pivot ${sourceRowCount.toLocaleString("pt-BR")}`,
    rowCount: sourceRowCount + 1,
    columnCount: 5,
    cells,
    merges: [],
    columns: {},
    rows: {}
  };
};

const createPivotInput = (sourceSheetId: string, sourceRowCount: number, executionMode?: PivotSheetInput["executionMode"]): PivotSheetInput => ({
  sourceSheetId,
  sourceRange: {
    start: { row: 0, col: 0 },
    end: { row: sourceRowCount, col: 4 }
  },
  rows: ["Region", "Channel"],
  columns: ["Quarter"],
  values: [
    { field: "Revenue", aggregate: "sum", as: "Revenue" },
    { field: "Units", aggregate: "sum", as: "Units" }
  ],
  executionMode
});

const createPivotBenchmarkEngine = (sourceRowCount: number): { engine: WorkbookEngine; sourceSheetId: string; snapshot: WorkbookModel } => {
  const engine = new WorkbookEngine({
    data: [createPivotBenchmarkSheet(sourceRowCount)]
  });

  return {
    engine,
    sourceSheetId: engine.getActiveSheet().id,
    snapshot: engine.getSnapshot()
  };
};

type PivotWorkerJob = {
  resolve: (value: { durationMs: number; pivotSheet: WorkbookDataInput; progressEvents: number }) => void;
  reject: (reason?: unknown) => void;
  progressEvents: number;
};

let pivotWorkerRequestSequence = 0;
const pivotWorkerJobs = new Map<string, PivotWorkerJob>();

pivotWorker.addEventListener("message", (event: MessageEvent<PivotBenchmarkWorkerResponse>) => {
  const job = pivotWorkerJobs.get(event.data.requestId);
  if (!job) {
    return;
  }

  if (event.data.kind === "progress") {
    job.progressEvents += 1;
    return;
  }

  pivotWorkerJobs.delete(event.data.requestId);

  if (event.data.kind === "error") {
    const error = new Error(event.data.message);
    error.name = event.data.code ?? "PivotWorkerError";
    job.reject(error);
    return;
  }

  job.resolve({
    durationMs: event.data.durationMs,
    pivotSheet: event.data.pivotSheet,
    progressEvents: job.progressEvents
  });
});

const runPivotWorkerJob = (
  workbook: WorkbookModel,
  input: PivotSheetInput
): Promise<{ durationMs: number; pivotSheet: WorkbookDataInput; progressEvents: number }> => {
  const requestId = `pivot-worker:${pivotWorkerRequestSequence++}`;

  return new Promise((resolve, reject) => {
    pivotWorkerJobs.set(requestId, {
      resolve,
      reject,
      progressEvents: 0
    });

    pivotWorker.postMessage({
      kind: "buildPivot",
      requestId,
      workbook,
      input
    } satisfies PivotBenchmarkWorkerRequest);
  });
};

const toPivotBenchmarkResult = (
  mode: PivotBenchmarkMode,
  label: string,
  sourceRowCount: number,
  durationMs: number,
  progressEvents: number,
  pivotSheet: WorkbookDataInput,
  remote: boolean
): PivotBenchmarkRunResult => ({
  mode,
  label,
  sourceRowCount,
  durationMs,
  progressEvents,
  resultRowCount: pivotSheet.rowCount ?? 0,
  resultColumnCount: pivotSheet.columnCount ?? 0,
  remote
});

const runPivotClientBenchmark = async (sourceRowCount: number): Promise<PivotBenchmarkRunResult> => {
  const { engine, sourceSheetId } = createPivotBenchmarkEngine(sourceRowCount);
  let progressEvents = 0;
  const startedAt = performance.now();
  const pivotSheet = await engine.createPivotSheetAsync(createPivotInput(sourceSheetId, sourceRowCount, "client"), {
    chunkSize: PIVOT_CHUNK_SIZE,
    yieldControl: async () => Promise.resolve(),
    onProgress: () => {
      progressEvents += 1;
    }
  });

  return toPivotBenchmarkResult(
    "client",
    `${sourceRowCount.toLocaleString("pt-BR")} linhas pivot client`,
    sourceRowCount,
    performance.now() - startedAt,
    progressEvents,
    pivotSheet,
    false
  );
};

const runPivotWorkerBenchmark = async (sourceRowCount: number): Promise<PivotBenchmarkRunResult> => {
  const { snapshot, sourceSheetId } = createPivotBenchmarkEngine(sourceRowCount);
  const result = await runPivotWorkerJob(snapshot, createPivotInput(sourceSheetId, sourceRowCount, "client"));

  return toPivotBenchmarkResult(
    "worker",
    `${sourceRowCount.toLocaleString("pt-BR")} linhas pivot worker`,
    sourceRowCount,
    result.durationMs,
    result.progressEvents,
    result.pivotSheet,
    false
  );
};

const runPivotServerBenchmark = async (sourceRowCount: number): Promise<PivotBenchmarkRunResult> => {
  const { engine, sourceSheetId, snapshot } = createPivotBenchmarkEngine(sourceRowCount);
  let progressEvents = 0;

  engine.setRowModel(
    sourceSheetId,
    new ServerSideRowModel({
      rowCount: "unknown",
      dataSource: {
        getRows: async (request: DataRequest) => {
          if (request.requestKind === "pivotSheet" && request.pivotInput) {
            const result = await runPivotWorkerJob(snapshot, request.pivotInput);
            progressEvents += result.progressEvents;

            return {
              rows: [],
              pivotSheet: result.pivotSheet
            };
          }

          return {
            rows: []
          };
        }
      }
    })
  );

  const startedAt = performance.now();
  const pivotSheet = await engine.createPivotSheetAsync(createPivotInput(sourceSheetId, sourceRowCount, "server"));

  return toPivotBenchmarkResult(
    "server",
    `${sourceRowCount.toLocaleString("pt-BR")} linhas pivot server`,
    sourceRowCount,
    performance.now() - startedAt,
    progressEvents,
    pivotSheet,
    true
  );
};

const logPivotBenchmarkRun = (result: PivotBenchmarkRunResult): void => {
  updatePivotSummary(result);
  appendMetric(result.label, result.durationMs);
  appendTextMetric(
    `pivot ${result.mode}: ${result.sourceRowCount.toLocaleString("pt-BR")} linhas fonte, ${result.resultRowCount}x${result.resultColumnCount} saída, ${result.progressEvents} eventos de progresso, remoto=${result.remote ? "sim" : "não"}`
  );
};

const runSequentialBenchmark = (mutationCount: number): BenchmarkRunResult => {
  const sheet = instance.engine.getActiveSheet();
  mutationRun += 1;
  const currentRun = mutationRun;
  const beforeMetrics = getEngineMetricsSnapshot();
  const beforeRendererMetrics = getRendererMetricsSnapshot();
  const startedAt = performance.now();

  for (let index = 0; index < mutationCount; index += 1) {
    const target = getMutationTarget(index);
    instance.engine.setCellValue({
      sheetId: sheet.id,
      row: target.row,
      col: target.col,
      value: buildNextValue(currentRun, index)
    });
  }

  const durationMs = performance.now() - startedAt;
  lastSequentialDurations.set(mutationCount, durationMs);
  return {
    mode: "sequencial",
    label: `${mutationCount.toLocaleString("pt-BR")} mutações sequenciais`,
    mutationCount,
    durationMs,
    ...getEngineMetricsDelta(beforeMetrics),
    ...getRendererMetricsDelta(beforeRendererMetrics),
    renderedCellCount: getRenderedCellCount()
  };
};

const runBatchBenchmark = (mutationCount: number): BenchmarkRunResult => {
  const sheet = instance.engine.getActiveSheet();
  mutationRun += 1;
  const currentRun = mutationRun;
  const updates: Parameters<typeof instance.engine.updateCells>[0]["updates"] = [];

  for (let index = 0; index < mutationCount; index += 1) {
    const target = getMutationTarget(index);
    const nextValue = buildNextValue(currentRun, index);
    updates.push({
      row: target.row,
      col: target.col,
      value: nextValue
    });
  }

  const beforeMetrics = getEngineMetricsSnapshot();
  const beforeRendererMetrics = getRendererMetricsSnapshot();
  const startedAt = performance.now();
  instance.engine.updateCells({
    sheetId: sheet.id,
    updates,
    affectedRanges: [
      {
        start: { row: 0, col: BENCHMARK_MUTATION_BASE_COLUMN },
        end: getMutationTarget(mutationCount - 1)
      }
    ]
  });
  const durationMs = performance.now() - startedAt;

  return {
    mode: "batch",
    label: `${mutationCount.toLocaleString("pt-BR")} mutações em batch`,
    mutationCount,
    durationMs,
    ...getEngineMetricsDelta(beforeMetrics),
    ...getRendererMetricsDelta(beforeRendererMetrics),
    renderedCellCount: getRenderedCellCount()
  };
};

const logBenchmarkRun = (result: BenchmarkRunResult): void => {
  updateSummary(result);
  appendMetric(result.label, result.durationMs);
  appendTextMetric(describeMutationWindow(result.mutationCount));
  appendTextMetric(
    `engine: ${result.commandCount} comandos, ${result.operationCount} ops, ${result.engineDurationMs.toFixed(2)}ms internos`
  );
  appendTextMetric(
    `renderer: ${result.renderCount} renders, ${result.renderDurationMs.toFixed(2)}ms acumulados, ${result.renderedCellCount} células visíveis`
  );
};

renderMetrics();
renderSummary();
renderPivotSummary();

sequentialButton.addEventListener("click", () => {
  const mutationCount = getMutationCount();
  const result = runSequentialBenchmark(mutationCount);
  logBenchmarkRun(result);
});

batchButton.addEventListener("click", () => {
  const mutationCount = getMutationCount();
  const result = runBatchBenchmark(mutationCount);
  logBenchmarkRun(result);
  const lastSequentialDuration = lastSequentialDurations.get(mutationCount);
  if (lastSequentialDuration !== undefined) {
    appendMetric(
      `ganho batch vs sequencial (${mutationCount.toLocaleString("pt-BR")})`,
      lastSequentialDuration - result.durationMs
    );
  }
});

runSuiteButton.addEventListener("click", () => {
  setControlsDisabled(true);
  appendTextMetric("suite: iniciando 1k/10k/100k sequencial + batch");

  try {
    for (const mutationCount of MUTATION_SCALES) {
      setScale(mutationCount);
      appendTextMetric(`suite: escala ${mutationCount.toLocaleString("pt-BR")}`);

      const sequentialResult = runSequentialBenchmark(mutationCount);
      logBenchmarkRun(sequentialResult);

      const batchResult = runBatchBenchmark(mutationCount);
      logBenchmarkRun(batchResult);

      appendMetric(
        `suite ganho batch vs sequencial (${mutationCount.toLocaleString("pt-BR")})`,
        sequentialResult.durationMs - batchResult.durationMs
      );
    }
  } finally {
    setControlsDisabled(false);
  }
});

clearMetricsButton.addEventListener("click", () => {
  metricsLog.splice(1);
  benchmarkSummary.clear();
  pivotBenchmarkSummary.clear();
  renderMetrics();
  renderSummary();
  renderPivotSummary();
});

pivotClientButton.addEventListener("click", () => {
  void (async () => {
    setControlsDisabled(true);

    try {
      logPivotBenchmarkRun(await runPivotClientBenchmark(getPivotScale()));
    } finally {
      setControlsDisabled(false);
    }
  })();
});

pivotWorkerButton.addEventListener("click", () => {
  void (async () => {
    setControlsDisabled(true);

    try {
      logPivotBenchmarkRun(await runPivotWorkerBenchmark(getPivotScale()));
    } finally {
      setControlsDisabled(false);
    }
  })();
});

pivotServerButton.addEventListener("click", () => {
  void (async () => {
    setControlsDisabled(true);

    try {
      logPivotBenchmarkRun(await runPivotServerBenchmark(getPivotScale()));
    } finally {
      setControlsDisabled(false);
    }
  })();
});

runPivotSuiteButton.addEventListener("click", () => {
  void (async () => {
    setControlsDisabled(true);
    appendTextMetric("pivot-suite: iniciando client + worker + server");

    try {
      for (const sourceRowCount of PIVOT_ROW_SCALES) {
        pivotScaleSelect.value = String(sourceRowCount);
        appendTextMetric(`pivot-suite: escala ${sourceRowCount.toLocaleString("pt-BR")} linhas`);
        logPivotBenchmarkRun(await runPivotClientBenchmark(sourceRowCount));
        logPivotBenchmarkRun(await runPivotWorkerBenchmark(sourceRowCount));
        logPivotBenchmarkRun(await runPivotServerBenchmark(sourceRowCount));
      }
    } finally {
      setControlsDisabled(false);
    }
  })();
});

for (const presetButton of presetButtons) {
  presetButton.addEventListener("click", () => {
    const nextScale = Number(presetButton.dataset.scalePreset);
    if (Number.isFinite(nextScale)) {
      setScale(nextScale);
    }
  });
}