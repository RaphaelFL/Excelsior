import "@excelsior/styles";
import { WorkbookEngine, type SpreadsheetOperation, type WorkbookModel } from "@excelsior/core";
import { BasicFormulaEngine } from "@excelsior/formulas";
import { DomSpreadsheetRenderer } from "@excelsior/renderer-dom";

const app = document.querySelector<HTMLDivElement>("#app");

if (!app) {
  throw new Error("Collaboration demo root was not found.");
}

const clientId = crypto.randomUUID();

app.innerHTML = `
  <div style="min-height:100vh;padding:24px;background:linear-gradient(180deg,#f7f2e8,#e8dfcc);font-family:Aptos,Segoe UI Variable,Segoe UI,sans-serif;">
    <div style="max-width:1360px;margin:0 auto;display:grid;gap:18px;">
      <div style="display:flex;justify-content:space-between;align-items:flex-end;gap:16px;flex-wrap:wrap;">
        <div>
          <h1 style="margin:0 0 6px;font-size:40px;color:#16202b;">Excelsior Collaboration Demo</h1>
          <p style="margin:0;color:#556170;max-width:760px;">Edite a planilha em duas janelas ao mesmo tempo. O backend replica <strong>Op[]</strong> em memoria e reaplica mudancas remotas no engine.</p>
        </div>
        <div style="display:flex;gap:10px;align-items:center;flex-wrap:wrap;">
          <span id="status" style="padding:8px 12px;border-radius:999px;background:rgba(255,255,255,.75);color:#0f766e;">conectando...</span>
          <button id="reset" style="padding:12px 18px;border:none;border-radius:999px;background:#0f766e;color:#fff;cursor:pointer;">Resetar estado</button>
        </div>
      </div>
      <div style="display:grid;grid-template-columns:minmax(0,1fr) 340px;gap:18px;align-items:stretch;">
        <div id="sheet-host" style="height:72vh;"></div>
        <aside style="padding:18px;border-radius:18px;background:rgba(255,255,255,.82);box-shadow:0 18px 45px rgba(31,42,55,.14);display:grid;gap:14px;align-content:start;">
          <div>
            <h2 style="margin:0 0 8px;color:#16202b;">Fluxo de replicacao</h2>
            <p style="margin:0;color:#556170;font-size:14px;line-height:1.45;">Mudancas locais viram <strong>Op[]</strong>, o servidor aplica no workbook em memoria e transmite para as outras sessoes via SSE.</p>
          </div>
          <div>
            <h3 style="margin:0 0 8px;color:#16202b;font-size:16px;">Ultimo pacote</h3>
            <pre id="ops-log" style="margin:0;white-space:pre-wrap;word-break:break-word;color:#556170;font-size:12px;"></pre>
          </div>
        </aside>
      </div>
    </div>
  </div>
`;

const host = document.querySelector<HTMLDivElement>("#sheet-host");
const status = document.querySelector<HTMLSpanElement>("#status");
const log = document.querySelector<HTMLPreElement>("#ops-log");
const resetButton = document.querySelector<HTMLButtonElement>("#reset");

if (!host || !status || !log || !resetButton) {
  throw new Error("Collaboration demo UI did not initialize.");
}

const setStatus = (text: string, tone: "ok" | "warn") => {
  status.textContent = text;
  status.style.color = tone === "ok" ? "#0f766e" : "#b42318";
};

const fetchWorkbook = async (): Promise<WorkbookModel> => {
  const response = await fetch("/api/workbook");
  if (!response.ok) {
    throw new Error("Failed to fetch workbook state.");
  }
  return response.json() as Promise<WorkbookModel>;
};

const start = async () => {
  const initialWorkbook = await fetchWorkbook();
  const engine = WorkbookEngine.fromJSON(initialWorkbook, new BasicFormulaEngine());
  const renderer = new DomSpreadsheetRenderer(host, engine, {
    chartLimits: {
      maxChartsPerSheet: 50,
      maxRangeCells: 100000,
      maxSeriesPerChart: 24,
      maxPointsPerChart: 50000
    },
    chartPerformance: {
      interactionThrottleMs: 16,
      offscreenMarginPx: 320,
      skipOffscreenPreview: true
    },
    chartInsertPreview: true,
    onChange: async (operations: SpreadsheetOperation[]) => {
      log.textContent = JSON.stringify(operations, null, 2);
      await fetch("/api/ops", {
        method: "POST",
        headers: {
          "content-type": "application/json"
        },
        body: JSON.stringify({ clientId, operations })
      });
    }
  });

  const events = new EventSource(`/api/events?clientId=${encodeURIComponent(clientId)}`);
  events.onopen = () => setStatus("conectado", "ok");
  events.onerror = () => setStatus("reconectando...", "warn");
  events.onmessage = (event) => {
    const payload = JSON.parse(event.data) as {
      type?: "ops" | "reset";
      clientId: string;
      operations: SpreadsheetOperation[];
    };

    if (payload.type === "reset") {
      void fetchWorkbook().then((nextWorkbook) => {
        engine.loadFromJSON(nextWorkbook);
        renderer.render();
      });
      return;
    }

    if (payload.clientId === clientId) {
      return;
    }

    log.textContent = JSON.stringify(payload.operations, null, 2);
    engine.applyOperations(payload.operations);
  };

  resetButton.addEventListener("click", async () => {
    await fetch("/api/reset", { method: "POST" });
    const nextWorkbook = await fetchWorkbook();
    engine.loadFromJSON(nextWorkbook);
    renderer.render();
  });
};

void start();