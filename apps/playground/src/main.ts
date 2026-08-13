import "@excelsior/styles";
import { createSpreadsheet } from "@excelsior/vanilla";

const app = document.querySelector<HTMLDivElement>("#app");

if (!app) {
  throw new Error("Playground root was not found.");
}

app.innerHTML = `
  <div style="min-height:100vh;padding:24px;background:linear-gradient(180deg,#fcfaf4,#f2eadc);font-family:Aptos,Segoe UI Variable,Segoe UI,sans-serif;">
    <div style="max-width:1200px;margin:0 auto;display:grid;gap:18px;">
      <div>
        <h1 style="margin:0 0 6px;font-size:40px;color:#16202b;">Excelsior Playground</h1>
        <p style="margin:0;color:#556170;max-width:720px;">Core sem framework, renderização DOM virtualizada, fórmulas básicas e emissão de operações serializáveis.</p>
      </div>
      <div>
        <div id="sheet-host" style="height:70vh;"></div>
      </div>
    </div>
  </div>
`;

const host = document.querySelector<HTMLDivElement>("#sheet-host");

if (!host) {
  throw new Error("Playground UI did not initialize.");
}

createSpreadsheet(host, {
  data: [
    {
      name: "Revenue",
      rowCount: 400,
      columnCount: 40,
      cells: {
        "0:0": { value: "Region", computedValue: "Region" },
        "0:1": { value: "Q1", computedValue: "Q1" },
        "0:2": { value: "Q2", computedValue: "Q2" },
        "1:0": { value: "LATAM", computedValue: "LATAM" },
        "1:1": { value: 120000, computedValue: 120000 },
        "1:2": { value: 140000, computedValue: 140000 },
        "1:3": { value: "=B2+C2", formula: "=B2+C2", computedValue: 260000 }
      }
    }
  ],
  settings: {
    clipboardPolicy: "safe-html",
    rowHeight: 28,
    columnWidth: 112
  },
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
  }
  ,
  chartInsertPreview: true,
});