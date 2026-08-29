# Exemplos Práticos

## Planilha editável no navegador

```ts
import { createSpreadsheet } from "@excelsior/vanilla";

const container = document.querySelector<HTMLElement>("#spreadsheet");
if (!container) throw new Error("Container não encontrado.");

const spreadsheet = createSpreadsheet(container, {
  data: [{ name: "Vendas" }]
});
const sheet = spreadsheet.engine.getActiveSheet();

spreadsheet.engine.setCellValue({ sheetId: sheet.id, row: 0, col: 0, value: 10 });
spreadsheet.engine.setCellValue({ sheetId: sheet.id, row: 1, col: 0, value: 20 });
spreadsheet.engine.setCellValue({ sheetId: sheet.id, row: 2, col: 0, value: "=MEDIAN(A1:A2)" });
spreadsheet.engine.setCellStyle({
  sheetId: sheet.id,
  row: 2,
  col: 0,
  style: { fontFamily: "Georgia", fontSize: 18, fontWeight: "bold" }
});
```

## Validação compatível com Excel

```ts
spreadsheet.engine.setCellValidation({
  sheetId: sheet.id,
  row: 0,
  col: 1,
  validation: {
    rules: [{ type: "number", min: 1, max: 100, message: "Use um valor de 1 a 100" }]
  }
});
```

Regras de número, intervalo, comprimento e listas são gravadas em OOXML nativo. Regras específicas do Excelsior continuam preservadas na metadata privada versionada.

## Exportação XLSX no Node.js

```ts
import { writeFile } from "node:fs/promises";
import { exportWorkbookEngineToXlsx } from "@excelsior/xlsx";

const bytes = await exportWorkbookEngineToXlsx(spreadsheet.engine);
await writeFile("vendas.xlsx", bytes);
```

Comentários simples são exportados como comentários legados do Excel. Threads, respostas e estado de resolução permanecem no fallback privado para roundtrip completo no Excelsior.

Ao desmontar a tela, chame `spreadsheet.destroy()` para liberar listeners e runtimes.