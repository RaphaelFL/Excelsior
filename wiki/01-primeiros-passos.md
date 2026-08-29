# Primeiros Passos

## Instalacao

No workspace raiz:

```bash
npm install
```

## Scripts principais

```bash
npm run typecheck
npm run test
npm run build
npm run audit:prod
npm run release:check
```

## Como escolher o pacote certo

- Use `@excelsior/core` quando voce quer a engine sem UI.
- Use `@excelsior/vanilla` quando voce quer a planilha pronta em um elemento DOM.
- Use `@excelsior/xlsx` para importacao/exportacao.
- Use `@excelsior/devtools` para inspecionar eventos do workbook.

## Exemplo minimo com o core

```ts
import { WorkbookEngine } from "@excelsior/core";
import { BasicFormulaEngine } from "@excelsior/formulas";

const engine = new WorkbookEngine(
  {
    data: [
      {
        name: "Sheet1",
        rowCount: 10,
        columnCount: 10,
        cells: {
          "0:0": { value: 10, computedValue: 10 },
          "1:0": { value: "=A1*2", formula: "=A1*2", computedValue: 20 }
        },
        merges: [],
        columns: {},
        rows: {}
      }
    ]
  },
  new BasicFormulaEngine()
);

const activeSheet = engine.getActiveSheet();
console.log(engine.getDisplayValue(activeSheet.id, 1, 0));
```

## Exemplo minimo com o wrapper vanilla

```ts
import { createSpreadsheet } from "@excelsior/vanilla";

const host = document.getElementById("app")!;

const instance = createSpreadsheet(host, {
  data: [
    {
      name: "Sales",
      rowCount: 20,
      columnCount: 8,
      cells: {
        "0:0": { value: "Region", computedValue: "Region" },
        "0:1": { value: "Revenue", computedValue: "Revenue" },
        "1:0": { value: "North", computedValue: "North" },
        "1:1": { value: 120, computedValue: 120 }
      },
      merges: [],
      columns: {},
      rows: {}
    }
  ],
  onChange: (operations) => {
    console.log("Ops serializaveis", operations);
  }
});

// quando desmontar
instance.destroy();
```

## Estrutura basica de workbook

Um workbook e composto por:

- `metadata`
- `settings`
- `sheets[]`
- `activeSheetId`

Cada sheet pode ter:

- `cells`
- `merges`
- `columns`
- `rows`
- `selection`
- `metadata`
- `frozenRows` e `frozenColumns`
- `conditionalFormats`

## Proximo passo

Leia [02-core-e-engine.md](./02-core-e-engine.md) para aprender a usar a engine como API publica.