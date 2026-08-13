# Dados Remotos, Pivot e Integracoes

## Row models

O core suporta quatro estrategias principais:

- `ClientSideRowModel`
- `ViewportRowModel`
- `InfiniteRowModel`
- `ServerSideRowModel`

### Associando um row model a uma sheet

```ts
import { ServerSideRowModel } from "@excelsior/core";

const sheetId = engine.getActiveSheet().id;

engine.setRowModel(
  sheetId,
  new ServerSideRowModel({
    rowCount: "unknown",
    dataSource: {
      async getRows(request) {
        return {
          rows: [{ index: request.startRow }],
          totalRows: 1000
        };
      }
    }
  })
);
```

APIs publicas relacionadas:

- `setRowModel(sheetId, rowModel)`
- `clearRowModel(sheetId)`
- `getRowModel(sheetId)`
- `updateRemoteRowModel(sheetId, update)`
- `getRemoteRowModelRequest(sheetId)`
- `setRemoteGroupExpanded(sheetId, path, expanded)`

## Contrato de DataSource

O datasource remoto e controlado pela aplicacao, nao pela biblioteca.

O request serializavel pode carregar, dependendo do caso:

- `startRow` e `endRow`
- `sortModel`
- `filterModel`
- `groupKeys`
- `expandedGroupPaths`
- `pivotModel`
- `aggregateModel`
- `visibleColumns`
- `requestKind`
- `pivotInput`
- `requestId`

Tipos uteis:

- `DataSource`
- `DataRequest`
- `DataResponse`
- `DataSortModel`
- `DataFilterModel`
- `DataAggregateModel`
- `DataPivotModel`

## Pivot no core

### Criacao client-side

```ts
const pivotSheet = engine.createPivotSheet({
  sourceSheetId,
  sourceRange: {
    start: { row: 0, col: 0 },
    end: { row: 100, col: 4 }
  },
  rows: ["Region"],
  columns: ["Quarter"],
  values: [{ field: "Revenue", aggregate: "sum", as: "Revenue" }]
});
```

### Criacao async com progresso

```ts
const pivotSheet = await engine.createPivotSheetAsync(input, {
  chunkSize: 500,
  onProgress: (progress) => {
    console.log(progress.phase, progress.completed, progress.total);
  },
  signal
});
```

### Inferencia de configuracao

```ts
const inferred = engine.inferPivotSheet({
  sourceSheetId,
  sourceRange
});
```

### Persistencia como derived view

Use quando voce quer manter a pivot como sheet viva do workbook:

- `addPivotSheet()`
- `addPivotSheetAsync()`
- `getPivotSheetDefinition()`
- `getPivotSheetViewDefinition()`
- `setPivotSheetAutoRefresh()`
- `updatePivotSheet()`
- `refreshPivotSheet()`

### Modos de execucao

- `client`
  - monta a pivot localmente.
- `server`
  - exige row model remoto com suporte a `requestKind: "pivotSheet"`.
- `auto`
  - usa remoto quando houver row model remoto compativel; caso contrario cai para client-side.

## Pivot server-side

Quando a source esta em row model remoto, a materializacao server-side pode devolver a sheet pronta via `DataResponse.pivotSheet`.

Isso e especialmente util quando:

- a fonte excede `maxPivotSourceRows`
- a agregacao e cara demais para o main thread
- o backend ja sabe materializar o resultado analitico

## Importacao e exportacao XLSX

O pacote `@excelsior/xlsx` oferece:

- `exportWorkbookToXlsx(workbook)`
- `importWorkbookFromXlsx(input)`
- `exportTableToXlsx(rows, options)`
- `importTableFromXlsx(input, options)`

### Workbook completo

```ts
import { exportWorkbookToXlsx, importWorkbookFromXlsx } from "@excelsior/xlsx";

const bytes = await exportWorkbookToXlsx(engine.getSnapshot());
const restored = await importWorkbookFromXlsx(bytes);
```

### Tabela com schema

```ts
import { exportTableToXlsx, importTableFromXlsx } from "@excelsior/xlsx";

const bytes = await exportTableToXlsx([{ name: "Ada", score: 99 }], {
  sheet: "People",
  columns: [
    {
      width: 24,
      header: { value: "Name", fontWeight: "bold" },
      cell: (row) => ({ value: row.name })
    },
    {
      width: 12,
      header: { value: "Score", fontWeight: "bold" },
      cell: (row) => ({ value: row.score, type: Number })
    }
  ]
});

const table = await importTableFromXlsx(bytes, {
  sheet: "People",
  schema: {
    name: { column: "Name", type: String, required: true },
    score: { column: "Score", type: Number, required: true }
  }
});
```

O adapter preserva formulas, estilos, merges, widths e heights.

## Devtools

Use `attachWorkbookDevtools()` para observar os eventos publicos do workbook.

```ts
import { attachWorkbookDevtools } from "@excelsior/devtools";

const session = attachWorkbookDevtools(engine, {
  onEvent: (event) => console.log(event.name, event.payload)
});

const snapshot = session.snapshot();
session.stop();
```

## Benchmark

## Proximo passo

Leia [05-seguranca-performance-e-operacao.md](./05-seguranca-performance-e-operacao.md) para limites, erros e operacao segura.