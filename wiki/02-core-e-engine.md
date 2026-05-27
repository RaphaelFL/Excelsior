# Core e Engine

## WorkbookEngine

`WorkbookEngine` e a API central do projeto. Ele concentra estado, historico, eventos, formulas, row models e pivot.

### Criacao

```ts
import { WorkbookEngine } from "@excelsior/core";
import { BasicFormulaEngine } from "@excelsior/formulas";

const engine = new WorkbookEngine(
  {
    data: [{ name: "Sheet1" }],
    settings: {
      maxRows: 1000,
      maxColumns: 100,
      maxCellLength: 5000,
      maxFormulaLength: 2048,
      maxPasteCells: 10000,
      maxRecalcCells: 10000,
      maxPivotSourceRows: 5000,
      rowHeight: 28,
      columnWidth: 120,
      viewportBuffer: 6,
      maxHistorySize: 100,
      enableFormulas: true,
      clipboardPolicy: "text-only"
    }
  },
  new BasicFormulaEngine()
);
```

### Leitura de estado

Metodos publicos mais usados:

- `getSnapshot()` ou `toJSON()` para serializar o workbook inteiro.
- `loadFromJSON(snapshot)` para reidratar.
- `getActiveSheet()` para ler a sheet ativa.
- `getCell(sheetId, row, col)` para ler a celula estruturada.
- `getDisplayValue(sheetId, row, col)` para ler o valor formatado para a UI.
- `getSelection(sheetId)`, `getMerge(sheetId, row, col)`, `getColumnSchema(sheetId, col)`, `getRowSchema(sheetId, row)`.
- `getFrozenPane(sheetId)` para panes congeladas.

## Edicao basica

### Valor de celula

```ts
engine.setCellValue({
  sheetId,
  row: 1,
  col: 1,
  value: "42"
});
```

### Batch de upserts

Use `updateCells()` quando varias celulas precisam ser atualizadas em uma unica operacao undoable.

```ts
engine.updateCells({
  sheetId,
  updates: [
    { row: 1, col: 1, value: 10 },
    { row: 1, col: 2, value: 20 }
  ]
});
```

### Transacoes keyed

Use `applyCellTransaction()` quando voce ja tem as chaves `row:col`.

```ts
engine.applyCellTransaction({
  sheetId,
  changes: [
    { type: "upsert", key: "1:1", value: 10 },
    { type: "remove", key: "4:2" }
  ]
});
```

### Operacoes serializaveis

Use `applyOperations()` para replay remoto ou colaboracao.

```ts
engine.applyOperations(operationsRecebidasDeOutroNo);
```

## Undo/Redo e no-op

- `undo()` e `redo()` fazem parte da API publica.
- Escritas redundantes sao tratadas como no-op em `setCellValue()`, `updateCells()` e `applyCellTransaction()`.

## Estrutura da planilha

### Sheets

- `addSheet()`
- `deleteSheet()`
- `setActiveSheet()`

### Linhas e colunas

- `insertRows()` e `deleteRows()`
- `insertColumns()` e `deleteColumns()`
- `resizeRow()` e `resizeColumn()`
- `freezeRows()` e `freezeColumns()`
- `setRowsHidden()` e `setColumnsHidden()`

### Merges

- `mergeCells({ sheetId, start, end })`
- `unmergeCells({ sheetId, row, col })`

## Estilo, validacao e conditional formatting

### Estilo

```ts
engine.setCellStyle({
  sheetId,
  row: 0,
  col: 0,
  style: {
    fontWeight: "bold",
    backgroundColor: "#FDE68A",
    align: "center"
  }
});
```

### Validacao

```ts
engine.setCellValidation({
  sheetId,
  row: 2,
  col: 1,
  validation: {
    rules: [
      { type: "required" },
      { type: "number", min: 0, max: 100 }
    ]
  }
});
```

Voce tambem pode registrar validadores seguros:

```ts
engine.registerValidator("score-policy", ({ value }) => {
  if (typeof value === "number" && value >= 0 && value <= 100) {
    return undefined;
  }

  return {
    code: "CUSTOM_SCORE_INVALID",
    message: "O score deve ficar entre 0 e 100.",
    area: "core",
    recoverable: true
  };
});
```

### Conditional formatting

```ts
engine.setConditionalFormattingRules(sheetId, [
  {
    id: "high-revenue",
    range: {
      start: { row: 1, col: 1 },
      end: { row: 100, col: 1 }
    },
    priority: 1,
    type: "formula",
    formula: "=B2>1000",
    style: {
      backgroundColor: "#DCFCE7",
      fontWeight: "bold"
    }
  }
]);
```

## Formulas seguras

O projeto nao usa `eval` nem `new Function`.

Capacidades atuais do motor padrao:

- referencias relativas e absolutas
- referencias entre sheets
- ranges
- funcoes como `SUM`, `MIN`, `MAX`, `AVERAGE`, `COUNT`, `COUNTA`, `ABS`, `ROUND`, `IF`, `AND`, `OR`, `NOT`

Exemplo:

```ts
engine.setCellValue({ sheetId, row: 0, col: 0, value: 10 });
engine.setCellValue({ sheetId, row: 1, col: 0, value: "=A1*2" });
```

## Eventos publicos

Use `engine.on(eventName, listener)` para assinar eventos.

Eventos importantes:

- `engine:created`
- `engine:disposed`
- `command:completed`
- `command:failed`
- `cell:updated`
- `selection:changed`
- `security:blocked-input`
- `formula:failed`
- `row-model:changed`

## Plugin engine

Plugins sao registrados por workbook e isolam estado por instancia.

```ts
engine.registerPlugin({
  id: "audit-plugin",
  setup(context) {
    context.on("cell:updated", (payload) => {
      console.log("Celula alterada", payload);
    });

    context.setState({ activations: 1 });
  }
});
```

APIs uteis:

- `registerPlugin()`
- `unregisterPlugin()`
- `enablePlugin()`
- `disablePlugin()`
- `isPluginEnabled()`
- `getRegisteredPlugins()`
- `getPluginState()`

## Serializacao e importacao de snapshots

```ts
const snapshot = engine.toJSON();
const restored = WorkbookEngine.fromJSON(snapshot, new BasicFormulaEngine());
```

## Proximo passo

Leia [03-renderer-e-experiencia.md](./03-renderer-e-experiencia.md) para aprender como essas capacidades aparecem na UI.