# Seguranca, Performance e Operacao

## Limites de seguranca e carga

O workbook expoe limites operacionais que devem ser ajustados por ambiente.

Campos principais em `WorkbookSettings`:

- `maxCellLength`
- `maxFormulaLength`
- `maxPasteCells`
- `maxRecalcCells`
- `maxPivotSourceRows`
- `maxRows`
- `maxColumns`
- `maxHistorySize`
- `clipboardPolicy`

Exemplo:

```ts
const engine = new WorkbookEngine({
  settings: {
    maxCellLength: 2000,
    maxFormulaLength: 1024,
    maxPasteCells: 5000,
    maxRecalcCells: 2000,
    maxPivotSourceRows: 3000,
    maxRows: 10000,
    maxColumns: 500,
    rowHeight: 28,
    columnWidth: 120,
    viewportBuffer: 6,
    maxHistorySize: 200,
    enableFormulas: true,
    clipboardPolicy: "text-only"
  }
});
```

## Formulas seguras

Pontos importantes:

- o parser nao usa execucao dinamica
- payloads hostis sao rejeitados como erro de formula
- referencias circulares geram erro controlado
- workloads de recalc acima do limite marcam celulas com `FORMULA_RECALC_LIMIT_EXCEEDED`

## Clipboard e input hostil

O renderer faz sanitizacao e pode emitir eventos de bloqueio de seguranca.

Voce pode observar isso com:

```ts
engine.on("security:blocked-input", (payload) => {
  console.warn("Entrada bloqueada", payload.reason, payload.details);
});
```

## Erros tipados

O projeto usa `SpreadsheetOperationError` para erros operacionais controlados.

Areas atualmente usadas:

- `core`
- `renderer`
- `formula`
- `wrapper`
- `security`
- `pivot`

Exemplos praticos:

- `CORE_SHEET_NOT_FOUND`
- `CORE_PIVOT_CLIENT_ROW_LIMIT_EXCEEDED`
- `FORMULA_PARSE_INVALID`
- `RENDERER_PIVOT_SOURCE_INVALID`
- `XLSX_SHEET_NOT_FOUND`

## Performance quente

Pontos relevantes na implementacao atual:

- `getSnapshot()` foi isolado de caminhos internos quentes por leitura direta controlada no engine
- `getDisplayValue()` usa cache por revisao do workbook
- `renderDebounceMs` coalesce bursts no renderer
- `updateCells()` e `applyCellTransaction()` evitam historico desnecessario em no-op
- pivot async usa chunking cooperativo com `yieldControl`
- row models remotos ignoram respostas atrasadas e abortam requests antigos

## Benchmark e budgets

O app benchmark agora mostra limiares de aceitacao na propria UI:

- budget para mutacao batch vs sequencial
- budget para pivot worker e server comparados ao client

Esse fluxo tambem esta automatizado em CI.

## Checklist de operacao

Antes de publicar uma alteracao grande no monorepo:

```bash
npm run typecheck
npm run test
npm run build
npm run audit:prod
```

Ou use o atalho completo:

```bash
npm run release:check
```

## Dicas praticas

- Para contratos publicos, consulte tambem `tests/compatibility/public-api.spec.ts`.

## Encerramento

Esta wiki foi escrita para a superficie publica atual do repositório. Quando uma feature nova entrar no core, no renderer ou nos wrappers, atualize primeiro o arquivo tematico correspondente aqui e depois o README raiz.