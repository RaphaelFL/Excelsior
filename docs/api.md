---
id: "excelsior-public-api"
title: "API Pública"
type: "ops"
status: "draft"
owner: "Raphael Lopes"
source_of_truth: true
related:
  - "reverse-engineering/api-matrix.md"
  - "../README.md"
---

# API Pública

## Visão geral

O contrato público do Excelsior é centrado em um core sem framework e um wrapper fino para browser puro. O ponto de entrada principal do domínio é `WorkbookEngine` no pacote `@excelsior/core`.

## Pacotes públicos

| Pacote | Responsabilidade | Entrada principal |
|---|---|---|
| `@excelsior/core` | engine, estado, comandos, eventos e tipos | `WorkbookEngine` |
| `@excelsior/formulas` | avaliação segura de fórmulas | `BasicFormulaEngine` |
| `@excelsior/renderer-dom` | renderer DOM e utilitários de clipboard/viewport | `DomSpreadsheetRenderer` |
| `@excelsior/vanilla` | inicialização sem framework | `createSpreadsheet()` |
| `@excelsior/xlsx` | import/export de workbook e fluxo tabular | `exportWorkbookToXlsx()`, `importWorkbookFromXlsx()` |
| `@excelsior/devtools` | inspeção leve de eventos do engine | `attachWorkbookDevtools()` |

O `BasicFormulaEngine` inclui agregações, lógica e funções numéricas comuns: `SUM`, `MIN`, `MAX`, `AVERAGE`, `MEDIAN`, `PRODUCT`, `COUNT`, `COUNTA`, `ABS`, `ROUND`, `ROUNDUP`, `ROUNDDOWN`, `SQRT`, `POWER`, `MOD`, `SIGN`, `IF`, `AND`, `OR` e `NOT`.

O adapter XLSX grava comentários legados e validações compatíveis em partes OOXML nativas. Threads, respostas e regras sem representação equivalente permanecem na metadata privada versionada para roundtrip completo no Excelsior.

`createCollaborationTransportAdapter(transport)` fornece a ponte oficial entre o engine e um transporte futuro. O contrato público inclui mensagens versionadas de `join`, `leave`, operações, presença, replay inicial e erro, sem acoplar o core a WebSocket ou backend específico. Consulte [collaboration.md](collaboration.md).

## `@excelsior/core`

### `WorkbookEngine`

Cria, consulta e altera um workbook serializável.

```ts
import { WorkbookEngine } from "@excelsior/core";
import { BasicFormulaEngine } from "@excelsior/formulas";

const engine = new WorkbookEngine(
  {
    data: [{ name: "Sheet1" }]
  },
  new BasicFormulaEngine()
);
```

### Construtor

```ts
new WorkbookEngine(config?: WorkbookConfig, formulaEngine?: FormulaEngine)
```

Quando a `FormulaEngine` expõe `collectReferences(expression)`, o core passa a recalcular apenas fórmulas impactadas por `setCellValue()`, `updateCells()` e `applyOperations()` baseadas em células; sem essa capability, o fallback continua sendo o recálculo completo do workbook.

No core, leituras quentes como `getDisplayValue()` usam cache por revisão interna do workbook, e getters granulares (`getCell()`, `getMerge()`, `getColumnSchema()`, `getRowSchema()`) passam a clonar apenas o fragmento retornado em vez de clonar o workbook inteiro a cada chamada.

### Métodos principais

| Método | Descrição | Retorno |
|---|---|---|
| `on(eventName, listener)` | assina eventos públicos tipados | função de unsubscribe |
| `getSnapshot()` | retorna snapshot serializável do workbook | `WorkbookModel` |
| `getActiveSheet()` | retorna a sheet ativa | `SheetModel` |
| `getCell(sheetId, row, col)` | lê a célula pelo endereço | `CellModel | undefined` |
| `getDisplayValue(sheetId, row, col)` | resolve o valor exibível de uma célula | `string` |
| `getMerge(sheetId, row, col)` | consulta merge que cobre a célula | `SheetMerge | undefined` |
| `getColumnSchema(sheetId, col)` | lê schema de coluna | `ColumnSchema | undefined` |
| `getRowSchema(sheetId, row)` | lê schema de linha | `RowSchema | undefined` |
| `registerPlugin(plugin, enabled?)` | registra plugin por id único nesta instância | `void` |
| `unregisterPlugin(pluginId)` | remove plugin e seu estado isolado | `void` |
| `enablePlugin(pluginId)` | ativa um plugin previamente registrado | `void` |
| `disablePlugin(pluginId)` | desativa um plugin sem removê-lo do registro | `void` |
| `isPluginEnabled(pluginId)` | informa se o plugin está ativo nesta instância | `boolean` |
| `getRegisteredPlugins()` | lista plugins registrados e seu estado de ativação | `RegisteredGridPlugin[]` |
| `getPluginState(pluginId)` | lê o estado isolado do plugin | `PluginState | undefined` |
| `setCellValidation(input)` | grava regras serializáveis de validação na célula | `SpreadsheetOperation[]` |
| `getCellValidation(sheetId, row, col)` | lê as regras de validação de uma célula | `CellValidationConfig | undefined` |
| `setCellNote(input)` | cria, edita ou remove uma nota textual serializável com histórico | `SpreadsheetOperation[]` |
| `getCellNote(sheetId, row, col)` | lê a nota textual de uma célula | `string | undefined` |
| `validateCellValue(input)` | avalia um valor sem mutar o workbook | `CellValidationResult` |
| `registerValidator(id, validator)` | registra um validador customizado seguro | `void` |
| `unregisterValidator(id)` | remove um validador customizado | `void` |
| `getRegisteredValidators()` | lista validadores customizados registrados | `RegisteredCellValidator[]` |
| `setConditionalFormattingRules(sheetId, rules)` | substitui as regras serializáveis de conditional formatting da sheet | `SpreadsheetOperation[]` |
| `getConditionalFormattingRules(sheetId)` | lê as regras de conditional formatting da sheet | `ConditionalFormattingRule[]` |
| `getConditionalStyle(sheetId, row, col)` | resolve o estilo efetivo derivado das regras condicionais | `CellStyle | undefined` |
| `freezeRows(sheetId, count)` | congela as primeiras linhas da sheet | `SpreadsheetOperation[]` |
| `freezeColumns(sheetId, count)` | congela as primeiras colunas da sheet | `SpreadsheetOperation[]` |
| `setRowsHidden(sheetId, start, end?, hidden?)` | oculta ou reexibe linhas por índice ou range | `SpreadsheetOperation[]` |
| `setColumnsHidden(sheetId, start, end?, hidden?)` | oculta ou reexibe colunas por índice ou range | `SpreadsheetOperation[]` |
| `getFrozenPane(sheetId)` | lê a configuração atual de panes congeladas | `{ rows: number; columns: number }` |
| `setCellValue(input)` | altera valor ou fórmula de célula; writes redundantes retornam `[]` e não criam histórico | `SpreadsheetOperation[]` |
| `updateCells(input)` | aplica upsert em lote de células, ignora no-ops redundantes e usa recálculo seletivo de fórmulas quando disponível | `SpreadsheetOperation[]` |
| `applyCellTransaction(input)` | aplica transação keyed por célula (`row:col`) com `upsert`/`remove` em uma única entrada de histórico | `SpreadsheetOperation[]` |
| `setCellStyle(input)` | aplica estilo à célula âncora | `SpreadsheetOperation[]` |
| `resizeColumn(sheetId, col, width)` | altera largura da coluna | `SpreadsheetOperation[]` |
| `resizeRow(sheetId, row, height)` | altera altura da linha | `SpreadsheetOperation[]` |
| `mergeCells(input)` | cria merge em um range | `SpreadsheetOperation[]` |
| `unmergeCells(input)` | remove merge que cobre a célula | `SpreadsheetOperation[]` |
| `selectRange(input)` | altera a seleção ativa | `void` |
| `undo()` | desfaz a última alteração | `boolean` |
| `redo()` | refaz a última alteração | `boolean` |
| `setActiveSheet(sheetId)` | troca a aba ativa | `void` |
| `reportSecurityEvent(reason, details?)` | registra bloqueio relacionado à segurança | `void` |
| `toJSON()` | serializa o workbook atual | `WorkbookModel` |
| `loadFromJSON(snapshot)` | reidrata um workbook serializado | `void` |
| `applyOperations(operations)` | aplica `Op[]` serializáveis e recalcula dependências | `void` |
| `applyBatchOperations(input)` | aplica batch com histórico usando `UpdateSheetOperationsCommand` | `SpreadsheetOperation[]` |
| `getSelection(sheetId)` | retorna o range selecionado na sheet | `CellRange` |
| `inferPivotSheet(input)` | infere uma configuração inicial de pivot client-side a partir de um range tabular com cabeçalho | `PivotSheetInput` |
| `createPivotSheet(input)` | gera uma sheet derivada client-side a partir de uma sheet tabular com cabeçalho na primeira linha do range | `WorkbookDataInput` |
| `createPivotSheetAsync(input, options?)` | gera a pivot por etapas assíncronas, cedendo o controle entre blocos para evitar travar a UI; `options` aceita `chunkSize`, `yieldControl`, `signal` e `onProgress` | `Promise<WorkbookDataInput>` |
| `addPivotSheet(input)` | materializa a pivot derivada como uma nova sheet no workbook | `string` |
| `addPivotSheetAsync(input, options?)` | materializa a pivot assíncrona como uma nova sheet no workbook, também com cancelamento/progresso via `options` | `Promise<string>` |
| `setRowModel(sheetId, rowModel)` | ativa explicitamente um row model para a sheet | `void` |
| `clearRowModel(sheetId)` | remove o row model explícito e volta ao client-side padrão | `void` |
| `getRowModel(sheetId)` | retorna o row model efetivo da sheet | `RowModel` |
| `updateRemoteRowModel(sheetId, update)` | atualiza `sortModel`, `filterModel`, `groupKeys`, `expandedGroupPaths`, `pivotModel` e `aggregateModel` de um row model remoto sem recriar a instância | `void` |
| `setRemoteGroupExpanded(sheetId, path, expanded)` | expande ou recolhe um caminho de grupo remoto preservando o restante do estado | `void` |
| `getRemoteRowModelRequest(sheetId)` | lê o request model remoto atual de um row model remoto explícito | `RemoteRowModelUpdate | undefined` |
| `addSheet(input?)` | adiciona nova sheet | `string` |
| `deleteSheet(sheetId)` | remove uma sheet | `void` |
| `insertRows(sheetId, index, count?)` | insere linhas | `void` |
| `deleteRows(sheetId, start, end?)` | remove linhas | `void` |
| `insertColumns(sheetId, index, count?)` | insere colunas | `void` |
| `deleteColumns(sheetId, start, end?)` | remove colunas | `void` |
| `dispose()` | limpa listeners e encerra a engine | `void` |

No renderer DOM e nos wrappers públicos, `renderDebounceMs` permite coalescer bursts de `selection:changed`, `command:completed`, scroll e resize em um único render agendado; omitido ou `0`, o comportamento continua imediato.

### Eventos públicos

`SpreadsheetEventMap` expõe os eventos públicos usados por renderer, wrappers e observabilidade.

| Evento | Quando ocorre |
|---|---|
| `engine:created` | após a inicialização do workbook |
| `engine:disposed` | ao descartar a engine |
| `command:completed` | após um comando bem-sucedido |
| `command:failed` | quando um comando falha |
| `cell:updated` | após alteração de célula |
| `selection:changed` | após mudança de seleção |
| `security:blocked-input` | quando conteúdo inseguro é bloqueado |
| `formula:failed` | quando uma fórmula retorna erro controlado |

### Plugin engine

O core expõe um sistema de plugins por instância de workbook.

```ts
import { WorkbookEngine, type GridPlugin } from "@excelsior/core";

const engine = new WorkbookEngine();

const plugin: GridPlugin = {
  id: "selection-audit",
  setup(context) {
    context.setState({ selections: 0 });
    return context.on("selection:changed", () => {
      context.setState<{ selections: number }>((previous) => ({
        selections: (previous?.selections ?? 0) + 1
      }));
    });
  }
};

engine.registerPlugin(plugin);
```

O contrato detalhado está em [docs/plugin-engine.md](docs/plugin-engine.md).

### Row Model Strategy

O core agora expõe uma base única de row models para permitir troca explícita da estratégia por sheet sem mudar a API consumida pelo renderer.

```ts
import { ViewportRowModel, WorkbookEngine } from "@excelsior/core";

const engine = new WorkbookEngine();
const sheet = engine.getActiveSheet();

engine.setRowModel(
  sheet.id,
  new ViewportRowModel({
    rowCount: sheet.rowCount,
    getRows: ({ startRow, endRow }) => ({
      rowCount: sheet.rowCount,
      rows: Array.from({ length: endRow - startRow + 1 }, (_value, index) => ({
        index: startRow + index
      }))
    })
  })
);
```

Tipos e implementações públicas desta etapa:

- `RowModel`: contrato base com `getRowCount()`, `getRows()`, `refresh()` e `dispose()`.
- `ClientSideRowModel`: estratégia padrão baseada no workbook local.
- `ViewportRowModel`: adaptador para fornecer a janela visível por callback.
- `InfiniteRowModel`: busca dados por blocos reutilizando cache já resolvido.
- `ServerSideRowModel`: delega cada janela ao datasource sem carregar tudo no browser.
- `RowModelRow`: cada linha resolvida pode carregar metadado opcional de grupo em `group`, reaproveitando `DataResponse.groupInfo` quando o datasource remoto informar esse contexto, incluindo `path` completo para expand/collapse de grupos aninhados.
- `updateRemoteRowModel(sheetId, update)`: aplica atualização incremental de `sortModel`, `filterModel`, `groupKeys`, `expandedGroupPaths`, `pivotModel` e `aggregateModel` sem trocar a instância do modelo remoto, e dispara invalidação do viewport no renderer.
- `setRemoteGroupExpanded(sheetId, path, expanded)`: helper público para manter o estado de expand/collapse de grupos remotos sem reconstruir manualmente `expandedGroupPaths`.
- `getRemoteRowModelRequest(sheetId)`: retorna o request model atual para a sheet quando ela usa `InfiniteRowModel` ou `ServerSideRowModel`.

### Pivot client-side

O core agora consegue derivar uma sheet de pivot client-side a partir de uma sheet-fonte tabular com cabeçalho na primeira linha do range. O `WorkbookEngine` usa um `pivotModule` configurável: por padrão ele aponta para a implementação interna atual, mas pode ser substituído por um módulo customizado ou desabilitado com `pivotModule: false` no `WorkbookConfig`.

```ts
const pivotInput = engine.inferPivotSheet({
  sourceSheetId: sheet.id,
  sourceRange: {
    start: { row: 0, col: 0 },
    end: { row: 20, col: 4 }
  }
});

const pivot = engine.createPivotSheet(pivotInput);
engine.addSheet(pivot);
```

`inferPivotSheet(input)` usa uma heurística inicial simples e pública: escolhe a última coluna numérica como medida (ou a última coluna do range quando não houver números), usa a primeira dimensão como `rows` e envia as demais dimensões para `columns`.

O contrato atual suporta `rows`, `columns`, múltiplos `values`, aliases por medida, agregadores `sum|avg|min|max|count`, totais, subtotais de linha e materialização como nova sheet derivada. O renderer DOM e o wrapper vanilla expõem uma UI na formula bar para configurar essas opções antes de criar a pivot, usando `addPivotSheetAsync()` no fluxo visual para ceder o controle entre blocos, reportar progresso e permitir cancelamento em datasets maiores. Row models remotos podem delegar a materialização ao datasource; uma view virtualizada específica de pivot permanece fora do escopo atual.

### Data Source API

Para fluxos remotos controlados pela própria aplicação, o core expõe um contrato serializável de datasource. Ele é independente de DOM e pode ser reutilizado tanto por `InfiniteRowModel` quanto por `ServerSideRowModel`. O core não faz `fetch`, não importa dados por conta própria e não embute transporte externo.

```ts
import { InfiniteRowModel, WorkbookEngine, type DataSource } from "@excelsior/core";

const rows = Array.from({ length: 500 }, (_value, index) => ({ index }));

const dataSource: DataSource = {
  async getRows(request) {
    const filtered = rows.filter((row) => {
      const statusFilter = request.filterModel?.status;
      if (!statusFilter || statusFilter.operator !== "equals") {
        return true;
      }

      return statusFilter.value === "active" ? row.index % 2 === 0 : row.index % 2 === 1;
    });
    const sorted = request.sortModel?.[0]?.direction === "desc" ? [...filtered].reverse() : filtered;
    return {
      totalRows: sorted.length,
      rows: sorted.slice(request.startRow, request.endRow + 1)
    };
  }
};

const engine = new WorkbookEngine();
const sheet = engine.getActiveSheet();

engine.setRowModel(
  sheet.id,
  new InfiniteRowModel({
    rowCount: "unknown",
    blockSize: 100,
    dataSource,
    visibleColumns: ["sku", "price"],
    sortModel: [{ field: "price", direction: "desc" }]
  })
);
```

Tipos públicos desta etapa:

- `DataSource`: contrato assíncrono com `getRows(request, context?)`.
- `DataRequest`: inclui `startRow`, `endRow`, `sortModel`, `filterModel`, `groupKeys`, `expandedGroupPaths`, `pivotModel`, `aggregateModel`, `visibleColumns` e `requestId`.
- `DataResponse`: retorna `rows`, `totalRows?`, `nextCursor?`, `groupInfo?` e `warnings?`. Quando `groupInfo` é usado, cada entrada pode incluir `key`, `path`, `level`, `childCount` e `expanded`.

O `context` opcional do datasource expõe `signal?: AbortSignal` para permitir cancelamento de requests em voo sem poluir o `DataRequest`, que continua serializável.

Os row models remotos ignoram respostas atrasadas quando uma requisição mais nova já está em curso, e agora também abortam requests em voo quando uma janela mais nova substitui a anterior ou quando o modelo é invalidado.

No `@excelsior/renderer-dom`, requests assíncronos do row model também atualizam a região de status já existente: enquanto a janela remota está pendente a UI anuncia carregamento, e falhas passam a aparecer no mesmo canal com tratamento de erro controlado, sem loop de retry automático no mesmo viewport. Quando a sheet ativa usa `InfiniteRowModel` ou `ServerSideRowModel`, a toolbar também expõe controles mínimos para a coluna ativa, aplicando `updateRemoteRowModel()` internamente sem recriar a instância do modelo.

Para a primeira fatia de 9.19, o mesmo request model remoto também já pode ser atualizado incrementalmente com `groupKeys` e `aggregateModel`, o que permite delegar agrupamento e agregação ao datasource controlado pela aplicação sem substituir a instância do row model. No renderer DOM isso aparece, por enquanto, como toggles mínimos de grouping e agregações `sum`, `avg`, `min`, `max` e `count` para a coluna ativa, ao lado dos controles já existentes de sort/filter.

### Data Validation

O core valida antes de confirmar a edição e também expõe uma API de pré-validação para fluxos como paste em lote.

```ts
engine.registerValidator("prefix", ({ value, params }) => {
  const prefix = String(params?.prefix ?? "");
  return typeof value === "string" && value.startsWith(prefix)
    ? undefined
    : {
        code: "CORE_VALIDATION_CUSTOM_PREFIX",
        message: `O valor deve começar com ${prefix}.`,
        ruleType: "custom",
        validator: "prefix"
      };
});

engine.setCellValidation({
  sheetId,
  row: 0,
  col: 0,
  validation: {
    rules: [
      { type: "required" },
      { type: "custom", validator: "prefix", params: { prefix: "SKU-" } }
    ]
  }
});

const validation = engine.validateCellValue({
  sheetId,
  row: 0,
  col: 0,
  value: "SKU-001"
});
```

Regras suportadas nesta etapa: `text`, `number`, `date`, `boolean`, `list`, `dropdown`, `checkbox`, `required`, `range`, `length`, `regex` e `custom`.

No renderer DOM, células com `list` ou `dropdown` exibem affordance visual e abrem um `select` inline no duplo clique. Células com `checkbox` exibem estado visual e alternam o valor no duplo clique, ainda passando pela mesma validação do core.

### Conditional Formatting

As regras de conditional formatting são serializáveis por sheet e o renderer DOM compõe o estilo resultante sem mutar o `CellModel.style` persistido. A prioridade é determinística: prioridades numéricas menores vencem conflitos no mesmo atributo e empates seguem a ordem declarada.

```ts
engine.setConditionalFormattingRules(sheetId, [
  {
    id: "hot-values",
    type: "greaterThan",
    range: {
      start: { row: 0, col: 0 },
      end: { row: 99, col: 0 }
    },
    value: 100,
    priority: 20,
    style: {
      backgroundColor: "#fee2e2",
      textColor: "#991b1b",
      fontWeight: "bold"
    }
  },
  {
    id: "heatmap",
    type: "colorScale",
    range: {
      start: { row: 0, col: 0 },
      end: { row: 99, col: 0 }
    },
    priority: 100,
    minColor: "#eff6ff",
    maxColor: "#1d4ed8"
  },
  {
    id: "formula-flag",
    type: "formula",
    range: {
      start: { row: 0, col: 1 },
      end: { row: 99, col: 1 }
    },
    formula: "=B1-10",
    priority: 10,
    style: {
      backgroundColor: "#fef3c7"
    }
  }
]);

const style = engine.getConditionalStyle(sheetId, 0, 0);
```

Regras suportadas nesta etapa: `greaterThan`, `lessThan`, `equal`, `notEqual`, `between`, `containsText`, `dateBefore`, `dateAfter`, `duplicates`, `colorScale` e `formula`.

## `@excelsior/vanilla`

### `createSpreadsheet(container, options?)`

Cria uma instância completa para browser puro.

```ts
import { ViewportRowModel } from "@excelsior/core";
import { createSpreadsheet } from "@excelsior/vanilla";

const instance = createSpreadsheet(container, {
  data: [{ name: "Sheet1" }],
  onChange: (operations) => {
    console.log(operations);
  },
  rowModel: new ViewportRowModel({
    rowCount: 200,
    getRows: ({ startRow, endRow }) => ({
      rowCount: 200,
      rows: Array.from({ length: endRow - startRow + 1 }, (_value, index) => ({
        index: startRow + index
      }))
    })
  }),
  localization: {
    direction: "rtl",
    locale: "pt-BR",
    messages: {
      findReplace: "Localizar"
    },
    shortcuts: {
      openFindReplace: ["Alt+F"]
    }
  }
});
```

### Retorno

| Campo | Descrição |
|---|---|
| `engine` | instância de `WorkbookEngine` |
| `renderer` | instância de `DomSpreadsheetRenderer` |
| `destroy()` | limpa renderer e engine |

`createSpreadsheet()` também aceita `rowModel`, aplicado inicialmente à sheet ativa criada pela engine.

## `@excelsior/renderer-dom`

### Principais exports

| Export | Descrição |
|---|---|
| `DomSpreadsheetRenderer` | renderer DOM do grid |
| `getVisibleCellWindow()` | helper de viewport para grids uniformes |
| `parseTabularText()` | parsing de TSV/CSV simples |
| `resolveClipboardText()` | política de clipboard seguro |
| `sanitizeClipboardHtml()` | sanitização de HTML colado |

### Renderers e editors customizados

`DomSpreadsheetRendererOptions` agora aceita `cellRenderers`, `cellEditors`, `includeHiddenCellsInClipboard`, `autofill` e `localization`.

- `cellRenderers` recebem contexto imutável e devem retornar apenas estrutura segura baseada em texto.
- `cellEditors` montam UI efêmera, mas o commit continua passando pelo fluxo normal de validação da engine.
- `includeHiddenCellsInClipboard` controla se copy/paste deve incluir linhas e colunas ocultas; por padrão, o renderer ignora ocultos.
- `autofill` habilita a alça de preenchimento, permite limitar a área máxima preenchida com `maxCells` e copiar estilo de origem com `copyStyle`.
- `localization` permite configurar `locale`, `direction`, mensagens da UI, atalhos do renderer e formatters de exibição para números e datas.
- o renderer DOM agora inclui um painel embutido de `Find/Replace` com escopo por sheet/workbook, navegação, case-sensitive, whole-cell, regex limitada e `replace-all` undoable.
- quando `direction` é `rtl`, a navegação horizontal do teclado é invertida no renderer.
- o renderer adia commits por `Enter` enquanto uma composição IME estiver ativa nos editores de texto e na barra de fórmulas.
- o viewport principal expõe semântica ARIA de `grid`, as células expõem `gridcell` com índices e labels de linha/coluna, e a célula ativa é anunciada por uma região `aria-live` interna.
- as abas de sheets usam `role="tablist"`/`role="tab"`, evitam controles interativos aninhados e podem ser navegadas por teclado com setas, `Home` e `End`.
- quando a engine recebe um `RowModel` explícito para a sheet ativa, o renderer passa a pedir a janela visível via `getRows()` e mantém fallback client-side para sheets sem estratégia customizada.
- quando a sheet ativa usa row model remoto, a toolbar ganha controles de sort asc/desc, filtro por igualdade e limpeza da query para a coluna ativa, todos apoiados em `getRemoteRowModelRequest()` e `updateRemoteRowModel()`.

Atalhos padrão do renderer DOM:

- `ArrowUp`, `ArrowDown`, `ArrowLeft`, `ArrowRight`: movem a seleção ativa no grid, respeitando RTL na horizontal quando configurado.
- `Enter` e `F2`: entram em modo de edição da célula ativa.
- `Ctrl+F` ou `Meta+F`: abrem o painel de `Find/Replace`.
- `F3` e `Shift+F3`: avançam e retrocedem pelos matches da busca.
- nas abas das sheets, `ArrowLeft`, `ArrowRight`, `Home` e `End` ativam navegação por teclado entre abas.

## `@excelsior/xlsx`

### Workbook API

| Função | Descrição |
|---|---|
| `exportWorkbookToXlsx(workbook)` | exporta workbook com fórmulas, merges, estilos e schema |
| `importWorkbookFromXlsx(input)` | importa workbook preservando fórmulas, merges e estilos suportados |

### Tabela orientada a schema

| Função | Descrição |
|---|---|
| `exportTableToXlsx(rows, options)` | exporta objetos a partir de um schema de colunas |
| `importTableFromXlsx(input, options)` | importa objetos com o schema tabular nativo do Excelsior |

## `@excelsior/devtools`

### `attachWorkbookDevtools(engine, options?)`

Anexa listeners de observabilidade a uma instância de `WorkbookEngine`.

```ts
import { attachWorkbookDevtools } from "@excelsior/devtools";

const session = attachWorkbookDevtools(engine, {
  onEvent: (event) => console.debug(event.name, event.payload)
});
```

### Retorno

| Campo | Descrição |
|---|---|
| `events` | buffer em memória com eventos observados |
| `snapshot()` | retorna o snapshot atual da engine |
| `stop()` | remove os listeners de devtools |