# Renderer e Experiencia de Uso

## Quando usar o renderer DOM

Use `@excelsior/renderer-dom` ou o wrapper `@excelsior/vanilla` quando voce quer a experiencia completa de planilha no navegador.

O wrapper vanilla aceita:

- `data`, `settings`, `metadata`
- `onChange`
- `cellRenderers`
- `cellEditors`
- `includeHiddenCellsInClipboard`
- `autofill`
- `localization`
- `renderDebounceMs`
- `rowModel`

## Interacoes basicas da UI

- edicao pelo grid
- edicao pela formula bar
- troca de abas
- adicao de sheets pela UI
- selecao com teclado
- merge renderizado com viewport virtualizada

## Find and Replace

O renderer oferece painel de busca e substituicao com:

- escopo por sheet ou workbook
- regex opcional
- navegacao entre resultados
- replace da ocorrencia atual
- replace-all com historico undoable

Atalhos padrao:

- `Ctrl+F` ou `Meta+F`
- `F3`
- `Shift+F3`

## Autofill e fill handle

O renderer suporta:

- sequencia numerica
- copia de formulas com ajuste de referencias relativas
- preservacao de referencias absolutas
- copia opcional de estilo
- limite configuravel de celulas preenchidas

Exemplo de configuracao via wrapper vanilla:

```ts
createSpreadsheet(host, {
  data: [{ name: "Sheet1" }],
  autofill: {
    maxCells: 5000,
    copyStyle: true
  }
});
```

## Clipboard seguro

O renderer aplica sanitizacao de HTML e bloqueia payloads hostis ou acima do limite configurado.

Comportamentos importantes:

- respeita `maxPasteCells`
- respeita linhas e colunas ocultas por padrao
- trabalha com politica de clipboard `text-only`

## Data validation na UI

O renderer integra validacao do core com:

- feedback visual ao falhar um commit
- dropdown inline para listas
- affordance de checkbox
- bloqueio de formula bar quando o valor e invalido
- validacao em lote antes de um paste

## Conditional formatting na UI

As regras serializadas no core sao refletidas visualmente sem mutar o estilo persistido da celula.

Isso inclui:

- regras por formula segura
- prioridades deterministicas
- duplicate highlight
- color scale simples

## Custom renderers e custom editors

O renderer DOM permite customizacao segura por estrutura de texto controlada.

Voce pode passar `cellRenderers` e `cellEditors` pelo wrapper vanilla.

Exemplo conceitual:

```ts
createSpreadsheet(host, {
  data: [{ name: "Tasks" }],
  cellRenderers: [
    {
      id: "priority-pill",
      match: ({ cell }) => cell?.value === "high",
      render: () => ({
        parts: [
          { type: "text", value: "Alta prioridade" }
        ]
      })
    }
  ]
});
```

Se um renderer customizado falhar, o renderer DOM faz fallback seguro em vez de quebrar a tela.

## Localizacao, RTL e atalhos

O wrapper aceita `localization` com:

- `direction`
- `messages`
- `shortcuts`
- `formatters`

Exemplo:

```ts
createSpreadsheet(host, {
  data: [{ name: "Sheet1" }],
  localization: {
    direction: "rtl",
    messages: {
      findReplace: "Localizar"
    },
    shortcuts: {
      openFindReplace: ["Alt+F"]
    }
  }
});
```

## Debounce de render

Se houver bursts de eventos, voce pode coalescer renders com `renderDebounceMs`.

```ts
createSpreadsheet(host, {
  data: [{ name: "Sheet1" }],
  renderDebounceMs: 8
});
```

## IME e acessibilidade

O renderer ja cobre:

- bloqueio de commit durante composicao IME
- `role="grid"` e `role="gridcell"`
- `aria-activedescendant`
- announcer de status
- abas com navegacao por `Setas`, `Home` e `End`
- selecao visual que nao depende apenas de cor

## Frozen panes e ocultacao

Na UI final, `freezeRows()`, `freezeColumns()`, `setRowsHidden()` e `setColumnsHidden()` se refletem no scroll, na navegacao e no clipboard.

## Pivot panel na UI

O renderer expõe `create-pivot` na toolbar.

Fluxo tipico:

1. selecione um range tabular com cabecalho ou use a area utilizada da sheet
2. clique em `create-pivot`
3. ajuste `rows`, `columns`, `values`, alias e totais
4. escolha o modo `auto`, `client` ou `server`
5. defina se a pivot derivada deve usar `autoRefresh`
6. clique em aplicar

Capacidades atuais do painel:

- multiplas dimensoes de linha
- multiplas medidas
- aliases por medida
- totais e subtotais
- progresso e cancelamento
- mensagem amigavel para fonte grande demais no client-side
- reabertura da configuracao para pivots persistidas
- update in-place da pivot derivada
- status de stale quando `autoRefresh` esta desligado

## Toolbar de dados remotos

Quando a sheet usa `InfiniteRowModel` ou `ServerSideRowModel`, a toolbar da coluna ativa permite:

- sort asc/desc
- filtro `equals`
- grouping
- expand/collapse de grupos
- pivot remoto por coluna
- agregacoes `sum`, `avg`, `min`, `max`, `count`

## Proximo passo

Leia [04-dados-remotos-pivot-e-integracoes.md](./04-dados-remotos-pivot-e-integracoes.md) para ver como ligar a UI a dados remotos, pivot server-side, XLSX e ferramentas auxiliares.