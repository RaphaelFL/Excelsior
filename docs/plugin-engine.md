---
id: "excelsior-plugin-engine"
title: "Plugin Engine"
type: "architecture"
status: "draft"
owner: "Raphael Lopes"
source_of_truth: true
related:
  - "./api.md"
  - "../.ia/future-improvements-ag-grid-handsontable.md"
---

# Plugin Engine

## Objetivo

Adicionar extensibilidade por workbook sem acoplar o core a framework, DOM global ou estado mutável externo.

## Alinhamento com a arquitetura atual

- usa o `WorkbookEngine` como ponto de entrada, sem criar uma segunda camada de estado;
- reaproveita o barramento de eventos já existente;
- força mutações por comandos públicos da engine, em vez de permitir acesso interno ao `CommandBus` ou ao snapshot mutável;
- mantém estado isolado por plugin e por instância de workbook;
- permite ativar, desativar e remover plugins sem deixar listeners pendurados.

## Contrato público

```ts
export interface GridPlugin {
  id: string;
  setup(context: PluginContext): void | PluginDisposer;
}
```

## `PluginContext`

O contexto entregue ao plugin contém apenas:

- `pluginId` e `workbookId`;
- `getSnapshot()` com cópia serializável do workbook;
- `getState()` e `setState()` para estado isolado do plugin;
- `on()` para assinar eventos públicos removíveis;
- `commands` com a superfície pública permitida do `WorkbookEngine`.

## Limites intencionais

- plugin não recebe referência ao estado interno mutável;
- plugin não recebe acesso direto ao DOM;
- plugin não registra dependências de framework no core;
- falha de listener não deve interromper o fluxo principal de edição.

## Próximos passos compatíveis com este desenho

- validadores como plugins opcionais;
- formatação condicional como plugin serializável;
- renderizadores e editores customizados adaptados pelos wrappers, não pelo core.