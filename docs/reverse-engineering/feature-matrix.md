---
id: "fortune-sheet-feature-matrix"
title: "Matriz de Features"
type: "workflow"
status: "draft"
owner: "Raphael Lopes"
source_of_truth: true
related:
  - "fortune-sheet-map.md"
---

# Matriz de Features

| Feature | FortuneSheet | MVP Excelsior | Status |
|---|---|---|---|
| Workbook com múltiplas sheets | Sim | Sim | Implementado |
| Edição de célula | Sim | Sim | Implementado |
| Seleção de range | Sim | Sim | Implementado |
| `Op[]` serializável | Sim | Sim | Implementado |
| Undo/redo | Sim | Sim | Implementado |
| Renderização virtualizada | Sim | Sim | Implementado |
| Wrapper vanilla | Parcial no ecossistema | Sim | Implementado |
| Wrapper React | Sim | Sim | Implementado |
| Wrapper Vue 3 | Sim | Sim | Implementado |
| Fórmulas básicas | Sim | Sim | Implementado no pacote formulas |
| Clipboard seguro | Sim | Sim | Implementado |
| Inserção/remoção de linhas e colunas | Sim | Sim | Implementado |
| Abas com adicionar/remover sheet | Sim | Sim | Implementado |
| Barra de fórmulas editável | Sim | Sim | Implementado |
| Serialização e reidratação JSON | Sim | Sim | Implementado |
| Import/export XLSX | Plugin | Sim | Implementado em `@excelsior/xlsx` |
| Colaboração backend | Sim | Não | Removido do repositório; a biblioteca continua aceitando `Op[]` remotos para integração externa |
| Aplicação de `Op[]` remotos | Sim | Sim | Implementado |
| Fórmulas entre sheets | Sim | Sim | Implementado |
| Estilos persistidos por célula | Sim | Sim | Implementado no core e renderer |
| Merge de células | Sim | Sim | Implementado no core, renderer e XLSX |
| Largura de coluna e altura de linha | Sim | Sim | Implementado no core, renderer e XLSX |
| Import/export XLSX com roundtrip de fórmulas e estilos | Plugin | Sim | Implementado em `@excelsior/xlsx` |
| Import/export tabular com schema de colunas | Plugin | Sim | Implementado em `@excelsior/xlsx` |