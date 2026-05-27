---
id: "acp-report-core-mvp"
title: "Relatório ACP - Core MVP"
type: "ops"
status: "draft"
owner: "Raphael Lopes"
source_of_truth: true
related:
  - "../reverse-engineering/fortune-sheet-map.md"
---

# Relatório ACP - Core MVP

## Fonte analisada no FortuneSheet

- Arquivo/pasta: README, `package.json`, `packages/core/package.json`, `packages/core/src/index.ts`, `packages/react/package.json`, `docs/guide/op.md`
- APIs públicas relacionadas: `Workbook`, `onOp`, `data`, exports do core
- Tipos/modelos relacionados: workbook, sheet, cell, `Op`
- Dependências usadas: parser de fórmulas, utilitários, datas, IDs

## Comportamento observado

- Entrada: dados iniciais em formato de workbook com sheets
- Processamento: mutações geram operações serializáveis e atualizam estado interno
- Saída: UI renderizada por wrappers e callbacks com operações
- Eventos/operações: `Op[]` com `op`, `id`, `path`, `value`
- Edge cases: múltiplas instâncias, fórmulas, clipboard e histórico

## Problemas do desenho original

- Performance: superfície grande e acoplada a muitos módulos
- Acoplamento: exports amplos do core e dependências pesadas para MVP
- Segurança: clipboard e fórmulas exigem hardening explícito
- Manutenção: compatibilidade histórica amplia a área de regressão

## Decisão para este projeto

- Pacote alvo: `packages/core`, `packages/formulas`
- Tipos novos: `WorkbookModel`, `SheetModel`, `CellModel`, `SpreadsheetOperation`
- Comandos novos: `SetCellValueCommand`, `SelectRangeCommand`
- Eventos novos: `engine:created`, `command:completed`, `selection:changed`, `cell:updated`
- Testes obrigatórios: criação de workbook, edição, undo/redo, fórmulas básicas, ciclo
- Dependências necessárias: nenhuma de produção no MVP

## Critérios de aceite

- [x] Core sem DOM
- [x] Workbook inicial com pelo menos uma sheet
- [x] Alteração de célula via comando tipado
- [x] Histórico simples com undo/redo
- [x] Fórmula sem `eval` e com erro controlado para ciclo