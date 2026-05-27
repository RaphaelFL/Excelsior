---
id: "fortune-sheet-api-matrix"
title: "Matriz de APIs Públicas"
type: "ops"
status: "draft"
owner: "Raphael Lopes"
source_of_truth: true
related:
  - "fortune-sheet-map.md"
---

# Matriz de APIs

| Origem | API observada | Papel no FortuneSheet | Decisão no Excelsior |
|---|---|---|---|
| README | `Workbook` React component | Entrada principal para uso em React | Criar wrapper React fino sobre `createSpreadsheet` |
| README / docs `op.md` | `onOp(ops)` | Emissão de operações serializáveis | Expor `onChange(operations)` no wrapper vanilla e eventos no core |
| `packages/core/src/index.ts` | reexport amplo do core | Superfície grande de APIs internas | Reduzir para `WorkbookEngine`, tipos, comandos e eventos |
| README | `data` inicial com sheets | Contrato de entrada do workbook | Manter `data: [{ name: "Sheet1" }]` como formato inicial |

## APIs alvo do MVP

- `new WorkbookEngine(config)`
- `engine.setCellValue(...)`
- `engine.selectRange(...)`
- `engine.undo()` / `engine.redo()`
- `engine.toJSON()`
- `createSpreadsheet(container, config)`