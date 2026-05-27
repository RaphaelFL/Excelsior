---
id: "fortune-sheet-dependency-map"
title: "Mapa de Dependências"
type: "ops"
status: "draft"
owner: "Raphael Lopes"
source_of_truth: true
related:
  - "fortune-sheet-map.md"
---

# Mapa de Dependências

## Dependências de produção observadas no FortuneSheet core

| Dependência | Uso observado | Tratamento no Excelsior |
|---|---|---|
| `@fortune-sheet/formula-parser` | Fórmulas | Não usada no MVP; substituída por engine própria mínima e segura |
| `dayjs` | Datas | Backlog |
| `immer` | Estado imutável | Não usada no MVP |
| `lodash` | Utilidades | Não usada no MVP |
| `numeral` | Formatação | Backlog |
| `uuid` | IDs | Substituída por `crypto.randomUUID` com fallback interno |

## Dependências adotadas agora

- Produção: nenhuma dependência externa no core do MVP.
- Desenvolvimento: TypeScript, Vite, Vitest, JSDOM, React e Vue para wrappers.

## Racional

- Começar sem dependências de produção reduz risco e simplifica auditoria.
- O desenho permanece compatível com a allowlist dos documentos da `.ia`.
- Dependências externas podem entrar depois por pacote e com revisão explícita.