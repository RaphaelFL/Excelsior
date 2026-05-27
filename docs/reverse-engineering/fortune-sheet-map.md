---
id: "fortune-sheet-map"
title: "Mapa de Engenharia Reversa do FortuneSheet"
type: "architecture"
status: "draft"
owner: "Raphael Lopes"
source_of_truth: true
related:
  - "api-matrix.md"
  - "dependency-map.md"
  - "feature-matrix.md"
---

# Mapa do FortuneSheet

## Fontes analisadas

- README do repositório FortuneSheet no GitHub.
- `package.json` raiz do FortuneSheet.
- `packages/core/package.json`.
- `packages/core/src/index.ts`.
- `packages/react/package.json`.
- `docs/guide/op.md`.

## Estrutura observada

- Monorepo com workspaces em `packages/*`.
- Pastas centrais: `packages`, `docs`, `stories`, `tests`, `backend-demo`.
- Core e wrapper React são publicados como pacotes independentes.
- API operacional baseada em `Op[]` emitido a cada mutação.

## Decisões extraídas para Excelsior

- Preservar monorepo com pacotes isolados por responsabilidade.
- Tratar o core como fonte da verdade, sem DOM e sem framework.
- Expor operações serializáveis inspiradas no formato de `Op`.
- Manter wrappers finos e dependências de framework como peer dependencies.

## Diferenças intencionais

- Escopo inicial menor e mais previsível.
- Fórmulas básicas próprias, sem copiar parser do projeto de referência.
- Histórico por patches de snapshot no MVP, com superfície de API pequena.
- Clipboard convertido para texto seguro antes de chegar ao modelo.