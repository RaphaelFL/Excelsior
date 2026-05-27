---
id: "excelsior-release-checklist"
title: "Checklist de Release"
type: "ops"
status: "draft"
owner: "Raphael Lopes"
source_of_truth: true
related:
  - "./reports/dependency-audit.md"
  - "../README.md"
---

# Checklist de Release

## Gate mínimo

Execute o gate abaixo na raiz do monorepo antes de publicar qualquer pacote ou demo:

```text
npm run release:check
```

## O que o gate valida

- `npm run typecheck`
- `npm run test`
- `npm run build`
- `npm run audit:prod`

## Publicação

1. Confirme que o gate terminou sem falhas.
2. Revise os artefatos gerados em `dist` apenas para os workspaces afetados.
3. Atualize versão e changelog conforme a estratégia de release adotada.
4. Publique somente depois de repetir o gate se houver qualquer alteração posterior.