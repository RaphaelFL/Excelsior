---
id: "excelsior-dependency-audit"
title: "Auditoria de Dependências"
type: "security"
status: "draft"
owner: "Raphael Lopes"
source_of_truth: true
related:
  - "../api.md"
  - "../../.ia/security.md"
---

# Auditoria de Dependências

## Escopo

Auditoria de dependências de produção do monorepo Excelsior.

## Comando usado

```text
npm audit --omit=dev
```

## Resultado mais recente

- Data: 2026-05-26
- Resultado: `found 0 vulnerabilities`

## Observações

- O comando cobre dependências de produção instaladas no lockfile atual.
- O workflow de release deve repetir a auditoria antes de publicação.