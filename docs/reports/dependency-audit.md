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

- Data: 2026-08-28
- `read-excel-file` e `write-excel-file` foram removidos do código, dos manifests e do lockfile.
- O pacote publicável `@excelsior/xlsx` não possui dependências externas nem pacotes embarcados.
- Leitura e gravação XLSX usam OOXML/ZIP próprios, `TextEncoder`, `TextDecoder` e `DecompressionStream`; não usam APIs Node no runtime.
- Os workspaces e peers de React e Vue foram removidos. A integração com aplicações de qualquer framework ocorre por `@excelsior/vanilla`.

## Observações

- O comando cobre dependências de produção instaladas no lockfile atual.
- Em monorepos com peers de workspaces, `--omit=dev` pode incluir transitivas de integrações opcionais; validar também cada artefato publicável.
- O workflow de release deve repetir a auditoria antes de publicação.