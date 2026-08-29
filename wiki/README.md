# Wiki do Excelsior

Esta wiki descreve como usar a superficie publica atual do monorepo Excelsior. O foco aqui e operacional: o que existe, quando usar cada pacote e qual API chamar em cada cenario.

## Estrutura

- [01-primeiros-passos.md](./01-primeiros-passos.md)
  - instalacao, scripts, escolha de pacote e exemplos minimos em core e vanilla
- [02-core-e-engine.md](./02-core-e-engine.md)
  - WorkbookEngine, formulas, historico, batch updates, validacao, conditional formatting, plugins, serializacao e eventos
- [03-renderer-e-experiencia.md](./03-renderer-e-experiencia.md)
  - renderer DOM, barra de formulas, pivot panel, find/replace, autofill, clipboard, localizacao, acessibilidade e customizacao visual
- [04-dados-remotos-pivot-e-integracoes.md](./04-dados-remotos-pivot-e-integracoes.md)
  - row models, datasource remoto, pivot client/server, derived views, XLSX, devtools, benchmark e demo colaborativa
- [05-seguranca-performance-e-operacao.md](./05-seguranca-performance-e-operacao.md)
  - limites de seguranca, erros tipados, performance, observabilidade e checklists de operacao

## Pacotes e quando usar

- `@excelsior/core`
  - use quando voce precisa da engine, do modelo de workbook, dos comandos, da serializacao, dos row models e das APIs de pivot sem UI.
- `@excelsior/renderer-dom`
  - use quando voce quer a experiencia completa de planilha no DOM, incluindo toolbar, formula bar, viewport virtualizada, pivot panel e recursos de acessibilidade.
- `@excelsior/vanilla`
  - use quando voce quer montar a planilha diretamente em um elemento HTML, inclusive dentro de componentes de qualquer framework.
- `@excelsior/formulas`
  - use quando voce quer o motor padrao de formulas seguras.
- `@excelsior/xlsx`
  - use para importacao/exportacao de workbooks completos ou tabelas com schema.
- `@excelsior/devtools`
  - use para logar eventos publicos e capturar snapshots da engine.

## Trilha recomendada

1. Comece em [01-primeiros-passos.md](./01-primeiros-passos.md) para subir a stack local e decidir entre core puro e vanilla.
2. Leia [02-core-e-engine.md](./02-core-e-engine.md) para entender o contrato do workbook e as APIs publicas principais.
3. Leia [03-renderer-e-experiencia.md](./03-renderer-e-experiencia.md) se voce vai expor a UI de planilha ao usuario final.
4. Leia [04-dados-remotos-pivot-e-integracoes.md](./04-dados-remotos-pivot-e-integracoes.md) se voce vai usar grandes volumes, pivot, XLSX, devtools ou benchmark.
5. Feche em [05-seguranca-performance-e-operacao.md](./05-seguranca-performance-e-operacao.md) para limites, erros e operacao segura em producao.

## Referencias complementares

- API publica resumida existente: [../docs/api.md](../docs/api.md)
- Plugin engine: [../docs/plugin-engine.md](../docs/plugin-engine.md)
- Aceite ACP charts+excel web: [../docs/charts-acceptance-report.md](../docs/charts-acceptance-report.md)
- Checklist de release: [../docs/release-checklist.md](../docs/release-checklist.md)