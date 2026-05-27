# Excelsior

Framework de planilha web em TypeScript puro, inspirado funcionalmente no FortuneSheet e implementado do zero.

## Funcionalidades atuais

- workbook com multiplas sheets e troca de aba
- edicao de celula por grid e barra de formulas
- formulas seguras com referencias entre sheets e funcoes `SUM`, `MIN`, `MAX`, `AVERAGE`, `COUNT`, `COUNTA`, `ABS`, `ROUND`, `IF`, `AND`, `OR`, `NOT`
- undo/redo
- insercao e exclusao de linhas e colunas
- estilos persistidos por celula, linha e coluna, com merges e larguras ou alturas customizadas
- data validation no core com validadores customizados seguros, dropdowns inline, checkbox toggle e feedback visual na UI
- conditional formatting serializavel com prioridade deterministica, formula segura, duplicates e color scale simples
- custom cell renderer/editor seguro no renderer DOM, com fallback e commit validado
- fill handle/autofill com sequencia numerica, ajuste de referencias relativas e preservacao de referencias absolutas
- find and replace com escopo por sheet ou workbook, regex opcional, navegacao entre resultados e replace-all undoable
- frozen rows/columns e hidden rows/columns com navegacao e clipboard respeitando ocultos por padrao
- renderer DOM com viewport virtualizada, merges renderizados, acoes basicas de formatacao, localization configuravel, RTL, protecao de commit durante IME, status de loading/erro para row models assincronos e controles mínimos de sort/filter/grouping/aggregation remotos na toolbar, incluindo `sum`, `avg`, `min`, `max` e `count`
- baseline de acessibilidade no renderer DOM com roles ARIA de grid, célula ativa anunciável, navegação de teclado nas abas e seleção visual que não depende só de cor
- row model strategy com contratos `clientSide`, `viewport`, `infinite` e `serverSide`, ativação explícita por sheet, datasource serializável controlado pela aplicação e atualização incremental de `sortModel`, `filterModel`, `groupKeys` e `aggregateModel` sem recriar a instância remota
- clipboard seguro com sanitizacao de HTML
- serializacao e reidratacao JSON
- aplicacao de `Op[]` remotos no engine com recálculo de formulas
- importacao e exportacao XLSX com formulas, estilos, merges, widths e heights preservados
- importacao e exportacao tabular com schema de colunas via `@excelsior/xlsx`
- devtools leve para inspecao de eventos publicos da engine
- plugin engine por workbook com estado isolado e hooks publicos
- demo backend colaborativo em memoria com SSE
- wrappers vanilla, React e Vue 3
- apps de playground, benchmark e colaboracao

## Pacotes

- `@excelsior/core`: engine, domínio, comandos, histórico, eventos, plugin engine por instância, contratos de row model e datasource controlado pela aplicação.
- `@excelsior/formulas`: parser e avaliador de fórmulas seguras sem execução dinâmica.
- `@excelsior/devtools`: inspeção leve de eventos e snapshots do workbook.
- `@excelsior/renderer-dom`: renderer DOM virtualizado com estilos, merges, localization configuravel, RTL e controles basicos de formatacao.
- `@excelsior/xlsx`: adapter para importacao/exportacao XLSX rica e fluxos tabulares com schema.
- `@excelsior/vanilla`: wrapper JavaScript puro.
- `@excelsior/react`: wrapper React fino.
- `@excelsior/vue`: wrapper Vue 3 fino.

## Scripts

- `npm install`
- `npm run typecheck`
- `npm run test`
- `npm run audit:prod`
- `npm run release:check`
- `npm run dev`
- `npm run bench`
- `npm run build && npm run start -w @excelsior/collab-demo`

## Atalhos do renderer DOM

- `Setas`: navegam pela seleção ativa no grid
- `Enter` ou `F2`: iniciam edição da célula ativa
- `Ctrl+F` ou `Meta+F`: abrem o painel de Find/Replace
- `F3` e `Shift+F3`: navegam entre matches de busca
- nas abas das sheets: `Setas`, `Home` e `End` trocam a aba ativa via teclado
- em row models remotos: a toolbar usa a coluna ativa para aplicar sort asc/desc, filtro por igualdade via `Enter`, toggle de grouping, agregações `sum`/`avg`/`min`/`max`/`count` e limpeza local da query

## Documentacao

- Wiki operacional: [wiki/README.md](wiki/README.md)
- API publica: [docs/api.md](docs/api.md)
- Plugin engine: [docs/plugin-engine.md](docs/plugin-engine.md)
- Reverse engineering: [docs/reverse-engineering/api-matrix.md](docs/reverse-engineering/api-matrix.md)
- Auditoria de dependencias: [docs/reports/dependency-audit.md](docs/reports/dependency-audit.md)
- Release checklist: [docs/release-checklist.md](docs/release-checklist.md)