# Excelsior

Framework de planilha web em TypeScript, com engine de dados, renderização de grid e módulo gráfico avançado.

## Visão Geral

- Núcleo de planilha implementado do zero, inspirado funcionalmente no FortuneSheet.
- Módulo gráfico `@excelsior/charts` inspirado funcionalmente no Plotly.js, com implementação própria (clean-room) em TypeScript Vanilla.
- Arquitetura modular por pacotes, sem acoplamento obrigatório entre planilha e gráficos.

## Capacidades Atuais da Planilha

- Workbook com múltiplas sheets, troca de aba e serialização JSON.
- Edição por grid e barra de fórmulas, com histórico de undo/redo.
- Fórmulas seguras com referências entre sheets e suporte a `SUM`, `MIN`, `MAX`, `AVERAGE`, `COUNT`, `COUNTA`, `ABS`, `ROUND`, `IF`, `AND`, `OR`, `NOT`.
- Inserção/exclusão de linhas e colunas, merges e estilos persistidos por célula/linha/coluna.
- Data validation, conditional formatting e find/replace com escopo por sheet ou workbook.
- Fill handle/autofill com ajuste de referências relativas e preservação de absolutas.
- Clipboard com sanitização de HTML.
- Importação/exportação XLSX (fórmulas, estilos, merges, widths e heights) e fluxo tabular via `@excelsior/xlsx`.
- Row models `clientSide`, `viewport`, `infinite` e `serverSide` com datasource serializável.
- Renderer DOM virtualizado, com RTL, controles básicos de formatação e baseline de acessibilidade.

## Módulo Gráfico (`@excelsior/charts`)

- Inspirado funcionalmente no Plotly.js para comportamento público e modelo declarativo.
- Implementado do zero (clean-room), sem importar Plotly.js, D3 ou qualquer biblioteca gráfica de runtime.
- API declarativa baseada em `data`, `layout` e `config`.
- Renderização nativa por `svg`, `canvas`, `webgl` e fallback híbrido.
- Cobertura de traces cartesianos, estatísticos, financeiros, domain, polar, ternary, geo e 3D.
- Recursos avançados: subplots, `xAxis2`/`yAxis2`, drilldown (`sunburst`/`treemap`), overlays, animação por frames e exportações (`svg`, `png`, `json`, tabela).
- Segurança: validação/sanitização centralizadas no `FigureValidator`, sem `eval` e sem `innerHTML` inseguro.
- Performance: `RenderScheduler`, sampling por `maxRenderPoints`, hover espacial opcional e otimizações de linha.

Documentação técnica completa: [`docs/charts.md`](docs/charts.md)

## Pacotes

- `@excelsior/core`: engine, domínio, comandos, histórico, eventos e plugin engine por instância.
- `@excelsior/formulas`: parser e avaliador de fórmulas seguras sem execução dinâmica.
- `@excelsior/devtools`: inspeção leve de eventos e snapshots do workbook.
- `@excelsior/renderer-dom`: renderer DOM virtualizado com merges, estilos e interações de edição.
- `@excelsior/xlsx`: importação/exportação XLSX rica e fluxo tabular com schema.
- `@excelsior/charts`: módulo gráfico declarativo TypeScript Vanilla, inspirado funcionalmente no Plotly.js (clean-room).
- `@excelsior/vanilla`: wrapper JavaScript puro.
- `@excelsior/react`: wrapper React.
- `@excelsior/vue`: wrapper Vue 3.

## Scripts Principais

- `npm install`
- `npm run typecheck`
- `npm run test`
- `npm run audit:prod`
- `npm run release:check`
- `npm run dev`
- `npm run bench`
- `npm run build && npm run start -w @excelsior/collab-demo`

## Atalhos do Renderer DOM

- `Setas`: navegam pela seleção ativa no grid.
- `Enter` ou `F2`: iniciam edição da célula ativa.
- `Ctrl+F` ou `Meta+F`: abrem o painel de find/replace.
- `F3` e `Shift+F3`: navegam entre resultados da busca.
- Nas abas das sheets: `Setas`, `Home` e `End` trocam a aba ativa via teclado.
- Em row models remotos: a toolbar usa a coluna ativa para sort/filter/grouping/aggregation e limpeza local da query.

## Documentação

- Wiki operacional: [wiki/README.md](wiki/README.md)
- API pública: [docs/api.md](docs/api.md)
- Plugin engine: [docs/plugin-engine.md](docs/plugin-engine.md)
- Módulo gráfico (`@excelsior/charts`): [docs/charts.md](docs/charts.md)
- Aceite ACP charts+excel web: [docs/charts-acceptance-report.md](docs/charts-acceptance-report.md)
- Especificação ACP do módulo gráfico: [.ia/Instrução ACP para Módulo Gráfico Vanilla TypeScript Inspirado no Plotly.js.md](.ia/Instrução ACP para Módulo Gráfico Vanilla TypeScript Inspirado no Plotly.js.md)
- Reverse engineering: [docs/reverse-engineering/api-matrix.md](docs/reverse-engineering/api-matrix.md)
- Auditoria de dependências: [docs/reports/dependency-audit.md](docs/reports/dependency-audit.md)
- Release checklist: [docs/release-checklist.md](docs/release-checklist.md)