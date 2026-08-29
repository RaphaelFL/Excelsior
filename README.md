# Excelsior

Framework de planilha web em TypeScript, com engine de dados, renderização de grid e módulo gráfico avançado.

## Visão Geral

- Núcleo de planilha implementado do zero, inspirado funcionalmente no FortuneSheet.
- Módulo gráfico `@excelsior/charts` inspirado funcionalmente no Plotly.js, com implementação própria (clean-room) em TypeScript Vanilla.
- Arquitetura modular por pacotes, sem acoplamento obrigatório entre planilha e gráficos.
- Runtime principal implementado somente com código próprio, APIs nativas e pacotes internos `@excelsior/*`; o suporte ZIP/XML/OOXML do adapter XLSX também é próprio.
- TypeScript, Vitest, jsdom, Vite e tsup permanecem restritos ao desenvolvimento; o código de produção não contém adapters ou imports de React/Vue.
- Aplicações consumidoras cuidam apenas de telas, integração com backend, payloads e configuração; grid, edição e recursos de planilha permanecem nos pacotes `@excelsior/*`.

O `package-lock.json` registra também essas ferramentas de compilação e teste, portanto nomes externos no lockfile não significam dependência do runtime. Os pacotes principais distribuídos não importam bibliotecas externas nem fazem chamadas de rede. `npm run audit:runtime` verifica essa regra e falha se uma dependência ou import externo for introduzido em `core`, `charts`, `formulas`, `renderer-dom`, `xlsx`, `vanilla`, `styles` ou `devtools`.

## Capacidades Atuais da Planilha

- Workbook com múltiplas sheets, troca de aba e serialização JSON.
- Edição por grid e barra de fórmulas, com histórico de undo/redo.
- Fórmulas seguras com referências entre sheets e suporte a `SUM`, `MIN`, `MAX`, `AVERAGE`, `MEDIAN`, `PRODUCT`, `COUNT`, `COUNTA`, `ABS`, `ROUND`, `ROUNDUP`, `ROUNDDOWN`, `SQRT`, `POWER`, `MOD`, `SIGN`, `IF`, `AND`, `OR` e `NOT`.
- Inserção/exclusão de linhas e colunas, merges e estilos persistidos por célula/linha/coluna.
- Data validation, conditional formatting e find/replace com escopo por sheet ou workbook.
- Painéis internos no renderer para links, divisão de colunas, validação, formatação condicional e localização de células especiais, sem depender de `prompt()` do navegador.
- Notas de célula serializáveis, com indicador visual, edição segura e undo/redo.
- Format painter de uso único, aplicado em batch com undo sem alterar valores ou fórmulas.
- Fill handle/autofill com ajuste de referências relativas e preservação de absolutas.
- Toolbar com família/tamanho/ênfase de fonte, cores, bordas, rotação, zoom da sheet e captura SVG da planilha.
- Clipboard com sanitização de HTML.
- Importação/exportação XLSX (fórmulas, estilos, merges, dimensões, comentários e validações OOXML) e fluxo tabular via `@excelsior/xlsx`; threads e regras não representáveis pelo Excel usam metadata privada versionada como fallback.
- Leitura ZIP/XLSX assíncrona e browser-native com `TextEncoder`, `TextDecoder` e `DecompressionStream`, sem `node:zlib`, `node:path`, `Buffer` ou parser XLSX externo.
- Row models `clientSide`, `viewport`, `infinite` e `serverSide` com datasource serializável.
- Renderer DOM virtualizado, com RTL e baseline de acessibilidade; cabeçalhos mantêm a largura das colunas e acompanham o scroll horizontal, inclusive com frozen/split columns.
- Colaboração opt-in por adapter, com protocolo versionado, fila pré-conexão, replay, presença, aplicação idempotente e funcionamento local independente do transporte.
- Imagens raster seguras como objetos flutuantes serializáveis, com inserção local e operações transacionais no core.

## Escopo da Colaboração

Este repositório fornece **somente a base cliente para colaboração**. O `@excelsior/core` inclui contratos transport-agnostic, `createCollaborationTransportAdapter()`, mensagens de `join`/`leave`, operações, presença, replay inicial e encaminhamento de erros.

**Nenhum backend colaborativo é criado ou distribuído neste projeto.** Não há servidor, WebSocket/SignalR concreto, banco de dados, autenticação, salas ou persistência remota. Uma aplicação futura poderá implementar esses componentes separadamente e conectá-los ao adapter público sem alterar o engine ou o renderer.

A planilha permanece totalmente funcional em modo local quando nenhum adapter é configurado. Consulte [docs/collaboration.md](docs/collaboration.md) para o contrato completo.

## Módulo Gráfico (`@excelsior/charts`)

- Inspirado funcionalmente no Plotly.js para comportamento público e modelo declarativo.
- Implementado do zero (clean-room), sem importar Plotly.js, D3 ou qualquer biblioteca gráfica de runtime.
- API declarativa baseada em `data`, `layout` e `config`.
- Renderização nativa por `svg`, `canvas`, `webgl` e fallback híbrido.
- Cobertura de traces cartesianos, estatísticos, financeiros, domain, polar, ternary, geo e 3D.
- Criação pela spreadsheet de line, bar, area, scatter, pie, histogram, box, violin, heatmap, contour, candlestick, waterfall, funnel, polar, ternary, treemap, sunburst, sankey, surface e scatter3d, com vínculo reativo ao range de origem.
- Recursos avançados: subplots, `xAxis2`/`yAxis2`, drilldown (`sunburst`/`treemap`), overlays, range selector/slider, animação por frames e exportações (`svg`, `png`, `json`, tabela).
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
- `@excelsior/vanilla`: integração JavaScript sem framework externo.

## Scripts Principais

- `npm install`
- `npm run build`
- `npm run typecheck`
- `npm run test`
- `npm run audit:prod`
- `npm run audit:runtime`
- `npm run release:check`

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
- Base de colaboração: [docs/collaboration.md](docs/collaboration.md)
- Módulo gráfico (`@excelsior/charts`): [docs/charts.md](docs/charts.md)
- Exemplos práticos: [docs/examples.md](docs/examples.md)
- Aceite ACP charts+excel web: [docs/charts-acceptance-report.md](docs/charts-acceptance-report.md)
- Especificação ACP do módulo gráfico: [.ia/Instrução ACP para Módulo Gráfico Vanilla TypeScript Inspirado no Plotly.js.md](.ia/Instrução ACP para Módulo Gráfico Vanilla TypeScript Inspirado no Plotly.js.md)
- Reverse engineering: [docs/reverse-engineering/api-matrix.md](docs/reverse-engineering/api-matrix.md)
- Auditoria de dependências: [docs/reports/dependency-audit.md](docs/reports/dependency-audit.md)
- Release checklist: [docs/release-checklist.md](docs/release-checklist.md)