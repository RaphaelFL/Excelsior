# `@excelsior/charts` - Documentacao Tecnica

Modulo grafico declarativo do Ecossistema Excelsior.

## 1) Objetivo

Fornecer uma engine de visualizacao modular, tipada e independente de framework, com API declarativa e capacidade de operar em diferentes backends de renderizacao.

## 2) Diretriz de implementacao (clean-room)

- Inspiracao funcional no Plotly.js para comportamento publico e experiencia do usuario.
- Implementacao propria, feita do zero em TypeScript Vanilla.
- Sem dependencia externa de runtime para renderizacao.
- Sem uso de CDN, script remoto ou servico externo para desenhar graficos.
- Sem copia de codigo, CSS, testes, assets ou estrutura interna de bibliotecas de terceiros.

## 3) Arquitetura em alto nivel

- **Modelo declarativo**: figura baseada em `data`, `layout` e `config`.
- **Nucleo**: `ChartEngine`, `FigureManager`, `FigureValidator`, `RenderScheduler`.
- **Renderizadores**: `SvgRenderer`, `CanvasRenderer`, `WebglRenderer`, `HybridRenderer`.
- **Interacoes**: hover/tooltip, zoom/pan, selecao, legenda, modebar.
- **Adaptadores**: integracao com dados de planilha sem acoplamento ao core.

## 4) API publica principal

- Criacao: `createFigure(target, figure)`.
- Atualizacao: `update`, `updateData`, `updateLayout`, `resize`.
- Ciclo de vida: `destroy`.
- Exportacao e serializacao: `exportSvg`, `exportPng`, `toJson`, `fromJson`, `exportDataTable`.
- Estado e eventos: `getSelection`, `clearSelection`, `on`, `off`.
- Animacao: `playFrames`, `stopFrames`, `isAnimating`.

## 5) Capacidades implementadas

### 5.1 Renderizacao

- `svg`, `canvas` e `webgl` nativos.
- Fallback hibrido para cenarios sem suporte WebGL adequado.

### 5.2 Traces e visualizacoes

- Cartesianos: `scatter`, `line`, `bar`, `area`.
- Estatisticos: `histogram`, `box`, `violin`, `density`, `distribution`.
- Financeiros: `candlestick`, `ohlc`, `waterfall`, `funnel`.
- Cientificos: `heatmap`, `contour`, `quiver`.
- Domain: `pie`, `donut`, `sunburst`, `treemap`, `sankey`, `parallel-categories`.
- Polar/Ternary/Geo e 3D: `polar`, `ternary`, `geo`, `scatter3d`, `surface`, `mesh3d`.

### 5.3 Layout e eixos

- Subplots com controle de sincronizacao (`syncSubplotZoom`).
- Eixos secundarios `xAxis2` e `yAxis2`.
- `multicategory` com rotulos hierarquicos via delimitadores.
- Overlays declarativos: `shapes`, `annotations`, `images`.

### 5.4 Interacao e UX

- Hover com tooltip, clique e selecao.
- Zoom por roda, zoom por eixo, zoom retangular e pan.
- Toggle de legenda, modebar e fullscreen.
- Range selector e range slider configuráveis para navegação no eixo X.
- Posicionamento defensivo de tooltip para não cobrir a modebar.
- Drilldown interativo em `sunburst` e `treemap`.

### 5.5 Exportacao e interoperabilidade

- Exportacao para `svg` e `png`.
- Serializacao para `json` com parse validado.
- Exportacao tabular por `exportDataTable`.

## 6) Seguranca

- Validacao e sanitizacao centralizadas no `FigureValidator`.
- Sem `eval`, sem `new Function` e sem `innerHTML` para entrada de usuario.
- Importacao JSON com validacao de schema/campos conhecidos.
- Limites defensivos para arrays e profundidade de entrada.
- Tratamento de erro controlado para dados invalidos (incluindo GeoJSON).

## 7) Performance

- `RenderScheduler` com consolidacao de updates por frame.
- Sampling por `maxRenderPoints`.
- Hover espacial opcional por `spatialHover`.
- Simplificacao e clipping de linhas.
- Benchmark em `packages/charts/bench/benchmark.mjs`.

### Snapshot de medicao (2026-08-12)

- Build (`tsup`): `dist/index.js` 475.67 KB, `dist/index.cjs` 478.24 KB, `dist/index.d.ts` 36.46 KB.
- Pacote (`npm pack --dry-run`): tarball 191.9 KB, tamanho desempacotado 1.1 MB, 6 arquivos.
- Benchmark:
  - 10k pontos: normalize 7.94 ms, parse JSON 0.64 ms
  - 100k pontos: normalize 52.52 ms, parse JSON 6.56 ms
  - 1M pontos: normalize 470.28 ms, parse JSON 64.88 ms

## 8) Acessibilidade

- `ariaDescription` no container do grafico.
- Controles da modebar com labels acessiveis.
- Tabela alternativa opcional via `accessibleTable`.

## 9) Limitacoes conhecidas

- Projecao geo e 3D simplificada para foco local/offline.
- Pipeline WebGL de `surface/mesh3d` leve (sem iluminacao fisica avancada).
- Em `parallel-categories`, filtro visual por match exato da linha selecionada.
- Em `multicategory`, parser textual por delimitadores (sem estrutura multi-nivel nativa por arrays aninhados).
- Em `sankey`, ciclos sao tratados como entrada invalida.
- Todos os tipos atualmente expostos na toolbar da spreadsheet possuem implementação habilitada.
- Exportacao XLSX usa fallback de placeholder visual quando o grafico nao possui range/série valida para exportacao estruturada.
- Importacao XLSX preserva metadados principais (titulo, legenda, eixos, range e ancora), mas recursos proprietarios avancados de Office continuam em `unsupportedFeatures`.
- A visualizacao embutida na sheet usa pausa de render para graficos fora da viewport; nesses casos e exibido placeholder leve ate o grafico voltar a area visivel.
- Em XLSX, `xAxis.type` com escala numerica/log em graficos categoriais pode ser degradado para categoria; para escala numerica completa, priorizar `scatter`.

## 10) Validacao e operacao

- Typecheck: `npm run typecheck`.
- Testes de charts: `npx vitest run packages/charts/tests`.
- Testes de integracao chart+sheet+xlsx:
  - `npx vitest run packages/renderer-dom/tests/dom-spreadsheet-renderer.spec.ts`
  - `npx vitest run packages/core/tests/workbook-engine.spec.ts packages/xlsx/tests/xlsx.spec.ts`

## 10.1) Matriz de testes de integracao (ACP)

- Fluxo funcional: criacao por UI/codigo, mover, redimensionar, excluir, editar tipo/range/titulo/legenda/eixos.
- Seguranca: sanitizacao de titulos/eixos/series, bloqueio de range excessivo, fallback seguro para payloads nao exportaveis.
- Performance (sheet): cobertura automatizada para `1/5/10/25/50` graficos simultaneos.
- Performance (range): cobertura automatizada para selecoes de `1k/10k/100k` celulas.
- Regressao: planilha sem charts e import/export sem charts mantidos estaveis.
- Ciclo de vida: `dispose()` validado com limpeza de runtimes e caches internos de charts.

## 11) Rollback operacional

1. Forcar `config.renderer = "svg"` em caso de regressao.
2. Desabilitar recursos opcionais por config (`spatialHover`, `fullscreen`, `accessibleTable`).
3. Reverter alteracoes de `packages/charts` sem impactar o core de planilha.

## 12) Aceite ACP (charts+excel web)

- Relatorio consolidado de aceite, matriz de cobertura e snapshot de performance:
  - `docs/charts-acceptance-report.md`
