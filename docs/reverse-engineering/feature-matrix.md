---
id: "fortune-sheet-feature-matrix"
title: "Matriz de Features"
type: "workflow"
status: "draft"
owner: "Raphael Lopes"
source_of_truth: true
related:
  - "fortune-sheet-map.md"
---

# Matriz de Features

| Feature | FortuneSheet | MVP Excelsior | Status |
|---|---|---|---|
| Workbook com múltiplas sheets | Sim | Sim | Implementado |
| Edição de célula | Sim | Sim | Implementado |
| Seleção de range | Sim | Sim | Implementado |
| `Op[]` serializável | Sim | Sim | Implementado |
| Undo/redo | Sim | Sim | Implementado |
| Renderização virtualizada | Sim | Sim | Implementado |
| Wrapper vanilla | Parcial no ecossistema | Sim | Implementado |
| Wrapper React | Sim | Sim | Implementado |
| Wrapper Vue 3 | Sim | Sim | Implementado |
| Fórmulas básicas | Sim | Sim | Implementado no pacote formulas |
| Clipboard seguro | Sim | Sim | Implementado |
| Inserção/remoção de linhas e colunas | Sim | Sim | Implementado |
| Abas com adicionar/remover sheet | Sim | Sim | Implementado |
| Barra de fórmulas editável | Sim | Sim | Implementado |
| Serialização e reidratação JSON | Sim | Sim | Implementado |
| Import/export XLSX | Plugin | Sim | Implementado em `@excelsior/xlsx` |
| Colaboração backend | Sim | Não | Removido do repositório; a biblioteca continua aceitando `Op[]` remotos para integração externa |
| Aplicação de `Op[]` remotos | Sim | Sim | Implementado |
| Fórmulas entre sheets | Sim | Sim | Implementado |
| Estilos persistidos por célula | Sim | Sim | Implementado no core e renderer |
| Merge de células | Sim | Sim | Implementado no core, renderer e XLSX |
| Largura de coluna e altura de linha | Sim | Sim | Implementado no core, renderer e XLSX |
| Import/export XLSX com roundtrip de fórmulas e estilos | Plugin | Sim | Implementado em `@excelsior/xlsx` |
| Import/export tabular com schema de colunas | Plugin | Sim | Implementado em `@excelsior/xlsx` |

## Paridade da toolbar

Esta seção compara recursos realmente acessíveis pela interface, não apenas tipos ou APIs disponíveis no core. A referência é a lista padrão `toolbarItems` do FortuneSheet e sua toolbar pública.

| Controle da toolbar FortuneSheet | Excelsior | Situação real |
|---|---|---|
| Undo/redo | Visível | Implementado |
| Pincel de formatação | Visível | Implementado |
| Limpar formatação | Visível | Implementado com substituição completa do estilo |
| Formato monetário | Visível | Comando BRL e formatação localizada implementados |
| Formato percentual | Visível | Comando e formatação localizada implementados |
| Diminuir/aumentar casas decimais | Visível | Implementado na toolbar |
| Seletor de formato geral/número/data | Visível | Seletor inclui geral, número, moeda, percentual e data |
| Família e tamanho da fonte | Visível | Implementado |
| Negrito/itálico/sublinhado | Visível | Implementado |
| Tachado | Visível | Implementado no estilo comum e na toolbar |
| Cor da fonte e preenchimento | Visível | Implementado |
| Bordas | Visível | Cor, todas as bordas, lados individuais e remoção disponíveis; estilo padrão fino |
| Mesclar/desmesclar | Visível | Inclui mesclagem total, horizontal, vertical e desmesclagem |
| Alinhamento horizontal | Visível | Implementado |
| Alinhamento vertical | Visível | Controles superior, central e inferior implementados |
| Quebra/overflow de texto | Visível | Implementado por controles separados |
| Rotação de texto | Visível | Controles de 45, -45 e remoção implementados |
| Congelar linhas/colunas | Visível | Congelar linhas, colunas e descongelar implementados |
| Formatação condicional | Visível | Fluxo visual para comparação, texto, duplicados, escala e remoção |
| Ordenação e filtro | Visível | Implementado com formulário próprio e limpeza global de filtros |
| Link | Visível | Inserção segura de HTTPS e mailto implementada |
| Imagem | Visível | Implementado |
| Comentário/nota | Visível | Implementado, com modelo próprio de nota/thread |
| Fórmula rápida/AutoSoma | Visível | AutoSoma usa a seleção ou o bloco numérico acima |
| Validação de dados | Visível | Fluxo visual para lista, número, data, checkbox e remoção |
| Dividir coluna | Visível | Divide a seleção por separador informado |
| Localizar condições/especiais | Visível | Localiza fórmulas, vazias, erros e constantes |
| Captura de tela | Visível | Exporta a área visível como SVG |
| Busca | Visível | Implementado |
| Zoom da planilha | Visível | Zoom de 50% a 200%, com incremento de 10% e redefinição |

Portanto, a existência de suporte no core não deve ser registrada como paridade de interface. Um recurso só é considerado visível quando pode ser descoberto e executado pela tela sem uso direto da API.