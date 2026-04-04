# ELCE ECG Diagnostics

Ferramenta corporativa de diagnóstico **read-only** para ambiente ECG, desenvolvida em **Windows PowerShell 5.1**, com foco em transformar coleta técnica em **laudo operacional HTML**.

## Status do produto

**Baseline operacional válida restaurada**

A solução está funcional e estável na baseline atual, com geração consistente de:
- `execution.log`
- `context.json`
- `benchmark.json`
- `timeline.json`
- `analysis.json`
- `ELCE_ECG_Diagnostics_Report.html`
- `ELCE_ECG_Diagnostics_Summary.txt`
- `ELCE_ECG_Diagnostics_Summary.json`

Este repositório deve ser tratado como **produto técnico em produção evolutiva**, com governança de baseline e patches mínimos.

## Objetivo

O ELCE ECG Diagnostics foi concebido para:
- coletar evidências técnicas relevantes do ambiente ECG;
- consolidar hipóteses e sinais operacionais;
- gerar um laudo HTML principal para triagem e decisão;
- manter artefatos estruturados por rodada;
- operar em modo **read-only**, sem remediação automática.

## Escopo atual

### Componentes ativos
- `src/ELCE_ECG_Diagnostics.ps1`
- `src/ELCE_ECG_Diagnostics_Menu.bat`
- `src/ExecutarDiagnostico.cmd`

### Componentes auxiliares de remediação
- `fixpacks/ECG-BDE-Fix.ps1`
- `fixpacks/ECG-BDE-Fix_Menu.bat`
- `docs/runbooks/runbookECG-BDE-Fix.txt`

### Componentes congelados
- GUI
- Console técnico
- Instalador

## Arquitetura vigente

A arquitetura em vigor segue uma linha pragmática, orientada à estabilidade:
- core único em PowerShell 5.1;
- launchers BAT/CMD autossuficientes;
- HTML como artefato principal;
- JSONs técnicos por rodada;
- summaries como apoio;
- sem benchmark assistido nesta fase;
- sem comparação com referência nesta fase;
- remediação mantida em trilha separada.

## Caminhos operacionais padrão

```text
C:\ECG\Tool
C:\ECG\Output
C:\ECG\Output\Runs\<RunId>
C:\ECG\Output\Latest
```

## Execução e empacotamento

Os launchers BAT/CMD devem resolver paths de forma relativa ao próprio arquivo sempre que possível.

Isso permite dois cenários suportados:
- execução diretamente do clone do repositório;
- execução após empacotamento/implantação em `C:\ECG\Tool`.

O output operacional continua padronizado em `C:\ECG\Output`.

A workflow oficial de release deve publicar **dois artefatos**:
- **source package**: snapshot do repositório, preservando `src/`, `docs/` e `fixpacks/`;
- **deploy package**: layout operacional com pasta `Tool/` pronta para cópia em `C:\ECG\Tool`.

No deploy package, o runbook canônico continua em `docs/runbooks/`, mas uma cópia operacional pode ser entregue em `Tool\runbookECG-BDE-Fix.txt` para compatibilidade do menu.

## Fluxo operacional

1. O técnico executa `ELCE_ECG_Diagnostics_Menu.bat` ou `ExecutarDiagnostico.cmd`
2. O launcher chama o core PowerShell
3. O core realiza a coleta técnica
4. O core consolida benchmark passivo e análise
5. O core gera o relatório HTML principal
6. O core publica os artefatos em `Runs\<RunId>` e atualiza `Latest`

## Artefatos por rodada

### Obrigatórios
- `execution.log`
- `context.json`
- `benchmark.json`
- `timeline.json`
- `analysis.json`
- `ELCE_ECG_Diagnostics_Report.html`

### Apoio
- `ELCE_ECG_Diagnostics_Summary.txt`
- `ELCE_ECG_Diagnostics_Summary.json`

## Estrutura do repositório

```text
src/                    -> core operacional versionado
fixpacks/               -> correções controladas e separadas do laudo
docs/                   -> governança, arquitetura, validação, runbooks e roadmap
samples/output-example/ -> exemplos sanitizados de saída
releases/notes/         -> notas de release
frozen/                 -> componentes legados congelados
.github/                -> automação leve
```

## Diretrizes de engenharia

- A baseline atual é a **golden copy**.
- Não misturar launcher BAT com core PS1 de linhas diferentes.
- Toda evolução deve partir da baseline operacional aprovada.
- Toda mudança deve ser mínima, reversível e validada isoladamente.
- Estabilidade tem precedência sobre expansão de escopo.
- A remediação deve permanecer separada do fluxo do laudo.

## Validação obrigatória antes de promover patch

1. Validar execução do Menu BAT
2. Validar criação do `RunId`
3. Validar `execution.log`
4. Validar `context.json`
5. Validar `timeline.json`
6. Validar `benchmark.json`
7. Validar `analysis.json`
8. Validar geração do HTML principal
9. Validar atualização de `Latest`
10. Validar abertura do fix menu e do runbook em layout de clone e layout implantado

## Limitações atuais

- ajustes de encoding / acentuação;
- refinamento do gráfico temporal embutido no HTML;
- classificação de tipo de máquina ainda incompleta;
- sem análise Defender/minifilter;
- sem benchmark ativo/assistido.

## Roadmap resumido

### Fase 1 — Hardening da baseline
- saneamento de encoding;
- ajuste do gráfico HTML;
- consolidação da classificação de máquinas.

### Fase 2 — Qualidade da análise
- enriquecimento de contexto;
- melhoria da análise consolidada;
- ampliação segura do conjunto de evidências.

### Fase 3 — Expansão controlada
- benchmark ativo;
- comparação com referência;
- Defender/minifilter;
- eventual reavaliação de GUI e console técnico.

## Versionamento

Este repositório adota versionamento semântico a partir da primeira baseline governada no GitHub.

Exemplo:
- `v1.0.0` -> baseline inicial publicada
- `v1.0.1` -> hotfix mínimo
- `v1.1.0` -> melhoria compatível
- `v2.0.0` -> mudança incompatível

## Política de commits

Padrão recomendado:
- `baseline:`
- `fix(core):`
- `fix(html):`
- `fix(context):`
- `fix(launcher):`
- `docs(...):`
- `chore(...):`

## Licenciamento

**Uso interno / proprietário**, salvo definição posterior explícita.

## Público-alvo

Times técnicos responsáveis por sustentação, triagem, investigação e análise operacional de ambiente ECG.

## Princípio central

O ELCE ECG Diagnostics não é uma suíte genérica de monitoramento.  
É um **motor de diagnóstico operacional com saída HTML principal**, desenhado para gerar evidência útil com baixo risco de interferência no ambiente.
