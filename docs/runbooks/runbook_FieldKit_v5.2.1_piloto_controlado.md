# Runbook — piloto controlado do FieldKit v5.2.1

Data: 2026-04-07
Escopo: validar pacote standalone `ECG Diagnostics Core v5.2.1` em 1 estação executante e 1 estação viewer antes de qualquer rollout.

## Objetivo

Validar quatro coisas sem contaminar produção:

1. execução do launcher e do core;
2. geração correta de HTML/JSON/logs;
3. aderência de `NETDIR` e caminhos do banco;
4. utilidade real do modo `Compare` entre executante e viewer.

## Amostra mínima recomendada

- **Estação 1 — executante**: uma máquina que abre e grava exame;
- **Estação 2 — viewer**: uma máquina problemática ou historicamente instável.

## Pré-condições

- pasta limpa do pacote em `C:\ECG\FieldKit`;
- `ECG_FieldKit.ini` revisado para o cenário real da unidade;
- acesso UNC funcional ao banco e ao `NetDir`;
- janela operacional de baixo impacto;
- para `Fix` e `Rollback`, executar o Hub elevado.

## Ordem de execução recomendada

### Etapa 1 — sanity check do pacote
Em cada estação:

1. validar presença de:
   - `ECG_Diagnostics_Core.ps1`
   - `ECG_Diagnostics_Hub.bat`
   - `ECG_FieldKit.ini`
   - `README_RELEASE.txt`
   - `SHA256SUMS.txt`
2. abrir o Hub;
3. confirmar que o menu sobe sem erro.

Critério de aceite:
- launcher abre normalmente;
- nenhum erro imediato de arquivo ausente.

### Etapa 2 — rodada Auto na executante
Na estação executante:

1. executar `Auto`;
2. aguardar a janela configurada no INI;
3. abrir o HTML gerado;
4. revisar o JSON da rodada.

Validar:
- `ECG_Report.html` existe;
- `ECG_Report.json` existe;
- `ECG_Fatal_Error.log` não existe;
- `DatabaseAccessible=true`;
- `NetDirAccessible=true`;
- tipo de máquina coerente com executante.

Critério de aceite:
- laudo gerado sem falha fatal;
- paths coerentes com o contrato esperado;
- sem divergência absurda de classificação.

### Etapa 3 — rodada Auto na viewer
Na estação viewer:

1. executar `Auto`;
2. abrir o HTML;
3. revisar o JSON.

Validar:
- geração normal do HTML/JSON;
- classificação da estação como viewer / visualização ou, no mínimo, sem ambiguidade operacional;
- `DatabaseAccessible` e `NetDirAccessible` coerentes;
- sinais de `NETDIR` divergente, fallback ou indisponibilidade, se existirem.

Critério de aceite:
- laudo útil para triagem;
- nenhum falso sucesso grosseiro;
- viewer identificada com contexto suficiente para decisão.

### Etapa 4 — Compare executante x viewer
Usar os dois JSONs gerados nas etapas anteriores.

Validar:
- o `Compare` usa explicitamente os laudos desejados;
- o HTML de comparação evidencia convergência ou drift;
- o técnico consegue entender rapidamente se a viewer está fora do padrão da executante.

Critério de aceite:
- comparação reproduzível;
- resultado legível e operacional.

### Etapa 5 — Fix somente se houver evidência
Executar `Fix` **apenas** se a rodada `Auto` indicar problema real de `NETDIR`/`IDAPI32.CFG`/variável de ambiente.

Validar após o `Fix`:
- criação do `BDE_Rollback.reg`;
- registro de alterações aplicadas no JSON;
- nova rodada `Auto` ou `Monitor` sem piora;
- melhoria de aderência entre `Current/Expected` quando aplicável.

Critério de aceite:
- alteração reversível;
- nenhuma regressão funcional do ECG;
- melhoria objetiva dos indicadores do laudo.

## Critérios de GO / NO-GO

### GO para ampliar piloto
- executante e viewer geram laudo sem erro fatal;
- caminhos de banco e `NetDir` refletem o ambiente real;
- `Compare` ajuda a distinguir drift real;
- `Fix` é reversível e não degrada a estação.

### NO-GO para rollout amplo
- `Auto` falha de forma intermitente;
- classificação de estação gera leitura errada com frequência;
- `Compare` induz comparação incorreta;
- `Fix` altera estação sem evidência forte ou sem melhora observável.

## Evidências mínimas a guardar

Por estação, guardar:
- `ECG_Report.html`;
- `ECG_Report.json`;
- `ECG_Fatal_Error.log`, se existir;
- `BDE_Rollback.reg`, se `Fix` for usado.

## Decisão recomendada após o piloto

### Cenário A — piloto limpo
Promover o pacote apenas para mais 1 executante e 1 viewer.

### Cenário B — piloto parcialmente bom
Manter uso restrito como ferramenta de diagnóstico de campo.

### Cenário C — piloto ruim
Segurar rollout, corrigir launcher/core e repetir o teste.

## Princípio

O FieldKit deve provar utilidade operacional em duas estações reais antes de ganhar crachá de solução corporativa. Sem isso, vira consultor bonito com slide forte e produção triste.