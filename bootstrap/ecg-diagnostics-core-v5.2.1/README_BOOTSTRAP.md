# ECG Diagnostics Core v5.2.1 — bootstrap para repositório novo

Este diretório foi criado como área de staging para extração da ferramenta para um repositório próprio.

## Objetivo

Separar a linha de produto `ECG Diagnostics Core` do repositório histórico `ELCE-ECG-Diagnostics`, reduzindo acoplamento entre:
- baseline read-only anterior
- fixpacks BDE legados
- toolkit unificado atual de diagnóstico, compare, monitor e rollback

## Nome recomendado do novo repositório

Preferência principal:
- `ECG-Diagnostics-Core`

Alternativas aceitáveis:
- `ECG-FieldKit`
- `ELCE-ECG-FieldKit`
- `ECG-Diagnostics-Unified`

## Escopo do novo repositório

Arquivos canônicos da release v5.2.1:
- `ECG_Diagnostics_Core.ps1`
- `ECG_Diagnostics_Hub.bat`
- `ECG_FieldKit.ini`
- `README_RELEASE.txt`
- `SHA256SUMS.txt`

## Observação importante

A criação do repositório novo não pôde ser executada diretamente por este conector GitHub porque a ação de `create_repository` não está exposta no ambiente atual.

Mesmo assim, este bootstrap deixa pronta a governança para corte imediato do novo repo sem redescoberta de contexto.
