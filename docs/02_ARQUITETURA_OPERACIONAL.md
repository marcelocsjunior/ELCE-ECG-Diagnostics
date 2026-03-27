# 02_ARQUITETURA_OPERACIONAL

## Arquitetura vigente

A arquitetura atual é simples por decisão de estabilidade:

- **Core**: `ELCE_ECG_Diagnostics.ps1`
- **Launcher**: `ELCE_ECG_Diagnostics_Menu.bat`
- **Execução**: Windows PowerShell 5.1
- **Artefato principal**: HTML
- **Artefatos técnicos**: JSON + log + summaries
- **Modo de operação**: read-only

## Caminhos padrão

```text
C:\ECG\Tool
C:\ECG\Output
C:\ECG\Output\Runs\<RunId>
C:\ECG\Output\Latest
```

## Racional da arquitetura

Esta linha foi mantida porque:
- reduz acoplamento;
- elimina superfícies de regressão;
- preserva fechamento da rodada;
- mantém previsibilidade operacional.
