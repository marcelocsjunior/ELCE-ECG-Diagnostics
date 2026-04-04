# 02_ARQUITETURA_OPERACIONAL

## Arquitetura vigente

A arquitetura atual foi mantida simples por decisão de estabilidade, mas hoje já contempla dois trilhos bem definidos:

- **Core de diagnóstico**: `src/ELCE_ECG_Diagnostics.ps1`
- **Entry points de diagnóstico**:
  - `src/ELCE_ECG_Diagnostics_Menu.bat`
  - `src/ExecutarDiagnostico.cmd`
- **Trilha separada de remediação**:
  - `fixpacks/ECG-BDE-Fix.ps1`
  - `fixpacks/ECG-BDE-Fix_Menu.bat`
  - `docs/runbooks/runbookECG-BDE-Fix.txt`
- **Execução**: Windows PowerShell 5.1
- **Artefato principal**: HTML
- **Artefatos técnicos**: JSON + log + summaries
- **Modo de operação do diagnóstico**: read-only
- **Remediação**: manual/controlada, separada do laudo

## Layouts suportados

### Clone do repositório
- execução diretamente a partir de `src/`
- acesso ao fix menu em `fixpacks/`
- runbook canônico em `docs/runbooks/`

### Layout implantado
- arquivos operacionais copiados para `C:\ECG\Tool`
- output em `C:\ECG\Output`
- runbook pode existir em `Tool\runbookECG-BDE-Fix.txt` por compatibilidade operacional do menu

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
- mantém previsibilidade operacional;
- separa diagnóstico e remediação por governança de produto.

## Regra importante

O comportamento real deve sempre ser lido do par:
- launcher BAT/CMD da baseline;
- core PS1 correspondente à mesma baseline.
