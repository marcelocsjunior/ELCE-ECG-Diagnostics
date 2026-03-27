# 03_FLUXO_EXECUCAO

## Sequência oficial

1. O técnico executa `ELCE_ECG_Diagnostics_Menu.bat`
2. O BAT localiza `C:\ECG\Tool\ELCE_ECG_Diagnostics.ps1`
3. O BAT chama o core em PowerShell com os parâmetros padrão
4. O core cria `RunId`
5. O core grava `execution.log`
6. O core coleta contexto, timeline e benchmark passivo
7. O core gera `analysis.json`
8. O core monta `ELCE_ECG_Diagnostics_Report.html`
9. O core publica artefatos em `Latest`

## Defaults atuais observados

- `StagePriority = ABRIR_EXAME`
- `SymptomCode = LENTIDAO_TRAVAMENTO`
- `ObservationMinutes = 10`
- `SampleIntervalSeconds = 20`

## Contrato importante

O comportamento real deve sempre ser lido do par BAT + PS1 da baseline.
