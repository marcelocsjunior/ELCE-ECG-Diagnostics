# 03_FLUXO_EXECUCAO

## Sequência oficial do diagnóstico

1. O técnico executa `src\ELCE_ECG_Diagnostics_Menu.bat` ou `src\ExecutarDiagnostico.cmd`
2. O launcher resolve o core de forma relativa ao próprio arquivo
3. Se necessário, o launcher usa fallback para `C:\ECG\Tool\ELCE_ECG_Diagnostics.ps1`
4. O launcher chama o core em PowerShell com os parâmetros padrão
5. O core cria `RunId`
6. O core grava `execution.log`
7. O core coleta contexto, timeline e benchmark passivo
8. O core gera `analysis.json`
9. O core monta `ELCE_ECG_Diagnostics_Report.html`
10. O core publica artefatos em `Latest`

## Fluxo oficial da trilha de remediação

1. Havendo indício de problema BDE, o técnico abre `ECG-BDE-Fix_Menu.bat`
2. No clone do repositório, o entry point é `fixpacks\ECG-BDE-Fix_Menu.bat`
3. No layout implantado, o entry point é `C:\ECG\Tool\ECG-BDE-Fix_Menu.bat`
4. O fix menu tenta localizar:
   - `BDE-Fix-Core.ps1`
   - `ECG_UnitProfiles.json`
   - `runbookECG-BDE-Fix.txt`
5. O runbook canônico é `docs/runbooks/runbookECG-BDE-Fix.txt`
6. No layout implantado, pode existir uma cópia operacional do runbook em `C:\ECG\Tool\runbookECG-BDE-Fix.txt`

## Defaults atuais observados

- `StagePriority = ABRIR_EXAME`
- `SymptomCode = LENTIDAO_TRAVAMENTO`
- `ObservationMinutes = 10`
- `SampleIntervalSeconds = 20`

## Contrato importante

O comportamento real deve sempre ser lido do par launcher + PS1 da baseline aprovada.

Não assumir:
- caminho fixo único de script;
- presença do fix pack dentro de `src\`;
- artefato de release como equivalente automático a pacote de deploy, sem validação explícita do layout.
