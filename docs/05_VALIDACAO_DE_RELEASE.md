# 05_VALIDACAO_DE_RELEASE

## Checklist oficial

### Contrato BAT + PS1
- [ ] `ELCE_ECG_Diagnostics_Menu.bat` aponta para o `ELCE_ECG_Diagnostics.ps1` correto
- [ ] `ExecutarDiagnostico.cmd` aponta para o `ELCE_ECG_Diagnostics.ps1` correto
- [ ] `ECG-BDE-Fix_Menu.bat` aponta para o `BDE-Fix-Core.ps1` correto
- [ ] Não há mistura de linhas diferentes

### Execução do diagnóstico
- [ ] O menu abre normalmente
- [ ] A opção de diagnóstico executa sem quebra
- [ ] O `RunId` é criado corretamente

### Artefatos
- [ ] `execution.log` gerado
- [ ] `context.json` gerado
- [ ] `timeline.json` gerado
- [ ] `benchmark.json` gerado
- [ ] `analysis.json` gerado
- [ ] `ELCE_ECG_Diagnostics_Report.html` gerado
- [ ] summaries gerados sem afetar o fechamento da rodada

### Publicação
- [ ] `Latest` atualizado corretamente
- [ ] `latest_run.txt` atualizado corretamente

### Trilha de remediação
- [ ] O fix menu abre normalmente
- [ ] O runbook abre normalmente
- [ ] `ECG_UnitProfiles.json` é resolvido corretamente
- [ ] O wrapper do fix pack encontra o catálogo canônico de profiles
- [ ] O fallback embutido de UN2 permanece alinhado ao catálogo canônico

### Empacotamento
- [ ] O **source package** contém `src/`, `docs/` e `fixpacks/`
- [ ] O **deploy package** contém `Tool/` com os arquivos operacionais esperados
- [ ] O runbook canônico está em `docs/runbooks/`
- [ ] A cópia operacional do runbook no deploy package abre pelo menu
- [ ] O clone do repositório funciona como layout suportado
- [ ] O layout implantado em `C:\ECG\Tool` funciona como layout suportado

### Qualidade
- [ ] HTML abre sem erro
- [ ] Conteúdo principal do laudo está legível
- [ ] Acentuação verificada
- [ ] Gráfico validado
- [ ] Regressão operacional não observada

## Regra final

Sem checklist completo, não há promoção de baseline.
